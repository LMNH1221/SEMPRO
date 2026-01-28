import streamlit as st
import pandas as pd
from datetime import date, timedelta
import calendar
from io import BytesIO
import json, os, shutil

DATA_DIR = "data"
EMP_FILE = f"{DATA_DIR}/employees.json"
CUTI_FILE = f"{DATA_DIR}/cuti.json"
STATE_FILE = f"{DATA_DIR}/state.json"
SCHED_DIR = f"{DATA_DIR}/schedules"

os.makedirs(SCHED_DIR, exist_ok=True) 


st.set_page_config(page_title="Jadwal Shift Indomaret", layout="wide")

# =========================================================
# Helper umum
# =========================================================
def month_dates(year: int, month: int):
    last_day = calendar.monthrange(year, month)[1]
    start = date(year, month, 1)
    return [start + timedelta(days=i) for i in range(last_day)]

def day_short(d: date):
    names = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
    return names[d.weekday()]

def is_libur(day_index: int, offset: int) -> bool:
    # pola 6 kerja 1 libur -> hari ke-7 libur
    return ((day_index + offset) % 7) == 6

def base_shift(day_index: int, start_shift_index: int) -> int:
    return ((start_shift_index + day_index) % 3) + 1  # 1..3

def color_schedule(val):
    if val == "1":
        return "background-color: #c6efce"
    if val == "2":
        return "background-color: #ffeb9c"
    if val == "3":
        return "background-color: #bdd7ee"
    if val == "Libur":
        return "background-color: #d9d9d9"
    if val == "Cuti":
        return "background-color: #f4cccc"
    if val == "Hari":
        return "background-color: #eeeeee; font-weight: bold"
    return ""

def safe_sched_view(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Nama" not in out.columns:
        out = out.reset_index()
        out = out.rename(columns={out.columns[0]: "Nama"})
    out = out.loc[:, ~out.columns.duplicated()]
    return out

def to_excel_bytes(sched_show: pd.DataFrame, recap: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_sched_view(sched_show).to_excel(writer, index=False, sheet_name="Jadwal")
        recap.to_excel(writer, index=False, sheet_name="RekapHarian")
        summary.to_excel(writer, index=False, sheet_name="RekapKaryawan")
    return output.getvalue()

# =========================================================
# TAMBAHAN: File Helpers
# =========================================================
def load_json(path, default):
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return default

def save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f, indent=2)

# =========================================================
# CORE GENERATOR (aturan lengkap)
# =========================================================
def generate_schedule_month(
    employees,
    year: int,
    month: int,
    cuti_set: set[tuple[str, str]],
    min_per_shift_warn: int = 2
):
    """
    ATURAN:
    - Pola libur: 6 kerja 1 libur (per karyawan, offset beda-beda)
    - Cuti => "Cuti"
    - Libur => "Libur"
    - Sisanya kerja shift 1/2/3
    - Rotasi shift default: 1->2->3 (berdasarkan hari)
    - Perempuan tidak boleh shift 3 (kalau default 3 -> jadi 1)
    - Aturan baru:
        (A) Shift 1 & 2 tiap hari harus ada minimal 1 wanita (kalau ada wanita yang bekerja hari itu)
        (B) Sehari sebelum libur harus shift 1
        (C) Sehari setelah libur harus shift 2
    - min_per_shift_warn hanya WARNING (bukan kuota)
    """
    if not employees:
        return None, None, None, ["Belum ada data karyawan."]

    dates = month_dates(year, month)
    names = [e["nama"] for e in employees]
    gender = {e["nama"]: e["gender"] for e in employees}

    # offset libur per orang
    offsets = {nama: i % 7 for i, nama in enumerate(names)}

    # =========================================================
    # SHIFT LANJUTAN ANTAR BULAN (AKTIF)
    # =========================================================
    state = load_json(STATE_FILE, {})
    last_shift_state = state.get("last_shift", {})

    start_shift = {}
    for i, nama in enumerate(names):
        if nama in last_shift_state:
            # lanjut dari shift terakhir bulan sebelumnya (0..2)
            start_shift[nama] = last_shift_state[nama]
        else:
            # fallback aman (bulan pertama / karyawan baru)
            start_shift[nama] = i % 3


    # jadwal mentah (kolom YYYY-MM-DD)
    sched = pd.DataFrame(index=names, columns=[d.strftime("%Y-%m-%d") for d in dates])
    warnings = []

    for day_i, d in enumerate(dates):
        day_str = d.strftime("%Y-%m-%d")

        # helper index tetangga hari
        prev_i = day_i - 1
        next_i = day_i + 1

        # siapa yang kerja hari ini
        working = []
        for nama in names:
            if (nama, day_str) in cuti_set:
                continue
            if is_libur(day_i, offsets[nama]):
                continue
            working.append(nama)

        # assignment awal (rotasi)
        assigned: dict[str, int] = {}
        forced: set[str] = set()  # yang shift-nya "dipaksa" karena aturan sebelum/after libur

        for nama in working:
            s = base_shift(day_i, start_shift[nama])

            # aturan lama: perempuan tidak boleh shift 3
            if gender.get(nama) == "P" and s == 3:
                s = 1

            # aturan baru: sebelum libur => shift 1
            if next_i < len(dates) and is_libur(next_i, offsets[nama]):
                s = 1
                forced.add(nama)

            # aturan baru: setelah libur => shift 2
            if prev_i >= 0 and is_libur(prev_i, offsets[nama]):
                s = 2
                forced.add(nama)

            assigned[nama] = s

        def count_shift(shift_no: int):
            return sum(1 for v in assigned.values() if v == shift_no)

        def females_in_shift(shift_no: int):
            return [n for n, s in assigned.items() if s == shift_no and gender.get(n) == "P"]

        def males_in_shift(shift_no: int):
            return [n for n, s in assigned.items() if s == shift_no and gender.get(n) == "L"]

        def can_move(nama: str, to_shift: int) -> bool:
            # yang forced tidak boleh dipindah
            if nama in forced:
                return False
            # perempuan tidak boleh shift 3
            if gender.get(nama) == "P" and to_shift == 3:
                return False
            return True

        # =====================================================
        # 1) pastikan shift 3 tidak kosong (kalau ada working)
        # =====================================================
        if working and count_shift(3) == 0:
            donors = [n for n in working if gender.get(n) == "L" and assigned[n] in (1, 2) and can_move(n, 3)]
            if donors:
                assigned[donors[0]] = 3
            else:
                warnings.append(f"{day_str}: WARNING shift 3 kosong (tidak ada kandidat laki-laki yang bisa dipindah).")

        # =====================================================
        # 2) aturan baru: Shift 1 & Shift 2 harus ada wanita
        # =====================================================
        women_working = [n for n in working if gender.get(n) == "P"]
        if women_working:
            # --- pastikan shift 1 ada wanita
            if len(females_in_shift(1)) == 0:
                # coba pindahin wanita dari shift 2 ke 1 (kalau shift 2 masih ada wanita setelah dipindah)
                candidates = [n for n in females_in_shift(2) if can_move(n, 1)]
                moved = False
                for n in candidates:
                    # setelah dipindah, shift 2 masih ada wanita?
                    if len(females_in_shift(2)) >= 2:
                        assigned[n] = 1
                        moved = True
                        break

                # kalau belum bisa, coba wanita dari shift 1/3? (wanita tidak ada di 3)
                if not moved:
                    # coba wanita dari shift 2 walaupun jadi kosong -> tapi aturan juga minta shift 2 ada wanita
                    # jadi hanya bisa kalau nanti kita bisa isi shift 2 dari wanita lain.
                    candidates2 = [n for n in females_in_shift(2) if can_move(n, 1)]
                    if candidates2 and len(women_working) >= 2:
                        assigned[candidates2[0]] = 1
                        moved = True

                if not moved:
                    warnings.append(f"{day_str}: WARNING aturan wanita di shift 1 tidak terpenuhi (kandidat tidak tersedia).")

            # --- pastikan shift 2 ada wanita
            if len(females_in_shift(2)) == 0:
                candidates = [n for n in females_in_shift(1) if can_move(n, 2)]
                moved = False
                for n in candidates:
                    # setelah dipindah, shift 1 masih ada wanita?
                    if len(females_in_shift(1)) >= 2:
                        assigned[n] = 2
                        moved = True
                        break

                if not moved:
                    candidates2 = [n for n in females_in_shift(1) if can_move(n, 2)]
                    if candidates2 and len(women_working) >= 2:
                        assigned[candidates2[0]] = 2
                        moved = True

                if not moved:
                    warnings.append(f"{day_str}: WARNING aturan wanita di shift 2 tidak terpenuhi (kandidat tidak tersedia).")
        else:
            warnings.append(f"{day_str}: WARNING tidak ada wanita yang bekerja hari ini, aturan shift 1&2 wanita tidak bisa diterapkan.")

        # =====================================================
        # 3) warning minimal orang per shift
        # =====================================================
        c1, c2, c3 = count_shift(1), count_shift(2), count_shift(3)
        if c1 < min_per_shift_warn or c2 < min_per_shift_warn or c3 < min_per_shift_warn:
            warnings.append(f"{day_str}: WARNING min {min_per_shift_warn}/shift -> S1={c1}, S2={c2}, S3={c3}")

        # =====================================================
        # 4) isi tabel hari ini
        # =====================================================
        for nama in names:
            if (nama, day_str) in cuti_set:
                sched.loc[nama, day_str] = "Cuti"
            elif is_libur(day_i, offsets[nama]):
                sched.loc[nama, day_str] = "Libur"
            else:
                # kalau working -> ambil assigned, kalau tidak -> default 1 (harusnya tidak kejadian)
                sched.loc[nama, day_str] = str(assigned.get(nama, 1))

    # =========================================================
    # REKAP HARIAN
    # =========================================================
    recap = []
    for col in sched.columns:
        c = sched[col]
        recap.append({
            "Tanggal": col,
            "Shift 1 (Pagi)": int((c == "1").sum()),
            "Shift 2 (Siang)": int((c == "2").sum()),
            "Shift 3 (Malam)": int((c == "3").sum()),
            "Libur": int((c == "Libur").sum()),
            "Cuti": int((c == "Cuti").sum()),
        })
    recap = pd.DataFrame(recap)

    # =========================================================
    # REKAP PER KARYAWAN
    # =========================================================
    summary_rows = []
    for nama in sched.index:
        row = sched.loc[nama]
        kerja = ((row == "1") | (row == "2") | (row == "3")).sum()
        summary_rows.append({
            "Nama": nama,
            "Gender": gender.get(nama, ""),
            "Hari Kerja": int(kerja),
            "Shift 1": int((row == "1").sum()),
            "Shift 2": int((row == "2").sum()),
            "Shift 3": int((row == "3").sum()),
            "Libur": int((row == "Libur").sum()),
            "Cuti": int((row == "Cuti").sum()),
        })
    summary = pd.DataFrame(summary_rows)

    # =========================================================
    # FORMAT TAMPILAN (kolom 1..31 + baris Hari)
    # =========================================================
    rename_cols = {}
    hari_row = {}
    for d in dates:
        old = d.strftime("%Y-%m-%d")
        new = str(d.day)
        rename_cols[old] = new
        hari_row[new] = day_short(d)

    sched2 = sched.rename(columns=rename_cols)
    header = pd.DataFrame([hari_row], index=["Hari"])
    sched_show = pd.concat([header, sched2])
    sched_show.index.name = None

    return sched_show, recap, summary, warnings

# =========================================================
# LOAD DATA SEKALI SAAT START (PRODUKSI)
# =========================================================
if "employees" not in st.session_state:
    data = load_json(EMP_FILE, None)
    st.session_state.employees = data if data is not None else [
        {"nama": "Yosep", "gender": "L"},
        {"nama": "Ghimel", "gender": "P"},
        {"nama": "Ilham", "gender": "L"},
        {"nama": "Shiddiq", "gender": "L"},
        {"nama": "Diana", "gender": "P"},
        {"nama": "Irma", "gender": "P"},
        {"nama": "Topik", "gender": "L"},
        {"nama": "Chaqi", "gender": "L"},
        {"nama": "Juna", "gender": "L"},
        {"nama": "Regi", "gender": "L"},
        {"nama": "Rajib", "gender": "L"},
    ]

if "cuti" not in st.session_state:
    st.session_state.cuti = set(tuple(x) for x in load_json(CUTI_FILE, []))

if "last_result" not in st.session_state:
    st.session_state.last_result = None


# =========================================================
# UI
# =========================================================
st.title("🗓️ Sistem Penjadwalan Shift Karyawan — Indomaret")
st.caption(
    "Aturan: 1=Pagi, 2=Siang, 3=Malam | Pola 6 kerja 1 Libur | Perempuan tidak boleh shift 3 | "
    "Shift 1 & 2 wajib ada wanita (jika ada wanita kerja hari itu) | "
    "Sebelum libur=Shift 1 | Setelah libur=Shift 2 | (Shift lanjutan antar bulan aktif) ✅"
)

# =========================================================
# TAMBAHAN: Inject start shift dari bulan sebelumnya
# =========================================================
last_state = load_json(STATE_FILE, {})


tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "👥 Data Karyawan",
    "🏖️ Data Cuti",
    "⚙️ Generate Jadwal",
    "📋 Hasil & Download",
    "📂 Arsip Bulan",
    "⚠️ Reset Sistem"
])


# TAB 1
with tab1:
    st.subheader("Data Karyawan")
    colA, colB = st.columns([1.2, 1.0])

    with colA:
        df_emp = pd.DataFrame(st.session_state.employees)
        st.dataframe(df_emp, use_container_width=True, hide_index=True)

    with colB:
        st.markdown("### Tambah Karyawan")
        nama_baru = st.text_input("Nama karyawan")
        gender_baru = st.selectbox("Gender", ["L", "P"])

        if st.button("➕ Tambah"):
            if nama_baru.strip():
                st.session_state.employees.append({
                    "nama": nama_baru.strip(),
                    "gender": gender_baru
                })
                save_json(EMP_FILE, st.session_state.employees)
                st.rerun()

        st.markdown("### Hapus Karyawan")
        names = [e["nama"] for e in st.session_state.employees]
        if names:
            hapus_nama = st.selectbox("Pilih nama", names)
            if st.button("🗑️ Hapus"):
                st.session_state.employees = [
                    e for e in st.session_state.employees
                    if e["nama"] != hapus_nama
                ]
                st.session_state.cuti = {
                    (n, t) for (n, t) in st.session_state.cuti
                    if n != hapus_nama
                }
                save_json(EMP_FILE, st.session_state.employees)
                save_json(CUTI_FILE, list(st.session_state.cuti))
                st.rerun()


# TAB 2
with tab2:
    st.subheader("Data Cuti (Sakit/Izin)")
    colC, colD = st.columns([1.1, 1.2])

    with colC:
        st.markdown("### Input Cuti")
        names = [e["nama"] for e in st.session_state.employees]
        if not names:
            st.info("Tambahkan karyawan dulu di tab Data Karyawan.")
        else:
            pilih_nama = st.selectbox("Nama karyawan", names, key="pilih_cuti_nama")
            tgl_cuti = st.date_input("Tanggal cuti", value=date.today(), key="tgl_cuti")
            tgl_str = tgl_cuti.strftime("%Y-%m-%d")

            c1, c2 = st.columns(2)
            with c1:
                if st.button("✅ Tambah Cuti", key="btn_tambah_cuti"):
                    st.session_state.cuti.add((pilih_nama, tgl_str))
                    st.success(f"{pilih_nama} cuti pada {tgl_str}")
            with c2:
                if st.button("❌ Hapus Cuti", key="btn_hapus_cuti"):
                    if (pilih_nama, tgl_str) in st.session_state.cuti:
                        st.session_state.cuti.remove((pilih_nama, tgl_str))
                        st.warning(f"Cuti dihapus: {pilih_nama} {tgl_str}")
                    else:
                        st.info("Data cuti tidak ditemukan.")

    with colD:
        st.markdown("### Daftar Cuti Saat Ini")
        if st.session_state.cuti:
            df_cuti = pd.DataFrame(sorted(list(st.session_state.cuti)), columns=["Nama", "Tanggal"])
            df_cuti.insert(0, "No", range(1, len(df_cuti) + 1))
            st.dataframe(df_cuti, use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada data cuti.")

# TAB 3
with tab3:
    st.subheader("Generate Jadwal Bulanan")
    colE, colF = st.columns([1.0, 1.0])

    with colE:
        today = date.today()
        year = st.number_input("Tahun", min_value=2020, max_value=2035, value=today.year, key="year_gen")
        month = st.number_input("Bulan (1-12)", min_value=1, max_value=12, value=today.month, key="month_gen")

    with colF:
        min_per_shift_warn = st.number_input("Minimal orang per shift (hanya WARNING)", min_value=0, max_value=50, value=2, key="min_per_shift_warn")
        hours_per_shift = st.number_input("Jam kerja per shift (untuk rekap jam)", min_value=1, max_value=12, value=8, key="hours_per_shift")

    st.markdown("---")
    if st.button("🧾 Generate Jadwal", key="btn_generate"):
        sched_show, recap, summary, warnings = generate_schedule_month(
            st.session_state.employees,
            int(year),
            int(month),
            st.session_state.cuti,
            int(min_per_shift_warn),
        )

        if sched_show is None:
            st.error("Belum ada karyawan.")
        else:
            summary2 = summary.copy()
            summary2["Estimasi Jam Kerja"] = summary2["Hari Kerja"] * int(hours_per_shift)

            st.session_state.last_result = {
                "sched_show": sched_show,
                "recap": recap,
                "summary2": summary2,
                "warnings": warnings,
                "meta": {
                    "year": int(year),
                    "month": int(month),
                    "min_per_shift_warn": int(min_per_shift_warn),
                    "hours_per_shift": int(hours_per_shift),
                }
            }

# =========================================================
# TAMBAHAN: Simpan hasil & state
# =========================================================
# simpan jadwal per bulan
            path = f"{SCHED_DIR}/{year}-{month:02d}.csv"
            safe_sched_view(sched_show).to_csv(path, index=False)
            
            st.success("Jadwal berhasil dibuat ✅")


# simpan karyawan & cuti
            save_json(EMP_FILE, st.session_state.employees)
            save_json(CUTI_FILE, list(st.session_state.cuti))

# simpan shift terakhir untuk bulan berikutnya
            last_shift = {}
            for nama in sched_show.index:
                if nama != "Hari":
                    last_val = sched_show.loc[nama].iloc[-1]
                    if last_val in ("1", "2", "3"):
                        last_shift[nama] = int(last_val) % 3

            save_json(STATE_FILE, {"last_shift": last_shift})

# TAB 4
with tab4:
    st.subheader("Hasil Jadwal & Download")

    result = st.session_state.last_result
    if not result:
        st.info("Belum ada hasil. Buka tab 'Generate Jadwal' lalu klik Generate.")
    else:
        meta = result["meta"]
        sched_show = result["sched_show"]
        recap = result["recap"]
        summary2 = result["summary2"]
        warnings = result["warnings"]

        st.markdown(f"### Jadwal Bulan: {calendar.month_name[int(meta['month'])]} {meta['year']}")

        if warnings:
            with st.expander("⚠️ Lihat WARNING (klik untuk buka)"):
                st.write("\n".join(warnings[:200]))
                if len(warnings) > 200:
                    st.write(f"...dan {len(warnings)-200} warning lainnya.")

        st.markdown("#### Tabel Jadwal (Berwarna)")
        sched_view = safe_sched_view(sched_show)
        st.dataframe(
            sched_view.style.applymap(color_schedule),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("#### Rekap Harian")
        st.dataframe(recap, use_container_width=True, hide_index=True)

        st.markdown("#### Rekap Per Karyawan + Estimasi Jam")
        df_sum = summary2.copy()
        df_sum.insert(0, "No", range(1, len(df_sum) + 1))
        st.dataframe(df_sum, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### Download File")

        csv = safe_sched_view(sched_show).to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Jadwal (CSV)",
            data=csv,
            file_name="jadwal_shift_bulanan.csv",
            mime="text/csv",
            key="dl_csv"
        )

        xlsx_bytes = to_excel_bytes(sched_show, recap, summary2)
        st.download_button(
            "⬇️ Download Jadwal (Excel .xlsx)",
            data=xlsx_bytes,
            file_name="jadwal_shift_bulanan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_xlsx"
        )

#TAB5
with tab5:
    files = sorted(os.listdir(SCHED_DIR))
    if files:
        pilih = st.selectbox("Pilih Arsip Bulan", files)
        df = pd.read_csv(f"{SCHED_DIR}/{pilih}")
        st.dataframe(df.style.applymap(color_schedule), use_container_width=True)
    else:
        st.info("Belum ada arsip bulan.")


#TAB6
with tab6:
    st.subheader("⚠️ RESET SISTEM (DEMO)")
    st.warning(
        "Tombol ini akan:\n"
        "- Menghapus semua jadwal bulan\n"
        "- Menghapus data cuti\n"
        "- Menghapus state shift lanjutan\n"
        "- Mengosongkan tampilan hasil\n\n"
        "❗ Data karyawan TIDAK dihapus."
    )

    confirm = st.checkbox("Saya mengerti dan ingin reset sistem")

    if st.button("🔥 RESET SEMUA DATA DEMO", disabled=not confirm):
        # hapus file
        if os.path.exists(CUTI_FILE):
            os.remove(CUTI_FILE)
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
        if os.path.exists(SCHED_DIR):
            shutil.rmtree(SCHED_DIR)
            os.makedirs(SCHED_DIR, exist_ok=True)

        # reset session
        st.session_state.cuti = set()
        st.session_state.last_result = None

        st.success("✅ Sistem berhasil di-reset total")
        st.rerun()
