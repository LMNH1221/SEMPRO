import streamlit as st
import pandas as pd
from datetime import date, timedelta
import calendar
from io import BytesIO
import json
import os
import shutil

st.set_page_config(page_title="Jadwal Shift Indomaret", layout="wide")

# =========================================================
# FILE SAVE (AMAN + BACKUP)
# =========================================================
SAVE_FILE = "saved_state.json"
BAK_FILE = "saved_state.bak.json"
TMP_FILE = "saved_state.tmp.json"
SAVE_SCHEMA_VERSION = 1

# =========================================================
# Helper aman untuk jadwal (FIX reset_index bentrok)
# =========================================================
def safe_sched_view(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ubah jadwal (yang index-nya Nama) jadi ada kolom "Nama" TANPA error bentrok.
    - Kalau sudah ada kolom Nama => tidak reset_index
    - Kalau belum => reset_index dan rename kolom pertama jadi "Nama"
    - Buang kolom duplikat
    """
    out = df.copy()

    if "Nama" in out.columns:
        out = out.loc[:, ~out.columns.duplicated()]
        return out

    out = out.reset_index()
    first_col = out.columns[0]
    if first_col != "Nama":
        out = out.rename(columns={first_col: "Nama"})

    out = out.loc[:, ~out.columns.duplicated()]
    return out

# =========================================================
# SAVE / LOAD (AUTO + BACKUP)
# =========================================================
def atomic_write_json(payload: dict, target_path: str, tmp_path: str):
    """Tulis JSON aman: tulis tmp dulu, lalu replace."""
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, target_path)

def save_state_to_disk():
    payload = {
        "schema_version": SAVE_SCHEMA_VERSION,
        "employees": st.session_state.get("employees", []),
        "cuti": sorted(list(st.session_state.get("cuti", set()))),
        "last_result": None
    }

    lr = st.session_state.get("last_result")
    if lr:
        sched_records = safe_sched_view(lr["sched_show"]).to_dict(orient="records")
        payload["last_result"] = {
            "meta": lr.get("meta", {}),
            "warnings": lr.get("warnings", []),
            "sched_show": sched_records,
            "recap": lr["recap"].to_dict(orient="records"),
            "summary2": lr["summary2"].to_dict(orient="records"),
        }

    # 1) Kalau file utama ada, bikin backup dulu
    if os.path.exists(SAVE_FILE):
        try:
            shutil.copy2(SAVE_FILE, BAK_FILE)
        except Exception:
            pass

    # 2) Atomic write ke file utama
    try:
        atomic_write_json(payload, SAVE_FILE, TMP_FILE)
    except Exception:
        # kalau gagal nulis utama, coba balikin dari backup (kalau ada)
        if os.path.exists(BAK_FILE):
            try:
                shutil.copy2(BAK_FILE, SAVE_FILE)
            except Exception:
                pass

def try_load_from(path: str) -> dict | None:
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def load_state_from_disk():
    # coba load utama
    payload = try_load_from(SAVE_FILE)

    # kalau utama rusak, coba backup
    if payload is None:
        payload = try_load_from(BAK_FILE)

    if payload is None:
        return False

    # cek schema version
    if payload.get("schema_version") != SAVE_SCHEMA_VERSION:
        # aman: load data dasar saja, hasil jadwal dibuang
        st.session_state.employees = payload.get("employees", [])
        st.session_state.cuti = set(tuple(x) for x in payload.get("cuti", []))
        st.session_state.last_result = None
        return True

    st.session_state.employees = payload.get("employees", [])
    st.session_state.cuti = set(tuple(x) for x in payload.get("cuti", []))

    lr = payload.get("last_result")
    if lr:
        sched_show = pd.DataFrame(lr.get("sched_show", []))
        if "Nama" in sched_show.columns:
            sched_show = sched_show.set_index("Nama")
            sched_show.index.name = None

        recap = pd.DataFrame(lr.get("recap", []))
        summary2 = pd.DataFrame(lr.get("summary2", []))

        st.session_state.last_result = {
            "sched_show": sched_show,
            "recap": recap,
            "summary2": summary2,
            "warnings": lr.get("warnings", []),
            "meta": lr.get("meta", {}),
        }
    else:
        st.session_state.last_result = None

    return True

# Auto-load sekali saat app dibuka
if "did_autoload" not in st.session_state:
    st.session_state.did_autoload = True
    load_state_from_disk()

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

def to_excel_bytes(sched_show: pd.DataFrame, recap: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_sched_view(sched_show).to_excel(writer, index=False, sheet_name="Jadwal")
        recap.to_excel(writer, index=False, sheet_name="RekapHarian")
        summary.to_excel(writer, index=False, sheet_name="RekapKaryawan")
    return output.getvalue()

# =========================================================
# Generator (revisi: habis cuti besok jangan libur)
# =========================================================
def generate_schedule_month_all_work(employees, year: int, month: int, cuti_set, min_per_shift: int = 2):
    """
    - Cuti => "Cuti"
    - Libur => "Libur" (pola 6 kerja 1 libur)
    - Selain itu => kerja shift 1/2/3
    - Rotasi 1->2->3
    - Perempuan tidak boleh shift 3, kalau kena 3 -> loncat ke 1
    - min_per_shift hanya WARNING (bukan kuota)
    - Revisi: habis cuti, besok jangan libur (liburnya digeser)
    """
    if not employees:
        return None, None, None, ["Belum ada data karyawan."]

    dates = month_dates(year, month)
    names = [e["nama"] for e in employees]
    gender = {e["nama"]: e["gender"] for e in employees}

    offsets_base = {nama: i % 7 for i, nama in enumerate(names)}
    offsets_now = offsets_base.copy()

    start_shift = {nama: i % 3 for i, nama in enumerate(names)}  # 0..2

    sched = pd.DataFrame(index=names, columns=[d.strftime("%Y-%m-%d") for d in dates])
    warnings = []

    for day_i, d in enumerate(dates):
        day_str = d.strftime("%Y-%m-%d")

        working = []
        for nama in names:
            if (nama, day_str) in cuti_set:
                continue
            if is_libur(day_i, offsets_now[nama]):
                continue
            working.append(nama)

        assigned = {}
        for nama in working:
            s = base_shift(day_i, start_shift[nama])
            if gender.get(nama) == "P" and s == 3:
                s = 1
            assigned[nama] = s

        def count_shift(x):
            return sum(1 for v in assigned.values() if v == x)

        # kalau shift 3 kosong, pindahkan 1 cowok dari shift 1/2
        if working and count_shift(3) == 0:
            donors = [n for n in working if gender.get(n) == "L" and assigned[n] in (1, 2)]
            if donors:
                assigned[donors[0]] = 3

        c1, c2, c3 = count_shift(1), count_shift(2), count_shift(3)
        if c1 < min_per_shift or c2 < min_per_shift or c3 < min_per_shift:
            warnings.append(f"{day_str}: WARNING min {min_per_shift}/shift -> S1={c1}, S2={c2}, S3={c3}")

        # isi tabel hari ini
        for nama in names:
            if (nama, day_str) in cuti_set:
                sched.loc[nama, day_str] = "Cuti"
            elif is_libur(day_i, offsets_now[nama]):
                sched.loc[nama, day_str] = "Libur"
            else:
                sched.loc[nama, day_str] = str(assigned.get(nama, 1))

        # Revisi: kalau hari ini Cuti, besok jangan Libur => geser offset libur
        if day_i < len(dates) - 1:
            next_i = day_i + 1
            next_str = dates[next_i].strftime("%Y-%m-%d")

            for nama in names:
                if (nama, day_str) not in cuti_set:
                    continue
                if (nama, next_str) in cuti_set:
                    continue
                if is_libur(next_i, offsets_now[nama]):
                    offsets_now[nama] += 1

    # rekap harian
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

    # rekap karyawan
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

    # tampilan ala foto: kolom 1..31 + baris Hari
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
# Session defaults
# =========================================================
if "employees" not in st.session_state:
    st.session_state.employees = [
        {"nama": "Yosep", "gender": "L"},
        {"nama": "Ghimel", "gender": "L"},
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
    st.session_state.cuti = set()

if "last_result" not in st.session_state:
    st.session_state.last_result = None

# =========================================================
# HEADER
# =========================================================
st.title("🗓️ Sistem Penjadwalan Shift Karyawan — Indomaret")
st.caption(
    "Aturan: 1=Pagi, 2=Siang, 3=Malam | Semua yang tidak Libur/Cuti = pasti kerja | "
    "Perempuan tidak boleh shift 3 (kalau kena 3 -> loncat ke shift 1) | Pola 6 kerja 1 Libur | "
    "Revisi: Habis Cuti besoknya jangan Libur | AUTO-SAVE + BACKUP aktif ✅"
)

tab1, tab2, tab3, tab4 = st.tabs(["👥 Data Karyawan", "🏖️ Data Cuti", "⚙️ Generate Jadwal", "📋 Hasil & Download"])

# -------------------------
# TAB 1: Data Karyawan
# -------------------------
with tab1:
    st.subheader("Data Karyawan")
    colA, colB = st.columns([1.2, 1.0])

    with colA:
        df_emp = pd.DataFrame(st.session_state.employees).copy()
        df_emp = df_emp.rename(columns={"nama": "Nama", "gender": "Gender"})
        df_emp.insert(0, "No", range(1, len(df_emp) + 1))
        st.dataframe(df_emp, use_container_width=True, hide_index=True)

    with colB:
        st.markdown("### Tambah Karyawan")
        nama_baru = st.text_input("Nama karyawan", key="nama_baru")
        gender_baru = st.selectbox("Gender", ["L", "P"], key="gender_baru")

        if st.button("➕ Tambah", key="btn_tambah_karyawan"):
            if nama_baru.strip():
                st.session_state.employees.append({"nama": nama_baru.strip(), "gender": gender_baru})
                save_state_to_disk()
                st.success(f"Ditambahkan: {nama_baru.strip()} ({gender_baru})")
            else:
                st.error("Nama tidak boleh kosong.")

        st.markdown("### Hapus Karyawan")
        names = [e["nama"] for e in st.session_state.employees]
        if names:
            hapus_nama = st.selectbox("Pilih nama", names, key="hapus_nama")
            if st.button("🗑️ Hapus", key="btn_hapus_karyawan"):
                st.session_state.employees = [e for e in st.session_state.employees if e["nama"] != hapus_nama]
                st.session_state.cuti = {(n, t) for (n, t) in st.session_state.cuti if n != hapus_nama}
                save_state_to_disk()
                st.warning(f"Dihapus: {hapus_nama}")

# -------------------------
# TAB 2: Data Cuti
# -------------------------
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
                    save_state_to_disk()
                    st.success(f"{pilih_nama} cuti pada {tgl_str}")
            with c2:
                if st.button("❌ Hapus Cuti", key="btn_hapus_cuti"):
                    if (pilih_nama, tgl_str) in st.session_state.cuti:
                        st.session_state.cuti.remove((pilih_nama, tgl_str))
                        save_state_to_disk()
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

# -------------------------
# TAB 3: Generate Jadwal
# -------------------------
with tab3:
    st.subheader("Generate Jadwal Bulanan")
    colE, colF = st.columns([1.0, 1.0])

    with colE:
        today = date.today()
        year = st.number_input("Tahun", min_value=2020, max_value=2035, value=today.year, key="year_gen")
        month = st.number_input("Bulan (1-12)", min_value=1, max_value=12, value=today.month, key="month_gen")

    with colF:
        min_per_shift = st.number_input("Minimal orang per shift (hanya WARNING)", min_value=0, max_value=50, value=2, key="min_per_shift")
        hours_per_shift = st.number_input("Jam kerja per shift (untuk rekap jam)", min_value=1, max_value=12, value=8, key="hours_per_shift")

    st.markdown("---")
    if st.button("🧾 Generate Jadwal", key="btn_generate"):
        sched_show, recap, summary, warnings = generate_schedule_month_all_work(
            st.session_state.employees,
            int(year),
            int(month),
            st.session_state.cuti,
            int(min_per_shift),
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
                    "min_per_shift": int(min_per_shift),
                    "hours_per_shift": int(hours_per_shift),
                }
            }

            save_state_to_disk()
            st.success("Jadwal dibuat & tersimpan otomatis ✅ (dengan backup juga)")

# -------------------------
# TAB 4: Hasil & Download
# -------------------------
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
                st.write("\n".join(warnings[:50]))
                if len(warnings) > 50:
                    st.write(f"...dan {len(warnings)-50} warning lainnya.")

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
