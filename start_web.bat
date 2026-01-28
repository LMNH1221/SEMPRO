@echo off
title Sistem Penjadwalan Shift Indomaret

cd /d "%~dp0"

echo ================================
echo MENJALANKAN APLIKASI STREAMLIT
echo Folder: %cd%
echo ================================

python -m streamlit run app.py

pause
