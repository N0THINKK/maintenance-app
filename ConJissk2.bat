@echo off
setlocal

set "input=C:\AC90HMI\Jissk.dat"
set "output=C:\AC90HMI\Records.csv"

echo Membaca file: %input%

powershell -Command ^
    "$in='%input%'; $out='%output%'; " ^
    "if (Test-Path $in) { " ^
    "  Get-Content -Path $in -Encoding Default -Tail 100 | Set-Content -Path $out -Encoding UTF8; " ^
    "  Write-Host 'Output selesai disimpan di ' $out " ^
    "} else { Write-Host 'File input tidak ditemukan' }"

