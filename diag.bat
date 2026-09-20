@echo off
rem ── levIT Office aktivalasi diagnosztika ────────────────────────────────
rem  Rendszergazdakent futtatni. Mindent kiment az Asztalra:
rem    office_diag.txt
rem
rem  Ez NEM aktival es NEM valtoztat semmit, csak kiolvas.

setlocal
set "OUT=%USERPROFILE%\Desktop\office_diag.txt"
set "OSPP="

if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs"      set "OSPP=%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set "OSPP=%ProgramFiles(x86)%\Microsoft Office\Office16"

> "%OUT%" echo === levIT Office diagnosztika ===
>>"%OUT%" echo Datum: %DATE% %TIME%
>>"%OUT%" echo.

>>"%OUT%" echo === Windows verzio ===
>>"%OUT%" ver
>>"%OUT%" echo.

>>"%OUT%" echo === Telepitett Office (ClickToRun) ===
>>"%OUT%" reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" 2>&1
>>"%OUT%" echo.
>>"%OUT%" reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" 2>&1
>>"%OUT%" echo.

if not defined OSPP (
    >>"%OUT%" echo === ospp.vbs NEM TALALHATO ===
    >>"%OUT%" echo Nincs MSI/C2R Office16 mappa. Talan masik verzio, vagy nincs telepitve.
    goto :done
)

>>"%OUT%" echo === ospp.vbs helye ===
>>"%OUT%" echo %OSPP%
>>"%OUT%" echo.

>>"%OUT%" echo === /dstatus  (licenc allapot) ===
pushd "%OSPP%"
>>"%OUT%" cscript //nologo ospp.vbs /dstatus 2>&1
>>"%OUT%" echo.
>>"%OUT%" echo === /dstatusall  (minden telepitett licenc) ===
>>"%OUT%" cscript //nologo ospp.vbs /dstatusall 2>&1
popd

:done
>>"%OUT%" echo.
>>"%OUT%" echo === Vege ===

echo.
echo   Kesz. A fajl itt van:
echo   %OUT%
echo.
pause
