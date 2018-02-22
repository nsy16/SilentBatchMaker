if %PROCESSOR_ARCHITECTURE% == x86 set _autoit=C:\Program Files\AutoIt3\Aut2Exe\Aut2exe.exe
if %PROCESSOR_ARCHITECTURE% == AMD64 set _autoit=C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe

"%_autoit%" /in "%~dp0silentbatchmaker-melba2.au3" /out "%~dp0sbm.exe" /pack /x86




ping localhost -n 5 >nul
