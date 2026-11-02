@echo off
setlocal
REM Developer-only registration. Run as administrator in an elevated prompt.
REM Select RegAsm from the .NET Framework matching the CorelDRAW process architecture.
if "%~1"=="" (
 echo Usage: register-host.cmd "full-path-to-RegAsm.exe" "full-path-to-CardMakerPro.Host.dll"
 exit /b 2
)
if "%~2"=="" (
 echo Missing DLL path.
 exit /b 2
)
if not exist "%~1" exit /b 3
if not exist "%~2" exit /b 4
"%~1" "%~2" /codebase /tlb
if errorlevel 1 exit /b 1
echo COM control registered. CorelDRAW docker XML and AddDocker integration are still required.
endlocal
