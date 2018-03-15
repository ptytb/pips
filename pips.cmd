@echo off

rem Minimize console window
if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

rem Set-Execution policy is required, otherwise won't be able to run any scripts at all.
rem Get-Help about_Execution_Policies
powershell -Sta -NonInteractive -NoLogo -Command "Set-ExecutionPolicy -Scope Process Bypass; Import-Module .\pips; Start-Main -HideConsole"

if %errorlevel% neq 0 (
	pause
)

rem Uncomment when debugging to see error messages
rem pause

exit
