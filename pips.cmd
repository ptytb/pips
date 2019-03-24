@echo off
setlocal enabledelayedexpansion

rem Minimize console window
if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

rem Set-Execution policy is required, otherwise won't be able to run any scripts at all.
rem Get-Help about_Execution_Policies
rem Main possible options: -HideConsole -Debug
rem -Mta will eventually make the app crash even if has started normally
set pwsh_args=-Sta -NoProfile -NonInteractive -NoLogo -WindowStyle Minimized -Command
powershell %pwsh_args% "Set-ExecutionPolicy -Scope Process Bypass; Import-Module .\pips; Main -Debug"

if !errorlevel! neq 0 (
	pause
)

rem Uncomment when debugging to see error messages
pause

exit
