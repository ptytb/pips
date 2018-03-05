@echo off

rem Minimize console window
if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

rem .ps1 does not belong to PATHEXT by default
powershell -Sta -NonInteractive -NoLogo -Command "using module .\pips.psd1; Start-Main -HideConsole"

exit