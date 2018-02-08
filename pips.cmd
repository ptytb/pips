@echo off
rem .ps1 does not belong to PATHEXT by default
powershell -Command "using module .\pips.psd1; Start-Main"
