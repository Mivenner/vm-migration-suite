@echo off
rem Esempio: Avvia-CLI.cmd -Action List
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0VMMigration.CLI.ps1" %*
