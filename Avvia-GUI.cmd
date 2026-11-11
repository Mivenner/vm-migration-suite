@echo off
rem Avvia l'interfaccia dai sorgenti (senza exe) ignorando la ExecutionPolicy
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0VMMigration.GUI.ps1"
