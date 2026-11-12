@echo off
rem Scarica l'ultima build (release "latest") senza aprire GitHub.
rem Richiede GitHub CLI installata una volta: winget install GitHub.cli   poi   gh auth login
set DEST=%USERPROFILE%\Downloads\VMMigrationSuite
if not exist "%DEST%" mkdir "%DEST%"
gh release download latest -R Mivenner/vm-migration-suite -p "VMMigrationSuite_Portable.zip" -D "%DEST%" --clobber
if errorlevel 1 ( echo Download fallito: verifica "gh auth status" & pause & exit /b 1 )
powershell -NoProfile -Command "Expand-Archive -Path '%DEST%\VMMigrationSuite_Portable.zip' -DestinationPath '%DEST%\Portable' -Force"
echo.
echo Pronto in %DEST%\Portable  -  avvia VMMigrationSuite.exe
pause
