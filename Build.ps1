#Requires -Version 5.1
<#
.SYNOPSIS
    Compila VMMigrationSuite.exe (PS2EXE) e, se presente Inno Setup 6, crea l'installer.
.NOTES
    Eseguire in Windows PowerShell 5.1 come amministratore (PS2EXE compila per .NET Framework = funziona ovunque c'e' Windows).
    Output: .\dist\VMMigrationSuite.exe + file di supporto, .\dist\VMMigrationSuite_Setup_3.0.0.exe
#>
[CmdletBinding()]
param([switch]$SkipInstaller, [string]$Version = '3.0.2')

$ErrorActionPreference = 'Stop'
$src  = $PSScriptRoot
$dist = Join-Path $src 'dist'
New-Item -ItemType Directory -Path $dist -Force | Out-Null

# 1. Verifica sintassi di tutti gli script prima di compilare
& (Join-Path $src 'Tests\Test-Syntax.ps1')

# 2. PS2EXE
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host '[*] Installo il modulo ps2exe (CurrentUser)...'
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe -Force

$exe = Join-Path $dist 'VMMigrationSuite.exe'
Write-Host "[*] Compilo $exe ..."
Invoke-PS2EXE -InputFile (Join-Path $src 'VMMigration.GUI.ps1') -OutputFile $exe -IconFile (Join-Path $src 'app.ico') `
    -NoConsole -RequireAdmin -STA -x64 -Title 'VM Migration Suite' -Description 'Migrazione profili utente Windows' `
    -Company 'vm technology' -Product 'VM Migration Suite' -Copyright "(c) $(Get-Date -Format yyyy) vm technology" -Version "$Version.0"

# File di supporto accanto all'exe (il motore e la GUI vengono letti a runtime)
foreach ($f in 'VMMigration.Engine.psm1','VMMigration.CLI.ps1','GUI.xaml','logo.png','app.ico','README.md','Avvia-CLI.cmd') { Copy-Item (Join-Path $src $f) $dist -Force }
Write-Host "[+] Compilazione completata: $dist"
# Zip portable (exe + file di supporto): basta copiarlo ed eseguirlo, nessuna installazione
Compress-Archive -Path (Join-Path $dist '*') -DestinationPath (Join-Path $dist "VMMigrationSuite_Portable_$Version.zip") -Force
Write-Host "[+] Zip portable: $dist\VMMigrationSuite_Portable_$Version.zip"

# 3. Inno Setup
if ($SkipInstaller) { return }
$iscc = @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe", "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { Write-Warning 'Inno Setup 6 non trovato (winget install JRSoftware.InnoSetup). Salto la creazione dell''installer.'; return }
Write-Host "[*] Creo l'installer con $iscc ..."
& $iscc "/DMyAppVersion=$Version" "/DSourceDir=$dist" "/O$dist" (Join-Path $src 'installer.iss')
if ($LASTEXITCODE -ne 0) { throw "ISCC ha restituito $LASTEXITCODE" }
Write-Host "[+] Installer creato in $dist"
