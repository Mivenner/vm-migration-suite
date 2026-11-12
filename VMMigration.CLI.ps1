#Requires -Version 5.1
<#
.SYNOPSIS
    VM Migration Suite - riga di comando (senza interfaccia). Per script, WinPE/Hiren's BootCD, RMM.
.EXAMPLE
    # Elenco profili
    .\VMMigration.CLI.ps1 -Action List
    # Elenco profili di un disco non avviato (WinPE)
    .\VMMigration.CLI.ps1 -Action List -OfflineRoot E:\
    # Esportazione con dataset consigliati
    .\VMMigration.CLI.ps1 -Action Export -User mario.rossi -Destination D:\Migrazioni
    # Esportazione offline di tutto, piu' una cartella extra
    .\VMMigration.CLI.ps1 -Action Export -User mario.rossi -OfflineRoot E:\ -Destination D:\Migrazioni -DataSets All -ExtraFolders E:\Lavoro
    # Importazione su utente esistente
    .\VMMigration.CLI.ps1 -Action Import -Package D:\Migrazioni\VMIG_mario.rossi_20261108-1500 -TargetUser mario.rossi
    # Importazione creando un nuovo utente locale
    .\VMMigration.CLI.ps1 -Action Import -Package ... -TargetUser mario -CreateUser -Password (Read-Host -AsSecureString) -Administrator
    # Ricezione / invio via LAN
    .\VMMigration.CLI.ps1 -Action Receive -Destination C:\Migrazioni -Code 123456
    .\VMMigration.CLI.ps1 -Action Send -Package D:\Migrazioni\VMIG_... -TargetHost 192.168.1.50 -Code 123456
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('List','Estimate','Export','Import','Receive','Send','Discover')][string]$Action,
    [string]$User,
    [string]$OfflineRoot,
    [string]$Destination,
    [string[]]$DataSets,            # Id (vedi -Action List) oppure 'All' / 'Default'
    [string[]]$RegistrySets,
    [string[]]$ExtraFolders,
    [switch]$NoSettings,
    [switch]$NoRegistry,
    [string]$Package,
    [string]$TargetUser,
    [switch]$CreateUser,
    [securestring]$Password,
    [switch]$Administrator,
    [switch]$Overwrite,              # importazione: sovrascrivi anche i file piu' recenti
    [switch]$OutlookProfile,         # importazione: importa anche il profilo MAPI di Outlook (stessa versione Office)
    [string]$TargetHost,
    [int]$Port = 8888,
    [string]$Code,
    [string]$LogFile
)

$root = $PSScriptRoot
if (-not $root) { $root = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }
$EngineText = [System.IO.File]::ReadAllText((Join-Path $root 'VMMigration.Engine.psm1'), [System.Text.Encoding]::UTF8)
Import-Module (New-Module -Name 'VMMigration.Engine' -ScriptBlock ([scriptblock]::Create($EngineText))) -Force

$colors = @{ INFO = 'Gray'; OK = 'Green'; WARN = 'Yellow'; ERROR = 'Red'; STEP = 'Cyan' }
$lastPct = -1
Initialize-MigLogging -LogPath $LogFile -OnLog { param($l, $m) Write-Host ("[{0}] {1}" -f $l, $m) -ForegroundColor $colors[$l] } -OnProgress {
    param($pct, $status, $done, $tot, $spd)
    $p = [int]$pct
    if ($p -ne $script:lastPct -or $status) { $script:lastPct = $p; Write-Progress -Activity 'VM Migration Suite' -Status $status -PercentComplete ([math]::Min(100, [math]::Max(0, $p))) }
}

try {
    switch ($Action) {
        'List' {
            Write-Host "`nPROFILI:" -ForegroundColor Cyan
            Get-MigUserProfiles -OfflineRoot $OfflineRoot | Format-Table Account, Tipo, Connesso, UltimoUso, Percorso, SID -AutoSize
            Write-Host "DATASET DISPONIBILI (Id - Nome [default]):" -ForegroundColor Cyan
            Get-MigDataSets | ForEach-Object { Write-Host ("  {0,-14} {1}{2}" -f $_.Id, $_.Nome, $(if ($_.Default) { '  [*]' } else { '' })) }
            Write-Host "SET REGISTRO:" -ForegroundColor Cyan
            Get-MigRegistrySets | ForEach-Object { Write-Host ("  {0,-14} {1}" -f $_.Id, $_.Nome) }
        }
        'Estimate' {
            $p = Get-MigUserProfiles -OfflineRoot $OfflineRoot | Where-Object { $_.Nome -eq $User -or $_.Account -eq $User -or $_.SID -eq $User } | Select-Object -First 1
            if (-not $p) { throw "Profilo '$User' non trovato." }
            $ids = $null; if ($DataSets -and $DataSets -ne 'Default') { $ids = if ($DataSets -eq 'All') { @(Get-MigDataSets | Select-Object -ExpandProperty Id) } else { $DataSets } }
            $e = Get-MigExportEstimate -ProfilePath $p.Percorso -DataSetIds $ids
            $e.Dettaglio | Where-Object Bytes -gt 0 | Sort-Object Bytes -Descending | Format-Table Nome, @{n='Dimensione';e={Format-MigBytes $_.Bytes}}, Files -AutoSize
            Write-Host ("TOTALE: {0} in {1} file" -f (Format-MigBytes $e.Bytes), $e.Files) -ForegroundColor Cyan
        }
        'Export' {
            if (-not $User -or -not $Destination) { throw 'Servono -User e -Destination.' }
            $p = Get-MigUserProfiles -OfflineRoot $OfflineRoot | Where-Object { $_.Nome -eq $User -or $_.Account -eq $User -or $_.SID -eq $User } | Select-Object -First 1
            if (-not $p) { throw "Profilo '$User' non trovato (usa -Action List)." }
            $ids = $null; if ($DataSets -and $DataSets -ne 'Default') { $ids = if ($DataSets -eq 'All') { @(Get-MigDataSets | Select-Object -ExpandProperty Id) } else { $DataSets } }
            $r = Start-MigExport -UserProfile $p -Destination $Destination -DataSetIds $ids -RegistrySetIds $RegistrySets -ExtraFolders $ExtraFolders -IncludeSettings (-not $NoSettings) -IncludeRegistry (-not $NoRegistry) -OfflineRoot $OfflineRoot
            Write-Host "`nPacchetto: $($r.Pacchetto)  (errori: $($r.Errori))" -ForegroundColor Cyan
            if ($TargetHost -and $Code) { Send-MigPackage -PackagePath $r.Pacchetto -TargetHost $TargetHost -Port $Port -Code $Code }
        }
        'Import' {
            if (-not $Package -or -not $TargetUser) { throw 'Servono -Package e -TargetUser.' }
            if ($CreateUser) { if (-not $Password) { throw '-CreateUser richiede -Password (SecureString).' }; New-MigLocalUser -Name $TargetUser -Password $Password -Administrator $Administrator.IsPresent | Out-Null }
            $r = Start-MigImport -PackagePath $Package -TargetUser $TargetUser -DataSetIds $DataSets -SkipNewer (-not $Overwrite) -RestoreOutlookProfile $OutlookProfile.IsPresent
            Write-Host ("`nImportazione completata su {0}: {1}, registro {2} ok / {3} falliti, errori copia {4}" -f $r.Destinazione.Account, (Format-MigBytes $r.ByteTotali), $r.RegistroOk, $r.RegistroFalliti, @($r.Errori).Count) -ForegroundColor Cyan
            $r.Note | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
        }
        'Receive' {
            if (-not $Destination) { throw 'Serve -Destination.' }
            if (-not $Code) { $Code = New-MigPairingCode }
            Write-Host "`n  CODICE DI ACCOPPIAMENTO: $Code`n" -ForegroundColor Green
            Enable-MigFirewall -TcpPort $Port
            $stop = [ref]$false
            $disc = [powershell]::Create(); $disc.AddScript({ param($e, $p, $s) Import-Module (New-Module -Name 'VMMigration.Engine' -ScriptBlock ([scriptblock]::Create($e))) -Force; Start-MigDiscoveryResponder -TcpPort $p -Stop $s }).AddArgument($EngineText).AddArgument($Port).AddArgument($stop) | Out-Null
            $null = $disc.BeginInvoke()
            try { $pkg = Receive-MigPackage -Destination $Destination -Port $Port -Code $Code -Stop $stop; Write-Host "`nPacchetto ricevuto: $pkg" -ForegroundColor Cyan }
            finally { $stop.Value = $true; try { $disc.Stop(); $disc.Dispose() } catch {}; Disable-MigFirewall }
        }
        'Send' {
            if (-not $Package -or -not $TargetHost -or -not $Code) { throw 'Servono -Package, -TargetHost e -Code.' }
            Send-MigPackage -PackagePath $Package -TargetHost $TargetHost -Port $Port -Code $Code
        }
        'Discover' {
            $f = @(Find-MigReceivers)
            if ($f.Count -eq 0) { Write-Host 'Nessun PC in ascolto.' -ForegroundColor Yellow } else { $f | Format-Table IP, Nome, Porta -AutoSize }
        }
    }
    Write-Progress -Activity 'VM Migration Suite' -Completed
    exit 0
} catch {
    Write-Progress -Activity 'VM Migration Suite' -Completed
    Write-Host "ERRORE: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
