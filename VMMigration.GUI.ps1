#Requires -Version 5.1
<#
.SYNOPSIS
    VM Migration Suite - Interfaccia grafica (WPF) v3.0
.DESCRIPTION
    Wizard in 4 passi: Modalita' -> Configurazione -> Esecuzione -> Riepilogo.
    Tutto il lavoro pesante gira in un runspace separato: la finestra resta reattiva,
    il log e la barra di avanzamento si aggiornano in tempo reale ed e' possibile annullare.
    Compilabile con PS2EXE (vedi Build.ps1). Richiede: VMMigration.Engine.psm1, GUI.xaml, logo.png nella stessa cartella.
#>
[CmdletBinding()]
param()

# --- Cartella di lavoro (script .ps1 oppure .exe PS2EXE) ------------------------------------------
$script:Root = $PSScriptRoot
if (-not $script:Root) { try { $script:Root = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch {} }
if (-not $script:Root) { $script:Root = [AppDomain]::CurrentDomain.BaseDirectory }

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

function Show-FatalError([string]$Message, $ErrorRecord) {
    $detail = $Message
    if ($ErrorRecord) { $detail += "`n`n" + $ErrorRecord.Exception.GetType().FullName + "`n" + $ErrorRecord.ScriptStackTrace + "`nRiga: " + $ErrorRecord.InvocationInfo.PositionMessage }
    $logPaths = @((Join-Path $script:Root 'vmig_error.log'), (Join-Path $env:TEMP 'vmig_error.log'))
    $written = $null
    foreach ($lp in $logPaths) { try { Add-Content -LiteralPath $lp -Value ("{0} {1}`n" -f (Get-Date -Format s), $detail) -Encoding UTF8; $written = $lp; break } catch {} }
    try { [System.Windows.MessageBox]::Show(($Message + "`n`nDettagli salvati in: " + $written), 'VM Migration Suite - errore', 'OK', 'Error') | Out-Null }
    catch { try { [System.Windows.Forms.MessageBox]::Show($Message + "`n`n" + $written, 'VM Migration Suite - errore') | Out-Null } catch {} }
}

# --- Privilegi ----------------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    # Rilancio elevato (funziona sia per .ps1 che per .exe)
    try {
        $exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($exe -match 'powershell(_ise)?\.exe$|pwsh\.exe$') { Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File', "`"$PSCommandPath`"") -Verb RunAs | Out-Null }
        else { Start-Process -FilePath $exe -Verb RunAs | Out-Null }
    } catch { Show-FatalError "Sono necessari privilegi di amministratore. Esegui l'applicazione con 'Esegui come amministratore'." }
    return
}

try {
# --- Motore + XAML ------------------------------------------------------------------------------
$EnginePath = Join-Path $script:Root 'VMMigration.Engine.psm1'
$XamlPath   = Join-Path $script:Root 'GUI.xaml'
foreach ($f in $EnginePath, $XamlPath) { if (-not (Test-Path -LiteralPath $f)) { Show-FatalError "File mancante: $f`nReinstallare l'applicazione."; return } }
# Caricamento del motore come testo (New-Module): non soggetto a ExecutionPolicy ne' al blocco "Mark of the Web"
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue } catch {}
try {
    $EngineText = [System.IO.File]::ReadAllText($EnginePath, [System.Text.Encoding]::UTF8)
    Import-Module (New-Module -Name 'VMMigration.Engine' -ScriptBlock ([scriptblock]::Create($EngineText))) -Force -ErrorAction Stop
} catch { Show-FatalError "Impossibile caricare il motore: $($_.Exception.Message)" $_; return }

try {
    [xml]$xaml = Get-Content -LiteralPath $XamlPath -Raw -Encoding UTF8
    $Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
} catch { Show-FatalError "Errore nel caricamento dell'interfaccia (GUI.xaml): $($_.Exception.Message)" $_; return }

# Tutti gli elementi con Name diventano variabili $ui.<Name>
$ui = @{}
$xaml.SelectNodes('//*[@Name]') | ForEach-Object { $ui[$_.Name] = $Window.FindName($_.Name) }

try {
    $logo = Join-Path $script:Root 'logo.png'
    if (Test-Path -LiteralPath $logo) {
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit(); $bmp.UriSource = New-Object System.Uri($logo); $bmp.CacheOption = 'OnLoad'; $bmp.EndInit()
        $ui.imgLogo.Source = $bmp
        $Window.Icon = $bmp
    }
} catch {}

$ui.lblComputer.Text = $env:COMPUTERNAME
$ui.lblLocalIp.Text  = 'IP: ' + ((Get-MigLocalIPs) -join ', ')

# --- Stato wizard -------------------------------------------------------------------------------
$script:Step = 1
$script:Mode = 'Export'
$script:LastPackage = $null
$script:LastLog = $null
$script:ReceivedPackage = $null
$script:Profiles = @()

function Get-Mode {
    if ($ui.rbImport.IsChecked)  { return 'Import' }
    if ($ui.rbReceive.IsChecked) { return 'Receive' }
    if ($ui.rbSend.IsChecked)    { return 'Send' }
    return 'Export'
}

function Set-Step([int]$n) {
    $script:Step = $n
    foreach ($p in 'pnlMode','pnlExport','pnlImport','pnlReceive','pnlSend','pnlRun','pnlDone') { $ui[$p].Visibility = 'Collapsed' }
    $ui.btnBack.IsEnabled = ($n -gt 1 -and $n -ne 3)
    $ui.btnNext.IsEnabled = $true
    switch ($n) {
        1 { $ui.pnlMode.Visibility = 'Visible'; $ui.lblStepTitle.Text = 'Passo 1 di 4 - Cosa vuoi fare?'; $ui.btnNext.Content = 'Avanti' }
        2 {
            $script:Mode = Get-Mode
            switch ($script:Mode) {
                'Export'  { $ui.pnlExport.Visibility = 'Visible';  $ui.lblStepTitle.Text = 'Passo 2 di 4 - Configura esportazione';  $ui.btnNext.Content = 'Avvia esportazione'; if (-not $script:Profiles) { Load-Profiles } ; Update-FreeSpace }
                'Import'  { $ui.pnlImport.Visibility = 'Visible';  $ui.lblStepTitle.Text = 'Passo 2 di 4 - Configura importazione';  $ui.btnNext.Content = 'Avvia importazione'; if ($ui.cmbTargetUsers.Items.Count -eq 0) { Load-TargetUsers } }
                'Receive' { $ui.pnlReceive.Visibility = 'Visible'; $ui.lblStepTitle.Text = 'Passo 2 di 4 - Ricezione dalla rete';   $ui.btnNext.Content = 'Mettiti in ascolto'; if ($ui.lblPairCode.Text -eq '------') { New-Code } ; $ui.lblRecvIps.Text = 'Questo PC: ' + $env:COMPUTERNAME + '  -  IP: ' + ((Get-MigLocalIPs) -join ', ') }
                'Send'    { $ui.pnlSend.Visibility = 'Visible';    $ui.lblStepTitle.Text = 'Passo 2 di 4 - Invio via rete';         $ui.btnNext.Content = 'Invia' }
            }
        }
        3 { $ui.pnlRun.Visibility = 'Visible'; $ui.lblStepTitle.Text = 'Passo 3 di 4 - Esecuzione'; $ui.btnNext.Content = 'Attendere...'; $ui.btnNext.IsEnabled = $false }
        4 { $ui.pnlDone.Visibility = 'Visible'; $ui.lblStepTitle.Text = 'Passo 4 di 4 - Riepilogo'; $ui.btnNext.Content = 'Chiudi'; $ui.btnBack.IsEnabled = $false }
    }
}

# --- Popolamento liste --------------------------------------------------------------------------
function New-Check([string]$Tag, [string]$Text, [bool]$Checked, [string]$Tip) {
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content = $Text; $cb.Tag = $Tag; $cb.IsChecked = $Checked
    if ($Tip) { $cb.ToolTip = $Tip }
    return $cb
}
function Add-GroupHeader($Panel, [string]$Text) {
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text; $tb.FontWeight = 'SemiBold'; $tb.Margin = '0,8,0,2'; $tb.Foreground = $Window.FindResource('Muted')
    [void]$Panel.Children.Add($tb)
}
$lastGroup = $null
foreach ($ds in Get-MigDataSets) {
    if ($ds.Gruppo -ne $lastGroup) { Add-GroupHeader $ui.pnlDataSets $ds.Gruppo; $lastGroup = $ds.Gruppo }
    [void]$ui.pnlDataSets.Children.Add((New-Check $ds.Id $ds.Nome $ds.Default ($ds.Percorsi -join "`n")))
}
foreach ($rs in Get-MigRegistrySets) { [void]$ui.pnlRegistrySets.Children.Add((New-Check $rs.Id ("Registro: " + $rs.Nome) $true ($rs.Chiavi -join "`n"))) }

function Get-CheckedTags($Panel) { @($Panel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked } | ForEach-Object { $_.Tag }) }
function Set-AllChecks($Panel, [bool]$Value) { $Panel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] } | ForEach-Object { $_.IsChecked = $Value } }

function Load-Profiles {
    $ui.cmbProfiles.Items.Clear()
    $root = $null
    if ($ui.chkOffline.IsChecked) { $root = $ui.txtOfflineRoot.Text.Trim(); if (-not $root) { return } }
    try { $script:Profiles = @(Get-MigUserProfiles -OfflineRoot $root) } catch { $script:Profiles = @(); $ui.lblProfileInfo.Text = "Errore: $($_.Exception.Message)"; return }
    foreach ($p in $script:Profiles) {
        $txt = '{0}  [{1}]  {2}{3}' -f $p.Account, $p.Tipo, $p.Percorso, $(if ($p.Connesso) { '  (connesso)' } else { '' })
        [void]$ui.cmbProfiles.Items.Add($txt)
    }
    if ($ui.cmbProfiles.Items.Count -gt 0) { $ui.cmbProfiles.SelectedIndex = 0 }
    $ui.lblProfileInfo.Text = "Trovati $($script:Profiles.Count) profili."
}
function Get-SelectedProfile { if ($ui.cmbProfiles.SelectedIndex -ge 0) { return $script:Profiles[$ui.cmbProfiles.SelectedIndex] } ; return $null }

function Load-TargetUsers {
    $ui.cmbTargetUsers.Items.Clear()
    try { Get-LocalUser | Where-Object { $_.Enabled -and $_.SID -like 'S-1-5-21-*' } | ForEach-Object { [void]$ui.cmbTargetUsers.Items.Add($_.Name) } } catch {}
    try { Get-MigUserProfiles | Where-Object { $_.Tipo -ne 'Locale' } | ForEach-Object { if ($ui.cmbTargetUsers.Items -notcontains $_.Account) { [void]$ui.cmbTargetUsers.Items.Add($_.Account) } } } catch {}
    if ($ui.cmbTargetUsers.Items.Count -gt 0) { $ui.cmbTargetUsers.SelectedIndex = 0 }
}
function New-Code { $ui.lblPairCode.Text = New-MigPairingCode }

function Update-FreeSpace {
    try {
        $d = $ui.txtExportDest.Text.Trim()
        if ($d.Length -ge 2) { $drv = Get-PSDrive -Name $d.Substring(0,1) -ErrorAction Stop; $ui.lblFreeSpace.Text = 'Spazio libero su {0}: {1}' -f $drv.Root, (Format-MigBytes $drv.Free) }
    } catch { $ui.lblFreeSpace.Text = '' }
}

function Pick-Folder([string]$Description, [string]$Initial) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $Description; $dlg.ShowNewFolderButton = $true
    if ($Initial -and (Test-Path -LiteralPath $Initial)) { $dlg.SelectedPath = $Initial }
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.SelectedPath }
    return $null
}

function Show-Manifest([string]$Path) {
    try {
        $m = Read-MigManifest -PackagePath $Path
        $ui.lblManifest.Text = ("Utente: {0} ({1})`nPC origine: {2}  -  creato: {3}`nDati: {4} in {5} file  -  dataset: {6}`nRegistro: {7} chiavi  -  errori export: {8}" -f $m.Sorgente.Account, $m.Sorgente.SID, $m.Sorgente.Computer, ([datetime]$m.Creato).ToString('g'), (Format-MigBytes $m.ByteTotali), $m.FileTotali, ($m.DataSet -join ', '), @($m.Registro).Count, @($m.Errori).Count)
        $ui.pnlImportSets.Children.Clear()
        foreach ($id in @($m.DataSet)) { $ds = Get-MigDataSets | Where-Object Id -eq $id; [void]$ui.pnlImportSets.Children.Add((New-Check $id $(if ($ds) { $ds.Nome } else { $id }) $true $null)) }
        if (@($m.Copie | Where-Object DataSet -eq '_Extra').Count -gt 0) { [void]$ui.pnlImportSets.Children.Add((New-Check '_Extra' 'Cartelle aggiuntive (-> Desktop\Cartelle migrate)' $true $null)) }
        return $true
    } catch { $ui.lblManifest.Text = "Pacchetto non valido: $($_.Exception.Message)"; $ui.pnlImportSets.Children.Clear(); return $false }
}

# --- Esecuzione in background (runspace) --------------------------------------------------------
$script:Sync = [hashtable]::Synchronized(@{})
$script:Sync.Log       = New-Object System.Collections.Concurrent.ConcurrentQueue[string]
$script:Sync.Percent   = 0.0
$script:Sync.Status    = ''
$script:Sync.Speed     = -1.0
$script:Sync.BytesDone = -1
$script:Sync.BytesTot  = -1
$script:Sync.Done      = $false
$script:Sync.Error     = $null
$script:Sync.Result    = $null
$script:Sync.CancelRef = [ref]$false
$script:Sync.Indeterminate = $false
$script:Runspace = $null
$script:PS = $null

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds(250)
$Timer.Add_Tick({
    $line = $null
    $added = $false
    while ($script:Sync.Log.TryDequeue([ref]$line)) { $ui.txtLog.AppendText($line + "`r`n"); $added = $true }
    if ($added) { $ui.txtLog.ScrollToEnd() }
    $ui.progressBar.IsIndeterminate = [bool]$script:Sync.Indeterminate
    if (-not $script:Sync.Indeterminate) { $ui.progressBar.Value = [double]$script:Sync.Percent }
    $ui.lblRunStatus.Text = [string]$script:Sync.Status
    $spd = ''
    if ($script:Sync.Speed -ge 0) { $spd = ('{0:N1} MB/s' -f $script:Sync.Speed) }
    if ($script:Sync.BytesDone -ge 0 -and $script:Sync.BytesTot -gt 0) { $spd = ('{0} / {1}   {2}' -f (Format-MigBytes $script:Sync.BytesDone), (Format-MigBytes $script:Sync.BytesTot), $spd) }
    elseif ($script:Sync.BytesDone -ge 0) { $spd = ('{0}   {1}' -f (Format-MigBytes $script:Sync.BytesDone), $spd) }
    $ui.lblRunSpeed.Text = $spd
    if ($script:Sync.Done) {
        $Timer.Stop()
        try { if ($script:PS) { $script:PS.Dispose() }; if ($script:Runspace) { $script:Runspace.Close(); $script:Runspace.Dispose() } } catch {}
        $ui.btnCancel.IsEnabled = $false
        Complete-Run
    }
})

function Start-Work([scriptblock]$Work, [hashtable]$WorkArgs) {
    $ui.txtLog.Clear()
    $script:Sync.Done = $false; $script:Sync.Error = $null; $script:Sync.Result = $null; $script:Sync.Percent = 0; $script:Sync.Status = ''; $script:Sync.Speed = -1; $script:Sync.BytesDone = -1; $script:Sync.BytesTot = -1
    $script:Sync.CancelRef = [ref]$false
    $script:Sync.Indeterminate = $false
    $ui.btnCancel.IsEnabled = $true
    $script:Runspace = [runspacefactory]::CreateRunspace()
    $script:Runspace.ApartmentState = 'STA'; $script:Runspace.ThreadOptions = 'ReuseThread'
    $script:Runspace.Open()
    $script:Runspace.SessionStateProxy.SetVariable('Sync', $script:Sync)
    $script:Runspace.SessionStateProxy.SetVariable('EnginePath', $EnginePath)
    $script:Runspace.SessionStateProxy.SetVariable('EngineText', $EngineText)
    $script:Runspace.SessionStateProxy.SetVariable('WorkArgs', $WorkArgs)
    $script:Runspace.SessionStateProxy.SetVariable('WorkText', $Work.ToString())
    $script:PS = [powershell]::Create()
    $script:PS.Runspace = $script:Runspace
    [void]$script:PS.AddScript({
        try {
            Import-Module (New-Module -Name 'VMMigration.Engine' -ScriptBlock ([scriptblock]::Create($EngineText))) -Force
            Initialize-MigLogging -OnLog {
                param($lvl, $msg)
                $Sync.Log.Enqueue(("{0} [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $lvl, $msg))
            } -OnProgress {
                param($pct, $status, $done, $tot, $spd)
                $Sync.Percent = $pct; $Sync.Status = $status; $Sync.BytesDone = $done; $Sync.BytesTot = $tot; $Sync.Speed = $spd
            } -Cancel $Sync.CancelRef
            $Work = [scriptblock]::Create($WorkText)
            $Sync.Result = & $Work $WorkArgs
        } catch {
            $Sync.Error = $_.Exception.Message
            $Sync.Log.Enqueue(("{0} [ERROR] {1}" -f (Get-Date -Format 'HH:mm:ss'), $_.Exception.Message))
            if ($_.ScriptStackTrace) { $Sync.Log.Enqueue($_.ScriptStackTrace) }
        } finally { $Sync.Done = $true }
    })
    [void]$script:PS.BeginInvoke()
    $Timer.Start()
}

# Lavori -----------------------------------------------------------------------------------------
$WorkExport = {
    param($a)
    $r = Start-MigExport -UserProfile $a.Profile -Destination $a.Dest -DataSetIds $a.DataSets -RegistrySetIds $a.RegSets -ExtraFolders $a.Extra -IncludeSettings $a.Settings -IncludeRegistry ($a.RegSets.Count -gt 0) -OfflineRoot $a.OfflineRoot
    if ($a.SendHost) {
        Write-MigLog -Level STEP -Message "Invio del pacchetto a $($a.SendHost)..."
        Send-MigPackage -PackagePath $r.Pacchetto -TargetHost $a.SendHost -Port $a.SendPort -Code $a.SendCode
    }
    return $r
}
$WorkImport = {
    param($a)
    if ($a.NewUser) { New-MigLocalUser -Name $a.NewUser -Password $a.NewPassword -FullName $a.NewFullName -Administrator $a.NewAdmin | Out-Null; $a.Target = $a.NewUser }
    return Start-MigImport -PackagePath $a.Pkg -TargetUser $a.Target -DataSetIds $a.DataSets -SkipNewer $a.SkipNewer -RestoreRegistry $a.Registry -RestoreWiFi $a.Wifi -RestorePrinters $a.Printers -RestoreDrives $a.Drives -RestoreWallpaper $a.Wallpaper -FixPermissions $a.Acl -RestoreOutlookProfile $a.OutlookProfile
}
$WorkReceive = {
    param($a)
    Enable-MigFirewall -TcpPort $a.Port
    $stop = $Sync.CancelRef
    $disc = [powershell]::Create()
    $disc.AddScript({ param($e, $p, $s) Import-Module (New-Module -Name 'VMMigration.Engine' -ScriptBlock ([scriptblock]::Create($e))) -Force; Start-MigDiscoveryResponder -TcpPort $p -Stop $s }).AddArgument($EngineText).AddArgument($a.Port).AddArgument($stop) | Out-Null
    $h = $disc.BeginInvoke()
    try {
        $Sync.Indeterminate = $true
        return Receive-MigPackage -Destination $a.Dest -Port $a.Port -Code $a.Code -Stop $stop
    } finally {
        $Sync.Indeterminate = $false
        $stop.Value = $true
        try { $disc.Stop(); $disc.Dispose() } catch {}
        Disable-MigFirewall
    }
}
$WorkSend = {
    param($a)
    Send-MigPackage -PackagePath $a.Pkg -TargetHost $a.Host -Port $a.Port -Code $a.Code
    return [pscustomobject]@{ Pacchetto = $a.Pkg }
}

function Complete-Run {
    Set-Step 4
    $ui.btnImportNow.Visibility = 'Collapsed'
    $ui.btnOpenPkg.IsEnabled = $false
    if ($script:Sync.Error) {
        $ui.lblDoneTitle.Text = 'Operazione terminata con errori'
        $ui.lblDoneText.Text = $script:Sync.Error + "`n`nControlla il log per i dettagli."
    } else {
        $ui.lblDoneTitle.Text = 'Operazione completata'
        switch ($script:Mode) {
            'Export'  { $r = $script:Sync.Result; $script:LastPackage = $r.Pacchetto; $script:LastLog = Join-Path $r.Pacchetto 'Logs\export.log'; $ui.lblDoneText.Text = ("Pacchetto creato in:`n{0}`n`n{1} in {2} file - errori di copia: {3}.`nPorta la cartella (o inviala via rete) sul nuovo PC e usa 'Importa un pacchetto'." -f $r.Pacchetto, (Format-MigBytes $r.Manifest.ByteTotali), $r.Manifest.FileTotali, $r.Errori) }
            'Import'  { $r = $script:Sync.Result; $script:LastPackage = $ui.txtImportPkg.Text; $script:LastLog = Join-Path $script:LastPackage 'Logs\import.log'; $ui.lblDoneText.Text = ("Dati ripristinati su '{0}' ({1}).`nRegistro: {2} chiavi importate, {3} fallite. Errori di copia: {4}.`n`nDisconnetti e riconnetti l'utente per vedere tutte le impostazioni." -f $r.Destinazione.Account, (Format-MigBytes $r.ByteTotali), $r.RegistroOk, $r.RegistroFalliti, @($r.Errori).Count) }
            'Receive' { $p = [string]$script:Sync.Result; $script:LastPackage = $p; $script:ReceivedPackage = $p; $script:LastLog = Join-Path $script:Root 'vmig_network.log'; $ui.lblDoneText.Text = "Pacchetto ricevuto in:`n$p"; $ui.btnImportNow.Visibility = 'Visible' }
            'Send'    { $script:LastPackage = $ui.txtSendPkg.Text; $script:LastLog = Join-Path $script:Root 'vmig_network.log'; $ui.lblDoneText.Text = 'Pacchetto inviato e verificato (SHA-256) sul PC di destinazione.' }
        }
    }
    if ($script:LastPackage -and (Test-Path -LiteralPath $script:LastPackage)) { $ui.btnOpenPkg.IsEnabled = $true }
    if ($script:Mode -in 'Receive','Send') { $ui.btnOpenLog.IsEnabled = $false } else { $ui.btnOpenLog.IsEnabled = $true }
}

# --- Validazione e avvio ------------------------------------------------------------------------
function Start-Mode {
    switch ($script:Mode) {
        'Export' {
            $p = Get-SelectedProfile
            if (-not $p) { [System.Windows.MessageBox]::Show('Seleziona un profilo da esportare.', 'VM Migration Suite') | Out-Null; return }
            $dest = $ui.txtExportDest.Text.Trim()
            if (-not $dest) { [System.Windows.MessageBox]::Show('Indica la cartella di destinazione.', 'VM Migration Suite') | Out-Null; return }
            if (-not (Test-Path -LiteralPath $dest)) { try { New-Item -ItemType Directory -Path $dest -Force | Out-Null } catch { [System.Windows.MessageBox]::Show("Impossibile creare $dest", 'VM Migration Suite') | Out-Null; return } }
            $ds = Get-CheckedTags $ui.pnlDataSets
            if ($ds.Count -eq 0) { [System.Windows.MessageBox]::Show('Seleziona almeno un gruppo di dati.', 'VM Migration Suite') | Out-Null; return }
            $opts = @{ Profile = $p; Dest = $dest; DataSets = $ds; RegSets = (Get-CheckedTags $ui.pnlRegistrySets); Extra = @($ui.lstExtra.Items); Settings = [bool]$ui.chkSettings.IsChecked; OfflineRoot = $(if ($ui.chkOffline.IsChecked) { $ui.txtOfflineRoot.Text.Trim() } else { $null }) }
            if ($ui.chkSendAfter.IsChecked) {
                if (-not $ui.txtSendAfterHost.Text.Trim() -or $ui.txtSendAfterCode.Text.Trim().Length -ne 6) { [System.Windows.MessageBox]::Show('Per inviare via rete servono PC/IP di destinazione e codice a 6 cifre.', 'VM Migration Suite') | Out-Null; return }
                $opts.SendHost = $ui.txtSendAfterHost.Text.Trim(); $opts.SendCode = $ui.txtSendAfterCode.Text.Trim(); $opts.SendPort = 8888
            }
            Set-Step 3; $ui.lblRunTitle.Text = "Esportazione di $($p.Account)"
            Start-Work $WorkExport $opts
        }
        'Import' {
            $pkg = $ui.txtImportPkg.Text.Trim()
            if (-not $pkg -or -not (Show-Manifest $pkg)) { [System.Windows.MessageBox]::Show('Seleziona un pacchetto valido.', 'VM Migration Suite') | Out-Null; return }
            $opts = @{ Pkg = $pkg; DataSets = (Get-CheckedTags $ui.pnlImportSets); SkipNewer = [bool]$ui.chkImpSkipNewer.IsChecked; Registry = [bool]$ui.chkImpRegistry.IsChecked; Wifi = [bool]$ui.chkImpWifi.IsChecked; Printers = [bool]$ui.chkImpPrinters.IsChecked; Drives = [bool]$ui.chkImpDrives.IsChecked; Wallpaper = [bool]$ui.chkImpWallpaper.IsChecked; Acl = [bool]$ui.chkImpAcl.IsChecked; OutlookProfile = [bool]$ui.chkImpOutlookProfile.IsChecked }
            if ($ui.rbTargetNew.IsChecked) {
                $n = $ui.txtNewUser.Text.Trim()
                if (-not $n -or $n -notmatch '^[^\\/:*?"<>|\[\]=,+;@]{1,20}$') { [System.Windows.MessageBox]::Show('Nome utente non valido (max 20 caratteri, senza simboli speciali).', 'VM Migration Suite') | Out-Null; return }
                if ($ui.pwdNewUser.SecurePassword.Length -lt 1) { [System.Windows.MessageBox]::Show('Inserisci una password per il nuovo utente.', 'VM Migration Suite') | Out-Null; return }
                if (Get-LocalUser -Name $n -ErrorAction SilentlyContinue) { [System.Windows.MessageBox]::Show("L'utente '$n' esiste gia': selezionalo come utente esistente.", 'VM Migration Suite') | Out-Null; return }
                $opts.NewUser = $n; $opts.NewPassword = $ui.pwdNewUser.SecurePassword.Copy(); $opts.NewFullName = $ui.txtNewFullName.Text.Trim(); $opts.NewAdmin = [bool]$ui.chkNewAdmin.IsChecked; $opts.Target = $n
            } else {
                $t = [string]$ui.cmbTargetUsers.Text; $t = $t.Trim()
                if (-not $t) { [System.Windows.MessageBox]::Show('Indica l''utente di destinazione.', 'VM Migration Suite') | Out-Null; return }
                $opts.Target = $t
            }
            $ok = [System.Windows.MessageBox]::Show("Importare il pacchetto su '$($opts.Target)'?`nI file esistenti con lo stesso nome verranno sovrascritti (salvo l'opzione 'non sovrascrivere piu' recenti').", 'Conferma importazione', 'YesNo', 'Question')
            if ($ok -ne 'Yes') { return }
            Set-Step 3; $ui.lblRunTitle.Text = "Importazione su $($opts.Target)"
            Start-Work $WorkImport $opts
        }
        'Receive' {
            $dest = $ui.txtRecvDest.Text.Trim(); $port = 0
            if (-not $dest) { [System.Windows.MessageBox]::Show('Indica la cartella di destinazione.', 'VM Migration Suite') | Out-Null; return }
            if (-not [int]::TryParse($ui.txtRecvPort.Text, [ref]$port) -or $port -lt 1024 -or $port -gt 65535) { [System.Windows.MessageBox]::Show('Porta non valida (1024-65535).', 'VM Migration Suite') | Out-Null; return }
            Set-Step 3; $ui.lblRunTitle.Text = "In ascolto - codice $($ui.lblPairCode.Text)"
            Start-Work $WorkReceive @{ Dest = $dest; Port = $port; Code = $ui.lblPairCode.Text }
        }
        'Send' {
            $pkg = $ui.txtSendPkg.Text.Trim(); $port = 0
            if (-not $pkg -or -not (Test-Path -LiteralPath (Join-Path $pkg 'manifest.json'))) { [System.Windows.MessageBox]::Show('Seleziona una cartella pacchetto valida (con manifest.json).', 'VM Migration Suite') | Out-Null; return }
            $h = [string]$ui.cmbSendHost.Text; $h = ($h -split '\s')[0].Trim()
            if (-not $h) { [System.Windows.MessageBox]::Show('Indica il PC di destinazione (nome o IP).', 'VM Migration Suite') | Out-Null; return }
            if ($ui.txtSendCode.Text.Trim().Length -ne 6) { [System.Windows.MessageBox]::Show('Inserisci il codice a 6 cifre mostrato sul PC di destinazione.', 'VM Migration Suite') | Out-Null; return }
            if (-not [int]::TryParse($ui.txtSendPort.Text, [ref]$port)) { $port = 8888 }
            Set-Step 3; $ui.lblRunTitle.Text = "Invio a $h"
            Start-Work $WorkSend @{ Pkg = $pkg; Host = $h; Port = $port; Code = $ui.txtSendCode.Text.Trim() }
        }
    }
}

# --- Eventi -------------------------------------------------------------------------------------
$ui.btnNext.Add_Click({
    switch ($script:Step) {
        1 { Set-Step 2 }
        2 { Start-Mode }
        4 { $Window.Close() }
    }
})
$ui.btnBack.Add_Click({ if ($script:Step -eq 2) { Set-Step 1 } })
$ui.btnRestart.Add_Click({ Set-Step 1 })
$ui.btnCancel.Add_Click({
    if ([System.Windows.MessageBox]::Show('Annullare l''operazione in corso?', 'VM Migration Suite', 'YesNo', 'Warning') -eq 'Yes') {
        $script:Sync.CancelRef.Value = $true; $ui.btnCancel.IsEnabled = $false; $script:Sync.Status = 'Annullamento in corso...'
    }
})
$ui.btnOpenPkg.Add_Click({ if ($script:LastPackage -and (Test-Path -LiteralPath $script:LastPackage)) { Start-Process explorer.exe -ArgumentList "`"$($script:LastPackage)`"" } })
$ui.btnOpenLog.Add_Click({ if ($script:LastLog -and (Test-Path -LiteralPath $script:LastLog)) { Start-Process notepad.exe -ArgumentList "`"$($script:LastLog)`"" } })
$ui.btnImportNow.Add_Click({
    $ui.rbImport.IsChecked = $true
    $ui.txtImportPkg.Text = $script:ReceivedPackage
    Set-Step 2
    Show-Manifest $script:ReceivedPackage | Out-Null
})

# Export
$ui.btnRefreshProfiles.Add_Click({ Load-Profiles })
$ui.chkOffline.Add_Checked({ $ui.pnlOffline.Visibility = 'Visible'; $ui.cmbProfiles.Items.Clear(); $ui.lblProfileInfo.Text = "Indica la radice del disco (es. D:\) e premi Aggiorna." })
$ui.chkOffline.Add_Unchecked({ $ui.pnlOffline.Visibility = 'Collapsed'; Load-Profiles })
$ui.btnOfflineBrowse.Add_Click({ $f = Pick-Folder 'Seleziona la radice del disco Windows (contiene \Windows e \Users)' $ui.txtOfflineRoot.Text; if ($f) { $ui.txtOfflineRoot.Text = $f; Load-Profiles } })
$ui.btnExportBrowse.Add_Click({ $f = Pick-Folder 'Cartella dove creare il pacchetto' $ui.txtExportDest.Text; if ($f) { $ui.txtExportDest.Text = $f; Update-FreeSpace } })
$ui.txtExportDest.Add_LostFocus({ Update-FreeSpace })
$ui.btnSelAll.Add_Click({ Set-AllChecks $ui.pnlDataSets $true })
$ui.btnSelDefault.Add_Click({ foreach ($cb in ($ui.pnlDataSets.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] })) { $cb.IsChecked = [bool](Get-MigDataSets | Where-Object Id -eq $cb.Tag | Select-Object -ExpandProperty Default) } })
$ui.btnExtraAdd.Add_Click({ $f = Pick-Folder 'Cartella aggiuntiva da includere' $null; if ($f -and $ui.lstExtra.Items -notcontains $f) { [void]$ui.lstExtra.Items.Add($f) } })
$ui.btnExtraRemove.Add_Click({ if ($ui.lstExtra.SelectedIndex -ge 0) { $ui.lstExtra.Items.RemoveAt($ui.lstExtra.SelectedIndex) } })
$ui.chkSendAfter.Add_Checked({ $ui.pnlSendAfter.Visibility = 'Visible' })
$ui.chkSendAfter.Add_Unchecked({ $ui.pnlSendAfter.Visibility = 'Collapsed' })
$ui.btnEstimate.Add_Click({
    $p = Get-SelectedProfile; if (-not $p) { return }
    $ui.lblEstimate.Text = 'Calcolo in corso...'; $ui.btnEstimate.IsEnabled = $false
    $Window.Dispatcher.Invoke([action]{}, 'Background')
    try {
        $est = Get-MigExportEstimate -ProfilePath $p.Percorso -DataSetIds (Get-CheckedTags $ui.pnlDataSets)
        $top = ($est.Dettaglio | Where-Object Bytes -gt 0 | Sort-Object Bytes -Descending | Select-Object -First 6 | ForEach-Object { '{0}: {1}' -f $_.Nome, (Format-MigBytes $_.Bytes) }) -join "`n"
        $ui.lblEstimate.Text = ("Totale: {0} in {1} file`n{2}" -f (Format-MigBytes $est.Bytes), $est.Files, $top)
    } catch { $ui.lblEstimate.Text = "Errore: $($_.Exception.Message)" }
    $ui.btnEstimate.IsEnabled = $true
})

# Import
$ui.btnImportBrowse.Add_Click({ $f = Pick-Folder 'Seleziona la cartella del pacchetto (VMIG_...)' $ui.txtImportPkg.Text; if ($f) { $ui.txtImportPkg.Text = $f; Show-Manifest $f | Out-Null } })
$ui.txtImportPkg.Add_LostFocus({ if ($ui.txtImportPkg.Text.Trim()) { Show-Manifest $ui.txtImportPkg.Text.Trim() | Out-Null } })
$ui.btnRefreshUsers.Add_Click({ Load-TargetUsers })
$ui.rbTargetNew.Add_Checked({ $ui.pnlNewUser.IsEnabled = $true; $ui.cmbTargetUsers.IsEnabled = $false })
$ui.rbTargetExisting.Add_Checked({ $ui.pnlNewUser.IsEnabled = $false; $ui.cmbTargetUsers.IsEnabled = $true })

# Receive / Send
$ui.btnRecvBrowse.Add_Click({ $f = Pick-Folder 'Cartella dove salvare il pacchetto ricevuto' $ui.txtRecvDest.Text; if ($f) { $ui.txtRecvDest.Text = $f } })
$ui.btnNewCode.Add_Click({ New-Code })
$ui.btnSendBrowse.Add_Click({ $f = Pick-Folder 'Seleziona la cartella del pacchetto da inviare' $ui.txtSendPkg.Text; if ($f) { $ui.txtSendPkg.Text = $f } })
$ui.btnScan.Add_Click({
    $ui.btnScan.IsEnabled = $false; $ui.btnScan.Content = 'Ricerca...'
    $Window.Dispatcher.Invoke([action]{}, 'Background')
    try {
        $found = @(Find-MigReceivers)
        $ui.cmbSendHost.Items.Clear()
        foreach ($r in $found) { [void]$ui.cmbSendHost.Items.Add(('{0}  ({1}:{2})' -f $r.IP, $r.Nome, $r.Porta)) }
        if ($found.Count -gt 0) { $ui.cmbSendHost.SelectedIndex = 0 } else { $ui.cmbSendHost.Text = ''; [System.Windows.MessageBox]::Show("Nessun PC in ascolto trovato.`nSull'altro PC avvia 'Ricevi un pacchetto dalla rete' e verifica che siano sulla stessa rete.", 'VM Migration Suite') | Out-Null }
    } catch { [System.Windows.MessageBox]::Show("Ricerca fallita: $($_.Exception.Message)", 'VM Migration Suite') | Out-Null }
    $ui.btnScan.IsEnabled = $true; $ui.btnScan.Content = 'Cerca in rete'
})

$Window.Add_Closing({
    if ($script:Step -eq 3 -and -not $script:Sync.Done) {
        if ([System.Windows.MessageBox]::Show('Un''operazione e'' in corso. Chiudere comunque?', 'VM Migration Suite', 'YesNo', 'Warning') -ne 'Yes') { $_.Cancel = $true; return }
        $script:Sync.CancelRef.Value = $true
        try { if ($script:PS) { $script:PS.Stop() } } catch {}
        try { Disable-MigFirewall } catch {}
    }
})

Set-Step 1
[void]$Window.ShowDialog()
} catch {
    Show-FatalError ("Errore imprevisto all'avvio: " + $_.Exception.Message) $_
}
