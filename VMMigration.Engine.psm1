#Requires -Version 5.1
<#
.SYNOPSIS
    VM Migration Suite - Motore (Engine) v3.0
.DESCRIPTION
    Funzioni core per la migrazione di profili utente Windows:
      - Inventario profili (da registro ProfileList, non da cartelle)
      - Esportazione dati con robocopy (progresso reale, codici di uscita gestiti)
      - Esportazione impostazioni: hive utente (chiavi selezionate), Wi-Fi, stampanti,
        unita' di rete, sfondo, elenco applicazioni, product key
      - IMPORTAZIONE completa su utente esistente o nuovo (creazione profilo via API)
        con rimappatura SID / percorso profilo nelle chiavi di registro
      - Trasferimento LAN: protocollo con autenticazione (codice di accoppiamento),
        integrita' SHA-256 per file, ripresa, regole firewall automatiche
    Compatibile Windows PowerShell 5.1 e PowerShell 7 (Windows 10/11).
    Nessun Write-Host: tutto passa da callback -> utilizzabile da GUI, CLI e runspace.
.NOTES
    (c) vm technology - uso interno. Richiede privilegi di amministratore.
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------------------
#  COSTANTI
# ------------------------------------------------------------------------------
$script:MigVersion      = '3.0.3'
$script:PackageMagic    = 'VMIG'
$script:ManifestName    = 'manifest.json'
$script:DefaultTcpPort  = 8888
$script:DefaultUdpPort  = 8889
$script:FirewallPrefix  = 'VM Migration Suite'

# Cartelle del profilo esportabili. Weight = usato solo per ordinare in GUI.
$script:DataSets = @(
    [pscustomobject]@{ Id='Desktop';    Nome='Desktop';                      Percorsi=@('Desktop');                                   Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Documents';  Nome='Documenti';                    Percorsi=@('Documents');                                 Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Downloads';  Nome='Download';                     Percorsi=@('Downloads');                                 Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Pictures';   Nome='Immagini';                     Percorsi=@('Pictures');                                  Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Videos';     Nome='Video';                        Percorsi=@('Videos');                                    Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Music';      Nome='Musica';                       Percorsi=@('Music');                                     Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Favorites';  Nome='Preferiti / Collegamenti';     Percorsi=@('Favorites','Links','Contacts','Searches');   Default=$true;  Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='SavedGames'; Nome='Giochi salvati';               Percorsi=@('Saved Games');                               Default=$false; Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='OneDrive';   Nome='OneDrive (solo file locali)';  Percorsi=@('OneDrive*');                                 Default=$false; Gruppo='Cartelle personali' }
    [pscustomobject]@{ Id='Chrome';     Nome='Google Chrome (profili, segnalibri, estensioni)'; Percorsi=@('AppData\Local\Google\Chrome\User Data'); Default=$true; Gruppo='Browser' }
    [pscustomobject]@{ Id='Edge';       Nome='Microsoft Edge';               Percorsi=@('AppData\Local\Microsoft\Edge\User Data');   Default=$true;  Gruppo='Browser' }
    [pscustomobject]@{ Id='Firefox';    Nome='Mozilla Firefox';              Percorsi=@('AppData\Roaming\Mozilla\Firefox','AppData\Local\Mozilla\Firefox'); Default=$true; Gruppo='Browser' }
    [pscustomobject]@{ Id='Brave';      Nome='Brave';                        Percorsi=@('AppData\Local\BraveSoftware\Brave-Browser\User Data'); Default=$false; Gruppo='Browser' }
    [pscustomobject]@{ Id='Opera';      Nome='Opera';                        Percorsi=@('AppData\Roaming\Opera Software');            Default=$false; Gruppo='Browser' }
    [pscustomobject]@{ Id='OutlookPst'; Nome='Outlook - archivi PST (no OST)'; Percorsi=@('Documents\Outlook Files','AppData\Local\Microsoft\Outlook','Documenti\File di Outlook'); Default=$true; Gruppo='Office / Posta'; Filtri=@('*.pst','*.nk2','*.srs','*.xml') }
    [pscustomobject]@{ Id='OfficeUser'; Nome='Office - firme, modelli, dizionari, QuickParts'; Percorsi=@('AppData\Roaming\Microsoft\Signatures','AppData\Roaming\Microsoft\Templates','AppData\Roaming\Microsoft\UProof','AppData\Roaming\Microsoft\Document Building Blocks','AppData\Roaming\Microsoft\Stationery','AppData\Roaming\Microsoft\Proof','AppData\Roaming\Microsoft\Excel\XLSTART','AppData\Roaming\Microsoft\AddIns','AppData\Roaming\Microsoft\Office'); Default=$true; Gruppo='Office / Posta' }
    [pscustomobject]@{ Id='Thunderbird';Nome='Mozilla Thunderbird';          Percorsi=@('AppData\Roaming\Thunderbird');               Default=$false; Gruppo='Office / Posta' }
    [pscustomobject]@{ Id='StickyNotes';Nome='Sticky Notes';                 Percorsi=@('AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState'); Default=$true; Gruppo='App di Windows' }
    [pscustomobject]@{ Id='UserFonts';  Nome='Caratteri installati dall''utente'; Percorsi=@('AppData\Local\Microsoft\Windows\Fonts'); Default=$true; Gruppo='App di Windows' }
    [pscustomobject]@{ Id='Themes';     Nome='Temi e sfondo (file)';         Percorsi=@('AppData\Roaming\Microsoft\Windows\Themes','AppData\Local\Microsoft\Windows\Themes'); Default=$true; Gruppo='App di Windows' }
    [pscustomobject]@{ Id='Taskbar';    Nome='Barra applicazioni e Start (pin)'; Percorsi=@('AppData\Roaming\Microsoft\Internet Explorer\Quick Launch','AppData\Roaming\Microsoft\Windows\Start Menu'); Default=$true; Gruppo='App di Windows' }
    [pscustomobject]@{ Id='WinTerminal';Nome='Windows Terminal / PowerShell profilo'; Percorsi=@('AppData\Local\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState','Documents\WindowsPowerShell','Documents\PowerShell'); Default=$false; Gruppo='App di Windows' }
    [pscustomobject]@{ Id='AppDataRoaming'; Nome='AppData\Roaming completo (avanzato)'; Percorsi=@('AppData\Roaming'); Default=$false; Gruppo='Avanzato' }
    [pscustomobject]@{ Id='AppDataLocal';   Nome='AppData\Local completo (avanzato, grande)'; Percorsi=@('AppData\Local'); Default=$false; Gruppo='Avanzato' }
)

# Esclusioni globali robocopy (cartelle / file che NON vanno mai migrati)
$script:ExcludeDirs = @(
    'AppData\Local\Temp','AppData\Local\Microsoft\Windows\INetCache','AppData\Local\Microsoft\Windows\WebCache',
    'AppData\Local\Microsoft\Windows\Explorer','AppData\Local\Microsoft\Windows\Caches','AppData\Local\Microsoft\Windows\Notifications',
    'AppData\Local\Packages\Microsoft.Windows.Search_cw5n1h2txyewy','AppData\Local\Microsoft\Teams\Cache','AppData\Local\Microsoft\OneDrive\logs',
    'AppData\Local\CrashDumps','AppData\Local\D3DSCache','AppData\Local\NVIDIA','AppData\Local\AMD','AppData\Local\Microsoft\Windows\WER',
    'Cache','Code Cache','GPUCache','ShaderCache','GrShaderCache','Service Worker\CacheStorage','Media Cache','DawnCache','DawnGraphiteCache','DawnWebGPUCache',
    'cache2','startupCache','thumbnails','crashes','minidumps','$RECYCLE.BIN','System Volume Information',
    'node_modules','.cache','__pycache__'
)
$script:ExcludeFiles = @(
    'NTUSER.DAT','NTUSER.DAT.LOG*','ntuser.ini','UsrClass.dat*','*.ost','*.tmp','*.temp','~$*','Thumbs.db','desktop.ini',
    'IconCache.db','*.blf','*.regtrans-ms','*.lock','lockfile','parent.lock','SingletonLock','SingletonCookie','SingletonSocket','hiberfil.sys','pagefile.sys'
)

# Chiavi HKCU da esportare (relative alla radice dell'hive utente)
$script:RegistryKeys = @(
    @{ Id='Desktop';   Nome='Desktop, schermo, mouse, tastiera, lingua';   Chiavi=@('Control Panel\Desktop','Control Panel\Mouse','Control Panel\Keyboard','Control Panel\International','Control Panel\Accessibility','Control Panel\Colors','Control Panel\Cursors','Control Panel\Sound') }
    @{ Id='Explorer';  Nome='Esplora file (viste, opzioni cartelle, Start)'; Chiavi=@('Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced','Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders','Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders','Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband','Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage','Software\Microsoft\Windows\CurrentVersion\Explorer\Streams','Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3','Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs','Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths','Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts','Software\Microsoft\Windows\CurrentVersion\Themes','Software\Microsoft\Windows\CurrentVersion\Themes\Personalize','Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags','Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU') }
    @{ Id='Network';   Nome='Unita'' di rete mappate';                     Chiavi=@('Network','Software\Microsoft\Windows\CurrentVersion\Explorer\Map Network Drive MRU','Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2') }
    @{ Id='Printers';  Nome='Stampanti di rete e predefinita';              Chiavi=@('Printers\Connections','Printers\Settings','Software\Microsoft\Windows NT\CurrentVersion\Windows','Software\Microsoft\Windows NT\CurrentVersion\Devices','Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts') }
    @{ Id='Office';    Nome='Microsoft Office (opzioni, account, MRU)';     Chiavi=@('Software\Microsoft\Office','Software\Microsoft\Shared Tools','Software\Microsoft\VBA') }
    @{ Id='Apps';      Nome='Impostazioni applicazioni comuni';             Chiavi=@('Software\7-Zip','Software\WinRAR','Software\Notepad++','Software\Adobe','Software\TeamViewer','Software\Google','Software\Microsoft\Edge','Software\Mozilla','Software\VideoLAN','Software\Microsoft\Notepad','Software\Microsoft\Windows\CurrentVersion\Run','Software\Microsoft\Windows\CurrentVersion\Internet Settings','Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders','Software\Microsoft\Windows\CurrentVersion\Search','Software\Microsoft\Windows\CurrentVersion\CloudStore','Software\Microsoft\Windows\CurrentVersion\Notifications\Settings','Software\Microsoft\Windows\CurrentVersion\Explorer\Accent','Software\Microsoft\Windows\CurrentVersion\Lock Screen','Software\Microsoft\Windows\DWM','Software\Microsoft\InputPersonalization','Software\Microsoft\TabletTip','Keyboard Layout','Environment','AppEvents\Schemes','Console') }
)

# ------------------------------------------------------------------------------
#  LOG / PROGRESSO
# ------------------------------------------------------------------------------
$script:LogFile      = $null
$script:LogCallback  = $null
$script:ProgCallback = $null
$script:CancelRef    = $null

function Initialize-MigLogging {
    <# Imposta file di log e callback. $Cancel e' un [ref] a bool controllato periodicamente. #>
    param([string]$LogPath, [scriptblock]$OnLog, [scriptblock]$OnProgress, [ref]$Cancel)
    $script:LogFile = $LogPath
    $script:LogCallback = $OnLog
    $script:ProgCallback = $OnProgress
    $script:CancelRef = $Cancel
    if ($LogPath) {
        $dir = Split-Path -Parent $LogPath
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }
}

function Write-MigLog {
    param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO','OK','WARN','ERROR','STEP')][string]$Level = 'INFO')
    $line = "{0} [{1,-5}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($script:LogFile) { try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch {} }
    if ($script:LogCallback) { try { & $script:LogCallback $Level $Message } catch {} }
}

function Write-MigProgress {
    <# Percent: 0-100 complessivo. Status: testo breve. BytesDone/BytesTotal/Speed: opzionali. #>
    param([double]$Percent, [string]$Status, [long]$BytesDone = -1, [long]$BytesTotal = -1, [double]$SpeedMBs = -1)
    if ($script:ProgCallback) { try { & $script:ProgCallback ([math]::Min(100, [math]::Max(0, $Percent))) $Status $BytesDone $BytesTotal $SpeedMBs } catch {} }
}

function Test-MigCancelled {
    if ($script:CancelRef -and $script:CancelRef.Value) { throw [System.OperationCanceledException]::new('Operazione annullata dall''utente.') }
}

function Test-MigAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Format-MigBytes {
    param([double]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0} B' -f [long]$Bytes)
}

# ------------------------------------------------------------------------------
#  PROFILI UTENTE
# ------------------------------------------------------------------------------
function Get-MigUserProfiles {
    <#
    .SYNOPSIS  Elenca i profili utente reali leggendo HKLM\...\ProfileList (SID -> percorso).
               Include profili su altre unita' (es. D:\Users) e rileva se l'hive e' caricato (utente connesso).
    .PARAMETER IncludeSystem  Include account di sistema / servizio.
    .PARAMETER OfflineRoot    Radice Windows di un disco secondario (es. "E:\") per leggere il SOFTWARE hive offline.
    #>
    [CmdletBinding()]
    param([switch]$IncludeSystem, [string]$OfflineRoot)

    $result = New-Object System.Collections.Generic.List[object]
    $listKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $mounted = $null
    try {
        if ($OfflineRoot) {
            $hive = Join-Path $OfflineRoot 'Windows\System32\config\SOFTWARE'
            if (-not (Test-Path -LiteralPath $hive)) { throw "SOFTWARE hive non trovato in $hive" }
            $mounted = 'HKLM\VMIG_OFFLINE_SW'
            $p = Start-Process reg.exe -ArgumentList @('load', $mounted, "`"$hive`"") -Wait -PassThru -WindowStyle Hidden
            if ($p.ExitCode -ne 0) { throw "Impossibile montare l'hive SOFTWARE offline (codice $($p.ExitCode))." }
            $listKey = 'HKLM:\VMIG_OFFLINE_SW\Microsoft\Windows NT\CurrentVersion\ProfileList'
        }
        foreach ($k in Get-ChildItem -Path $listKey -ErrorAction SilentlyContinue) {
            $sid = $k.PSChildName
            $isUser = $sid -like 'S-1-5-21-*' -or $sid -like 'S-1-12-1-*'   # locali/dominio + Entra ID
            if (-not $isUser -and -not $IncludeSystem) { continue }
            $props = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
            $path = [string]$props.ProfileImagePath
            if (-not $path) { continue }
            $path = [Environment]::ExpandEnvironmentVariables($path)
            if ($OfflineRoot) { $path = ($path -replace '^[A-Za-z]:', $OfflineRoot.TrimEnd('\')) }
            $name = Split-Path -Leaf $path
            $account = $null
            try { $account = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value } catch {}
            $loaded = Test-Path -LiteralPath "Registry::HKEY_USERS\$sid"
            $ntuser = Join-Path $path 'NTUSER.DAT'
            $lastUse = $null
            if (Test-Path -LiteralPath $ntuser) { try { $lastUse = (Get-Item -LiteralPath $ntuser -Force).LastWriteTime } catch {} }
            $kind = if ($sid -like 'S-1-12-1-*') { 'Entra ID' } elseif ($account -and $account -like "$env:COMPUTERNAME\*") { 'Locale' } elseif ($account) { 'Dominio' } else { 'Sconosciuto' }
            $result.Add([pscustomobject]@{
                Nome        = $name
                Account     = if ($account) { $account } else { $name }
                SID         = $sid
                Percorso    = $path
                Tipo        = $kind
                Connesso    = $loaded
                UltimoUso   = $lastUse
                Esiste      = (Test-Path -LiteralPath $path)
                Temporaneo  = ($path -like '*TEMP*' -or $path -like '*.bak')
            })
        }
    } finally {
        if ($mounted) { [gc]::Collect(); Start-Process reg.exe -ArgumentList @('unload', $mounted) -Wait -WindowStyle Hidden | Out-Null }
    }
    return $result | Where-Object { $_.Esiste } | Sort-Object -Property @{Expression='Connesso';Descending=$true}, @{Expression='UltimoUso';Descending=$true}
}

function Get-MigFolderSize {
    <# Dimensione veloce via robocopy /L (molto piu' rapido di Get-ChildItem -Recurse). Ritorna [pscustomobject] Bytes, Files. #>
    param([Parameter(Mandatory)][string]$Path, [string[]]$Filtri)
    if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]@{ Bytes = 0L; Files = 0L } }
    $null = New-Item -ItemType Directory -Path "$env:TEMP\vmig_null" -Force
    $rcArgs = @("`"$Path`"", "`"$env:TEMP\vmig_null`"") + $(if ($Filtri) { $Filtri } else { @() }) + @('/L','/E','/BYTES','/NFL','/NDL','/NJH','/NP','/R:0','/W:0','/XJ')
    foreach ($d in $script:ExcludeDirs) { $rcArgs += '/XD'; $rcArgs += "`"$d`"" }
    $rcArgs += '/XF'; $rcArgs += ($script:ExcludeFiles | ForEach-Object { "`"$_`"" })
    $out = & robocopy.exe @rcArgs 2>$null
    $bytes = 0L; $files = 0L
    foreach ($line in $out) {
        if ($line -match '^\s*(Files|File)\s*:\s*(\d+)') { $files = [long]$Matches[2] }
        if ($line -match '^\s*(Bytes|Byte)\s*:\s*([\d\.]+)') { $bytes = [long]($Matches[2] -replace '\.','') }
    }
    return [pscustomobject]@{ Bytes = $bytes; Files = $files }
}

function Get-MigDataSets { return $script:DataSets }
function Get-MigRegistrySets { return $script:RegistryKeys | ForEach-Object { [pscustomobject]$_ } }

function Resolve-MigDataSetPaths {
    <# Ritorna i percorsi assoluti esistenti per un dataset nel profilo dato (gestisce wildcard tipo OneDrive*). #>
    param([Parameter(Mandatory)]$DataSet, [Parameter(Mandatory)][string]$ProfilePath)
    $paths = @()
    foreach ($rel in $DataSet.Percorsi) {
        if ($rel -like '*`**') {
            $parent = Join-Path $ProfilePath (Split-Path -Parent $rel)
            $leaf = Split-Path -Leaf $rel
            if (Test-Path -LiteralPath $parent) { Get-ChildItem -LiteralPath $parent -Directory -Filter $leaf -Force -ErrorAction SilentlyContinue | ForEach-Object { $paths += $_.FullName } }
        } else {
            $full = Join-Path $ProfilePath $rel
            if (Test-Path -LiteralPath $full) { $paths += $full }
        }
    }
    return $paths
}

function Get-MigExportEstimate {
    <# Stima dimensione/numero file per i dataset selezionati. #>
    param([Parameter(Mandatory)][string]$ProfilePath, [string[]]$DataSetIds)
    $total = 0L; $files = 0L; $detail = @()
    foreach ($ds in $script:DataSets) {
        if ($DataSetIds -and $DataSetIds -notcontains $ds.Id) { continue }
        $b = 0L; $f = 0L
        foreach ($p in (Resolve-MigDataSetPaths -DataSet $ds -ProfilePath $ProfilePath)) {
            $filtri = $null; if ($ds.PSObject.Properties['Filtri']) { $filtri = $ds.Filtri }
            $s = Get-MigFolderSize -Path $p -Filtri $filtri
            $b += $s.Bytes; $f += $s.Files
        }
        $detail += [pscustomobject]@{ Id = $ds.Id; Nome = $ds.Nome; Bytes = $b; Files = $f }
        $total += $b; $files += $f
    }
    return [pscustomobject]@{ Bytes = $total; Files = $files; Dettaglio = $detail }
}

# ------------------------------------------------------------------------------
#  ROBOCOPY CON PROGRESSO REALE
# ------------------------------------------------------------------------------
function Invoke-MigRobocopy {
    <#
    .SYNOPSIS  Esegue robocopy leggendo l'output in streaming per calcolare il progresso in byte.
    .OUTPUTS   [pscustomobject] ExitCode, Success, BytesCopied, FilesCopied, Failed
    .NOTES     Codici robocopy: 0-7 = OK (bit 1 copiato, 2 extra, 4 mismatch), >=8 = errori.
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [string[]]$FileFilters,
        [string]$LogFile,
        [long]$ExpectedBytes = 0,
        [long]$BaseBytes = 0,        # byte gia' copiati prima di questa chiamata (per il % complessivo)
        [long]$GrandTotal = 0,       # totale complessivo per il % globale
        [string]$Label = '',
        [switch]$SkipNewerAtDest,    # /XO: non sovrascrivere file piu' recenti a destinazione
        [switch]$Mirror,             # /MIR (solo per usi speciali, pericoloso)
        [int]$Threads = 16
    )
    if (-not (Test-Path -LiteralPath $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
    $a = New-Object System.Collections.Generic.List[string]
    $a.Add("`"$Source`""); $a.Add("`"$Destination`"")
    if ($FileFilters) { foreach ($f in $FileFilters) { $a.Add($f) } }
    if ($Mirror) { $a.Add('/MIR') } else { $a.Add('/E') }
    $a.AddRange([string[]]@('/COPY:DAT','/DCOPY:DAT','/B','/XJ','/R:2','/W:2',"/MT:$Threads",'/NP','/NDL','/NJH','/NJS','/NC','/BYTES','/FFT','/256'))
    if ($SkipNewerAtDest) { $a.Add('/XO') }
    foreach ($d in $script:ExcludeDirs) { $a.Add('/XD'); $a.Add("`"$d`"") }
    $a.Add('/XF'); foreach ($x in $script:ExcludeFiles) { $a.Add("`"$x`"") }
    if ($LogFile) { $a.Add("/UNILOG+:`"$LogFile`""); $a.Add('/TEE') }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'robocopy.exe'
    $psi.Arguments = ($a -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    try { $psi.StandardOutputEncoding = [System.Console]::OutputEncoding } catch {}

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $bytes = 0L; $files = 0L; $failed = 0L
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastReport = 0
    [void]$proc.Start()
    $proc.PriorityClass = 'BelowNormal'
    $reader = $proc.StandardOutput
    while (-not $reader.EndOfStream) {
        $line = $reader.ReadLine()
        if ($null -eq $line) { break }
        # Formato con /NC /BYTES /NDL: "<spazi><dimensione>\t<percorso>"  (file copiato/nuovo) ; righe errore contengono "ERROR"
        if ($line -match '^\s*(\d+)\s+(.+)$') {
            $bytes += [long]$Matches[1]; $files++
        } elseif ($line -match 'ERROR\s+\d+\s+\(0x') {
            $failed++
            Write-MigLog -Level WARN -Message ("robocopy: " + $line.Trim())
        }
        if ($sw.ElapsedMilliseconds - $lastReport -gt 400) {
            $lastReport = $sw.ElapsedMilliseconds
            $pct = if ($GrandTotal -gt 0) { (($BaseBytes + $bytes) / $GrandTotal) * 100 } else { 0 }
            $speed = if ($sw.Elapsed.TotalSeconds -gt 0) { ($bytes / 1MB) / $sw.Elapsed.TotalSeconds } else { 0 }
            Write-MigProgress -Percent $pct -Status ("{0}: {1} file, {2}" -f $Label, $files, (Format-MigBytes $bytes)) -BytesDone ($BaseBytes + $bytes) -BytesTotal $GrandTotal -SpeedMBs $speed
            if ($script:CancelRef -and $script:CancelRef.Value) { try { $proc.Kill() } catch {}; throw [System.OperationCanceledException]::new('Copia annullata dall''utente.') }
        }
    }
    $proc.WaitForExit()
    $code = $proc.ExitCode
    $ok = $code -lt 8
    if ($ok) { Write-MigLog -Level OK -Message ("{0}: completato ({1} file, {2}, codice {3})" -f $Label, $files, (Format-MigBytes $bytes), $code) }
    else { Write-MigLog -Level ERROR -Message ("{0}: robocopy ha restituito codice {1} (errori). Vedi log: {2}" -f $Label, $code, $LogFile) }
    return [pscustomobject]@{ ExitCode = $code; Success = $ok; BytesCopied = $bytes; FilesCopied = $files; Failed = $failed }
}

# ------------------------------------------------------------------------------
#  HIVE DI REGISTRO UTENTE (NTUSER.DAT)
# ------------------------------------------------------------------------------
function Mount-MigUserHive {
    <#
    .SYNOPSIS  Restituisce la radice HKU dell'utente. Se l'utente e' connesso usa HKU\<SID>,
               altrimenti monta NTUSER.DAT con un nome univoco. Smontare SEMPRE con Dismount-MigUserHive.
    #>
    param([Parameter(Mandatory)][string]$SID, [Parameter(Mandatory)][string]$ProfilePath)
    if (Test-Path -LiteralPath "Registry::HKEY_USERS\$SID") {
        return [pscustomobject]@{ Root = "HKEY_USERS\$SID"; Mounted = $false; MountName = $null }
    }
    $dat = Join-Path $ProfilePath 'NTUSER.DAT'
    if (-not (Test-Path -LiteralPath $dat)) { throw "NTUSER.DAT non trovato in $ProfilePath" }
    $name = 'VMIG_' + ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $p = Start-Process reg.exe -ArgumentList @('load', "HKU\$name", "`"$dat`"") -Wait -PassThru -WindowStyle Hidden
    if ($p.ExitCode -ne 0) { throw "Impossibile caricare l'hive $dat (codice $($p.ExitCode)). L'utente potrebbe avere processi attivi: disconnetterlo e riprovare." }
    Write-MigLog "Hive utente montato come HKU\$name"
    return [pscustomobject]@{ Root = "HKEY_USERS\$name"; Mounted = $true; MountName = $name }
}

function Dismount-MigUserHive {
    param([Parameter(Mandatory)]$Hive)
    if (-not $Hive.Mounted) { return }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    for ($i = 1; $i -le 5; $i++) {
        $p = Start-Process reg.exe -ArgumentList @('unload', "HKU\$($Hive.MountName)") -Wait -PassThru -WindowStyle Hidden
        if ($p.ExitCode -eq 0) { Write-MigLog "Hive HKU\$($Hive.MountName) smontato."; return }
        Start-Sleep -Milliseconds (500 * $i)
        [gc]::Collect()
    }
    Write-MigLog -Level WARN -Message "Impossibile smontare HKU\$($Hive.MountName): verra' rilasciato al riavvio."
}

function Export-MigRegistry {
    <# Esporta le chiavi selezionate in file .reg (uno per set). Ritorna elenco file creati. #>
    param([Parameter(Mandatory)]$Hive, [Parameter(Mandatory)][string]$OutFolder, [string[]]$SetIds)
    if (-not (Test-Path -LiteralPath $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null }
    $created = @()
    foreach ($set in $script:RegistryKeys) {
        if ($SetIds -and $SetIds -notcontains $set.Id) { continue }
        $i = 0
        foreach ($key in $set.Chiavi) {
            $full = "$($Hive.Root)\$key"
            if (-not (Test-Path -LiteralPath "Registry::$full")) { continue }
            $i++
            $file = Join-Path $OutFolder ("{0}_{1:D2}.reg" -f $set.Id, $i)
            $p = Start-Process reg.exe -ArgumentList @('export', "`"$full`"", "`"$file`"", '/y') -Wait -PassThru -WindowStyle Hidden
            if ($p.ExitCode -eq 0 -and (Test-Path -LiteralPath $file)) { $created += [pscustomobject]@{ Set = $set.Id; Chiave = $key; File = (Split-Path -Leaf $file) } }
            else { Write-MigLog -Level WARN -Message "Export registro fallito: $key" }
        }
    }
    Write-MigLog -Level OK -Message ("Registro: esportate {0} chiavi." -f $created.Count)
    return $created
}

function Import-MigRegistryFile {
    <#
    .SYNOPSIS  Importa un .reg esportato da un altro utente/PC riscrivendo:
               radice HKEY_USERS\<vecchia> -> HKEY_USERS\<nuova>, vecchio SID -> nuovo SID,
               vecchio percorso profilo -> nuovo (sia "C:\Users\x" che la forma escaped "C:\\Users\\x").
    #>
    param(
        [Parameter(Mandatory)][string]$RegFile,
        [Parameter(Mandatory)][string]$OldRoot,   # es. HKEY_USERS\S-1-5-21-... o HKEY_USERS\VMIG_xxxx
        [Parameter(Mandatory)][string]$NewRoot,
        [string]$OldSID, [string]$NewSID,
        [string]$OldProfilePath, [string]$NewProfilePath,
        [string]$OldUserName, [string]$NewUserName
    )
    $text = Get-Content -LiteralPath $RegFile -Raw -Encoding Unicode
    $text = $text -replace [regex]::Escape("[$OldRoot"), "[$NewRoot"
    $text = $text -replace [regex]::Escape("[-$OldRoot"), "[-$NewRoot"
    if ($OldSID -and $NewSID -and $OldSID -ne $NewSID) { $text = $text -replace [regex]::Escape($OldSID), $NewSID }
    if ($OldProfilePath -and $NewProfilePath -and $OldProfilePath -ne $NewProfilePath) {
        $text = $text -replace [regex]::Escape($OldProfilePath.Replace('\', '\\')), $NewProfilePath.Replace('\', '\\')
        $text = $text -replace [regex]::Escape($OldProfilePath), $NewProfilePath
        # forma hex(2) (REG_EXPAND_SZ) non viene riscritta: contiene %USERPROFILE% nella maggior parte dei casi.
    }
    $tmp = Join-Path $env:TEMP ('vmig_' + [guid]::NewGuid().ToString('N') + '.reg')
    [System.IO.File]::WriteAllText($tmp, $text, [System.Text.Encoding]::Unicode)
    try {
        $p = Start-Process reg.exe -ArgumentList @('import', "`"$tmp`"") -Wait -PassThru -WindowStyle Hidden
        return ($p.ExitCode -eq 0)
    } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

# ------------------------------------------------------------------------------
#  IMPOSTAZIONI DI SISTEMA / UTENTE
# ------------------------------------------------------------------------------
$script:OfflineSwMount = $null
function Mount-MigOfflineSoftwareHive {
    <# Monta <OfflineRoot>\Windows\System32\config\SOFTWARE come HKLM\VMIG_SW e ritorna la radice (es. HKEY_LOCAL_MACHINE\VMIG_SW). #>
    param([Parameter(Mandatory)][string]$OfflineRoot)
    if ($script:OfflineSwMount) { return $script:OfflineSwMount }
    $hive = Join-Path $OfflineRoot 'Windows\System32\config\SOFTWARE'
    if (-not (Test-Path -LiteralPath $hive)) { throw "SOFTWARE hive non trovato: $hive" }
    $p = Start-Process reg.exe -ArgumentList @('load', 'HKLM\VMIG_SW', "`"$hive`"") -Wait -PassThru -WindowStyle Hidden
    if ($p.ExitCode -ne 0) { throw "Impossibile montare l'hive SOFTWARE offline (codice $($p.ExitCode))." }
    $script:OfflineSwMount = 'HKEY_LOCAL_MACHINE\VMIG_SW'
    return $script:OfflineSwMount
}
function Dismount-MigOfflineSoftwareHive {
    if (-not $script:OfflineSwMount) { return }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    for ($i = 1; $i -le 5; $i++) { $p = Start-Process reg.exe -ArgumentList @('unload', 'HKLM\VMIG_SW') -Wait -PassThru -WindowStyle Hidden; if ($p.ExitCode -eq 0) { break }; Start-Sleep -Milliseconds (400 * $i) }
    $script:OfflineSwMount = $null
}

function Get-MigWindowsProductKey {
    param([string]$OfflineRoot)
    $key = $null
    if (-not $OfflineRoot) { try { $key = (Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop).OA3xOriginalProductKey } catch {} }
    else { try { $key = (Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop).OA3xOriginalProductKey; if ($key) { return [pscustomobject]@{ Origine = 'UEFI/OEM (firmware di questa macchina)'; Chiave = $key } } } catch {} }
    if ($key) { return [pscustomobject]@{ Origine = 'UEFI/OEM'; Chiave = $key } }
    try {
        $cvKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        if ($OfflineRoot) { $cvKey = 'Registry::' + (Mount-MigOfflineSoftwareHive -OfflineRoot $OfflineRoot) + '\Microsoft\Windows NT\CurrentVersion' }
        $dpid = (Get-ItemProperty -LiteralPath $cvKey -Name DigitalProductId -ErrorAction Stop).DigitalProductId
        $map = 'BCDFGHJKMPQRTVWXY2346789'; $id = $dpid[52..66]; $isN = ([math]::Truncate($id[14] / 6)) -band 1
        $id[14] = ($id[14] -band 0xF7) -bor (($isN -band 2) * 4)
        $out = ''
        for ($i = 24; $i -ge 0; $i--) { $cur = 0; for ($j = 14; $j -ge 0; $j--) { $cur = $cur * 256 -bxor $id[$j]; $id[$j] = [math]::Truncate($cur / 24); $cur = $cur % 24 }; $out = $map[$cur] + $out }
        if ($isN) { $out = $out.Substring(1).Insert($cur, 'N') }
        $out = ($out -replace '(.{5})(?=.)', '$1-')
        return [pscustomobject]@{ Origine = 'Registro (potrebbe essere chiave generica se licenza digitale)'; Chiave = $out }
    } catch { return [pscustomobject]@{ Origine = 'n/d'; Chiave = 'Licenza digitale: nessuna chiave leggibile' } }
}

function Get-MigInstalledApps {
    param([string]$OfflineRoot)
    $root = 'HKLM:\SOFTWARE'
    if ($OfflineRoot) { $root = 'Registry::' + (Mount-MigOfflineSoftwareHive -OfflineRoot $OfflineRoot) }
    $paths = @("$root\Microsoft\Windows\CurrentVersion\Uninstall\*","$root\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    $apps = Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and -not $_.SystemComponent -and -not $_.ParentKeyName -and ($_.ReleaseType -ne 'Update') } |
        Select-Object @{n='Nome';e={$_.DisplayName}}, @{n='Versione';e={$_.DisplayVersion}}, @{n='Produttore';e={$_.Publisher}}, @{n='Data';e={$_.InstallDate}}, @{n='Disinstalla';e={$_.UninstallString}} |
        Sort-Object Nome -Unique
    return $apps
}

function Export-MigSettings {
    <# Esporta Wi-Fi, stampanti, unita' mappate, app installate, product key, info sistema nella cartella Settings. #>
    param([Parameter(Mandatory)][string]$OutFolder, [Parameter(Mandatory)]$Hive, [Parameter(Mandatory)][string]$ProfilePath, [string]$OfflineRoot)
    if (-not (Test-Path -LiteralPath $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null }
    $summary = @{}
    $offline = [bool]$OfflineRoot

    # Wi-Fi (chiavi in chiaro: il pacchetto va trattato come riservato) - solo sistema avviato
    if ($offline) { Write-MigLog -Level INFO -Message 'Modalita'' offline (WinPE): Wi-Fi, stampanti locali e winget non disponibili; leggo quanto possibile dagli hive del disco.' }
    if (-not $offline) { try {
        $wifi = Join-Path $OutFolder 'WiFi'; New-Item -ItemType Directory -Path $wifi -Force | Out-Null
        $null = & netsh.exe wlan export profile key=clear folder="$wifi" 2>&1
        $summary.WiFi = @(Get-ChildItem -LiteralPath $wifi -Filter *.xml -ErrorAction SilentlyContinue).Count
        Write-MigLog -Level OK -Message "Wi-Fi: esportati $($summary.WiFi) profili."
    } catch { Write-MigLog -Level WARN -Message "Wi-Fi: $($_.Exception.Message)" } }

    # Stampanti (locali/IP e connessioni di rete dell'utente)
    try {
        $printers = @()
        foreach ($pr in @($(if ($offline) { @() } else { Get-Printer -ErrorAction SilentlyContinue }))) {
            $port = $null; try { $port = Get-PrinterPort -Name $pr.PortName -ErrorAction SilentlyContinue } catch {}
            $printers += [pscustomobject]@{ Nome = $pr.Name; Driver = $pr.DriverName; Porta = $pr.PortName; Tipo = [string]$pr.Type; Condivisa = $pr.Shared; Server = $pr.ComputerName; IndirizzoIP = $(if ($port -and $port.PSObject.Properties['PrinterHostAddress']) { $port.PrinterHostAddress } else { $null }); Predefinita = $false }
        }
        $connKey = "Registry::$($Hive.Root)\Printers\Connections"
        $netPrinters = @()
        if (Test-Path -LiteralPath $connKey) { $netPrinters = @(Get-ChildItem -LiteralPath $connKey | ForEach-Object { ($_.PSChildName -replace ',', '\') }) }
        $defPrinter = $null
        try { $dev = (Get-ItemProperty -LiteralPath "Registry::$($Hive.Root)\Software\Microsoft\Windows NT\CurrentVersion\Windows" -Name Device -ErrorAction Stop).Device; $defPrinter = ($dev -split ',')[0] } catch {}
        $obj = [pscustomobject]@{ Locali = $printers; DiRete = $netPrinters; Predefinita = $defPrinter }
        $obj | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutFolder 'printers.json') -Encoding UTF8
        $summary.Stampanti = $printers.Count + $netPrinters.Count
        Write-MigLog -Level OK -Message "Stampanti: $($printers.Count) locali/IP, $($netPrinters.Count) di rete. Predefinita: $defPrinter"
    } catch { Write-MigLog -Level WARN -Message "Stampanti: $($_.Exception.Message)" }

    # Unita' di rete mappate (dalla chiave Network dell'hive)
    try {
        $drives = @()
        $nk = "Registry::$($Hive.Root)\Network"
        if (Test-Path -LiteralPath $nk) {
            foreach ($d in Get-ChildItem -LiteralPath $nk) {
                $p = Get-ItemProperty -LiteralPath $d.PSPath
                $drives += [pscustomobject]@{ Lettera = $d.PSChildName.ToUpper(); Percorso = $p.RemotePath; Utente = $p.UserName; Provider = $p.ProviderName }
            }
        }
        $drives | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $OutFolder 'drives.json') -Encoding UTF8
        $summary.Unita = $drives.Count
        Write-MigLog -Level OK -Message "Unita' di rete: $($drives.Count) mappature."
    } catch { Write-MigLog -Level WARN -Message "Unita' di rete: $($_.Exception.Message)" }

    # Sfondo desktop
    try {
        $wp = (Get-ItemProperty -LiteralPath "Registry::$($Hive.Root)\Control Panel\Desktop" -Name Wallpaper -ErrorAction SilentlyContinue).Wallpaper
        $style = (Get-ItemProperty -LiteralPath "Registry::$($Hive.Root)\Control Panel\Desktop" -Name WallpaperStyle -ErrorAction SilentlyContinue).WallpaperStyle
        $src = $null
        if ($wp) { $src = [Environment]::ExpandEnvironmentVariables($wp) -replace '^%USERPROFILE%', $ProfilePath }
        $trans = Join-Path $ProfilePath 'AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper'
        if ($src -and (Test-Path -LiteralPath $src)) { Copy-Item -LiteralPath $src -Destination (Join-Path $OutFolder ('wallpaper' + [IO.Path]::GetExtension($src))) -Force }
        elseif (Test-Path -LiteralPath $trans) { Copy-Item -LiteralPath $trans -Destination (Join-Path $OutFolder 'wallpaper.jpg') -Force }
        [pscustomobject]@{ Originale = $wp; Stile = $style } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutFolder 'wallpaper.json') -Encoding UTF8
        Write-MigLog -Level OK -Message "Sfondo: $wp"
    } catch { Write-MigLog -Level WARN -Message "Sfondo: $($_.Exception.Message)" }

    # Applicazioni installate (inventario + winget export per reinstallazione)
    try {
        $apps = @(Get-MigInstalledApps -OfflineRoot $OfflineRoot)
        $apps | ConvertTo-Json -Depth 2 | Set-Content -LiteralPath (Join-Path $OutFolder 'apps.json') -Encoding UTF8
        $apps | Export-Csv -LiteralPath (Join-Path $OutFolder 'apps.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $summary.Applicazioni = $apps.Count
        if (-not $offline -and (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
            $null = & winget.exe export -o (Join-Path $OutFolder 'winget.json') --accept-source-agreements 2>&1
            if (Test-Path -LiteralPath (Join-Path $OutFolder 'winget.json')) { Write-MigLog -Level OK -Message "winget: elenco pacchetti esportato (reinstallabile con 'winget import')." }
        }
        Write-MigLog -Level OK -Message "Applicazioni: inventario di $($apps.Count) programmi."
    } catch { Write-MigLog -Level WARN -Message "Applicazioni: $($_.Exception.Message)" }

    # Sistema / licenza / fuso / layout
    try {
        $pk = Get-MigWindowsProductKey -OfflineRoot $OfflineRoot
        if ($offline) {
            $sw = Mount-MigOfflineSoftwareHive -OfflineRoot $OfflineRoot
            try {
                $cv = Get-ItemProperty -LiteralPath "Registry::$sw\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
                $sys = [pscustomobject]@{ Computer = '(offline)'; Dominio = $null; Produttore = $null; Modello = $null; SO = $cv.ProductName; Versione = "$($cv.CurrentMajorVersionNumber).$($cv.CurrentMinorVersionNumber)"; Build = $cv.CurrentBuildNumber; Architettura = $null; FusoOrario = $null; ProductKey = $pk.Chiave; OrigineKey = $pk.Origine; Lingua = $null; LayoutTastiera = @(); DiscoOffline = $OfflineRoot }
            } finally { Dismount-MigOfflineSoftwareHive }
        } else {
            $os = Get-CimInstance Win32_OperatingSystem
            $cs = Get-CimInstance Win32_ComputerSystem
            $sys = [pscustomobject]@{
                Computer = $env:COMPUTERNAME; Dominio = $cs.Domain; Produttore = $cs.Manufacturer; Modello = $cs.Model
                SO = $os.Caption; Versione = $os.Version; Build = $os.BuildNumber; Architettura = $os.OSArchitecture
                FusoOrario = (Get-TimeZone).Id; ProductKey = $pk.Chiave; OrigineKey = $pk.Origine
                Lingua = (Get-Culture).Name; LayoutTastiera = @(Get-WinUserLanguageList -ErrorAction SilentlyContinue | ForEach-Object { $_.LanguageTag })
            }
        }
        $sys | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $OutFolder 'system.json') -Encoding UTF8
        Write-MigLog -Level OK -Message "Sistema: $($sys.SO) build $($sys.Build) - key $($pk.Origine)"
    } catch { Write-MigLog -Level WARN -Message "Sistema: $($_.Exception.Message)" }

    # Account Outlook (solo inventario: le password restano in Credential Manager e NON sono migrabili)
    try {
        $prof = "Registry::$($Hive.Root)\Software\Microsoft\Office\16.0\Outlook\Profiles"
        if (Test-Path -LiteralPath $prof) {
            $names = @(Get-ChildItem -LiteralPath $prof | Select-Object -ExpandProperty PSChildName)
            $names | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutFolder 'outlook-profiles.json') -Encoding UTF8
            Write-MigLog -Level INFO -Message "Outlook: profili trovati: $($names -join ', '). Le password andranno reinserite al primo avvio."
        }
    } catch {}

    return $summary
}

# ------------------------------------------------------------------------------
#  ESPORTAZIONE
# ------------------------------------------------------------------------------
function Start-MigExport {
    <#
    .SYNOPSIS  Crea un pacchetto di migrazione: <Destination>\VMIG_<utente>_<data>\{manifest.json, Data\, Registry\, Settings\, Logs\}
    .PARAMETER UserProfile      Oggetto da Get-MigUserProfiles.
    .PARAMETER DataSetIds   Id dei dataset da includere (default: quelli con Default=$true).
    .PARAMETER RegistrySetIds  Id dei set di registro (default: tutti).
    .PARAMETER ExtraFolders Cartelle assolute aggiuntive (es. D:\Lavoro) -> finiscono in Data\_Extra\<nome>.
    .PARAMETER IncludeSettings  Esporta Wi-Fi, stampanti, unita', sfondo, app.
    .PARAMETER PackageName  Nome cartella pacchetto (default auto).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$UserProfile,
        [Parameter(Mandatory)][string]$Destination,
        [string[]]$DataSetIds,
        [string[]]$RegistrySetIds,
        [string[]]$ExtraFolders,
        [bool]$IncludeSettings = $true,
        [bool]$IncludeRegistry = $true,
        [string]$PackageName,
        [string]$OfflineRoot      # es. 'E:\' quando si esporta da WinPE un disco Windows non avviato
    )
    if (-not (Test-MigAdmin)) { throw 'Sono necessari privilegi di amministratore.' }
    if (-not $DataSetIds) { $DataSetIds = @($script:DataSets | Where-Object Default | Select-Object -ExpandProperty Id) }
    if (-not $PackageName) { $PackageName = 'VMIG_{0}_{1}' -f ($UserProfile.Nome -replace '[^\w\-]', '_'), (Get-Date -Format 'yyyyMMdd-HHmm') }
    $pkg = Join-Path $Destination $PackageName
    if ($pkg.TrimEnd('\').ToLower().StartsWith($UserProfile.Percorso.TrimEnd('\').ToLower())) { throw "La destinazione non puo' trovarsi dentro il profilo da esportare." }
    foreach ($d in 'Data','Registry','Settings','Logs') { New-Item -ItemType Directory -Path (Join-Path $pkg $d) -Force | Out-Null }
    $script:LogFile = Join-Path $pkg 'Logs\export.log'
    $started = Get-Date
    Write-MigLog -Level STEP -Message "=== ESPORTAZIONE profilo '$($UserProfile.Account)' ($($UserProfile.SID)) -> $pkg ==="
    Write-MigLog "Motore v$script:MigVersion - $env:COMPUTERNAME - connesso: $($UserProfile.Connesso)"
    if ($UserProfile.Connesso -and $UserProfile.SID -eq ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)) {
        Write-MigLog -Level WARN -Message "Stai esportando l'utente attualmente connesso: chiudi browser/Outlook per evitare file bloccati o incoerenti."
    }

    # 1. Stima + spazio
    Write-MigProgress -Percent 0 -Status 'Analisi dimensioni...'
    $est = Get-MigExportEstimate -ProfilePath $UserProfile.Percorso -DataSetIds $DataSetIds
    $extraBytes = 0L
    foreach ($ef in @($ExtraFolders | Where-Object { $_ })) { $extraBytes += (Get-MigFolderSize -Path $ef).Bytes }
    $grand = $est.Bytes + $extraBytes
    $free = $null; $driveName = 'destinazione'
    try { $drv = (Get-Item -LiteralPath $Destination).PSDrive; if ($drv) { $driveName = $drv.Name + ':'; $free = (Get-PSDrive -Name $drv.Name).Free } } catch {}
    if ($null -eq $free -and $Destination -like '\\*') { try { $d = New-Object System.IO.DriveInfo($Destination); $free = $d.AvailableFreeSpace } catch {} }
    Write-MigLog ("Da copiare: {0} in {1} file. Spazio libero su {2}: {3}" -f (Format-MigBytes $grand), $est.Files, $driveName, $(if ($null -ne $free) { Format-MigBytes $free } else { 'n/d' }))
    if ($null -ne $free -and $free -lt ($grand * 1.05 + 200MB)) { throw ("Spazio insufficiente sulla destinazione: servono almeno {0}, disponibili {1}." -f (Format-MigBytes ($grand * 1.05 + 200MB)), (Format-MigBytes $free)) }
    Test-MigCancelled

    # 2. Dati
    $done = 0L; $results = @()
    foreach ($ds in $script:DataSets) {
        if ($DataSetIds -notcontains $ds.Id) { continue }
        $filtri = $null; if ($ds.PSObject.Properties['Filtri']) { $filtri = $ds.Filtri }
        foreach ($src in (Resolve-MigDataSetPaths -DataSet $ds -ProfilePath $UserProfile.Percorso)) {
            Test-MigCancelled
            $rel = $src.Substring($UserProfile.Percorso.Length).TrimStart('\')
            $dst = Join-Path (Join-Path $pkg 'Data') $rel
            Write-MigLog -Level STEP -Message "Copio $($ds.Nome): $rel"
            $r = Invoke-MigRobocopy -Source $src -Destination $dst -FileFilters $filtri -LogFile (Join-Path $pkg 'Logs\robocopy.log') -BaseBytes $done -GrandTotal $grand -Label $ds.Nome
            $done += $r.BytesCopied
            $results += [pscustomobject]@{ DataSet = $ds.Id; Relativo = $rel; File = $r.FilesCopied; Byte = $r.BytesCopied; Codice = $r.ExitCode; Ok = $r.Success }
        }
    }
    foreach ($ef in @($ExtraFolders | Where-Object { $_ -and (Test-Path -LiteralPath $_) })) {
        Test-MigCancelled
        $leaf = ($ef -replace '[:\\]+', '_').Trim('_')
        $dst = Join-Path $pkg "Data\_Extra\$leaf"
        Write-MigLog -Level STEP -Message "Copio cartella extra: $ef"
        $r = Invoke-MigRobocopy -Source $ef -Destination $dst -LogFile (Join-Path $pkg 'Logs\robocopy.log') -BaseBytes $done -GrandTotal $grand -Label (Split-Path -Leaf $ef)
        $done += $r.BytesCopied
        $results += [pscustomobject]@{ DataSet = '_Extra'; Relativo = "_Extra\$leaf"; Origine = $ef; File = $r.FilesCopied; Byte = $r.BytesCopied; Codice = $r.ExitCode; Ok = $r.Success }
    }

    # 3. Registro + impostazioni
    $hive = $null; $regFiles = @(); $settings = @{}
    try {
        if ($IncludeRegistry -or $IncludeSettings) {
            Write-MigProgress -Percent 96 -Status 'Registro e impostazioni...'
            $hive = Mount-MigUserHive -SID $UserProfile.SID -ProfilePath $UserProfile.Percorso
            if ($IncludeRegistry) { $regFiles = @(Export-MigRegistry -Hive $hive -OutFolder (Join-Path $pkg 'Registry') -SetIds $RegistrySetIds) }
            if ($IncludeSettings) { $settings = Export-MigSettings -OutFolder (Join-Path $pkg 'Settings') -Hive $hive -ProfilePath $UserProfile.Percorso -OfflineRoot $OfflineRoot }
        }
    } catch { Write-MigLog -Level ERROR -Message "Registro/impostazioni: $($_.Exception.Message)" }
    finally { if ($hive) { Dismount-MigUserHive -Hive $hive }; Dismount-MigOfflineSoftwareHive }

    # 4. Manifest
    $hiveRoot = if ($hive) { $hive.Root } else { "HKEY_USERS\$($UserProfile.SID)" }
    $manifest = [ordered]@{
        Magic = $script:PackageMagic; Versione = $script:MigVersion; Creato = $started.ToString('o'); Completato = (Get-Date).ToString('o')
        Sorgente = [ordered]@{ Computer = $(if ($OfflineRoot) { "(offline: $OfflineRoot)" } else { $env:COMPUTERNAME }); Utente = $UserProfile.Nome; Account = $UserProfile.Account; SID = $UserProfile.SID; Percorso = $UserProfile.Percorso; Tipo = $UserProfile.Tipo; HiveRoot = $hiveRoot; Offline = [bool]$OfflineRoot }
        DataSet = $DataSetIds; Copie = $results; ByteTotali = $done; FileTotali = ($results | Measure-Object -Property File -Sum).Sum
        Registro = $regFiles; Impostazioni = $settings
        Errori = @($results | Where-Object { -not $_.Ok } | ForEach-Object { $_.Relativo })
    }
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $pkg $script:ManifestName) -Encoding UTF8
    $errCount = @($manifest.Errori).Count
    Write-MigProgress -Percent 100 -Status 'Esportazione completata'
    Write-MigLog -Level STEP -Message ("=== COMPLETATO in {0:mm\:ss}: {1} in {2} file, {3} errori. Pacchetto: {4} ===" -f ((Get-Date) - $started), (Format-MigBytes $done), $manifest.FileTotali, $errCount, $pkg)
    return [pscustomobject]@{ Pacchetto = $pkg; Manifest = $manifest; Errori = $errCount }
}

# ------------------------------------------------------------------------------
#  UTENTE E PROFILO DI DESTINAZIONE
# ------------------------------------------------------------------------------
$script:UserEnvType = $null
function Initialize-MigNative {
    if ($script:UserEnvType) { return }
    $script:UserEnvType = Add-Type -Namespace VMig -Name Native -PassThru -MemberDefinition @'
[DllImport("userenv.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern int CreateProfile([MarshalAs(UnmanagedType.LPWStr)] string pszUserSid, [MarshalAs(UnmanagedType.LPWStr)] string pszUserName, [Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszProfilePath, uint cchProfilePath);
'@
}

function New-MigLocalUser {
    <# Crea un utente locale (nessun Read-Host: password come SecureString, adatto a GUI). #>
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][securestring]$Password, [string]$FullName, [bool]$Administrator = $false, [bool]$PasswordNeverExpires = $true)
    if (Get-LocalUser -Name $Name -ErrorAction SilentlyContinue) { throw "L'utente '$Name' esiste gia'." }
    $u = New-LocalUser -Name $Name -Password $Password -FullName $(if ($FullName) { $FullName } else { $Name }) -PasswordNeverExpires:$PasswordNeverExpires -AccountNeverExpires -ErrorAction Stop
    # Gruppi: usa i SID per funzionare anche su Windows in italiano (Users = S-1-5-32-545, Administrators = S-1-5-32-544)
    Add-LocalGroupMember -SID 'S-1-5-32-545' -Member $u -ErrorAction SilentlyContinue
    if ($Administrator) { Add-LocalGroupMember -SID 'S-1-5-32-544' -Member $u -ErrorAction Stop }
    Write-MigLog -Level OK -Message "Utente locale '$Name' creato (amministratore: $Administrator)."
    return $u
}

function Resolve-MigTargetUser {
    <#
    .SYNOPSIS  Dato un nome (locale, DOMINIO\utente o Entra "AzureAD\utente"), ritorna SID + percorso profilo,
               creando il profilo con l'API CreateProfile se l'utente non ha mai eseguito l'accesso.
    #>
    param([Parameter(Mandatory)][string]$UserName)
    $acct = New-Object System.Security.Principal.NTAccount($UserName)
    $sid = $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value
    $plKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
    $path = $null
    if (Test-Path -LiteralPath $plKey) { $path = [Environment]::ExpandEnvironmentVariables((Get-ItemProperty -LiteralPath $plKey).ProfileImagePath) }
    if (-not $path -or -not (Test-Path -LiteralPath $path)) {
        Write-MigLog "L'utente '$UserName' non ha ancora un profilo: lo creo con CreateProfile()..."
        Initialize-MigNative
        $sb = New-Object System.Text.StringBuilder 260
        $leaf = ($UserName -split '\\')[-1]
        $hr = [VMig.Native]::CreateProfile($sid, $leaf, $sb, 260)
        if ($hr -ne 0) { throw ("CreateProfile fallito (HRESULT 0x{0:X8}) per {1}" -f $hr, $UserName) }
        $path = $sb.ToString()
        Write-MigLog -Level OK -Message "Profilo creato: $path"
    }
    $loaded = Test-Path -LiteralPath "Registry::HKEY_USERS\$sid"
    return [pscustomobject]@{ Nome = ($UserName -split '\\')[-1]; Account = $UserName; SID = $sid; Percorso = $path; Connesso = $loaded }
}

function Repair-MigProfileAcl {
    <# Imposta proprietario e permessi completi all'utente sui dati ripristinati (necessario quando si copia come admin/SYSTEM). #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$SID)
    $sidArg = "*$SID"
    $null = & icacls.exe "$Path" /setowner $sidArg /T /C /Q 2>&1
    $null = & icacls.exe "$Path" /grant "${sidArg}:(OI)(CI)F" /T /C /Q 2>&1
}

# ------------------------------------------------------------------------------
#  IMPORTAZIONE
# ------------------------------------------------------------------------------
function Read-MigManifest {
    param([Parameter(Mandatory)][string]$PackagePath)
    $mf = Join-Path $PackagePath $script:ManifestName
    if (-not (Test-Path -LiteralPath $mf)) { throw "manifest.json non trovato in ${PackagePath}: non e' un pacchetto VM Migration valido." }
    $m = Get-Content -LiteralPath $mf -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($m.Magic -ne $script:PackageMagic) { throw 'Formato pacchetto non riconosciuto.' }
    return $m
}

function Start-MigImport {
    <#
    .SYNOPSIS  Ripristina un pacchetto su un utente di questo PC.
    .PARAMETER TargetUser     Nome utente di destinazione (deve esistere; usare New-MigLocalUser prima se serve).
    .PARAMETER DataSetIds     Sottoinsieme dei dataset del pacchetto (default: tutti quelli presenti).
    .PARAMETER SkipNewer      Non sovrascrivere file piu' recenti gia' presenti a destinazione.
    .PARAMETER Restore*       Flag per registro / Wi-Fi / stampanti / unita' / sfondo.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PackagePath,
        [Parameter(Mandatory)][string]$TargetUser,
        [string[]]$DataSetIds,
        [bool]$SkipNewer = $true,
        [bool]$RestoreRegistry = $true,
        [bool]$RestoreWiFi = $true,
        [bool]$RestorePrinters = $true,
        [bool]$RestoreDrives = $true,
        [bool]$RestoreWallpaper = $true,
        [bool]$FixPermissions = $true
    )
    if (-not (Test-MigAdmin)) { throw 'Sono necessari privilegi di amministratore.' }
    $m = Read-MigManifest -PackagePath $PackagePath
    $script:LogFile = Join-Path $PackagePath 'Logs\import.log'
    $started = Get-Date
    $target = Resolve-MigTargetUser -UserName $TargetUser
    Write-MigLog -Level STEP -Message "=== IMPORTAZIONE $PackagePath -> '$($target.Account)' ($($target.SID)) in $($target.Percorso) ==="
    Write-MigLog "Sorgente: $($m.Sorgente.Computer)\$($m.Sorgente.Utente) ($($m.Sorgente.SID)) creato $($m.Creato)"
    if ($target.Connesso -and $target.SID -eq ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)) {
        Write-MigLog -Level WARN -Message "Stai importando sull'utente attualmente connesso: chiudi Explorer-intensive app, browser e Outlook. Alcune impostazioni saranno visibili dopo nuovo accesso."
    }

    # 1. Dati
    $copies = @($m.Copie)
    if ($DataSetIds) { $copies = @($copies | Where-Object { $DataSetIds -contains $_.DataSet }) }
    $grand = ($copies | Measure-Object -Property Byte -Sum).Sum; if (-not $grand) { $grand = 1 }
    $free = $null; try { $free = (Get-PSDrive -Name (Get-Item -LiteralPath $target.Percorso).PSDrive.Name).Free } catch {}
    if ($null -ne $free -and $free -lt ($grand + 500MB)) { throw ("Spazio insufficiente su disco di destinazione: servono {0}, disponibili {1}." -f (Format-MigBytes ($grand + 500MB)), (Format-MigBytes $free)) }
    $done = 0L; $results = @()
    foreach ($c in $copies) {
        Test-MigCancelled
        $src = Join-Path (Join-Path $PackagePath 'Data') $c.Relativo
        if (-not (Test-Path -LiteralPath $src)) { Write-MigLog -Level WARN -Message "Manca nel pacchetto: $($c.Relativo)"; continue }
        if ($c.DataSet -eq '_Extra') {
            $dst = Join-Path $target.Percorso ("Desktop\Cartelle migrate\" + (Split-Path -Leaf $c.Relativo))
        } else {
            $dst = Join-Path $target.Percorso $c.Relativo
            # Localizzazione: 'Documenti\File di Outlook' -> 'Documents\Outlook Files' etc. e' gestita dai nomi reali in inglese su disco.
        }
        Write-MigLog -Level STEP -Message "Ripristino $($c.DataSet): $($c.Relativo)"
        $r = Invoke-MigRobocopy -Source $src -Destination $dst -LogFile (Join-Path $PackagePath 'Logs\robocopy-import.log') -BaseBytes $done -GrandTotal $grand -Label $c.DataSet -SkipNewerAtDest:$SkipNewer
        $done += $r.BytesCopied
        $results += [pscustomobject]@{ DataSet = $c.DataSet; Relativo = $c.Relativo; File = $r.FilesCopied; Byte = $r.BytesCopied; Codice = $r.ExitCode; Ok = $r.Success }
    }

    # 2. Permessi
    if ($FixPermissions) {
        Write-MigProgress -Percent 93 -Status 'Correzione permessi NTFS...'
        Write-MigLog -Level STEP -Message 'Imposto proprietario/permessi sulle cartelle ripristinate...'
        foreach ($c in $results) {
            $p = if ($c.DataSet -eq '_Extra') { Join-Path $target.Percorso 'Desktop\Cartelle migrate' } else { Join-Path $target.Percorso (($c.Relativo -split '\\')[0]) }
            if (Test-Path -LiteralPath $p) { Repair-MigProfileAcl -Path $p -SID $target.SID }
        }
        Write-MigLog -Level OK -Message 'Permessi aggiornati.'
    }

    # 3. Registro + impostazioni
    $hive = $null; $regOk = 0; $regFail = 0
    try {
        Write-MigProgress -Percent 95 -Status 'Registro e impostazioni...'
        $hive = Mount-MigUserHive -SID $target.SID -ProfilePath $target.Percorso
        if ($RestoreRegistry -and $m.Registro) {
            foreach ($rf in @($m.Registro)) {
                $file = Join-Path (Join-Path $PackagePath 'Registry') $rf.File
                if (-not (Test-Path -LiteralPath $file)) { continue }
                $ok = Import-MigRegistryFile -RegFile $file -OldRoot $m.Sorgente.HiveRoot -NewRoot $hive.Root -OldSID $m.Sorgente.SID -NewSID $target.SID -OldProfilePath $m.Sorgente.Percorso -NewProfilePath $target.Percorso -OldUserName $m.Sorgente.Utente -NewUserName $target.Nome
                if ($ok) { $regOk++ } else { $regFail++; Write-MigLog -Level WARN -Message "Import registro fallito: $($rf.Chiave)" }
            }
            Write-MigLog -Level OK -Message "Registro: importate $regOk chiavi, $regFail fallite."
        }
        $set = Join-Path $PackagePath 'Settings'
        if ($RestoreWiFi -and (Test-Path -LiteralPath (Join-Path $set 'WiFi'))) {
            $n = 0
            foreach ($x in Get-ChildItem -LiteralPath (Join-Path $set 'WiFi') -Filter *.xml) { $o = & netsh.exe wlan add profile filename="$($x.FullName)" user=all 2>&1; if ($LASTEXITCODE -eq 0) { $n++ } }
            Write-MigLog -Level OK -Message "Wi-Fi: ripristinati $n profili."
        }
        if ($RestoreDrives -and (Test-Path -LiteralPath (Join-Path $set 'drives.json'))) {
            $drives = @(Get-Content -LiteralPath (Join-Path $set 'drives.json') -Raw | ConvertFrom-Json)
            foreach ($d in $drives) {
                if (-not $d.Lettera -or -not $d.Percorso) { continue }
                $k = "Registry::$($hive.Root)\Network\$($d.Lettera.ToLower())"
                if (-not (Test-Path -LiteralPath $k)) { New-Item -Path $k -Force | Out-Null }
                Set-ItemProperty -LiteralPath $k -Name RemotePath -Value $d.Percorso
                Set-ItemProperty -LiteralPath $k -Name UserName -Value $(if ($d.Utente) { $d.Utente } else { '' })
                Set-ItemProperty -LiteralPath $k -Name ProviderName -Value $(if ($d.Provider) { $d.Provider } else { 'Microsoft Windows Network' })
                Set-ItemProperty -LiteralPath $k -Name ProviderType -Value 0x20000 -Type DWord
                Set-ItemProperty -LiteralPath $k -Name ConnectionType -Value 1 -Type DWord
                Set-ItemProperty -LiteralPath $k -Name DeferFlags -Value 4 -Type DWord
            }
            Write-MigLog -Level OK -Message "Unita' di rete: $($drives.Count) mappature scritte (attive al prossimo accesso)."
        }
        if ($RestorePrinters -and (Test-Path -LiteralPath (Join-Path $set 'printers.json'))) {
            $pr = Get-Content -LiteralPath (Join-Path $set 'printers.json') -Raw | ConvertFrom-Json
            $n = 0
            foreach ($np in @($pr.DiRete)) {
                try { Add-Printer -ConnectionName $np -ErrorAction Stop; $n++ } catch { Write-MigLog -Level WARN -Message "Stampante di rete non aggiunta ($np): $($_.Exception.Message)" }
            }
            foreach ($lp in @($pr.Locali | Where-Object { $_.IndirizzoIP })) {
                try {
                    if (-not (Get-PrinterPort -Name $lp.Porta -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $lp.Porta -PrinterHostAddress $lp.IndirizzoIP -ErrorAction Stop }
                    if (Get-PrinterDriver -Name $lp.Driver -ErrorAction SilentlyContinue) { Add-Printer -Name $lp.Nome -DriverName $lp.Driver -PortName $lp.Porta -ErrorAction Stop; $n++ }
                    else { Write-MigLog -Level WARN -Message "Stampante IP '$($lp.Nome)': driver '$($lp.Driver)' non installato su questo PC. Porta $($lp.IndirizzoIP) creata; installare il driver e aggiungere la stampante." }
                } catch { Write-MigLog -Level WARN -Message "Stampante IP '$($lp.Nome)': $($_.Exception.Message)" }
            }
            if ($pr.Predefinita) {
                try { Set-ItemProperty -LiteralPath "Registry::$($hive.Root)\Software\Microsoft\Windows NT\CurrentVersion\Windows" -Name Device -Value "$($pr.Predefinita),winspool,Ne00:" -ErrorAction Stop } catch {}
            }
            Write-MigLog -Level OK -Message "Stampanti: $n aggiunte. Predefinita: $($pr.Predefinita)"
        }
        if ($RestoreWallpaper) {
            $wpFile = Get-ChildItem -LiteralPath $set -Filter 'wallpaper.*' -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne '.json' } | Select-Object -First 1
            if ($wpFile) {
                $destPic = Join-Path $target.Percorso ('Pictures\Sfondo migrato' + $wpFile.Extension)
                New-Item -ItemType Directory -Path (Split-Path $destPic) -Force | Out-Null
                Copy-Item -LiteralPath $wpFile.FullName -Destination $destPic -Force
                Set-ItemProperty -LiteralPath "Registry::$($hive.Root)\Control Panel\Desktop" -Name Wallpaper -Value $destPic
                Write-MigLog -Level OK -Message "Sfondo impostato: $destPic (visibile al prossimo accesso)."
            }
        }
    } catch { Write-MigLog -Level ERROR -Message "Registro/impostazioni: $($_.Exception.Message)" }
    finally { if ($hive) { Dismount-MigUserHive -Hive $hive } }

    # 4. Report
    $report = [ordered]@{
        Completato = (Get-Date).ToString('o'); Durata = ((Get-Date) - $started).ToString(); Destinazione = $target
        Copie = $results; ByteTotali = $done; RegistroOk = $regOk; RegistroFalliti = $regFail
        Errori = @($results | Where-Object { -not $_.Ok } | ForEach-Object { $_.Relativo })
        Note = @(
            'Le password salvate nei browser (Chrome/Edge) sono cifrate con DPAPI del vecchio utente e NON sono recuperabili: usare la sincronizzazione account Google/Microsoft.',
            'Le password di Outlook/posta e le credenziali di rete (Credential Manager) vanno reinserite.',
            'I programmi NON vengono installati: usare Settings\apps.csv come checklist o "winget import -i Settings\winget.json".',
            'Per vedere tutte le impostazioni, disconnettere e riconnettere l''utente di destinazione.'
        )
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PackagePath 'Logs\import-report.json') -Encoding UTF8
    Write-MigProgress -Percent 100 -Status 'Importazione completata'
    Write-MigLog -Level STEP -Message ("=== IMPORTAZIONE COMPLETATA in {0:mm\:ss}: {1}, errori: {2} ===" -f ((Get-Date) - $started), (Format-MigBytes $done), @($report.Errori).Count)
    return [pscustomobject]$report
}

# ------------------------------------------------------------------------------
#  RETE: FIREWALL, DISCOVERY, TRASFERIMENTO
# ------------------------------------------------------------------------------
function Enable-MigFirewall {
    param([int]$TcpPort = $script:DefaultTcpPort, [int]$UdpPort = $script:DefaultUdpPort)
    try {
        Get-NetFirewallRule -DisplayName "$script:FirewallPrefix*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue
        New-NetFirewallRule -DisplayName "$script:FirewallPrefix TCP $TcpPort" -Direction Inbound -Protocol TCP -LocalPort $TcpPort -Action Allow -Profile Any | Out-Null
        New-NetFirewallRule -DisplayName "$script:FirewallPrefix UDP $UdpPort" -Direction Inbound -Protocol UDP -LocalPort $UdpPort -Action Allow -Profile Any | Out-Null
        Write-MigLog -Level OK -Message "Firewall: regole temporanee create (TCP $TcpPort, UDP $UdpPort)."
    } catch { Write-MigLog -Level WARN -Message "Firewall: impossibile creare le regole ($($_.Exception.Message)). Il trasferimento potrebbe essere bloccato." }
}
function Disable-MigFirewall {
    try { Get-NetFirewallRule -DisplayName "$script:FirewallPrefix*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue; Write-MigLog 'Firewall: regole temporanee rimosse.' } catch {}
}

function Get-MigLocalIPs {
    try { return @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' -and $_.PrefixOrigin -ne 'WellKnown' } | Sort-Object InterfaceMetric | Select-Object -ExpandProperty IPAddress) }
    catch { return @([System.Net.Dns]::GetHostAddresses($env:COMPUTERNAME) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.ToString() }) }
}

function New-MigPairingCode { return ('{0:D6}' -f (Get-Random -Minimum 0 -Maximum 999999)) }

function Get-MigHmac {
    param([byte[]]$Data, [string]$Code)
    $h = New-Object System.Security.Cryptography.HMACSHA256 (, [System.Text.Encoding]::UTF8.GetBytes("VMIG:$Code"))
    try { return $h.ComputeHash($Data) } finally { $h.Dispose() }
}

function Start-MigDiscoveryResponder {
    <# Risponde ai broadcast UDP "VMIG_DISCOVER" con "VMIG_HERE|<nome>|<porta tcp>". Loop finche' $Stop.Value e' $false. #>
    param([int]$UdpPort = $script:DefaultUdpPort, [int]$TcpPort = $script:DefaultTcpPort, [Parameter(Mandatory)][ref]$Stop)
    $udp = New-Object System.Net.Sockets.UdpClient
    $udp.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::Socket, [System.Net.Sockets.SocketOptionName]::ReuseAddress, $true)
    $udp.Client.Bind((New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, $UdpPort)))
    $udp.Client.ReceiveTimeout = 500
    try {
        while (-not $Stop.Value) {
            $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
            try { $data = $udp.Receive([ref]$ep) } catch { continue }
            if ([System.Text.Encoding]::ASCII.GetString($data) -eq 'VMIG_DISCOVER') {
                $resp = [System.Text.Encoding]::ASCII.GetBytes("VMIG_HERE|$env:COMPUTERNAME|$TcpPort")
                [void]$udp.Send($resp, $resp.Length, $ep)
            }
        }
    } finally { $udp.Close() }
}

function Find-MigReceivers {
    <# Invia broadcast su tutte le interfacce e raccoglie i PC in ascolto per ~2s. #>
    param([int]$UdpPort = $script:DefaultUdpPort, [int]$TimeoutMs = 2000)
    $found = @{}
    $udp = New-Object System.Net.Sockets.UdpClient
    $udp.EnableBroadcast = $true
    $udp.Client.ReceiveTimeout = 400
    try {
        $req = [System.Text.Encoding]::ASCII.GetBytes('VMIG_DISCOVER')
        [void]$udp.Send($req, $req.Length, (New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Broadcast, $UdpPort)))
        try { foreach ($ip in (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object { $_.PrefixLength -lt 32 -and $_.IPAddress -notlike '127.*' })) {
            $bytes = ([System.Net.IPAddress]::Parse($ip.IPAddress)).GetAddressBytes(); [array]::Reverse($bytes); $n = [BitConverter]::ToUInt32($bytes, 0)
            $mask = [uint32]([math]::Pow(2, 32) - [math]::Pow(2, 32 - $ip.PrefixLength)); $b = ($n -bor (-bnot $mask -band 0xFFFFFFFF)); $bb = [BitConverter]::GetBytes([uint32]$b); [array]::Reverse($bb)
            [void]$udp.Send($req, $req.Length, (New-Object System.Net.IPEndPoint((New-Object System.Net.IPAddress(, $bb)), $UdpPort)))
        } } catch {}
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
            $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
            try { $data = $udp.Receive([ref]$ep) } catch { continue }
            $msg = [System.Text.Encoding]::ASCII.GetString($data)
            if ($msg -like 'VMIG_HERE|*') { $parts = $msg -split '\|'; $found[$ep.Address.ToString()] = [pscustomobject]@{ IP = $ep.Address.ToString(); Nome = $parts[1]; Porta = [int]$parts[2] } }
        }
    } finally { $udp.Close() }
    return @($found.Values)
}

function Send-MigPackage {
    <#
    .SYNOPSIS  Invia una cartella pacchetto al PC in ascolto (Receive-MigPackage).
    .DESCRIPTION
      Protocollo VMIG/1 (TCP):
        C->S  "VMIG1"
        S->C  nonce[16]
        C->S  HMAC-SHA256(nonce, codice)[32]      (autenticazione con codice di accoppiamento)
        S->C  "OK" | "NO"
        C->S  per ogni file: [int32 lenJSON][JSON {r=relpath,s=size,m=mtime}] [bytes] [sha256[32]]
              S->C  "A" (ack) | "E" (hash errato -> ritrasmesso una volta)
        C->S  [int32 0]  (fine)
        S->C  "DONE"
      Un file gia' presente a destinazione con stessa dimensione/mtime viene saltato (S->C "S" al posto di "A" dopo l'header):
      consente la RIPRESA di un trasferimento interrotto.
    #>
    param([Parameter(Mandatory)][string]$PackagePath, [Parameter(Mandatory)][string]$TargetHost, [int]$Port = $script:DefaultTcpPort, [Parameter(Mandatory)][string]$Code)
    # Normalizza a nome lungo: %TEMP% o percorsi 8.3 (es. RUNNER~1) non coinciderebbero con i FullName enumerati
    $PackagePath = (Get-Item -LiteralPath $PackagePath).FullName.TrimEnd('\')
    $files = @(Get-ChildItem -LiteralPath $PackagePath -Recurse -File -Force)
    if ($files.Count -gt 0 -and -not $files[0].FullName.StartsWith($PackagePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        try { $PackagePath = (New-Object -ComObject Scripting.FileSystemObject).GetFolder($PackagePath).Path.TrimEnd('\') } catch {}
    }
    $total = ($files | Measure-Object -Property Length -Sum).Sum; if (-not $total) { $total = 1 }
    Write-MigLog -Level STEP -Message ("Invio {0} file ({1}) a {2}:{3}" -f $files.Count, (Format-MigBytes $total), $TargetHost, $Port)
    $tcp = New-Object System.Net.Sockets.TcpClient
    $tcp.SendBufferSize = 1MB; $tcp.NoDelay = $true
    $tcp.Connect($TargetHost, $Port)
    $ns = $tcp.GetStream(); $ns.ReadTimeout = 60000; $ns.WriteTimeout = 60000
    $enc = [System.Text.Encoding]::UTF8
    try {
        $ns.Write($enc.GetBytes('VMIG1'), 0, 5)
        $nonce = New-Object byte[] 16; $read = 0; while ($read -lt 16) { $n = $ns.Read($nonce, $read, 16 - $read); if ($n -le 0) { throw 'Connessione chiusa durante l''handshake.' }; $read += $n }
        $mac = Get-MigHmac -Data $nonce -Code $Code
        $ns.Write($mac, 0, 32)
        $resp = New-Object byte[] 2; $ns.Read($resp, 0, 2) | Out-Null
        if ($enc.GetString($resp) -ne 'OK') { throw 'Codice di accoppiamento rifiutato dal PC di destinazione.' }
        Write-MigLog -Level OK -Message 'Autenticazione riuscita.'
        $sent = 0L; $skipped = 0; $sw = [System.Diagnostics.Stopwatch]::StartNew(); $buf = New-Object byte[] (1MB)
        $pkgLeaf = Split-Path -Leaf $PackagePath
        $i = 0
        foreach ($f in $files) {
            Test-MigCancelled
            $i++
            $rel = $pkgLeaf + '\' + $f.FullName.Substring($PackagePath.Length + 1)
            $hdr = $enc.GetBytes((@{ r = $rel; s = $f.Length; m = $f.LastWriteTimeUtc.Ticks } | ConvertTo-Json -Compress))
            $ns.Write([BitConverter]::GetBytes([int32]$hdr.Length), 0, 4); $ns.Write($hdr, 0, $hdr.Length)
            $ack = New-Object byte[] 1; if ($ns.Read($ack, 0, 1) -le 0) { throw 'Connessione persa.' }
            if ($enc.GetString($ack) -eq 'S') { $skipped++; $sent += $f.Length; continue }   # gia' presente -> ripresa
            $attempt = 0
            do {
                $attempt++
                $sha = [System.Security.Cryptography.SHA256]::Create()
                $fs = [System.IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
                try {
                    $fileSent = 0L
                    while (($n = $fs.Read($buf, 0, $buf.Length)) -gt 0) {
                        $ns.Write($buf, 0, $n); $sha.TransformBlock($buf, 0, $n, $null, 0) | Out-Null
                        $fileSent += $n
                        if ($sw.ElapsedMilliseconds % 300 -lt 20) {
                            $spd = (($sent + $fileSent) / 1MB) / [math]::Max(0.001, $sw.Elapsed.TotalSeconds)
                            Write-MigProgress -Percent ((($sent + $fileSent) / $total) * 100) -Status ("Invio {0}/{1}: {2}" -f $i, $files.Count, $f.Name) -BytesDone ($sent + $fileSent) -BytesTotal $total -SpeedMBs $spd
                        }
                    }
                    $sha.TransformFinalBlock($buf, 0, 0) | Out-Null
                    $ns.Write($sha.Hash, 0, 32)
                } finally { $fs.Close(); $sha.Dispose() }
                $ns.Read($ack, 0, 1) | Out-Null
                $okAck = ($enc.GetString($ack) -eq 'A')
                if (-not $okAck) { Write-MigLog -Level WARN -Message "Hash errato per $rel, ritrasmetto (tentativo $attempt)..." }
            } while (-not $okAck -and $attempt -lt 2)
            if (-not $okAck) { throw "Integrita' non verificata per $rel dopo 2 tentativi." }
            $sent += $f.Length
        }
        $ns.Write([BitConverter]::GetBytes([int32]0), 0, 4)
        $fin = New-Object byte[] 4; $ns.Read($fin, 0, 4) | Out-Null
        if ($enc.GetString($fin) -ne 'DONE') { throw 'Il ricevitore non ha confermato il completamento.' }
        Write-MigProgress -Percent 100 -Status 'Trasferimento completato'
        Write-MigLog -Level STEP -Message ("Trasferimento completato: {0} in {1:mm\:ss} ({2} file saltati perche' gia' presenti)." -f (Format-MigBytes $total), $sw.Elapsed, $skipped)
    } finally { $ns.Close(); $tcp.Close() }
}

function Receive-MigPackage {
    <#
    .SYNOPSIS  Si mette in ascolto, autentica il mittente con il codice e salva il pacchetto in <Destination>\<nomePacchetto>.
               Ritorna il percorso del pacchetto ricevuto. Usa $Stop ([ref] bool) per annullare l'attesa.
    #>
    param([Parameter(Mandatory)][string]$Destination, [int]$Port = $script:DefaultTcpPort, [Parameter(Mandatory)][string]$Code, [ref]$Stop)
    if (-not (Test-Path -LiteralPath $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
    $enc = [System.Text.Encoding]::UTF8
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Any, $Port)
    $listener.Start()
    Write-MigLog -Level STEP -Message ("In ascolto su TCP {0} - IP: {1} - codice: {2}" -f $Port, ((Get-MigLocalIPs) -join ', '), $Code)
    $client = $null; $pkgRoot = $null
    try {
        while (-not $listener.Pending()) { if ($Stop -and $Stop.Value) { throw [System.OperationCanceledException]::new('Attesa annullata.') }; Start-Sleep -Milliseconds 250 }
        $client = $listener.AcceptTcpClient(); $client.ReceiveBufferSize = 1MB
        $remote = $client.Client.RemoteEndPoint.ToString()
        $ns = $client.GetStream(); $ns.ReadTimeout = 120000; $ns.WriteTimeout = 60000
        $magic = New-Object byte[] 5; $ns.Read($magic, 0, 5) | Out-Null
        if ($enc.GetString($magic) -ne 'VMIG1') { throw "Client $remote non valido (magic errato)." }
        $nonce = New-Object byte[] 16; [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($nonce)
        $ns.Write($nonce, 0, 16)
        $mac = New-Object byte[] 32; $r = 0; while ($r -lt 32) { $n = $ns.Read($mac, $r, 32 - $r); if ($n -le 0) { throw 'Handshake interrotto.' }; $r += $n }
        $expected = Get-MigHmac -Data $nonce -Code $Code
        $same = $true; for ($k = 0; $k -lt 32; $k++) { if ($mac[$k] -ne $expected[$k]) { $same = $false } }
        if (-not $same) { $ns.Write($enc.GetBytes('NO'), 0, 2); throw "Codice errato dal client $remote. Connessione rifiutata." }
        $ns.Write($enc.GetBytes('OK'), 0, 2)
        Write-MigLog -Level OK -Message "Client $remote autenticato. Ricezione in corso..."
        $buf = New-Object byte[] (1MB); $recv = 0L; $count = 0; $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $lenB = New-Object byte[] 4
        while ($true) {
            $r = 0; while ($r -lt 4) { $n = $ns.Read($lenB, $r, 4 - $r); if ($n -le 0) { throw 'Connessione persa durante la ricezione.' }; $r += $n }
            $len = [BitConverter]::ToInt32($lenB, 0)
            if ($len -eq 0) { break }
            $hb = New-Object byte[] $len; $r = 0; while ($r -lt $len) { $n = $ns.Read($hb, $r, $len - $r); if ($n -le 0) { throw 'Connessione persa.' }; $r += $n }
            $h = $enc.GetString($hb) | ConvertFrom-Json
            if ($h.r -match '\.\.|^[\\/]|:') { throw "Percorso non sicuro rifiutato: $($h.r)" }
            $dest = Join-Path $Destination $h.r
            if (-not $pkgRoot) { $pkgRoot = Join-Path $Destination (($h.r -split '\\')[0]) }
            $dir = Split-Path -Parent $dest; if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            if ((Test-Path -LiteralPath $dest) -and ((Get-Item -LiteralPath $dest -Force).Length -eq [long]$h.s) -and ((Get-Item -LiteralPath $dest -Force).LastWriteTimeUtc.Ticks -eq [long]$h.m)) {
                $ns.Write($enc.GetBytes('S'), 0, 1); $recv += [long]$h.s; $count++; continue
            }
            $ns.Write($enc.GetBytes('G'), 0, 1)
            $attempt = 0
            do {
                $attempt++
                $sha = [System.Security.Cryptography.SHA256]::Create()
                $fs = [System.IO.File]::Create($dest + '.part')
                $remaining = [long]$h.s
                try {
                    while ($remaining -gt 0) {
                        $n = $ns.Read($buf, 0, [int][math]::Min($buf.Length, $remaining)); if ($n -le 0) { throw 'Connessione persa durante il file.' }
                        $fs.Write($buf, 0, $n); $sha.TransformBlock($buf, 0, $n, $null, 0) | Out-Null; $remaining -= $n
                        if ($sw.ElapsedMilliseconds % 300 -lt 20) {
                            $spd = (($recv + ($h.s - $remaining)) / 1MB) / [math]::Max(0.001, $sw.Elapsed.TotalSeconds)
                            Write-MigProgress -Percent 0 -Status ("Ricevuti {0} file, {1}" -f $count, (Format-MigBytes ($recv + ($h.s - $remaining)))) -BytesDone ($recv + ($h.s - $remaining)) -BytesTotal -1 -SpeedMBs $spd
                        }
                    }
                    $sha.TransformFinalBlock($buf, 0, 0) | Out-Null
                } finally { $fs.Close() }
                $theirs = New-Object byte[] 32; $r = 0; while ($r -lt 32) { $n = $ns.Read($theirs, $r, 32 - $r); if ($n -le 0) { throw 'Connessione persa (hash).' }; $r += $n }
                $ok = $true; for ($k = 0; $k -lt 32; $k++) { if ($theirs[$k] -ne $sha.Hash[$k]) { $ok = $false } }
                $sha.Dispose()
                if ($ok) {
                    Move-Item -LiteralPath ($dest + '.part') -Destination $dest -Force
                    [System.IO.File]::SetLastWriteTimeUtc($dest, (New-Object DateTime([long]$h.m, 'Utc')))
                    $ns.Write($enc.GetBytes('A'), 0, 1)
                } else { $ns.Write($enc.GetBytes('E'), 0, 1); Write-MigLog -Level WARN -Message "Hash errato: $($h.r) (tentativo $attempt)" }
            } while (-not $ok -and $attempt -lt 2)
            if (-not $ok) { throw "Integrita' non verificata per $($h.r)." }
            $recv += [long]$h.s; $count++
        }
        $ns.Write($enc.GetBytes('DONE'), 0, 4)
        Write-MigProgress -Percent 100 -Status 'Ricezione completata'
        Write-MigLog -Level STEP -Message ("Ricevuti {0} file ({1}) in {2:mm\:ss}. Pacchetto: {3}" -f $count, (Format-MigBytes $recv), $sw.Elapsed, $pkgRoot)
        return $pkgRoot
    } finally {
        if ($client) { $client.Close() }
        $listener.Stop()
    }
}

Export-ModuleMember -Function *-Mig*, Format-MigBytes, Write-MigLog, Write-MigProgress, Initialize-MigLogging, Test-MigAdmin
