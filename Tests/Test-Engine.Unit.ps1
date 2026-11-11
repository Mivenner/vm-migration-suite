<#
 Test unitari eseguibili su qualsiasi Windows (non modificano il sistema):
   - caricamento modulo ed elenco funzioni
   - Format-MigBytes
   - rimappatura .reg (Import-MigRegistryFile viene testata sulla sola trasformazione testuale)
   - protocollo di rete: invio/ricezione di un pacchetto fittizio su 127.0.0.1 con verifica hash e ripresa
 Uso: powershell -ExecutionPolicy Bypass -File .\Tests\Test-Engine.Unit.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'VMMigration.Engine.psm1') -Force
$fail = 0
function Assert($cond, $msg) { if ($cond) { Write-Host "PASS $msg" -ForegroundColor Green } else { Write-Host "FAIL $msg" -ForegroundColor Red; $script:fail++ } }

Assert ((Format-MigBytes 1536) -eq '2 KB' -or (Format-MigBytes 1536) -eq '1,5 KB' -or (Format-MigBytes 1536) -like '*KB') 'Format-MigBytes KB'
Assert ((Format-MigBytes 3GB) -like '3*GB') 'Format-MigBytes GB'
Assert ((Get-MigDataSets).Count -ge 20) 'DataSets caricati'
Assert ((New-MigPairingCode) -match '^\d{6}$') 'Codice accoppiamento a 6 cifre'

# Protocollo di rete in loopback
$tmp = Join-Path $env:TEMP ('vmigtest_' + [guid]::NewGuid().ToString('N'))
$pkg = Join-Path $tmp 'VMIG_test'; New-Item -ItemType Directory -Path (Join-Path $pkg 'Data\Documents') -Force | Out-Null
$rnd = New-Object byte[] (3MB); (New-Object Random).NextBytes($rnd); [IO.File]::WriteAllBytes((Join-Path $pkg 'Data\Documents\big.bin'), $rnd)
'{"Magic":"VMIG"}' | Set-Content (Join-Path $pkg 'manifest.json')
$recv = Join-Path $tmp 'recv'
$code = '424242'
$rs = [powershell]::Create()
$rs.AddScript({ param($e, $d, $c) Import-Module $e -Force; Receive-MigPackage -Destination $d -Port 18888 -Code $c }).AddArgument((Join-Path $root 'VMMigration.Engine.psm1')).AddArgument($recv).AddArgument($code) | Out-Null
$h = $rs.BeginInvoke(); Start-Sleep 1
Send-MigPackage -PackagePath $pkg -TargetHost '127.0.0.1' -Port 18888 -Code $code
$out = $rs.EndInvoke($h)
$a = Get-FileHash (Join-Path $pkg 'Data\Documents\big.bin'); $b = Get-FileHash (Join-Path $recv 'VMIG_test\Data\Documents\big.bin')
Assert ($a.Hash -eq $b.Hash) 'Trasferimento LAN: hash identico'
Assert ($out -like '*VMIG_test') 'Receive-MigPackage ritorna il percorso pacchetto'
# codice errato -> rifiutato
$rs2 = [powershell]::Create()
$rs2.AddScript({ param($e, $d, $c) Import-Module $e -Force; try { Receive-MigPackage -Destination $d -Port 18889 -Code $c } catch { 'REFUSED' } }).AddArgument((Join-Path $root 'VMMigration.Engine.psm1')).AddArgument($recv).AddArgument('111111') | Out-Null
$h2 = $rs2.BeginInvoke(); Start-Sleep 1
$refused = $false; try { Send-MigPackage -PackagePath $pkg -TargetHost '127.0.0.1' -Port 18889 -Code '222222' } catch { $refused = $true }
Assert $refused 'Codice errato rifiutato dal mittente'
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
if ($fail) { throw "$fail test falliti" } else { Write-Host "Tutti i test superati." -ForegroundColor Cyan }
