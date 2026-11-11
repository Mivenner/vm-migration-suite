# Verifica sintassi (parser) di tutti gli script e il XAML. Esce con codice 1 in caso di errori.
$root = Split-Path -Parent $PSScriptRoot
$errors = 0
foreach ($f in Get-ChildItem -Path $root -Include *.ps1, *.psm1 -Recurse) {
    $t = $null; $e = $null
    [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$t, [ref]$e) | Out-Null
    if ($e) { $errors += $e.Count; $e | ForEach-Object { Write-Host ("ERRORE {0}:{1} {2}" -f $f.Name, $_.Extent.StartLineNumber, $_.Message) -ForegroundColor Red } }
    else { Write-Host "OK  $($f.Name)" -ForegroundColor Green }
}
try { [xml](Get-Content (Join-Path $root 'GUI.xaml') -Raw) | Out-Null; Write-Host 'OK  GUI.xaml' -ForegroundColor Green } catch { $errors++; Write-Host "ERRORE GUI.xaml: $($_.Exception.Message)" -ForegroundColor Red }
if ($errors) { throw "$errors errori di sintassi." }
