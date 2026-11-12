# VM Migration Suite 3.0

Strumento interno **vm technology** per migrare profili utente Windows (10/11) fra PC: dati, browser, Office/Outlook, impostazioni di registro, Wi-Fi, stampanti, unità di rete, sfondo. Funziona su disco/USB/cartella di rete, via LAN con codice di accoppiamento, e da WinPE/Hiren's BootCD su dischi non avviati.

```
VM-Migration-Suite\
├── VMMigration.Engine.psm1   motore (export/import/rete) - nessuna UI
├── VMMigration.GUI.ps1       interfaccia WPF (wizard 4 passi)
├── VMMigration.CLI.ps1       riga di comando (script, WinPE, RMM)
├── GUI.xaml                  layout interfaccia
├── app.ico / logo.png        branding vm technology
├── Build.ps1                 compila VMMigrationSuite.exe (PS2EXE) + installer (Inno Setup 6)
├── installer.iss             script Inno Setup
├── Tests\Test-Syntax.ps1     controllo sintassi (eseguito anche dal build)
├── Tests\Test-Engine.Unit.ps1 test protocollo di rete in loopback
└── ANALISI-E-CONFRONTO.md    difetti della versione precedente e confronto con USMT/Transwiz/PCmover/...
```

## Requisiti
- Windows 10/11 (x64), PowerShell 5.1 (incluso) o 7.x. Privilegi di **amministratore** (l'app si rilancia elevata da sola).
- Per compilare: `Install-Module ps2exe`, facoltativo Inno Setup 6 (`winget install JRSoftware.InnoSetup`).
- Nessuna dipendenza esterna a runtime (robocopy, reg, netsh, icacls sono di Windows).

## Tre modi per distribuirla
1. **Portable dai sorgenti (nessuna compilazione)**: copia la cartella e lancia `Avvia-GUI.cmd`. Nessun exe, nessun avviso SmartScreen.
2. **Exe portable**: `.\Build.ps1 -SkipInstaller` → `dist\VMMigrationSuite_Portable_<ver>.zip` (richiede solo il modulo ps2exe, installato in automatico).
3. **Link fissi all'ultima build** (release `latest`, aggiornata a ogni push su `main`):
   - Portable: <https://github.com/Mivenner/vm-migration-suite/releases/latest/download/VMMigrationSuite_Portable.zip>
   - Installer: <https://github.com/Mivenner/vm-migration-suite/releases/latest/download/VMMigrationSuite_Setup.exe>
   - Senza browser: `Aggiorna-VMMigration.cmd` (usa GitHub CLI: `winget install GitHub.cli`, poi `gh auth login` una volta sola).
4. **Compilazione automatica su GitHub**: con `.github\workflows\build.yml` ogni push su `main` produce exe, zip portable e `Setup.exe` (scaricabili da *Actions → Artifacts*); un tag `v3.0.2` crea una *Release* con i file allegati. Non serve avere PS2EXE né Inno Setup sul proprio PC.

## Avvio rapido
**Senza compilare**: doppio clic su `Avvia-GUI.cmd` (imposta da solo `-ExecutionPolicy Bypass`), oppure
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\VMMigration.GUI.ps1
```
Il motore viene caricato in memoria (`New-Module`), quindi l'exe funziona anche con ExecutionPolicy `Restricted` o imposta da GPO.
**Compilato**: `.\Build.ps1` → `dist\VMMigrationSuite.exe` (+ `dist\VMMigrationSuite_Setup_3.0.0.exe` se Inno Setup è installato). L'exe legge `VMMigration.Engine.psm1`, `GUI.xaml` e `logo.png` dalla propria cartella: tenerli insieme.

## Flusso tipico (due PC in LAN)
1. **Nuovo PC** → *Ricevi un pacchetto dalla rete* → annota il **codice a 6 cifre** e l'IP mostrati → *Mettiti in ascolto* (le regole firewall vengono create e poi rimosse automaticamente).
2. **Vecchio PC** → *Esporta un profilo utente* → scegli profilo, dati, impostazioni, destinazione locale (serve spazio pari ai dati) → spunta *Invia subito il pacchetto* e inserisci IP + codice → *Avvia esportazione*.
3. **Nuovo PC** → al termine della ricezione premi *Importa ora* → scegli utente di destinazione (esistente, di dominio/Entra o nuovo locale) → *Avvia importazione* → disconnetti/riconnetti l'utente.

**Con USB/disco esterno**: esporta su `E:\Migrazioni`, porta il disco sul nuovo PC, *Importa un pacchetto*.

**Da WinPE / disco secondario**: *Esporta* → spunta *Disco Windows non avviato* → radice `D:\` (dove si trovano `\Windows` e `\Users`) → *Aggiorna*. In WinPE senza WPF usare la CLI: `VMMigration.CLI.ps1 -Action Export -User nome -OfflineRoot D:\ -Destination E:\Migrazioni`.

## Riga di comando
```powershell
.\VMMigration.CLI.ps1 -Action List [-OfflineRoot E:\]
.\VMMigration.CLI.ps1 -Action Estimate -User mario.rossi -DataSets All
.\VMMigration.CLI.ps1 -Action Export -User mario.rossi -Destination D:\Migrazioni [-DataSets Desktop,Documents,Chrome] [-ExtraFolders D:\Lavoro] [-TargetHost 192.168.1.50 -Code 123456]
.\VMMigration.CLI.ps1 -Action Import -Package D:\Migrazioni\VMIG_mario.rossi_20261108-1500 -TargetUser mario.rossi [-Overwrite]
.\VMMigration.CLI.ps1 -Action Import -Package ... -TargetUser mario -CreateUser -Password (Read-Host -AsSecureString) -Administrator
.\VMMigration.CLI.ps1 -Action Receive -Destination C:\Migrazioni [-Code 123456] [-Port 8888]
.\VMMigration.CLI.ps1 -Action Send -Package ... -TargetHost PC-NUOVO -Code 123456
.\VMMigration.CLI.ps1 -Action Discover
```
Codice di uscita 0 = ok, 1 = errore (per RMM/script).

## Struttura del pacchetto `VMIG_<utente>_<data>\`
```
manifest.json      origine (PC, utente, SID, percorso, hive), dataset, byte/file, errori
Data\              copia 1:1 delle cartelle del profilo (+ _Extra\ per cartelle fuori profilo)
Registry\*.reg     chiavi HKCU esportate (rimappate in import su nuovo SID/percorso)
Settings\          WiFi\*.xml (con chiave), printers.json, drives.json, wallpaper.*, apps.csv/json, winget.json, system.json
Logs\              export.log, robocopy.log, import.log, import-report.json
```
Il pacchetto è una cartella normale: si può ispezionare, copiare con qualunque mezzo o ripristinare a mano. **Contiene password Wi-Fi in chiaro**: trattarlo come riservato e cancellarlo a migrazione conclusa.

## Stampanti e Outlook: come funzionano in v3.1
- **Stampanti locali/IP**: il pacchetto driver viene esportato dal PC di origine (`Settings\PrinterDrivers\`, via `pnputil /export-driver`) e reinstallato sul nuovo PC; porta IP e stampante vengono ricreate. USB/WSD: driver installato, la stampante si aggiunge al collegamento/rilevamento.
- **Stampanti condivise da server, predefinita, unità di rete**: sono *per utente*, quindi vengono collegate da `VMIG-PrimoAccesso.cmd` (RunOnce) **al primo accesso dell'utente di destinazione**, non durante l'importazione.
- **Outlook**: per default il profilo MAPI **non** viene importato (si corrompe con OST mancante e credenziali DPAPI del vecchio utente): al primo avvio Outlook ricrea il profilo con autodiscover; gli account da configurare sono in `Settings\outlook-accounts.json`. Firme, PST, completamento automatico (RoamCache), modelli e impostazioni vengono migrati. Spuntare "Importa anche il profilo Outlook" solo tra PC con la stessa versione di Office.

## Cosa NON viene migrato (per design di Windows)
- Password salvate in browser/Outlook/Credential Manager (DPAPI legata al vecchio account) → usare la sincronizzazione account o esportare dal browser prima.
- Programmi: inventario in `Settings\apps.csv`; reinstallazione rapida con `winget import -i Settings\winget.json --accept-package-agreements`.
- File OST di Outlook (si rigenerano), cache browser, Temp.
Per migrazioni che richiedono le impostazioni interne di applicazioni non coperte si può affiancare **USMT** (`scanstate /i:MigApp.xml`) o **Transwiz**; vedi `ANALISI-E-CONFRONTO.md`.

## Risoluzione problemi
| Sintomo | Causa / rimedio |
|---------|-----------------|
| "L'esecuzione di script è disabilitata nel sistema" | Versioni < 3.0.2: aggiornare. Per la CLI usare `Avvia-CLI.cmd` o `powershell -ExecutionPolicy Bypass -File`. |
| SmartScreen "PC protetto da Windows" / Defender segnala l'exe | Eseguibile non firmato (PS2EXE e' spesso un falso positivo). Click su *Ulteriori informazioni → Esegui comunque*, oppure firmare l'exe con un certificato code-signing (`signtool sign /fd SHA256 /a dist\VMMigrationSuite.exe`) e aggiungere un'esclusione Defender per la cartella di installazione. In alternativa usare `Avvia-GUI.cmd` dai sorgenti, che non passa da PS2EXE. |
| "Impossibile caricare l'hive" | L'utente sorgente ha ancora processi attivi: disconnetterlo (o esportare da un altro account admin / WinPE). |
| Nessun PC trovato con *Cerca in rete* | Reti diverse / VLAN / Wi-Fi guest con isolamento client: inserire l'IP a mano. Il ricevitore deve già essere *in ascolto*. |
| Codice rifiutato | Codice scaduto: su *Ricevi* premere *Genera nuovo codice* e riavviare l'ascolto. |
| robocopy codice ≥ 8 | File bloccati o percorsi > 256 caratteri: vedere `Logs\robocopy.log`; rilanciare l'export (copia solo il mancante). |
| Impostazioni non visibili dopo import | Serve disconnessione/riconnessione dell'utente di destinazione (Explorer rilegge l'hive). |
| Stampanti IP non aggiunte | Driver assente sul nuovo PC: la porta viene creata; installare il driver e aggiungere la stampante. |

## Test
```powershell
.\Tests\Test-Syntax.ps1          # parser di tutti gli script + XAML
.\Tests\Test-Engine.Unit.ps1     # trasferimento LAN in loopback: hash, ripresa, codice errato
```
