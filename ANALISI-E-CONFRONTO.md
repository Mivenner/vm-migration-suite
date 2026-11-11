# VM Migration Suite 3.0 — Analisi del pacchetto precedente e confronto con il mercato

Data: 8 novembre 2026 · Autore: vm technology (con assistenza Claude)

## 1. Perché la v10 "non funzionava come dovrebbe"

Ho smontato `Windows_Migration_Suite_Installer_v10.zip` (6 file: `MigrationEngine.ps1`, `MigrationGUI.ps1`, `GUI.xaml`, `Build-Installer.ps1`, `installer.iss`, `app.ico`) e l'export della chat Gemini (`Chat_Export.md`, che descrive esattamente lo stesso codice). Questi i difetti, in ordine di gravità:

| # | Difetto | Effetto per l'utente | Dove |
|---|---------|----------------------|------|
| 1 | **Manca del tutto la fase di importazione.** Esiste solo l'export; la modalità "Server" riceve uno zip e lo lascia lì. `New-MigrationTargetUser` non viene mai chiamata e usa `Read-Host`, che in un exe `-noConsole` blocca o fallisce. | Il tool non migra: copia tre cartelle e basta. | `MigrationGUI.ps1` r.170-222, `MigrationEngine.ps1` r.78-87 |
| 2 | **Firewall mai configurato.** Nessuna regola per TCP 8888 / UDP 8889: Windows Firewall blocca `TcpListener`/`UdpClient` in ingresso. | Il "Scan LAN" non trova nulla e il trasferimento non parte: è la causa più probabile del "non funziona" in rete. | `Start-ServerDiscoveryService`, `Receive-TCPFileWithProgress` |
| 3 | **Tutto gira sul thread della GUI** con `DoEvents`; `AcceptTcpClient()` e `robocopy -Wait` bloccano la finestra ("Non risponde"), nessun Annulla. | Finestra congelata per minuti/ore, impossibile annullare. | `MigrationGUI.ps1` r.176-221 |
| 4 | **Progresso finto** (10 → 30 → 70 → 90 cablati), nessuna lettura dell'output di robocopy, codice di uscita robocopy ignorato (≥ 8 = errori). | Nessun feedback reale, errori di copia invisibili. | `Invoke-DataMigration` |
| 5 | **Export incompleto**: solo Desktop, Documenti, Download. Niente Immagini/Video/Musica/Preferiti, niente profili browser (segnalibri!), firme Outlook, PST, Sticky Notes, font, pin barra applicazioni. | Il nuovo PC è "vuoto" per l'utente. | `FoldersToCopy @("Desktop","Documents","Downloads")` |
| 6 | **Registro quasi inutile**: `reg load` fallisce se l'utente è connesso (hive bloccato), esporta solo 3 chiavi (`Software\Microsoft\Edge`, `Mozilla`, `Google\Chrome`) che non contengono i dati dei browser (stanno in AppData), nessuna rimappatura SID/percorso in fase di import, hive non smontato in caso di errore. | Impostazioni perse; hive "zombie" fino al riavvio. | `Invoke-OfflineRegistryExport` |
| 7 | **Protocollo TCP senza protezione**: nessun nome/dimensione/hash, nessuna autenticazione; chiunque in LAN può connettersi e scrivere il file; nessuna ripresa; zip intero creato su disco (spazio doppio, limite .NET Zip su alcune versioni). | Trasferimenti corrotti non rilevabili, rischio sicurezza. | `Send/Receive-TCPFileWithProgress` |
| 8 | **Profili rilevati dalle cartelle** (`C:\Users\*`) e non da `ProfileList`: nessun SID, cartelle orfane incluse, nessuna distinzione locale/dominio/Entra ID. | Impossibile rimappare correttamente al nuovo utente. | `Load-UserProfiles` |
| 9 | `trap { ... Continue }` globale che ingoia ogni errore; `Get-WmiObject` (deprecato, assente in PS7); `{pf}` deprecato in Inno Setup; versione installer 2.0 vs "v10"; UI metà inglese metà italiana; nessun log operativo, nessun controllo spazio disco, nessun manifest del pacchetto. | Debug impossibile, incoerenze. | vari |

## 2. Cosa esiste già (ricerca di mercato, novembre 2026)

| Strumento | Costo | Punti di forza | Limiti rilevanti per noi |
|-----------|-------|----------------|--------------------------|
| **USMT (ScanState/LoadState)** — Microsoft, nel Windows ADK | Gratis | Standard enterprise, migra anche impostazioni app via `MigApp.xml`, supporta hard-link per upgrade in place, store cifrato. | Solo riga di comando con sintassi rigida; richiede ADK installato; la versione per Windows 11 ha perso sfondo personalizzato e stampanti di rete, problemi con OneDrive; `LoadState` su Entra ID sconsigliato da Microsoft. |
| **Windows Backup / PC-to-PC Migration** — Microsoft | Gratis | Era la soluzione nativa locale (accoppiamento con codice, Wi-Fi/Ethernet) durante l'OOBE. | **Rimossa** con le build 26100/26200/26300.9550 (KB5124010); oggi la via guidata passa solo da OneDrive (5 GB gratis) con account Microsoft. Non adatta a PC aziendali/dominio. |
| **ForensiT Transwiz** | Freeware (anche in azienda) | Wizard semplice, archivio profilo portabile 32↔64 bit, molto affidabile, 1 exe. | Nessun trasferimento di rete integrato, non supporta account Entra ID, nessun inventario app/stampanti/Wi-Fi, chiuso. |
| **ForensiT User Profile Wizard** | Freeware / Corporate | Riassegna un profilo esistente a un nuovo account (dominio/Entra/locale) **sullo stesso PC** senza copiare. | Non sposta dati fra PC: complementare, non alternativo. |
| **Laplink PCmover** | A pagamento (per licenza) | Unico raccomandato da Microsoft, migra anche i programmi, cavo/LAN. | Costo per PC, successo parziale su app complesse, nessuna automazione da script. |
| **EaseUS Todo PCTrans** | Freemium | GUI amichevole, migrazione app locali, immagine file. | La versione gratuita è di fatto una demo (limiti di GB/app), licenze annuali. |
| **Zinstall WinWin / Migration Kit Pro** | A pagamento | Porta "tutto" incluso software, anche da disco non avviabile. | Costoso, chiuso, poco controllabile. |
| **USMTGUI / Super Grate** | Pagamento / open source | Front-end per USMT (Super Grate fa anche migrazioni remote). | Ereditano i limiti di USMT. |

**Conclusione di posizionamento.** Dopo la rimozione di PC-to-PC Migration non esiste più uno strumento Microsoft gratuito, grafico e locale: lo spazio per un tool interno c'è. La scelta giusta non è reinventare USMT ma coprire il caso d'uso del tecnico in azienda (Caffè Aiello e clienti): wizard chiaro in italiano, pacchetto trasparente (cartelle + JSON, ispezionabile e copiabile anche a mano), trasferimento LAN protetto, funzionamento anche da WinPE/Hiren's su disco non avviato, e automazione da riga di comando. Per i casi in cui serve migrare le *impostazioni interne di applicazioni* oltre le nostre coperture, il README indica come affiancare USMT o Transwiz.

## 3. Cosa fa la v3.0 (riassunto delle scelte progettuali)

1. **Motore separato (`VMMigration.Engine.psm1`)** senza `Write-Host`: usato da GUI, CLI e runspace, testabile.
2. **Export completo**: 24 dataset selezionabili (cartelle personali, Chrome/Edge/Firefox/Brave/Opera, Outlook PST/firme/modelli, Thunderbird, Sticky Notes, font utente, temi, pin Start/barra, Windows Terminal, AppData interi in modalità avanzata) + cartelle extra fuori dal profilo + esclusioni intelligenti (cache, Temp, OST, lock file).
3. **Impostazioni**: 6 set di chiavi di registro (desktop, Esplora file, unità di rete, stampanti, Office, app), Wi-Fi con chiavi, stampanti locali/IP/di rete + predefinita, unità mappate, sfondo, inventario app (CSV/JSON) + `winget export`, product key (UEFI o registro), info sistema.
4. **Import vero**: utente esistente, di dominio/Entra (profilo creato con `CreateProfile()` API se mai loggato) o nuovo utente locale creato dalla GUI; robocopy con `/XO` opzionale; **rimappatura SID + percorso profilo** nei `.reg`; `icacls` per proprietario/permessi; ripristino Wi-Fi, unità (chiave `Network`), stampanti (`Add-Printer`, porte IP), sfondo; report JSON finale con note su cosa non è migrabile (password DPAPI, Credential Manager, programmi).
5. **Rete**: protocollo `VMIG/1` con nonce + HMAC-SHA256 del codice a 6 cifre, SHA-256 per ogni file con ritrasmissione, ripresa (file già presenti saltati), discovery UDP su tutte le sottoreti, **regole firewall create e rimosse automaticamente**, nessuno zip intermedio.
6. **GUI WPF** in italiano, 4 passi, lavoro in runspace separato (finestra sempre reattiva), log live, velocità MB/s, byte copiati/totali, **Annulla**, stima dimensioni e spazio libero prima di partire, "Importa ora" dopo la ricezione.
7. **WinPE / disco offline**: lettura di `ProfileList`, `SOFTWARE` hive e `NTUSER.DAT` da `E:\` (o altra radice) senza avviare Windows; CLI per ambienti senza WPF.
8. **Qualità**: controllo sintassi automatico in build, test unitari del protocollo di rete, log in ogni pacchetto (`Logs\export.log`, `robocopy.log`, `import-report.json`), manifest con SID/percorsi/hash di provenienza.

## 4. Limiti noti (documentati, non nascosti)

- Password dei browser e credenziali salvate: cifrate con DPAPI dell'utente di origine → non migrabili da nessun tool che non conosca la password originale. Soluzione: sincronizzazione account Google/Microsoft o esportazione CSV dal browser prima della migrazione.
- Programmi: non vengono installati; il pacchetto contiene `apps.csv` e `winget.json` (`winget import -i winget.json`).
- Account Outlook: i profili MAPI vengono ripristinati dal registro ma la password va reinserita; OST esclusi di proposito (si rigenerano).
- Licenze applicative legate all'hardware (Adobe, AutoCAD, Office a licenza perpetua) vanno disattivate sul vecchio PC.
- Export dell'utente attualmente connesso: possibile ma meglio da un secondo account amministratore o da WinPE per evitare file aperti.

## Fonti

- Microsoft Learn — USMT overview / requirements / common issues: <https://learn.microsoft.com/windows/deployment/usmt/usmt-overview>, <https://technet.microsoft.com/itpro/windows/deploy/usmt-requirements>
- Microsoft Support — "Transfer your files and settings to a new Windows PC" (feature no longer available): <https://support.microsoft.com/en-US/Windows/Experience/Backup-Recovery/transfer-your-files-and-settings-to-a-new-windows-pc>
- Windows Latest, 2 nov 2026 — rimozione di PC-to-PC Migration: <https://windowslatest.com/2026/11/02/microsoft-doesnt-care-if-you-hate-onedrive-kills-windows-11s-best-local-way-to-move-to-a-new-pc-and-only-gives-you-5gb-free>
- ElevenForum — build 26100/26200/26300.9550: <https://www.elevenforum.com/t/transfer-files-and-settings-to-new-windows-11-pc-using-windows-backup-app.38145/post-622025>
- Microsoft Tech Community — USMT Windows 11 perde sfondo/stampanti: <https://techcommunity.microsoft.com/discussions/windows-servicing/usmt-windows-11/3180110>
- ForensiT — Transwiz / User Profile Wizard R24: <https://www.forensit.com/Downloads.html>, <https://forum.forensit.com/printer_friendly_posts.asp?TID=2160>, <https://forum.forensit.com/printer_friendly_posts.asp?TID=1996>
- cBackup — EaseUS Todo PCTrans vs PCmover (2026): <https://www.cbackup.com/best/eu-todo-pctrans-vs-pcmr-2578.html>
- AlternativeTo — USMT / Transwiz / PCmover alternatives (USMTGUI, Super Grate, Zinstall): <https://alternativeto.net/software/user-state-migration-tool/>
