# FIUTO 🔍

**FIUTO** (**F**orensic **I**nvestigation **U**tility **T**ool for **O**ffline) è un toolkit DFIR (Digital Forensics and Incident Response) unificato per l'analisi completa di dischi Windows offline. Automatizza l'estrazione e l'analisi di 38+ artefatti critici in ottica forense digitale, generando report HTML dettagliati per investigazioni rapide ed efficaci.

---

## 📋 Caratteristiche Principali

### Copertura Completa di Artefatti Windows
FIUTO raccoglie e analizza:
- **Cronologie di esecuzione** (Prefetch, AmCache, ShimCache, BAM)
- **Artefatti di persistenza** (Run Keys, Scheduled Tasks, Services, WMI)
- **Cronologia di navigazione** (Browser history, URL MRU, TypedPaths)
- **Attività utente** (UserAssist, ShellBags, LNK files, JumpLists)
- **Artefatti di rete** (Interfacce, cache DNS, profili WLAN/VPN)
- **Log eventi di Windows** (Security, System, PowerShell, RDP)
- **Memoria virtuale** (Pagefile, Hibernation, SRUM)
- **Dispositivi rimovibili** (Cronologia USB, storia delle connessioni)
- **Active Directory** (NTDS.dit, hash domain, PAD offline analysis)
- **E molto altro...**

### Modalità di Esecuzione Flessibili
```bash
./fiuto.sh                          # Menu interattivo
./fiuto.sh /mnt/windows             # Specifica la root del volume Windows
./fiuto.sh /mnt/windows --all       # Esegui tutti i moduli
./fiuto.sh /mnt/windows --module 3  # Esegui un modulo specifico
```

### Output Professionale
- Report HTML interattivi e navigabili
- Timeline cronologiche aggregate
- Log dettagliati per audit e tracciamento
- Riepilogati batch con metriche di esecuzione

---

## 🛠️ Dipendenze

### Requisiti di Sistema
- **Bash 4.0+**
- **Python 3.9+** (con module di compatibilità multiple versioni)
- **Linux o macOS** (per montare/analizzare dischi Windows offline)

### Moduli Python Richiesti
```bash
pip install regipy          # Parsing degli hive di registro offline
pip install python-evtx     # Lettura dei file .evtx
```

### Script di Supporto
Lo script utilizza internamente helper bash per:
- Parsing di file binari Windows
- Decodifica di timestamp FILETIME
- Estrazione di metadati da formati proprietari

---

## 📦 Installazione

1. **Clona il repository**
   ```bash
   git clone https://github.com/fabriren/fiuto.git
   cd fiuto
   ```

2. **Rendi eseguibile lo script**
   ```bash
   chmod +x fiuto.sh
   ```

3. **Installa le dipendenze Python**
   ```bash
   pip install regipy python-evtx
   # Su macOS con Homebrew:
   # python3 -m pip install --user regipy python-evtx
   ```

4. **Monta il disco Windows (esempio)**
   ```bash
   # Su Linux con ntfs-3g
   sudo mount -t ntfs-3g -o ro /dev/sda3 /mnt/windows
   
   # Su macOS
   sudo mount -t ntfs -o rdonly /dev/disk0s3 /mnt/windows
   ```

---

## 🚀 Guida Rapida

### Analisi Interattiva
```bash
./fiuto.sh /mnt/windows
```
Lo script presenterà un menu numerato con i 38 moduli disponibili. Seleziona il numero del modulo o digita `--all` per eseguirli tutti.

### Analisi Batch Automatica
```bash
./fiuto.sh /mnt/windows --all
```
Esegue tutti i moduli sequenzialmente, genera un riepilogo finale e salva i report in `fiuto_reports/`.

### Singolo Modulo
```bash
./fiuto.sh /mnt/windows --module 12
```
Esegue solo il modulo 12 (Event Log nel nostro esempio).

### Output
I report generati si trovano in:
```
fiuto_reports/
├── 001_PSReadLine_History.html
├── 002_IFEO.html
├── 012_Event_Log.html
├── ...
└── 38_PAD_Offline.html
```

---

## 📊 I 38 Moduli di Analisi

| # | Nome Modulo | Artefatto Windows | Utilizzo |
|---|---|---|---|
| 1 | PowerShell PSReadLine History | PSReadLine history file | Recupera comandi eseguiti e stringhe sensibili |
| 2 | IFEO | Registry SOFTWARE | Rilevare manipolazioni di eseguibili di sistema |
| 3 | BAM | Registry SYSTEM | Timeline di esecuzioni in background |
| 4 | RDP Cache | Terminal Server Client Cache | Ricostruire sessioni RDP |
| 5 | Run Keys & Persistenza | Registry NTUSER.DAT, SOFTWARE | Rilevare backdoor e malware |
| 6 | Prefetch | Windows/Prefetch/*.pf | Cronologia esecuzione programmi |
| 7 | Notepad TabState | AppData Packages | Recuperare testo non salvato |
| 8 | Scheduled Tasks | Windows/System32/Tasks | Trovar dropper temporizzati |
| 9 | USB/Dispositivi | Registry SYSTEM | Storico inserimento supporti USB |
| 10 | LNK & JumpList | AppData Recent | File usati recentemente |
| 11 | Services | Registry SYSTEM | Servizi anomali o custom |
| 12 | Event Log | Windows/System32/winevt/Logs/*.evtx | Login, RDP, attività sospette |
| 13 | Amcache + Shimcache | Amcache.hve, Registry | Cronologia esecuzione con hash |
| 14 | Recycle Bin | $Recycle.Bin | Recuperare file cancellati |
| 15 | WMI Subscriptions | OBJECTS.DATA | Persistenza fileless |
| 16 | SRUM | SRUDB.dat | Consumo risorse per app |
| 17 | Browser History | SQLite History/places | Navigazione web |
| 18 | UserAssist + MRU | Registry NTUSER.DAT | App avviate via GUI |
| 19 | ShellBags | Registry NTUSER.DAT/UsrClass | Navigazione cartelle |
| 20 | SAM Hash | Windows/System32/config/SAM | NTLM hash account locali |
| 21 | MFT Timeline | $MFT | Granular NTFS changes |
| 22 | OpenSaveMRU | Registry NTUSER.DAT | Cronologia file aperti/salvati |
| 23 | USN Journal | $UsnJrnl:$J | Attività file system massicce |
| 24 | NTDS.dit | ntds.dit, Registry SYSTEM | Hash utenti dominio |
| 25 | Hibernation/Pagefile | hiberfil.sys, pagefile.sys | Memory dump analysis |
| 26 | WER Files | WER Report Archive | Crash e errori di sistema |
| 27 | Credential Manager | Credentials DPAPI | Credenziali di rete salvate |
| 28 | WLAN & VPN Profiles | Wlansvc Profiles | SSID e profili VPN |
| 29 | AppX / UWP Packages | AppData/Local/Packages | App Modern installate |
| 30 | Browser Downloads & Logins | SQLite databases | Download e credenziali |
| 31 | Clipboard History | ActivitiesCache.db | Cronologia appunti |
| 32 | Office MRU & Recent | NTUSER.DAT, AppData | Documenti aperti recenti |
| 33 | Defender Quarantine | Quarantine folder | Minacce identificate |
| 34 | PowerShell Script Block | Event Log 4104 | Script eseguiti |
| 35 | JumpLists | Recent Destinations | Interazione utente app |
| 36 | Network Artifacts | Registry SYSTEM | DNS, interfacce, reti |
| 37 | Master Timeline | (Aggregato) | Timeline cross-artefatto |
| 38 | PAD Offline | NTDS.dit | Active Directory avanzato |

---

## 🔍 Casi d'Uso

### Indagini Malware
Utilizza i moduli di persistenza (Run Keys, Services, WMI) per trovare backdoor, e i log di PowerShell per tracciare esecuzione di payload offuscati.

### Incident Response
Combina BAM, Prefetch e Event Logs per costruire una timeline precisa di cosa è stato eseguito e quando, aiutando a identificare il paziente zero.

### Compliance & Audit
Estrai l'attività di rete (SRUM, DNS) e gli accessi RDP (RDP Cache, Event Logs) per dimostrare chi ha accesso a quali sistemi.

### Ransomware Forensics
Analizza USN Journal, MFT Timeline, Recycle Bin e Browser Downloads per tracciare la diffusione e le origini dell'infezione.

### Threat Hunting
Usa la cronologia USB, i profili WLAN/VPN e la navigazione web per scoprire esfiltrazione di dati o comunicazioni con indirizzi IP sospetti.

---

## ⚙️ Opzioni Avanzate

### Specificare Utenti
```bash
./fiuto.sh /mnt/windows --user Administrator
```
Focalizza l'analisi su un utente specifico.

### Carica Lista IoC
```bash
./fiuto.sh /mnt/windows --ioc /path/to/ioc_list.txt
```
Scansiona gli artefatti per trovare match con indicatori di compromissione.

### Modalità Silenziosa
```bash
./fiuto.sh /mnt/windows --all --silent
```
Esegui senza output interattivo (utile per script automatici).

---

## 📄 Output Report

Ogni modulo genera un report HTML con:
- **Tabella dati** strutturata e ordinabile
- **Timeline** con timestamp UTC e valori FILETIME decodificati
- **Evidenziazione** di elementi sospetti (password, token, comandi offuscati)
- **Metadata** (hash, percorsi assoluti, account coinvolti)
- **Note forensiche** su come interpretare i risultati

### Esempio Report
```
┌─ Report: Event Log (Module 12)
│
├─ Timespan: 2025-03-15 08:30:00 UTC → 2025-04-14 17:42:00 UTC
├─ Total Events: 12,847
├─ Critical Events: 8
│  ├─ Suspicious PowerShell: 3
│  ├─ Failed RDP Logins: 4
│  └─ Service Installs: 1
│
└─ Top Events
   ├─ [08:45:32] EID 4688 - Process Created: powershell.exe -NoP -W H -C "IEX ..."
   ├─ [14:22:15] EID 4768 - Kerberos Auth Failure: Administrator
   └─ ...
```

---

## 🐛 Troubleshooting

### Errore: "regipy module not found"
```bash
# Installa per la versione Python corretta
python3 -m pip install regipy
# O controlla quale Python ha regipy:
which python3 | xargs python3 -c "import regipy; print('OK')"
```

### File system di sola lettura
Lo script apre il disco in modalità read-only (`-o ro`) per impostazione predefinita. Non modifica mai il sistema analizzato.

### Permessi negati su alcuni file
Alcuni artefatti (SAM, NTDS.dit) potrebbero richiedere privilegi elevati:
```bash
sudo ./fiuto.sh /mnt/windows --all
```

### Report non generati
Verifica che la directory `fiuto_reports/` esista ed è scrivibile:
```bash
mkdir -p fiuto_reports
chmod 755 fiuto_reports
```

---

## 📝 Logging

Lo script crea un log dettagliato di ogni sessione in:
```
fiuto_reports/session_YYYY-MM-DD_HH-MM-SS.log
```

Utile per debugging e audit trail:
```bash
tail -f fiuto_reports/session_*.log
```

---

## 🤝 Contributi

Se hai miglioramenti, segnalazioni di bug o moduli aggiuntivi:

1. Fork il repository
2. Crea un branch per la tua feature (`git checkout -b feature/nuovo-modulo`)
3. Commit i tuoi cambiamenti (`git commit -am 'Add nuovo modulo'`)
4. Push al branch (`git push origin feature/nuovo-modulo`)
5. Apri una Pull Request

---

## ⚖️ Licenza

Questo progetto è distribuito sotto licenza **MIT**.

---

## 👤 Autore

Creato da **zi®iginal** per la community DFIR.

---

## 📚 Risorse Consigliate

- [SANS Windows Artifact Analysis](https://www.sans.org)
- [Plaso - Timeline log2timeline](https://plaso.readthedocs.io)
- [Registry Explorer](https://www.sans.org/tools/registry-explorer/)
- [KAPE - Kroll Artifact Parser](https://www.kroll.com/en/services/cyber-risk/incident-response-forensics/kape)
- [Hayabusa - Windows Event Log Analysis](https://github.com/Yamato-Security/hayabusa)

---

## ⚠️ Disclaimer Legale

FIUTO è uno strumento per velocizzare le analisi forensi digitale legittimo, da impiegarsi per analisi offline autorizzate. Deve essere utilizzato **solo** su sistemi di cui hai il diritto legale di analizzare. L'uso non autorizzato potrebbe violare le leggi sulla privacy e protezione dei dati.

**Non siamo responsabili per:**
- Uso non autorizzato o illegale dello strumento
- Violazioni di privacy o protezione dati
- Danni diretti o indiretti derivanti dall'uso di fiuto

---

**Ultima modifica:** 2026-04-14 | **Versione:** 2.6

<img width="982" height="574" alt="image" src="https://github.com/user-attachments/assets/0ddc9024-0736-4fe0-ae47-1e6fb6c7c003" />
<img width="983" height="934" alt="image" src="https://github.com/user-attachments/assets/3d8a3c09-fe3c-4711-baca-9d3e8c65bff7" />


