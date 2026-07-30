# FIUTO 🔍

<!-- markdownlint-disable MD033 MD060 -->

## 🇬🇧 English Version

**FIUTO** (**F**orensic **I**nvestigation **U**tility **T**ool for **O**ffline) is a unified DFIR (Digital Forensics and Incident Response) toolkit for comprehensive **offline disk analysis of Windows, Linux and macOS** volumes. It automatically **detects the operating system of each mounted volume** and proposes the relevant module set, generating detailed HTML reports for rapid and effective investigations.

- **Windows** volumes → the 39 Windows artifact modules.
- **Linux** volumes → 16 dedicated modules (system logs, journal, logins, shell/AI history, browsers, accounts, persistence, SSH, network, packages, trash, filesystem timeline, auditd, containers, master timeline).
- **macOS** volumes → 13 dedicated modules (system logs, dslocal accounts, persistence, login items/BTM, quarantine, TCC, KnowledgeC, browsers, shell/AI history, recent items, FSEvents, Spotlight, master timeline).

Everything runs strictly **offline**, parsing the read-only mounted filesystem.

---

## 📋 Key Features

### Automatic per-volume OS detection

At startup FIUTO lists the mounted volumes with an OS badge (`[Windows]` / `[Linux]` / `[macOS]`) and, once a volume is selected, shows only the modules that apply to that operating system.

### Comprehensive Windows Artifact Coverage

On Windows volumes FIUTO collects and analyzes:

- **Execution histories** (Prefetch, AmCache, ShimCache, BAM)
- **Persistence artifacts** (Run Keys, Scheduled Tasks, Services, WMI)
- **Navigation history** (Browser history, URL MRU, TypedPaths)
- **User activity** (UserAssist, ShellBags, LNK files, JumpLists)
- **Network artifacts** (Interfaces, DNS cache, WLAN/VPN profiles)
- **Windows event logs** (Security, System, PowerShell, RDP)
- **Virtual memory** (Pagefile, Hibernation, SRUM)
- **Removable devices** (USB history, connection timeline)
- **Active Directory** (NTDS.dit, domain hashes, PAD offline analysis)
- **AI assistant chat history** (ChatGPT, Copilot, Claude, Cursor, Gemini, Codex, Windsurf, Continue) — *new in v1.2*
- **And much more...**

### Linux Artifact Coverage — *new in v2.0*

On Linux volumes FIUTO collects and analyzes:

- **System logs** (`/var/log`: syslog, auth, kern, secure…) and **systemd journal** (`*.journal`)
- **Login history** (`wtmp` / `btmp` / `lastlog` — successful and failed logins with source IP)
- **Shell history** (bash/zsh/sh + python/mysql/psql) and **AI CLI history** (claude, aider, aichat, ollama…)
- **Browser history** (Firefox `places.sqlite`, Chrome/Chromium/Brave/Edge — incl. snap/flatpak paths)
- **User accounts** (`passwd`, `shadow`, `group`, `sudoers`, `sudoers.d`)
- **Persistence** (cron, systemd units/timers, autostart, `rc.local`, shell init, `ld.so.preload`)
- **SSH artifacts** (`authorized_keys`, `known_hosts`, `sshd_config`, private-key presence)
- **Network config** (`hosts`, NetworkManager Wi-Fi PSK, netplan, iptables/nftables)
- **Installed packages** (dpkg, rpm, apt history, snap — installation timeline)
- **Trash & recent files** (`~/.local/share/Trash` with deletion timestamps, `recently-used.xbel`)
- **Filesystem MAC-time timeline** of forensically sensitive areas
- **auditd** (`/var/log/audit`) — syscalls, authentications, EXECVE with hex-decoded arguments, policy violations
- **Containers** (Docker/Podman) — offline inventory from on-disk metadata, with detection of escape-prone configurations (privileged, host root or Docker socket bind-mounted, `CAP_SYS_ADMIN`, host PID/network namespace)
- **PAM** — authentication backdoors: suspicious directives and `.so` modules owned by no package
- **Kernel modules** — LKM rootkits, `install` directives in `modprobe.d`, unpackaged `.ko`
- **Web server logs** — nginx/apache, ranked by HTTP status so a successful attack stands out from the background noise
- **Cloud and development credentials** — `~/.aws`, `~/.kube`, `~/.docker`, `~/.ssh`: presence and non-secret identifiers, never the secrets themselves
- **SUID/SGID, capabilities and world-writable files** — the privilege escalation surface

### macOS Artifact Coverage — *new in v2.0*

On macOS volumes FIUTO collects and analyzes:

- **System logs** (`system.log`, `install.log`, ASL) — *unified `.tracev3` logs are out of scope*
- **User accounts** (dslocal `*.plist` — UID, home, shell, password-hash presence)
- **Persistence** (`LaunchAgents` / `LaunchDaemons` system & per-user, cron) — binary plists rendered readable
- **Login Items / BTM** (`backgrounditems.btm`)
- **Quarantine / downloads** (`QuarantineEventsV2` — download URL + timestamp)
- **TCC privacy** (`TCC.db` — camera/mic/disk permissions)
- **KnowledgeC** (`knowledgeC.db` — app usage / device activity)
- **Browser history** (Safari `History.db`, Chrome, Firefox)
- **Shell & AI history** (`.zsh_history`, `.bash_history`, AI CLI)
- **Recent items** (`SFL`/`SFL2`, `~/.Trash`)
- **FSEvents** (`/.fseventsd`) — filesystem change history, the macOS counterpart of the USN Journal
- **Spotlight** (`.Spotlight-V100/store.db`) — heuristic extraction of download URLs and user paths
- **Messages** (`chat.db`) — iMessage and forwarded SMS, with links and credential references flagged
- **Cookies and downloads** — Safari `Cookies.binarycookies` (values never printed, only their length) and `Downloads.plist`
- **XProtect and Gatekeeper** — signature version, manually granted authorisations, third-party kernel extensions
- **Application inventory** — signature presence and bundles outside the standard directories
- **Time Machine and snapshots** — earlier versions of files; their absence is reported as an indicator
- **Unified logs** (`*.tracev3`) — LZ4 chunk decompression and string extraction; **partial by design**, see the note below

### Flexible Execution Modes

```bash
./fiuto.sh                          # Interactive menu (auto-detects mounted volumes + OS)
./fiuto.sh /mnt/disk                # Specify a volume root (Windows / Linux / macOS)
./fiuto.sh /mnt/disk --all          # Run all modules for the detected OS
./fiuto.sh /mnt/disk --module 3     # Run a specific module (numbering depends on the OS)
./fiuto.sh /mnt/disk --modules 1,4,6-8   # Run a list/range of modules
./fiuto.sh /mnt/disk --all --jsonl       # Also export JSONL (Timesketch schema)
./fiuto.sh /mnt/disk --all --since 2026-03-01 --until 2026-03-08  # Restrict to a time window
./fiuto.sh /mnt/disk --all --no-log-replay   # Do not replay registry .LOG1/.LOG2
```

The module numbers shown by `--module`/`--modules` always refer to the **menu of the detected OS**.

### Professional Output

- Interactive and navigable HTML reports
- One-click CSV export of the complete data for every module
- Aggregated chronological timelines
- Detailed logs for audit and tracking
- Batch summaries with execution metrics

---

## 🛠️ Dependencies

### System Requirements

- **Bash 4.0+**
- **Python 3.9+** (with multi-version compatibility)
- **Linux (or WSL) or macOS** (host for mounting/analyzing the offline target disks)
- **Optional:** `journalctl` (Linux journal parsing), `rpm` CLI (offline RPM dump). `sqlite3` and `plistlib` ship with Python 3 — no extra install needed for Linux/macOS modules.

### Required Python Modules

```bash
pip install regipy          # Offline registry hive parsing + .LOG1/.LOG2 replay
pip install python-evtx     # Reading .evtx files
```

### Recommended for full Windows coverage

These libraries unlock the full output of some Windows modules. They are **optional**: if missing, the affected module falls back to a partial parser instead of failing.

```bash
pip install impacket          # SAM / NTDS.dit / AD hashes (modules 20, 24, 38)
pip install libesedb-python   # ESE database parsing — SRUM (module 16)
pip install mft               # MFT timeline (module 21)
pip install python-snappy     # ChatGPT LevelDB decompression (module 39) — or: pip install cramjam
pip install libpff-python     # Outlook PST/OST parsing (module 44)
pip install libesedb-python   # WebCacheV01 and Windows.edb (modules 49, 50)
```

> **Linux and macOS modules need no extra packages** — they rely only on the Python standard library (`sqlite3`, `plistlib`, …). Optionally, `journalctl` (for systemd journal) and the `rpm` CLI (for offline RPM dumps) improve coverage if present.

### Support Scripts

The script uses internal bash helpers for:

- Binary Windows file parsing
- FILETIME timestamp decoding
- Proprietary format metadata extraction

---

## 📦 Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/fabriren/fiuto.git
   cd fiuto
   ```

2. **Make the script executable**

   ```bash
   chmod +x fiuto.sh
   ```

3. **Install Python dependencies**

   ```bash
   pip install regipy python-evtx
   # On macOS with Homebrew:
   # python3 -m pip install --user regipy python-evtx
   ```

4. **Mount the target disk read-only (examples)**

   ```bash
   # Windows (NTFS) on Linux with ntfs-3g
   sudo mount -t ntfs-3g -o ro /dev/sda3 /mnt/disk

   # Linux (ext4)
   sudo mount -o ro /dev/sda2 /mnt/disk

   # macOS (APFS/HFS+) — read-only
   sudo mount -o ro /dev/sdb2 /mnt/disk
   ```

   FIUTO auto-detects the OS of the mounted volume; no need to tell it which OS it is.

---

## 🚀 Quick Start Guide

### Interactive Analysis

```bash
./fiuto.sh /mnt/disk
```

The script detects the volume's OS and presents a numbered menu with the relevant modules (52 for Windows, 23 for Linux, 21 for macOS). Select the module number or type `--all` to run them all.

### Automated Batch Analysis

```bash
./fiuto.sh /mnt/windows --all
```

Runs all modules sequentially, generates a final summary, and saves reports to `fiuto_reports/`.

### Single Module

```bash
./fiuto.sh /mnt/windows --module 12
```

Runs only module 12 (Event Log in our example).

### Output

Generated reports are located in:

```text
fiuto_reports/
├── 001_PSReadLine_History.html
├── 002_IFEO.html
├── 012_Event_Log.html
├── ...
└── 38_PAD_Offline.html
```

---

## 📊 The 54 Windows analysis modules

| #  | Module Name                   | Windows Artifact                    | Usage                                                           |
| -- | ----------------------------- | ----------------------------------- | --------------------------------------------------------------- |
| 1  | PowerShell PSReadLine History | PSReadLine history file             | Retrieve executed commands and sensitive strings                |
| 2  | IFEO                          | Registry SOFTWARE                   | Detect executable manipulation                                  |
| 3  | BAM                           | Registry SYSTEM                     | Timeline of background executions                               |
| 4  | RDP Cache                     | Terminal Server Client Cache        | Reconstruct RDP sessions                                        |
| 5  | Run Keys & Persistence        | Registry NTUSER.DAT, SOFTWARE       | Detect backdoors and malware                                    |
| 6  | Prefetch                      | Windows/Prefetch/*.pf               | Program execution history                                       |
| 7  | Notepad TabState              | AppData Packages                    | Recover unsaved text                                            |
| 8  | Scheduled Tasks               | Windows/System32/Tasks              | Find scheduled droppers                                         |
| 9  | USB/Devices                   | Registry SYSTEM                     | USB device insertion history                                    |
| 10 | LNK & JumpList                | AppData Recent                      | Recently used files                                             |
| 11 | Services                      | Registry SYSTEM                     | Anomalous or custom services                                    |
| 12 | Event Log                     | Windows/System32/winevt/Logs/*.evtx | Logins, RDP, suspicious activity                                |
| 13 | Amcache + Shimcache           | Amcache.hve, Registry               | Execution history with hashes                                   |
| 14 | Recycle Bin                   | $Recycle.Bin                        | Recover deleted files                                           |
| 15 | WMI Subscriptions             | OBJECTS.DATA                        | Fileless persistence                                            |
| 16 | SRUM                          | SRUDB.dat                           | Resource consumption by app                                     |
| 17 | Browser History               | SQLite History/places               | Web navigation                                                  |
| 18 | UserAssist + MRU              | Registry NTUSER.DAT                 | GUI-launched apps                                               |
| 19 | ShellBags                     | Registry NTUSER.DAT/UsrClass        | Folder navigation                                               |
| 20 | SAM Hash                      | Windows/System32/config/SAM         | NTLM hashes of local accounts                                   |
| 21 | MFT Timeline                  | $MFT                                | Granular NTFS changes                                           |
| 22 | OpenSaveMRU                   | Registry NTUSER.DAT                 | Open/save file history                                          |
| 23 | USN Journal                   | $UsnJrnl:$J                       | Massive file system activity                                    |
| 24 | NTDS.dit                      | ntds.dit, Registry SYSTEM           | Domain user hashes                                              |
| 25 | Hibernation/Pagefile          | hiberfil.sys, pagefile.sys          | Memory dump analysis                                            |
| 26 | WER Files                     | WER Report Archive                  | System crashes and errors                                       |
| 27 | Credential Manager            | Credentials DPAPI                   | Saved network credentials                                       |
| 28 | WLAN & VPN Profiles           | Wlansvc Profiles                    | SSIDs and VPN profiles                                          |
| 29 | AppX / UWP Packages           | AppData/Local/Packages              | Installed modern apps                                           |
| 30 | Browser Downloads & Logins    | SQLite databases                    | Downloads and credentials                                       |
| 31 | Clipboard History             | ActivitiesCache.db                  | Clipboard history                                               |
| 32 | Office MRU & Recent           | NTUSER.DAT, AppData                 | Recently opened documents                                       |
| 33 | Defender Quarantine           | Quarantine folder                   | Identified threats                                              |
| 34 | PowerShell Script Block       | Event Log 4104                      | Executed scripts                                                |
| 35 | JumpLists                     | Recent Destinations                 | App user interaction                                            |
| 36 | Network Artifacts             | Registry SYSTEM                     | DNS, interfaces, networks                                       |
| 37 | Master Timeline               | (Aggregated)                        | Cross-artifact timeline                                         |
| 38 | PAD Offline                   | NTDS.dit                            | Advanced Active Directory                                       |
| 39 | AI Chat / Query History       | LevelDB / JSON / SQLite             | Recover AI assistant conversations (ChatGPT, Copilot, Claude…) |
| 40 | SetupAPI Device Log           | Windows/INF/setupapi.dev.log        | **First** install time of USB devices                           |
| 41 | PowerShell Transcript         | PowerShell_transcript*.txt          | Full sessions: commands **and** output                          |
| 42 | LSA Secrets & DCC2            | Registry SECURITY                   | Service-account passwords, cached domain credentials            |
| 43 | Volume Shadow Copies          | System Volume Information           | Earlier volume snapshots, differential analysis                 |
| 44 | Outlook PST / OST             | *.pst, *.ost (libpff)               | Local mail, attachments, deleted items                          |
| 45 | Cloud Sync                    | OneDrive, Dropbox, Google Drive     | Synced files — the modern exfiltration path                     |
| 46 | BITS Jobs                     | qmgr.db                             | Background downloads used as LOLBin (T1197)                     |
| 47 | Thumbcache                    | thumbcache_*.db                     | Thumbnails of **deleted** files, carved                         |
| 48 | Chat Desktop                  | Slack / Teams / Discord (LevelDB)   | Internal social engineering, files shared in private chats      |
| 49 | WebCacheV01                   | WebCacheV01.dat (ESE)               | IE/Edge Legacy **and everything using WinINET**                 |
| 50 | Windows Search Index          | Windows.edb (ESE)                   | Paths and content excerpts of **deleted** files                 |
| 51 | SQLite Recovery *(cross-OS)*  | freelist, unallocated space         | Content of **deleted** records still on disk                    |
| 52 | EFI System Partition *(cross-OS)* | ESP                             | Bootkits: code running **before** the OS and any EDR            |
| 53 | YARA *(cross-OS)*                 | YAR                             | External rules applied to a declared, bounded scope (`--yara`)  |
| 54 | Sigma                             | SIG                             | Community detections over EVTX, declared subset (`--sigma`)     |

---

## 🐧 The 24 Linux analysis modules

| #  | Module Name         | Linux Artifact                                                     | Usage                                                           |
| -- | ------------------- | ------------------------------------------------------------------ | --------------------------------------------------------------- |
| 1  | System Logs         | `/var/log/{syslog,messages,auth.log,secure,kern.log,…}`         | System/auth events, failures, sudo                              |
| 2  | systemd Journal     | `/var/log/journal/*.journal`                                     | Structured journal (via `journalctl` or `strings` fallback) |
| 3  | Login History       | `wtmp` / `btmp` / `lastlog`                                  | Successful/failed logins with source IP                         |
| 4  | Shell History       | `.bash_history`, `.zsh_history`, `.python_history`, …       | Executed commands and sensitive strings                         |
| 5  | AI CLI History      | `~/.claude`, aider, aichat, ollama, copilot                      | AI assistant CLI conversations                                  |
| 6  | Browser History     | Firefox `places.sqlite`, Chrome/Chromium/Brave/Edge              | Web navigation (incl. snap/flatpak paths)                       |
| 7  | User Accounts       | `passwd` / `shadow` / `group` / `sudoers`                  | Accounts, UID 0, empty/locked passwords                         |
| 8  | Persistence         | cron, systemd units/timers, autostart,`rc.local`, shell init     | Backdoors and auto-start mechanisms                             |
| 9  | SSH Artifacts       | `authorized_keys`, `known_hosts`, `sshd_config`              | Remote access trust and config                                  |
| 10 | Network Config      | `hosts`, NetworkManager (Wi-Fi PSK), netplan, iptables/nftables  | Network configuration and Wi-Fi keys                            |
| 11 | Installed Packages  | dpkg, rpm, apt history, snap                                       | Installation timeline / suspicious packages                     |
| 12 | Trash & Recent      | `~/.local/share/Trash` (+`.trashinfo`), `recently-used.xbel` | Deleted files with deletion timestamp                           |
| 13 | Filesystem Timeline | MAC times of sensitive areas (`find`/`stat`)                   | Cross-area chronological timeline                               |
| 14 | auditd              | `/var/log/audit/audit.log*`                                      | Syscall, authentication, EXECVE, policy violations              |
| 15 | Container           | `/var/lib/docker`, `/var/lib/containers`                        | Docker/Podman inventory + container-escape indicators           |
| 16 | Master Timeline     | (Aggregated)                                                       | Cross-module chronological timeline (runs last with `--all`)    |
| 17 | PAM                 | `/etc/pam.d`, `security/*.so`                                    | Authentication backdoors, unpackaged modules                    |
| 18 | Kernel Modules      | `lib/modules`, `modprobe.d`, initramfs                           | LKM rootkits, `install` directives, unpackaged `.ko`            |
| 19 | Web Server Logs     | nginx / apache access+error                                        | Webshell, traversal, SQLi — ranked by HTTP status               |
| 20 | Cloud Credentials   | `~/.aws`, `~/.kube`, `~/.docker`, `~/.ssh`                     | Keys granting access to the wider infrastructure                |
| 21 | SUID & Capabilities | filesystem scan                                                    | Privilege escalation surface, unexpected SUID                   |
| 22 | SQLite Recovery *(cross-OS)* | freelist, unallocated space                               | Content of **deleted** records still on disk                    |
| 23 | EFI System Partition *(cross-OS)* | ESP                                                  | Bootkits: code running **before** the OS                        |
| 24 | YARA *(cross-OS)*                 | YAR                                                  | External rules, declared scope (`--yara`)                       |

---

## 🍎 The 22 macOS analysis modules

| #  | Module Name            | macOS Artifact                                                 | Usage                                                         |
| -- | ---------------------- | -------------------------------------------------------------- | ------------------------------------------------------------- |
| 1  | System Logs            | `system.log`, `install.log`, ASL                           | System/install events (*`.tracev3` out of scope*)         |
| 2  | User Accounts          | dslocal `*.plist`                                            | Accounts, UID, shell, password-hash presence                  |
| 3  | Persistence            | `LaunchAgents` / `LaunchDaemons` (system & per-user), cron | Persistence (binary plists rendered readable)                 |
| 4  | Login Items (BTM)      | `backgrounditems.btm`                                        | Background/login items                                        |
| 5  | Quarantine / Downloads | `QuarantineEventsV2`                                         | Download provenance (URL + timestamp)                         |
| 6  | TCC Privacy            | `TCC.db`                                                     | Camera/mic/disk privacy permissions                           |
| 7  | KnowledgeC             | `knowledgeC.db`                                              | App usage / device activity                                   |
| 8  | Browser History        | Safari `History.db`, Chrome, Firefox                         | Web navigation                                                |
| 9  | Shell & AI History     | `.zsh_history`, `.bash_history`, AI CLI                    | Executed commands and AI conversations                        |
| 10 | Recent Items           | `SFL`/`SFL2`, `~/.Trash`                                 | Recent apps/docs/servers and trash                            |
| 11 | FSEvents               | `/.fseventsd`                                                | Filesystem change history (creations, renames, deletions)     |
| 12 | Spotlight              | `.Spotlight-V100/store.db`                                   | Download provenance and names of deleted files (heuristic)    |
| 13 | Master Timeline        | (Aggregated)                                                   | Cross-module chronological timeline (runs last with `--all`)  |
| 14 | Messages               | `chat.db`                                                    | iMessage/SMS: smishing, 2FA codes, off-channel talk           |
| 15 | Cookie & Download      | `Cookies.binarycookies`, `Downloads.plist`                 | Authenticated services; download provenance                   |
| 16 | XProtect / Gatekeeper  | XProtect, SystemPolicy, KextPolicy                             | Signature version; manually granted authorisations            |
| 17 | Applications           | `/Applications`, user homes                                  | Inventory, signature presence, non-standard locations         |
| 18 | Time Machine / Snapshot| `com.apple.TimeMachine.plist`, `Backups.backupdb`          | Earlier versions of files; absence as an indicator            |
| 19 | Unified Logs           | `*.tracev3`                                                  | LZ4 chunk decompression + string extraction (**partial**)     |
| 20 | SQLite Recovery *(cross-OS)* | freelist, unallocated space                            | Content of **deleted** records still on disk                  |
| 21 | EFI System Partition *(cross-OS)* | ESP                                               | Bootkits: code running **before** the OS                      |
| 22 | YARA *(cross-OS)*                 | YAR                                               | External rules, declared scope (`--yara`)                     |

---

## 🔍 Use Cases

### Malware Investigations

Use persistence modules (Run Keys, Services, WMI) to find backdoors, and PowerShell logs to track obfuscated payload execution.

### Incident Response

Combine BAM, Prefetch, and Event Logs to build an accurate timeline of what was executed and when, helping identify patient zero.

### Compliance & Audit

Extract network activity (SRUM, DNS) and RDP access (RDP Cache, Event Logs) to demonstrate who accessed which systems.

### Ransomware Forensics

Analyze USN Journal, MFT Timeline, Recycle Bin, and Browser Downloads to trace infection spread and origins.

### Threat Hunting

Use USB history, WLAN/VPN profiles, and web navigation to uncover data exfiltration or communication with suspicious IPs.

---

## ⚙️ Advanced Options

### Load IoC List

```bash
./fiuto.sh /mnt/windows --ioc /path/to/ioc_list.txt
```

Scan artifacts for matches with indicators of compromise.

### Registry transaction log replay (Windows)

Windows does not write registry changes straight into the primary hive: it
queues them in the transaction logs (`.LOG1` / `.LOG2`) and consolidates them
only on a clean unmount. A hive taken from a machine powered off abruptly, from
a disk image or from a snapshot is therefore almost always *dirty* — the most
recent writes, frequently the attacker's, exist **only** in the logs.

FIUTO replays those logs by default onto a **temporary copy**; the evidence
volume is never written to. On a real test hive this recovered 2,456 additional
keys and 3,905 additional values that raw parsing simply does not see.

```bash
./fiuto.sh /mnt/windows --all --no-log-replay   # disable (not recommended)
```

Requires `regipy`. If it is missing, FIUTO warns once and falls back to the raw
hive rather than failing.

### Chain of custody

Every session writes `evidence_manifest.json` into the report directory. It
records what the analysis was based on, so the reports can be verified later:

- tool version, exact command line, operator and analysis host;
- session start and end in UTC;
- the analysed volume, its detected OS and the hostname read from artefacts;
- **every evidence file consulted**, with size, mtime and SHA-256;
- **every generated report**, with its SHA-256;
- the outcome of the registry transaction-log replay.

Coverage is automatic: files are recorded inside the functions every module
goes through to reach an artefact (`ci_find_file`, `query_sqlite`,
`read_plist`, `recover_hive`), so modules added later are covered without
touching them.

```bash
./fiuto.sh /mnt/disk --all --no-hash        # record files without hashing (faster)
./fiuto.sh /mnt/disk --all --hash-limit 256 # skip hashing above 256 MB
./fiuto.sh /mnt/disk --all --no-custody     # disable the manifest entirely
```

Files above the size limit (1 GB by default — `pagefile.sys`, `$MFT`,
`Windows.edb`) are still listed, with the reason the hash is missing rather
than silently omitting it.

### Redacted copies (`--redact`, `--defang`)

Reports hold material that belongs in the office and nowhere else: NTLM hashes
from SAM and NTDS, Wi-Fi PSKs, OAuth tokens, private keys, LSA secrets. As long
as the only options are "send the whole report or send nothing", that material
ends up in a client email or a ticket attachment.

```bash
./fiuto.sh /mnt/disk --all --redact     # report.redacted.html beside each report
./fiuto.sh /mnt/disk --all --defang     # also renders URLs and IPs inert
```

**The original is never touched.** A `report.redacted.html` appears next to it.
Redacting in place would destroy evidence for a communication need — a trade the
tool has no business making on your behalf. The custody manifest lists both,
with distinct roles.

**Redaction is by context, not by shape.** A SHA-256 and an NTLM hash are both
hex strings: the first is a exhibit's integrity fingerprint and removing it would
break the chain of custody in the very file you are about to share; the second is
a credential. What tells them apart is the label beside them, so the rules look at
the field, not at the value. Labels stay — knowing *that* there was a password is
part of the analysis — and only values go.

Every copy declares in its header that it is not the original, and lists how
many occurrences each rule removed. A file that looks like a report but is not
quite one, without saying so, is worse than no file.

### IoC engine (`--ioc`)

```bash
./fiuto.sh /mnt/disk --all --ioc indicators.txt     # flat list
./fiuto.sh /mnt/disk --all --ioc bundle.json        # STIX 2.x or MISP export
```

Indicators are **typed on load** and matched with the boundaries their type
deserves. Plain substring matching, which is what FIUTO did until 2.2, was wrong
in two opposite directions:

- **false positives** — `10.0.0.5` matches inside `110.0.0.55`, `evil.com`
  matches inside `notevil.com.au`. In a report with a hundred thousand rows that
  is not annoying noise: it is noise that hides the real hits;
- **false negatives** — indicators almost always arrive defanged (`hxxp://`,
  `1[.]2[.]3[.]4`, `evil[.]com`), because that is how they are written in a mail
  or an advisory. Searched literally they match nothing, and the analyst
  concludes the disk is clean.

Both are handled: defanged input is normalised, and each type gets its own
boundaries. A domain matches its subdomains (`mail.evil.com` for `evil.com`) but
not a longer suffix (`evil.com.other.net`). Hashes are case-insensitive tokens.
A value between slashes (`/inv[o0]ke/`) is a regular expression.

Types are inferred from the value; `type:value` overrides the guess where the
guess cannot win (`file:payload.com` — a filename or a domain?). The count per
type is printed on load, which is what makes a misread file obvious: forty
`literal` where forty hashes were expected is a format problem, and without that
line it would only surface as an absence of matches.

`--ioc` also reads STIX 2.x bundles and MISP exports, nested MISP objects
included. A JSON file that is neither is **refused**, not read line by line:
loading braces as a literal indicator would make it match everywhere.

### Parallel execution (`--jobs`)

```bash
./fiuto.sh /mnt/disk --all --jobs 4
```

Opt-in, not the default. Modules are independent — different files in, different
folders out — but on a volume mounted from a spinning disk or over the network,
N readers are *slower* than one. The bottleneck here is almost always I/O rather
than CPU, and only the person running the analysis knows which one they have.

The guarantee is that **the outcome does not depend on the degree of
parallelism**: same reports, same order in the summary. A result that changes
with `--jobs` is a result you cannot rely on, and in expert-witness work that is
worse than a slow tool. There are tests for exactly this.

Modules flagged `defer` — the Master Timeline, which aggregates the others — run
last, alone, once the pool has drained *and* the other results have been
collected. They read the list of generated reports, so running them any earlier
would produce an empty timeline without saying so.

Registry hive recovery takes a lock: with several modules asking for the same
hive at once, the second would otherwise read the reconstructed copy while the
first is still writing it. A truncated hive raises no error — it just yields
partial results, which is worse.

ESC to skip a module is not available in parallel: intercepting it needs
exclusive control of the terminal. This is stated at startup rather than letting
the key quietly stop working.

### YARA (`--yara`)

YARA is how the industry ships signatures: a threat-intel feed, a vendor IR
team or a national CERT hands you `.yar` files. FIUTO applies them.

```bash
./fiuto.sh /mnt/disk --all --yara /rules/               # a file or a directory
./fiuto.sh /mnt/disk --module 24 --yara r.yar --yara-scan /mnt/disk/Users
./fiuto.sh /mnt/disk --all --yara /rules/ --yara-max-mb 256
```

Requires `yara-python`. There is no fallback if it is missing — a YARA scan
without YARA is not a scan, and the module says so instead of pretending.

**The delicate part is scope.** Walking a terabyte volume file by file is not
practical on a forensic workstation, so the module covers a bounded set of
locations: the ones writable without privileges, where code that no package
manager installed nearly always ends up (`Temp`, `AppData`, `ProgramData`,
`/tmp`, `/dev/shm`, `~`, `/var/www`, LaunchAgents, the ESP…). `--yara-scan`
replaces that set with a path of your choosing.

The report then **lists exactly what was scanned**, per location and with file
counts, plus what was skipped and why: files over the per-file cap, unreadable
files, and whether the overall file cap was hit — in which case the scan is
labelled PARTIAL. Symlinks are never followed: on a mounted volume they would
lead out of the evidence and into the analysis workstation's own file system.
Without that accounting, "no match" would read as "the disk is clean"; with it,
it reads as what it is.

A rule file that does not compile is reported and skipped; the others still run.

### Sigma (`--sigma`, Windows)

Sigma is how the community publishes detections: SigmaHQ, CERTs and vendors
ship thousands of YAML rules. Applying them to the EVTX of an acquired disk is
what Chainsaw and Hayabusa do, and it is the step that turns a pile of logs
into a triage.

```bash
./fiuto.sh /mnt/windows --all --sigma /sigma/rules/windows/
```

Requires `pyyaml` and `python-evtx`.

**The supported subset is declared, not implied.** Sigma is a broad language —
base64 modifiers, CIDR, parenthesised conditions, temporal aggregations.
Implementing a part of it and pretending to support all of it would mean a rule
that was never evaluated shows up as a rule that did not fire: a silent false
negative, the worst defect a detection can have. Rules the engine cannot
evaluate are **counted and listed in the report with the reason**, grouped by
cause.

| | |
|---|---|
| Selections | field/value maps, value lists (OR), lists of maps (OR) |
| Modifiers | `contains`, `startswith`, `endswith`, `re`, `all`, `cased` |
| Conditions | `sel`, `a and b`, `a or b`, `a and not b`, `not a`, `1 of x*`, `all of x*`, `1 of them`, `all of them` |
| Values | `null` = field absent or empty |
| **Rejected** | parenthesised or aggregating conditions, `base64offset`/`utf16`/`wide`/`cidr`/`gt`/`lt`, unmappable logsource |

A rule whose `logsource` cannot be mapped to an EVTX channel present on the
volume is **not** run against every log as a fallback: it would be evaluated
against fields that channel does not have, and "did not fire" would mean
nothing. It is rejected and listed.

The report also states how many records were read per channel, and marks the
evaluation PARTIAL if the record cap was reached.

### Executive summary

Ninety HTML reports are a dump, not an analysis. At the end of `--all` — and
from `[S]` in the menu, for the reports produced so far — FIUTO writes an
`executive_summary.html` and a machine-readable `findings.json` next to them.

It answers two questions: **where do I start** and **what happened alongside
what else**. It does not answer "is this machine compromised": that is the
analyst's conclusion, and the page says so at the top.

Findings come from three substrates, in decreasing order of reliability:

1. **rows the modules already flagged** — every module passes its own keyword
   list to the log renderer; a flagged row is a judgement by someone who knows
   the artefact;
2. **IoC matches**, when `--ioc` was used;
3. **an explicit rule table**, deliberately small, each rule carrying a MITRE
   ATT&CK technique and an explanation of why that data matters.

Rules run **only on the data** — table rows and log lines — never on the
explanatory notes FIUTO itself writes into the reports. A detection engine
reading the whole page would fire on its own prose; there is a test for this.

**Cross-module correlation** is what no single report can show. Events from all
reports are clustered in 30-minute windows; when a window contains modules
matching a known scenario, the summary states the hypothesis — for example
removable media connected while LNK files were opened and the USN journal
recorded changes. These are **hypotheses to verify**, and the page labels them
as such: temporal coincidence is not causation.

The **priority score** is the sum of finding weights (critical 40, high 15,
medium 5, low 1) capped at 100. It orders the work queue; it does not measure
compromise, and the report prints the formula so the number is never taken for
more than it is. Equally, no findings is not a clean bill of health — it means
a small, conservative rule set found nothing.

### Time window (`--since` / `--until`)

On a large disk a report can run to hundreds of thousands of rows spanning
years, while the incident sits in three days.

```bash
./fiuto.sh /mnt/disk --all --since 2026-03-01 --until 2026-03-08
./fiuto.sh /mnt/disk --all --since -7d      # last 7 days
./fiuto.sh /mnt/disk --all --since "2026-03-01 14:30"
```

Accepted forms: `YYYY-MM-DD`, `YYYY-MM-DD HH:MM[:SS]`, the same with `T` and/or
a trailing `Z`, and the relative `-7d` / `-36h` / `-90m`. A bound that cannot be
parsed **stops the run** instead of being ignored: silently dropping the filter
would produce a report that claims a coverage it does not have.

The filter applies to tables and to `<pre>` log blocks alike, and to the JSONL
export, so the two views of a module cannot contradict each other. Two rules
keep it from removing evidence:

- **rows carrying no date are always kept** — they cannot be evaluated;
- a row carrying **several** dates is kept if *any* of them falls in the window
  (a file created before the window but used inside it stays).

Every filtered table and log block declares how many rows it hid, and the
manifest records the window along with the total, so a filtered report can never
be mistaken for an empty one. Line numbers in `<pre>` blocks stay those of the
original file: the gaps are the visible sign that something was removed.

**On timezones.** FIUTO detects the volume's timezone (`/etc/timezone`,
`/etc/localtime`, `SYSTEM\Control\TimeZoneInformation`) and declares it — at
startup, in the manifest and in every JSONL record — but does **not** convert
anything. Artefacts on a single volume mix UTC (registry, Windows event logs)
and local time (syslog, shell history); converting blindly would shift events by
hours, which is far worse than dates that are honestly ambiguous. The window
therefore compares dates *as they appear in the artefact*: for windows of days
this is irrelevant, for windows of hours widen them by the volume's offset.

### JSONL export (Timesketch / plaso)

```bash
./fiuto.sh /mnt/disk --all --jsonl
```

Alongside every HTML report writes a `report.jsonl`, plus a single
`fiuto_timeline.jsonl` for the whole session, using the fields expected by
Timesketch (`datetime`, `timestamp_desc`, `message`) enriched with volume,
hostname and module. Events whose year had to be inferred from a syslog-style
timestamp are flagged with `year_inferred: true` rather than being presented as
certain.

---

## 📄 Report Output

Each module generates an HTML report with:

- **Structured data table** that is sortable and filterable
- **Timeline** with UTC timestamps and decoded FILETIME values
- **Highlighting** of suspicious elements (passwords, tokens, obfuscated commands)
- **Metadata** (hashes, absolute paths, involved accounts)
- **Forensic notes** on how to interpret results

### Example Report

```text
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

### Error: "regipy module not found"

```bash
# Install for correct Python version
python3 -m pip install regipy
# Or check which Python has regipy:
which python3 | xargs python3 -c "import regipy; print('OK')"
```

### Read-only file system

The script opens the disk in read-only mode (`-o ro`) by default. It never modifies the analyzed system.

### Permission denied on some files

Some artifacts (SAM, NTDS.dit) may require elevated privileges:

```bash
sudo ./fiuto.sh /mnt/windows --all
```

### Reports not generated

Verify the `fiuto_reports/` directory exists and is writable:

```bash
mkdir -p fiuto_reports
chmod 755 fiuto_reports
```

---

## 📝 Logging

The script creates a detailed session log in:

```text
fiuto_reports/session_YYYY-MM-DD_HH-MM-SS.log
```

Useful for debugging and audit trail:

```bash
tail -f fiuto_reports/session_*.log
```

---

## 🧪 Development and tests

```bash
bats tests/                              # test suite (bats-core)
shellcheck -S warning -x fiuto.sh        # lint
python3 tests/lint_embedded_python.py fiuto.sh   # compile the embedded parsers
```

CI ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)) runs on every push:
bash syntax, ShellCheck, the bats suite, and compilation of the ~99 Python
parsers embedded as heredocs on both Python 3.9 and 3.12.

That last job is not decoration: `bash -n` treats heredocs as opaque text, so a
syntax error inside an embedded parser stays invisible until that module runs on
a real disk. It is how a long-standing defect in module 38 (PAD Offline) was
found — the block never compiled, so the module had never produced output.

## 🤝 Contributing

If you have improvements, bug reports, or additional modules:

1. Fork the repository
2. Create a branch for your feature (`git checkout -b feature/new-module`)
3. Commit your changes (`git commit -am 'Add new module'`)
4. Push to the branch (`git push origin feature/new-module`)
5. Open a Pull Request

---

## ⚖️ License

This project is **free** and distributed under the **MIT License**.

If FIUTO has been useful to you and you'd like to support its development, a small donation via PayPal is always appreciated — but entirely optional! 🙏

[![Donate via PayPal](https://img.shields.io/badge/Donate-PayPal-blue.svg)](https://paypal.me/rendina)

---

## 👤 Author

Created by **zi®iginal** for the DFIR community.

---

## 📚 Recommended Resources

- [SANS Windows Artifact Analysis](https://www.sans.org)
- [Plaso - Timeline log2timeline](https://plaso.readthedocs.io)
- [Registry Explorer](https://www.sans.org/tools/registry-explorer/)
- [KAPE - Kroll Artifact Parser](https://www.kroll.com/en/services/cyber-risk/incident-response-forensics/kape)
- [Hayabusa - Windows Event Log Analysis](https://github.com/Yamato-Security/hayabusa)

---

## ⚠️ Legal Disclaimer

FIUTO is a tool to accelerate legitimate digital forensic analysis, intended for authorized offline analysis. It must be used **only** on systems you have the legal right to analyze. Unauthorized use may violate privacy and data protection laws.

**We are not responsible for:**

- Unauthorized or illegal tool usage
- Privacy or data protection violations
- Direct or indirect damages from using fiuto

---

---

## 🇮🇹 Versione Italiana

**FIUTO** (**F**orensic **I**nvestigation **U**tility **T**ool for **O**ffline) è un toolkit DFIR (Digital Forensics and Incident Response) unificato per l'analisi offline di dischi **Windows, Linux e macOS**. **Rileva automaticamente il sistema operativo di ogni volume montato** e propone il set di moduli pertinente, generando report HTML dettagliati per investigazioni rapide ed efficaci.

- Volumi **Windows** → i 39 moduli per artefatti Windows.
- Volumi **Linux** → 16 moduli dedicati (log di sistema, journal, login, history shell/AI, browser, account, persistenza, SSH, rete, pacchetti, cestino, timeline filesystem, auditd, container, master timeline).
- Volumi **macOS** → 13 moduli dedicati (log, account dslocal, persistenza, login items/BTM, quarantine, TCC, KnowledgeC, browser, history shell/AI, recenti, FSEvents, Spotlight, master timeline).

Tutto rigorosamente **offline**, sul filesystem montato in sola lettura.

---

## 📋 Caratteristiche Principali

### Rilevamento automatico dell'OS per volume

All'avvio FIUTO elenca i volumi montati con un badge OS (`[Windows]` / `[Linux]` / `[macOS]`) e, una volta selezionato il volume, mostra solo i moduli applicabili a quel sistema operativo.

### Copertura Completa di Artefatti Windows

Sui volumi Windows FIUTO raccoglie e analizza:

- **Cronologie di esecuzione** (Prefetch, AmCache, ShimCache, BAM)
- **Artefatti di persistenza** (Run Keys, Scheduled Tasks, Services, WMI)
- **Cronologia di navigazione** (Browser history, URL MRU, TypedPaths)
- **Attività utente** (UserAssist, ShellBags, LNK files, JumpLists)
- **Artefatti di rete** (Interfacce, cache DNS, profili WLAN/VPN)
- **Log eventi di Windows** (Security, System, PowerShell, RDP)
- **Memoria virtuale** (Pagefile, Hibernation, SRUM)
- **Dispositivi rimovibili** (Cronologia USB, storia delle connessioni)
- **Active Directory** (NTDS.dit, hash domain, PAD offline analysis)
- **Cronologia chat con assistenti AI** (ChatGPT, Copilot, Claude, Cursor, Gemini, Codex, Windsurf, Continue) — *novità v1.2*
- **E molto altro...**

### Copertura Artefatti Linux — *novità v2.0*

Sui volumi Linux: log di sistema (`/var/log`) e **systemd journal**; **login** (`wtmp`/`btmp`/`lastlog`); **history shell** (bash/zsh/sh + python/mysql/psql) e **AI CLI**; **browser** (Firefox/Chrome/Chromium/Brave, anche snap/flatpak); **account** (`passwd`/`shadow`/`group`/`sudoers`); **persistenza** (cron, systemd, autostart, `rc.local`, init shell, `ld.so.preload`); **SSH** (`authorized_keys`, `known_hosts`, `sshd_config`); **rete** (`hosts`, PSK Wi-Fi NetworkManager, netplan, iptables/nftables); **pacchetti** (dpkg/rpm/apt history/snap); **cestino e recenti** (`Trash` con data cancellazione, `recently-used.xbel`); **timeline MAC-time**.

### Copertura Artefatti macOS — *novità v2.0*

Sui volumi macOS: **log** (`system.log`, `install.log`, ASL — i unified log `.tracev3` sono fuori scope); **account** dslocal (`*.plist`); **persistenza** (`LaunchAgents`/`LaunchDaemons` di sistema e per-utente, cron — plist binari resi leggibili); **Login Items/BTM**; **quarantine** (`QuarantineEventsV2`, URL+data download); **TCC** (permessi privacy); **KnowledgeC** (uso app); **browser** (Safari/Chrome/Firefox); **history shell & AI**; **recenti** (`SFL`/`SFL2`, `.Trash`).

### Modalità di Esecuzione Flessibili

```bash
./fiuto.sh                          # Menu interattivo (rileva volumi montati + OS)
./fiuto.sh /mnt/disk                # Specifica la root di un volume (Windows/Linux/macOS)
./fiuto.sh /mnt/disk --all          # Esegui tutti i moduli dell'OS rilevato
./fiuto.sh /mnt/disk --module 3     # Esegui un modulo specifico (numerazione per OS)
./fiuto.sh /mnt/disk --all --jsonl  # Esporta anche in JSONL (schema Timesketch)
./fiuto.sh /mnt/disk --all --since 2026-03-01 --until 2026-03-08  # Limita a una finestra
./fiuto.sh /mnt/disk --all --no-log-replay   # Non applicare i .LOG1/.LOG2 del registro
```

### Output Professionale

- Report HTML interattivi e navigabili
- Esportazione CSV dei dati completi per ogni modulo con un click
- Timeline cronologiche aggregate
- Log dettagliati per audit e tracciamento
- Riepilogati batch con metriche di esecuzione

---

## 🛠️ Dipendenze

### Requisiti di Sistema

- **Bash 4.0+**
- **Python 3.9+** (con module di compatibilità multiple versioni)
- **Linux (anche WSL) o macOS** (per montare/analizzare dischi Windows offline)

### Moduli Python Richiesti

```bash
pip install regipy          # Parsing hive di registro offline + replay .LOG1/.LOG2
pip install python-evtx     # Lettura dei file .evtx
```

### Consigliati per la copertura Windows completa

Queste librerie abilitano l'output completo di alcuni moduli Windows. Sono **opzionali**: se mancano, il modulo interessato ripiega su un parser parziale invece di fallire.

```bash
pip install impacket          # Hash SAM / NTDS.dit / AD (moduli 20, 24, 38)
pip install libesedb-python   # Parsing database ESE — SRUM (modulo 16)
pip install mft               # MFT timeline (modulo 21)
pip install python-snappy     # Decompressione LevelDB ChatGPT (modulo 39) — oppure: pip install cramjam
pip install libpff-python     # Parsing PST/OST di Outlook (modulo 44)
pip install libesedb-python   # WebCacheV01 e Windows.edb (moduli 49, 50)
```

> **I moduli Linux e macOS non richiedono pacchetti aggiuntivi** — usano solo la libreria standard di Python (`sqlite3`, `plistlib`, …). Facoltativamente, `journalctl` (per il journal systemd) e la CLI `rpm` (per il dump RPM offline) migliorano la copertura se presenti.

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

4. **Monta il disco target in sola lettura (esempi)**

   ```bash
   # Windows (NTFS) su Linux con ntfs-3g
   sudo mount -t ntfs-3g -o ro /dev/sda3 /mnt/disk

   # Linux (ext4)
   sudo mount -o ro /dev/sda2 /mnt/disk

   # macOS (APFS/HFS+) in sola lettura
   sudo mount -o ro /dev/sdb2 /mnt/disk
   ```

   FIUTO rileva automaticamente l'OS del volume montato: non serve indicarglielo.

---

## 🚀 Guida Rapida

### Analisi Interattiva

```bash
./fiuto.sh /mnt/disk
```

Lo script rileva l'OS del volume e presenta un menu numerato con i moduli pertinenti (52 per Windows, 23 per Linux, 21 per macOS). Seleziona il numero del modulo o digita `--all` per eseguirli tutti.

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

### Output Generato

I report generati si trovano in:

```text
fiuto_reports/
├── 001_PSReadLine_History.html
├── 002_IFEO.html
├── 012_Event_Log.html
├── ...
└── 38_PAD_Offline.html
```

---

## 📊 I 54 Moduli di Analisi Windows

| #  | Nome Modulo                   | Artefatto Windows                   | Utilizzo                                                              |
| -- | ----------------------------- | ----------------------------------- | --------------------------------------------------------------------- |
| 1  | PowerShell PSReadLine History | PSReadLine history file             | Recupera comandi eseguiti e stringhe sensibili                        |
| 2  | IFEO                          | Registry SOFTWARE                   | Rilevare manipolazioni di eseguibili di sistema                       |
| 3  | BAM                           | Registry SYSTEM                     | Timeline di esecuzioni in background                                  |
| 4  | RDP Cache                     | Terminal Server Client Cache        | Ricostruire sessioni RDP                                              |
| 5  | Run Keys & Persistenza        | Registry NTUSER.DAT, SOFTWARE       | Rilevare backdoor e malware                                           |
| 6  | Prefetch                      | Windows/Prefetch/*.pf               | Cronologia esecuzione programmi                                       |
| 7  | Notepad TabState              | AppData Packages                    | Recuperare testo non salvato                                          |
| 8  | Scheduled Tasks               | Windows/System32/Tasks              | Trovar dropper temporizzati                                           |
| 9  | USB/Dispositivi               | Registry SYSTEM                     | Storico inserimento supporti USB                                      |
| 10 | LNK & JumpList                | AppData Recent                      | File usati recentemente                                               |
| 11 | Services                      | Registry SYSTEM                     | Servizi anomali o custom                                              |
| 12 | Event Log                     | Windows/System32/winevt/Logs/*.evtx | Login, RDP, attività sospette                                        |
| 13 | Amcache + Shimcache           | Amcache.hve, Registry               | Cronologia esecuzione con hash                                        |
| 14 | Recycle Bin                   | $Recycle.Bin                        | Recuperare file cancellati                                            |
| 15 | WMI Subscriptions             | OBJECTS.DATA                        | Persistenza fileless                                                  |
| 16 | SRUM                          | SRUDB.dat                           | Consumo risorse per app                                               |
| 17 | Browser History               | SQLite History/places               | Navigazione web                                                       |
| 18 | UserAssist + MRU              | Registry NTUSER.DAT                 | App avviate via GUI                                                   |
| 19 | ShellBags                     | Registry NTUSER.DAT/UsrClass        | Navigazione cartelle                                                  |
| 20 | SAM Hash                      | Windows/System32/config/SAM         | NTLM hash account locali                                              |
| 21 | MFT Timeline                  | $MFT                                | Granular NTFS changes                                                 |
| 22 | OpenSaveMRU                   | Registry NTUSER.DAT                 | Cronologia file aperti/salvati                                        |
| 23 | USN Journal                   | $UsnJrnl:$J                       | Attività file system massicce                                        |
| 24 | NTDS.dit                      | ntds.dit, Registry SYSTEM           | Hash utenti dominio                                                   |
| 25 | Hibernation/Pagefile          | hiberfil.sys, pagefile.sys          | Memory dump analysis                                                  |
| 26 | WER Files                     | WER Report Archive                  | Crash e errori di sistema                                             |
| 27 | Credential Manager            | Credentials DPAPI                   | Credenziali di rete salvate                                           |
| 28 | WLAN & VPN Profiles           | Wlansvc Profiles                    | SSID e profili VPN                                                    |
| 29 | AppX / UWP Packages           | AppData/Local/Packages              | App Modern installate                                                 |
| 30 | Browser Downloads & Logins    | SQLite databases                    | Download e credenziali                                                |
| 31 | Clipboard History             | ActivitiesCache.db                  | Cronologia appunti                                                    |
| 32 | Office MRU & Recent           | NTUSER.DAT, AppData                 | Documenti aperti recenti                                              |
| 33 | Defender Quarantine           | Quarantine folder                   | Minacce identificate                                                  |
| 34 | PowerShell Script Block       | Event Log 4104                      | Script eseguiti                                                       |
| 35 | JumpLists                     | Recent Destinations                 | Interazione utente app                                                |
| 36 | Network Artifacts             | Registry SYSTEM                     | DNS, interfacce, reti                                                 |
| 37 | Master Timeline               | (Aggregato)                         | Timeline cross-artefatto                                              |
| 38 | PAD Offline                   | NTDS.dit                            | Active Directory avanzato                                             |
| 39 | AI Chat / Query History       | LevelDB / JSON / SQLite             | Recupera conversazioni con assistenti AI (ChatGPT, Copilot, Claude…) |
| 40 | SetupAPI Device Log           | Windows/INF/setupapi.dev.log        | **Prima** installazione dei dispositivi USB                           |
| 41 | PowerShell Transcript         | PowerShell_transcript*.txt          | Sessioni complete: comandi **e** output                               |
| 42 | LSA Secrets & DCC2            | Registry SECURITY                   | Password account di servizio, credenziali di dominio in cache         |
| 43 | Volume Shadow Copies          | System Volume Information           | Snapshot precedenti del volume, analisi differenziale                 |
| 44 | Outlook PST / OST             | *.pst, *.ost (libpff)               | Posta locale, allegati, item cancellati                               |
| 45 | Cloud Sync                    | OneDrive, Dropbox, Google Drive     | File sincronizzati — la via di esfiltrazione moderna                  |
| 46 | BITS Jobs                     | qmgr.db                             | Download in background usati come LOLBin (T1197)                      |
| 47 | Thumbcache                    | thumbcache_*.db                     | Miniature di file **cancellati**, estratte per carving                |
| 48 | Chat Desktop                  | Slack / Teams / Discord (LevelDB)   | Social engineering interno, file condivisi in chat private            |
| 49 | WebCacheV01                   | WebCacheV01.dat (ESE)               | IE/Edge Legacy **e tutto cio' che usa WinINET**                       |
| 50 | Windows Search Index          | Windows.edb (ESE)                   | Percorsi ed estratti di contenuto di file **cancellati**              |
| 51 | SQLite Recovery *(cross-OS)*  | freelist, spazio non allocato       | Contenuto di record **cancellati** ancora sul disco                   |
| 52 | EFI System Partition *(cross-OS)* | ESP                             | Bootkit: codice eseguito **prima** del sistema e di ogni EDR          |
| 53 | YARA *(cross-OS)*                 | YAR                             | Regole esterne su un ambito limitato e dichiarato (`--yara`)          |
| 54 | Sigma                             | SIG                             | Detection della comunita' sugli EVTX, sottoinsieme dichiarato (`--sigma`) |

---

## 🐧 I 24 Moduli di Analisi Linux

| #  | Nome Modulo         | Artefatto Linux                                                    | Utilizzo                                                        |
| -- | ------------------- | ------------------------------------------------------------------ | --------------------------------------------------------------- |
| 1  | System Logs         | `/var/log/{syslog,messages,auth.log,secure,kern.log,…}`         | Eventi sistema/auth, errori, sudo                               |
| 2  | systemd Journal     | `/var/log/journal/*.journal`                                     | Journal strutturato (via `journalctl` o fallback `strings`) |
| 3  | Login History       | `wtmp` / `btmp` / `lastlog`                                  | Login riusciti/falliti con IP sorgente                          |
| 4  | Shell History       | `.bash_history`, `.zsh_history`, `.python_history`, …       | Comandi eseguiti e stringhe sensibili                           |
| 5  | AI CLI History      | `~/.claude`, aider, aichat, ollama, copilot                      | Conversazioni CLI con assistenti AI                             |
| 6  | Browser History     | Firefox `places.sqlite`, Chrome/Chromium/Brave/Edge              | Navigazione web (anche path snap/flatpak)                       |
| 7  | User Accounts       | `passwd` / `shadow` / `group` / `sudoers`                  | Account, UID 0, password vuote/bloccate                         |
| 8  | Persistence         | cron, unit/timer systemd, autostart,`rc.local`, init shell       | Backdoor e meccanismi di avvio automatico                       |
| 9  | SSH Artifacts       | `authorized_keys`, `known_hosts`, `sshd_config`              | Fiducia e configurazione accesso remoto                         |
| 10 | Network Config      | `hosts`, NetworkManager (PSK Wi-Fi), netplan, iptables/nftables  | Configurazione di rete e chiavi Wi-Fi                           |
| 11 | Installed Packages  | dpkg, rpm, apt history, snap                                       | Timeline installazioni / pacchetti sospetti                     |
| 12 | Trash & Recent      | `~/.local/share/Trash` (+`.trashinfo`), `recently-used.xbel` | File cancellati con data di cancellazione                       |
| 13 | Filesystem Timeline | MAC times aree sensibili (`find`/`stat`)                       | Timeline cronologica cross-area                                 |
| 14 | auditd              | `/var/log/audit/audit.log*`                                      | Syscall, autenticazioni, EXECVE, violazioni di policy           |
| 15 | Container           | `/var/lib/docker`, `/var/lib/containers`                        | Inventario Docker/Podman + indicatori di fuga dal container     |
| 16 | Master Timeline     | (Aggregato)                                                        | Timeline cronologica cross-modulo (con `--all` gira per ultima) |
| 17 | PAM                 | `/etc/pam.d`, `security/*.so`                                    | Backdoor di autenticazione, moduli non pacchettizzati           |
| 18 | Kernel Modules      | `lib/modules`, `modprobe.d`, initramfs                           | Rootkit LKM, direttive `install`, `.ko` non pacchettizzati      |
| 19 | Web Server Logs     | nginx / apache access+error                                        | Webshell, traversal, SQLi — ordinati per stato HTTP             |
| 20 | Cloud Credentials   | `~/.aws`, `~/.kube`, `~/.docker`, `~/.ssh`                     | Chiavi che danno accesso all'infrastruttura                     |
| 21 | SUID & Capabilities | scansione filesystem                                               | Superficie di privilege escalation, SUID inattesi               |
| 22 | SQLite Recovery *(cross-OS)* | freelist, spazio non allocato                             | Contenuto di record **cancellati** ancora sul disco             |
| 23 | EFI System Partition *(cross-OS)* | ESP                                                  | Bootkit: codice eseguito **prima** del sistema                  |
| 24 | YARA *(cross-OS)*                 | YAR                                                  | Regole esterne, ambito dichiarato (`--yara`)                    |

---

## 🍎 I 22 Moduli di Analisi macOS

| #  | Nome Modulo            | Artefatto macOS                                                   | Utilizzo                                                        |
| -- | ---------------------- | ----------------------------------------------------------------- | --------------------------------------------------------------- |
| 1  | System Logs            | `system.log`, `install.log`, ASL                              | Eventi sistema/installazioni (*`.tracev3` fuori scope*)     |
| 2  | User Accounts          | dslocal `*.plist`                                               | Account, UID, shell, presenza hash password                     |
| 3  | Persistence            | `LaunchAgents` / `LaunchDaemons` (sistema e per-utente), cron | Persistenza (plist binari resi leggibili)                       |
| 4  | Login Items (BTM)      | `backgrounditems.btm`                                           | Elementi di background/login                                    |
| 5  | Quarantine / Downloads | `QuarantineEventsV2`                                            | Provenienza download (URL + data)                               |
| 6  | TCC Privacy            | `TCC.db`                                                        | Permessi privacy camera/mic/disco                               |
| 7  | KnowledgeC             | `knowledgeC.db`                                                 | Uso app / attività dispositivo                                 |
| 8  | Browser History        | Safari `History.db`, Chrome, Firefox                            | Navigazione web                                                 |
| 9  | Shell & AI History     | `.zsh_history`, `.bash_history`, AI CLI                       | Comandi eseguiti e conversazioni AI                             |
| 10 | Recent Items           | `SFL`/`SFL2`, `~/.Trash`                                    | App/documenti/server recenti e cestino                          |
| 11 | FSEvents               | `/.fseventsd`                                                   | Storico modifiche al filesystem (creazioni, rinomine, cancellazioni) |
| 12 | Spotlight              | `.Spotlight-V100/store.db`                                      | Provenienza download e nomi di file cancellati (euristico)      |
| 13 | Master Timeline        | (Aggregato)                                                       | Timeline cronologica cross-modulo (con `--all` gira per ultima) |
| 14 | Messages               | `chat.db`                                                       | iMessage/SMS: smishing, codici 2FA, comunicazioni fuori canale  |
| 15 | Cookie & Download      | `Cookies.binarycookies`, `Downloads.plist`                    | Servizi a cui si era autenticati; provenienza download          |
| 16 | XProtect / Gatekeeper  | XProtect, SystemPolicy, KextPolicy                                | Versione firme; autorizzazioni concesse a mano                  |
| 17 | Applications           | `/Applications`, home utenti                                    | Inventario, presenza firma, posizioni non standard              |
| 18 | Time Machine / Snapshot| `com.apple.TimeMachine.plist`, `Backups.backupdb`             | Versioni precedenti dei file; l'assenza come indicatore         |
| 19 | Unified Logs           | `*.tracev3`                                                     | Decompressione chunk LZ4 + estrazione stringhe (**parziale**)   |
| 20 | SQLite Recovery *(cross-OS)* | freelist, spazio non allocato                             | Contenuto di record **cancellati** ancora sul disco             |
| 21 | EFI System Partition *(cross-OS)* | ESP                                                  | Bootkit: codice eseguito **prima** del sistema                  |
| 22 | YARA *(cross-OS)*                 | YAR                                                  | Regole esterne, ambito dichiarato (`--yara`)                    |

---

## 🔍 Casi d'Uso

### Indagini Malware

Utilizza i moduli di persistenza (Run Keys, Services, WMI) per trovare backdoor, e i log di PowerShell per tracciare esecuzione di payload offuscati.

### Incident Response (IR)

Combina BAM, Prefetch e Event Logs per costruire una timeline precisa di cosa è stato eseguito e quando, aiutando a identificare il paziente zero.

### Conformità & Audit

Estrai l'attività di rete (SRUM, DNS) e gli accessi RDP (RDP Cache, Event Logs) per dimostrare chi ha accesso a quali sistemi.

### Forense Ransomware

Analizza USN Journal, MFT Timeline, Recycle Bin e Browser Downloads per tracciare la diffusione e le origini dell'infezione.

### Caccia alle Minacce (Threat Hunting)

Usa la cronologia USB, i profili WLAN/VPN e la navigazione web per scoprire esfiltrazione di dati o comunicazioni con indirizzi IP sospetti.

---

## ⚙️ Opzioni Avanzate

### Carica Lista IoC

```bash
./fiuto.sh /mnt/windows --ioc /path/to/ioc_list.txt
```

Scansiona gli artefatti per trovare match con indicatori di compromissione.

### Replay dei transaction log del registro (Windows)

Windows non scrive subito le modifiche nell'hive primario: le accoda nei
transaction log (`.LOG1` / `.LOG2`) e le consolida solo a uno smontaggio
pulito. Un hive acquisito da una macchina spenta a caldo, da un'immagine o da
uno snapshot è quindi quasi sempre *dirty*: le scritture più recenti — spesso
proprio quelle dell'attaccante — esistono **solo** nei log.

FIUTO li riapplica di default su una **copia temporanea**; il volume di
evidenza non viene mai toccato. Su un hive di test reale questo ha recuperato
2.456 chiavi e 3.905 valori in più, invisibili al parsing dell'hive grezzo.

```bash
./fiuto.sh /mnt/windows --all --no-log-replay   # disattiva (sconsigliato)
```

Richiede `regipy`. Se manca, FIUTO avvisa una volta e ripiega sull'hive
originale invece di fallire.

### Catena di custodia

Ogni sessione scrive `evidence_manifest.json` nella cartella dei report.
Registra su cosa si è basata l'analisi, così i report restano verificabili:

- versione del tool, comando esatto, operatore e host di analisi;
- inizio e fine sessione in UTC;
- volume analizzato, OS rilevato e hostname letto dagli artefatti;
- **ogni file di evidenza consultato**, con dimensione, data e SHA-256;
- **ogni report generato**, con il suo SHA-256;
- l'esito del replay dei transaction log del registro.

La copertura è automatica: i file vengono annotati dentro le funzioni che
tutti i moduli attraversano per arrivare a un artefatto (`ci_find_file`,
`query_sqlite`, `read_plist`, `recover_hive`), quindi i moduli aggiunti in
futuro sono coperti senza doverli toccare.

```bash
./fiuto.sh /mnt/disk --all --no-hash        # annota i file senza calcolare gli hash
./fiuto.sh /mnt/disk --all --hash-limit 256 # niente hash oltre i 256 MB
./fiuto.sh /mnt/disk --all --no-custody     # disattiva del tutto il manifesto
```

I file oltre la soglia (1 GB di default — `pagefile.sys`, `$MFT`,
`Windows.edb`) restano elencati, con il motivo per cui manca l'hash invece di
ometterlo in silenzio.

### Copie oscurate (`--redact`, `--defang`)

I report contengono materiale che deve restare nello studio: hash NTLM da SAM e
NTDS, PSK Wi-Fi, token OAuth, chiavi private, segreti LSA. Finché l'unica scelta
è "mandare il report intero o non mandarlo", quel materiale finisce in una mail
a un cliente o nell'allegato di un ticket.

```bash
./fiuto.sh /mnt/disk --all --redact     # report.redacted.html accanto a ogni report
./fiuto.sh /mnt/disk --all --defang     # rende inerti anche URL e IP
```

**L'originale non viene mai toccato.** Accanto nasce un `report.redacted.html`.
Oscurare sul posto distruggerebbe evidenza per una necessità di comunicazione:
un baratto che non spetta al tool fare al posto tuo. Il manifesto di custodia
elenca entrambi, con ruoli distinti.

**Si oscura per contesto, non per forma.** Un SHA-256 e un hash NTLM sono
entrambe stringhe esadecimali: il primo è l'impronta di integrità di un reperto
e rimuoverlo spezzerebbe la catena di custodia proprio nel file che stai per
condividere, il secondo è una credenziale. A distinguerli è l'etichetta accanto,
quindi le regole guardano il campo e non il valore. Le etichette restano —
sapere *che* c'era una password è un dato dell'analisi — e spariscono solo i
valori.

Ogni copia dichiara in testa di non essere l'originale, ed elenca quante
occorrenze ha rimosso ciascuna regola. Un file che sembra un report ma non lo è
del tutto, senza dirlo, è peggio di nessun file.

### Motore IoC (`--ioc`)

```bash
./fiuto.sh /mnt/disk --all --ioc indicatori.txt     # lista piatta
./fiuto.sh /mnt/disk --all --ioc bundle.json        # STIX 2.x o export MISP
```

Gli indicatori vengono **tipizzati al caricamento** e confrontati con i confini
che il loro tipo merita. Il match a sottostringa, che è quello che FIUTO faceva
fino alla 2.2, sbagliava in due direzioni opposte:

- **falsi positivi** — `10.0.0.5` corrisponde dentro `110.0.0.55`, `evil.com`
  dentro `notevil.com.au`. In un report da centomila righe non è rumore
  fastidioso: è rumore che nasconde i match veri;
- **falsi negativi** — gli indicatori arrivano quasi sempre defanged
  (`hxxp://`, `1[.]2[.]3[.]4`, `evil[.]com`), perché è così che si scrivono in
  una mail o in un bollettino. Cercati alla lettera non corrispondono a niente,
  e l'analista conclude che il disco è pulito.

Sono coperti entrambi: l'input defanged viene normalizzato e ogni tipo ha i suoi
confini. Un dominio corrisponde nei propri sottodomini (`mail.evil.com` per
`evil.com`) ma non in un suffisso più lungo (`evil.com.altro.net`). Gli hash
sono token case-insensitive. Un valore fra slash (`/inv[o0]ke/`) è una regex.

Il tipo si deduce dal valore; `tipo:valore` forza la mano dove l'euristica non
può vincere (`file:payload.com` — nome di file o dominio?). Il conteggio per
tipo viene stampato al caricamento, ed è ciò che rende evidente un file letto
male: quaranta `literal` al posto di quaranta hash sono un errore di formato, e
senza quella riga lo si scoprirebbe solo dall'assenza di match.

`--ioc` legge anche bundle STIX 2.x ed export MISP, oggetti annidati compresi.
Un JSON che non è né l'uno né l'altro viene **rifiutato**, non letto riga per
riga: caricare le graffe come indicatore letterale lo farebbe corrispondere
ovunque.

### Esecuzione parallela (`--jobs`)

```bash
./fiuto.sh /mnt/disk --all --jobs 4
```

Opt-in, non il default. I moduli sono indipendenti — leggono file diversi e
scrivono in cartelle diverse — ma su un volume montato da disco meccanico o via
rete N lettori vanno *più piano* di uno. Il collo di bottiglia qui è quasi
sempre l'I/O e non la CPU, e qual è lo sa solo chi sta analizzando.

La garanzia è che **l'esito non dipende dal grado di parallelismo**: stessi
report, stesso ordine nel riepilogo. Un risultato che cambia con `--jobs` è un
risultato di cui non ci si può fidare, e in ambito peritale è peggio di un tool
lento. Ci sono test dedicati esattamente a questo.

I moduli marcati `defer` — la Master Timeline, che aggrega gli altri — girano
per ultimi, da soli, quando il pool si è svuotato *e* gli esiti degli altri sono
stati raccolti. Leggono l'elenco dei report generati, quindi eseguirli prima
produrrebbe una timeline vuota senza dirlo.

Il recupero degli hive di registro prende un lock: con più moduli che chiedono
lo stesso hive insieme, il secondo leggerebbe la copia ricostruita mentre il
primo la sta ancora scrivendo. Un hive troncato non dà errore — dà risultati
parziali, che è peggio.

In parallelo l'interruzione con ESC non è disponibile: intercettarla richiede il
controllo esclusivo del terminale. Viene dichiarato all'avvio, invece di lasciare
che il tasto smetta di funzionare senza spiegazione.

### YARA — regole esterne (`--yara`)

YARA è il formato con cui l'industria distribuisce le firme: un feed di threat
intelligence, l'IR di un vendor o il CERT nazionale mandano file `.yar`. FIUTO
li applica.

```bash
./fiuto.sh /mnt/disk --all --yara /regole/              # un file o una directory
./fiuto.sh /mnt/disk --module 24 --yara r.yar --yara-scan /mnt/disk/Users
./fiuto.sh /mnt/disk --all --yara /regole/ --yara-max-mb 256
```

Richiede `yara-python`. Se manca non c'è alcun ripiego — una scansione YARA
senza YARA non è una scansione, e il modulo lo dice invece di fingere.

**Il punto delicato è l'ambito.** Scansionare un volume da un terabyte file per
file non è praticabile su una workstation forense, quindi il modulo copre un
insieme limitato di posizioni: quelle scrivibili senza privilegi, dove il codice
che nessun gestore di pacchetti ha installato finisce quasi sempre (`Temp`,
`AppData`, `ProgramData`, `/tmp`, `/dev/shm`, `~`, `/var/www`, i LaunchAgents,
la ESP…). Con `--yara-scan` quell'insieme viene sostituito da un percorso
indicato da te.

Il report **elenca esattamente cosa è stato scansionato**, posizione per
posizione e con i conteggi, più cosa è stato saltato e perché: i file oltre il
tetto per file, quelli illeggibili, e se il tetto complessivo è stato raggiunto
— nel qual caso la scansione è marcata PARZIALE. I symlink non vengono mai
seguiti: su un volume montato porterebbero fuori dall'evidenza, fino al file
system della workstation di analisi. Senza questa contabilità un "nessun match"
si leggerebbe come "il disco è pulito"; con essa si legge per quello che è.

Un file di regole che non compila viene segnalato e saltato; gli altri girano.

### Sigma — detection della comunità (`--sigma`, Windows)

Sigma è il formato in cui la comunità pubblica le detection: SigmaHQ, i CERT e
i vendor distribuiscono migliaia di regole YAML. Applicarle agli EVTX di un
disco acquisito è quello che fanno Chainsaw e Hayabusa, ed è il passo che
trasforma una raccolta di log in un triage.

```bash
./fiuto.sh /mnt/windows --all --sigma /sigma/rules/windows/
```

Richiede `pyyaml` e `python-evtx`.

**Il sottoinsieme supportato è dichiarato, non implicito.** Sigma è un
linguaggio ampio — modificatori base64, CIDR, condizioni con parentesi,
aggregazioni temporali. Implementarne una parte e far finta di supportarlo
tutto significherebbe che una regola mai valutata compare come una regola che
non è scattata: un falso negativo silenzioso, il difetto peggiore che una
detection possa avere. Le regole che il motore non sa valutare vengono
**contate ed elencate nel report con il motivo**, raggruppate per causa.

| | |
|---|---|
| Selezioni | mappe campo/valore, liste di valori (OR), liste di mappe (OR) |
| Modificatori | `contains`, `startswith`, `endswith`, `re`, `all`, `cased` |
| Condizioni | `sel`, `a and b`, `a or b`, `a and not b`, `not a`, `1 of x*`, `all of x*`, `1 of them`, `all of them` |
| Valori | `null` = campo assente o vuoto |
| **Scartate** | condizioni con parentesi o aggregazioni, `base64offset`/`utf16`/`wide`/`cidr`/`gt`/`lt`, logsource non mappabile |

Una regola la cui `logsource` non è mappabile a un canale EVTX presente sul
volume **non** viene fatta girare a tappeto su tutti i log come ripiego:
verrebbe valutata su campi che quel canale non ha, e il "non scattata" sarebbe
privo di significato. Viene scartata ed elencata.

Il report riporta anche quanti record sono stati letti per canale, e marca la
valutazione PARZIALE se il tetto è stato raggiunto.

### Executive summary (riepilogo di sessione)

Novanta report HTML sono un dump, non un'analisi. Al termine di `--all` — e da
`[S]` nel menu, sui report prodotti fino a quel momento — FIUTO scrive accanto
a loro un `executive_summary.html` e un `findings.json` leggibile da programma.

Risponde a due domande: **da dove comincio** e **cosa è successo insieme a
cos'altro**. Non risponde a "la macchina è compromessa": quella è una
conclusione dell'analista, e la pagina lo dice in testa.

I riscontri vengono da tre sostrati, in ordine di affidabilità decrescente:

1. **le righe che i moduli hanno già marcato** — ogni modulo passa le proprie
   parole chiave al renderer dei log, e una riga marcata è un giudizio di chi
   conosce l'artefatto;
2. **le corrispondenze con gli IoC**, se è stato usato `--ioc`;
3. **una tabella di regole esplicite**, deliberatamente piccola, ognuna con la
   sua tecnica MITRE ATT&CK e la spiegazione del perché quel dato conta.

Le regole girano **solo sul dato** — righe di tabella e righe di log — mai sui
cartigli esplicativi che FIUTO stesso scrive nei report. Un motore di detection
che leggesse l'intera pagina scatterebbe sulla propria prosa: c'è un test che
lo presidia.

La **correlazione cross-modulo** è ciò che nessun report singolo può mostrare.
Gli eventi di tutti i report vengono raggruppati in finestre di 30 minuti;
quando una finestra contiene moduli che corrispondono a uno scenario noto, il
riepilogo enuncia l'ipotesi — per esempio un supporto rimovibile collegato
mentre venivano aperti file LNK e il journal USN registrava modifiche. Sono
**ipotesi da verificare**, e la pagina le presenta come tali: la coincidenza
temporale non è un nesso di causa.

Il **punteggio di priorità** è la somma dei pesi dei riscontri (critico 40, alto
15, medio 5, basso 1) limitata a 100. Ordina la coda di lavoro; non misura la
compromissione, e il report stampa la formula perché il numero non venga preso
per più di quello che è. Allo stesso modo, l'assenza di riscontri non è un
attestato di pulizia: significa che un insieme piccolo e conservativo di regole
non ha trovato nulla.

### Finestra temporale (`--since` / `--until`)

Su un disco grande un report può contenere centinaia di migliaia di righe che
coprono anni, mentre l'incidente sta in tre giorni.

```bash
./fiuto.sh /mnt/disk --all --since 2026-03-01 --until 2026-03-08
./fiuto.sh /mnt/disk --all --since -7d      # ultimi 7 giorni
./fiuto.sh /mnt/disk --all --since "2026-03-01 14:30"
```

Formati ammessi: `YYYY-MM-DD`, `YYYY-MM-DD HH:MM[:SS]`, le stesse con `T` e/o
`Z` finale, e le forme relative `-7d` / `-36h` / `-90m`. Un limite che non si
riesce a interpretare **ferma l'esecuzione** invece di essere ignorato:
lasciar cadere il filtro in silenzio produrrebbe un report che dichiara una
copertura che non ha.

Il filtro vale sia per le tabelle sia per i blocchi di log `<pre>`, e per
l'export JSONL, così le due viste di uno stesso modulo non possono
contraddirsi. Due regole gli impediscono di cancellare evidenza:

- **le righe prive di data vengono sempre mantenute** — non sono valutabili;
- una riga con **più** date resta se *almeno una* cade nella finestra (un file
  creato prima della finestra ma usato dentro non sparisce).

Ogni tabella e ogni blocco di log filtrato dichiara quante righe ha nascosto, e
il manifesto registra la finestra insieme al totale: un report filtrato non può
essere scambiato per un report vuoto. Nei blocchi `<pre>` i numeri di riga
restano quelli del file originale, quindi i salti di numerazione sono il segnale
visibile che qualcosa è stato tolto.

**Sui fusi orari.** FIUTO rileva il fuso del volume (`/etc/timezone`,
`/etc/localtime`, `SYSTEM\Control\TimeZoneInformation`) e lo dichiara — all'avvio,
nel manifesto e in ogni record JSONL — ma **non converte niente**. Gli artefatti
di uno stesso volume mescolano UTC (registro, log eventi Windows) e ora locale
(syslog, shell history): una conversione applicata alla cieca sposterebbe gli
eventi di ore, che è molto peggio di date dichiarate ambigue. Il confronto usa
quindi le date *così come compaiono nell'artefatto*: su finestre di giorni la
differenza è irrilevante, su finestre di ore vanno allargate dell'offset del
volume.

### Export JSONL (Timesketch / plaso)

```bash
./fiuto.sh /mnt/disk --all --jsonl
```

Accanto a ogni report HTML scrive un `report.jsonl` e, per l'intera sessione,
un unico `fiuto_timeline.jsonl` con i campi attesi da Timesketch (`datetime`,
`timestamp_desc`, `message`) arricchiti con volume, hostname e modulo. Gli
eventi il cui anno è stato dedotto da un timestamp in stile syslog vengono
marcati con `year_inferred: true` invece di essere presentati come certi.

---

## 📄 Output Report

Ogni modulo genera un report HTML con:

- **Tabella dati** strutturata e ordinabile
- **Timeline** con timestamp UTC e valori FILETIME decodificati
- **Evidenziazione** di elementi sospetti (password, token, comandi offuscati)
- **Metadata** (hash, percorsi assoluti, account coinvolti)
- **Note forensiche** su come interpretare i risultati

### Esempio Report

```text
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

## 🐛 Risoluzione dei Problemi

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

## 📝 Log di Sessione

Lo script crea un log dettagliato di ogni sessione in:

```text
fiuto_reports/session_YYYY-MM-DD_HH-MM-SS.log
```

Utile per debugging e audit trail:

```bash
tail -f fiuto_reports/session_*.log
```

---

## 🧪 Sviluppo e test

```bash
bats tests/                              # suite di test (bats-core)
shellcheck -S warning -x fiuto.sh        # lint
python3 tests/lint_embedded_python.py fiuto.sh   # compila i parser incorporati
```

La CI ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)) gira a ogni push:
sintassi bash, ShellCheck, suite bats e compilazione dei ~99 parser Python
incorporati come heredoc, sia su Python 3.9 sia su 3.12.

Quest'ultimo job non è un ornamento: `bash -n` tratta gli heredoc come testo
opaco, quindi un errore di sintassi dentro un parser incorporato resta
invisibile finché quel modulo non gira su un disco reale. È così che è emerso un
difetto di lunga data nel modulo 38 (PAD Offline): il blocco non compilava, e il
modulo non aveva mai prodotto output.

## 🤝 Contributi

Se hai miglioramenti, segnalazioni di bug o moduli aggiuntivi:

1. Fork il repository
2. Crea un branch per la tua feature (`git checkout -b feature/nuovo-modulo`)
3. Commit i tuoi cambiamenti (`git commit -am 'Add nuovo modulo'`)
4. Push al branch (`git push origin feature/nuovo-modulo`)
5. Apri una Pull Request

---

## ⚖️ Licenza

Questo progetto è **gratuito** e distribuito sotto licenza **MIT**.

Se FIUTO ti è stato utile e vuoi supportarne lo sviluppo, una piccola donazione tramite PayPal è sempre apprezzata — ma assolutamente facoltativa! 🙏

[![Dona con PayPal](https://img.shields.io/badge/Dona-PayPal-blue.svg)](https://paypal.me/rendina)

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

## 📝 Changelog

**Date:** 2026-07-30 | **Version:** 2.2

**Twenty-four new modules.** Windows: eleven. **SetupAPI Device Log** — the only source dating a USB device's *first* connection (the USBSTOR registry keeps the last one). **PowerShell Transcript** — full sessions including command output, invisible to PSReadLine (module 1) when commands come from scripts, `-EncodedCommand` or remoting. **LSA Secrets & DCC2** — cleartext service-account passwords and cached domain credentials from the SECURITY hive, complementing SAM (module 20). **Volume Shadow Copies** — inventory and differential-analysis workflow; their *absence* is reported as an indicator, since deleting them is a standard ransomware step. **Outlook PST/OST** — first local-mail coverage, with risky-attachment and IoC flagging. **Cloud Sync** — OneDrive/Dropbox/Google Drive accounts and synced files: the modern exfiltration path, which leaves no USB artefact. **BITS Jobs** — background downloads abused as a LOLBin (T1197), flagging non-Microsoft hosts, cleartext HTTP and risky targets. **Thumbcache** — thumbnails of deleted files, recovered by signature carving and shown as a gallery beside the report. **Chat Desktop** — Slack/Teams/Discord message fragments carved from LevelDB, flagging sensitive terms. **WebCacheV01** — IE/Edge Legacy history and, more importantly, everything routed through the WinINET APIs, including non-browser code. **Windows Search Index** — indexed paths and content excerpts, which survive file deletion.

Both ESE-based modules fall back to string extraction when libesedb cannot open the database — the normal case for a file acquired from a running machine — and state in the report which parser actually produced the data.

**Linux: five new modules** — PAM (authentication backdoors), kernel modules and LKM rootkits, web server logs, cloud/development credentials, SUID/capabilities.

**macOS: six new modules** — Messages, Safari cookies and downloads, XProtect/Gatekeeper, application inventory, Time Machine/snapshots, and unified logs.

**Chain of custody.** Every session now writes an `evidence_manifest.json` recording tool version, command line, operator, analysed volume, and the SHA-256 of every evidence file consulted and every report produced. Until 2.1 hashing was scattered across a handful of modules and there was no way to answer "which files were read, in what state, and are the attached reports the ones produced then?".

**Executive summary.** At the end of `--all`, an `executive_summary.html` plus a machine-readable `findings.json`: findings ordered by severity, each with its MITRE ATT&CK technique, the data it matched and why that data matters — and cross-module time correlations, which is what no single report can show (removable media connected while LNK files were opened and USN recorded changes, and so on). Findings come mostly from what the modules already flagged themselves; the explicit rule table is deliberately small, and it runs only on the data, never on the explanatory prose FIUTO writes into its own reports. The priority score orders the work queue and prints its own formula: it does not measure compromise, and no findings is not a clean bill of health.

**Time window `--since` / `--until`.** On a large disk the incident is three days inside years of artefacts. The window applies to tables, to `<pre>` log blocks and to the JSONL export together, so no two views of a module can disagree. It never removes what it cannot judge: rows without a date are kept, and a row with several dates survives if any of them falls in the window. Every filtered block declares how many rows it hid, and the manifest records the window — a filtered report must not be mistakable for an empty one. The volume timezone is detected and declared, but **nothing is converted**: artefacts on one volume mix UTC and local time, and a blind conversion would shift events by hours.

**Two cross-OS modules**, available on all three systems. **SQLite Recovery** — nearly every modern artefact is a SQLite database, and every module reading one sees only the *live* records; a deleted record stays in the file until overwritten, in the freelist or in a page's unallocated space. This module carves it back, which is often the only place a "cleared" history still exists. **EFI System Partition** — code in the ESP runs before the OS, the kernel and any EDR, and survives a full system reinstall; the module inventories it, hashes everything and flags structural anomalies.

**On unified logs.** `.tracev3` was previously declared out of scope. The new module decompresses the LZ4 (`bv41`) chunks the format is built from and extracts the readable strings that emerge — paths, bundle ids, URLs. It does **not** reconstruct log messages: that requires interpreting the catalogue and resolving string references in `.uuidtext` and the dyld shared cache, which is a project of its own. The report states this rather than implying full support; for complete analysis use `log show --archive` on a Mac.

**Note on Linux module numbering.** Adding auditd and Containers in this release shifted the Linux Master Timeline from 14 to 16. Scripts pinning `--module 14` on Linux volumes need updating. The registry now supports a `defer` flag so the Master Timeline keeps its number while still running last under `--all`: further modules can be appended without renumbering anything again.

**Internals.** Module dispatch is now data-driven for Windows too: three parallel dispatchers were removed (a hand-written menu and two separate 39-branch `case` statements). Registry entries support bilingual labels and optional guards. Module numbering is unchanged — `--module N` keeps invoking the same modules.

**Date:** 2026-07-29 | **Version:** 2.1

**Forensic correctness.** Windows registry **transaction logs (`.LOG1`/`.LOG2`) are now replayed by default** onto a temporary copy before parsing — the evidence volume is never written to. Without this step the most recent hive writes are invisible: on a real test hive the replay recovered **+2,456 keys and +3,905 values**. Applies to the four system hives and to every per-user `NTUSER.DAT` / `UsrClass.dat`. Disable with `--no-log-replay`.

**New modules.** Linux: **auditd** (`/var/log/audit`, hex-decoded EXECVE arguments and proctitle, breakdown by record type) and **Container forensics** (offline Docker/Podman inventory with container-escape indicators: privileged, host root or Docker socket bind-mounted, `CAP_SYS_ADMIN`, host PID/network namespace). macOS: **FSEvents** (`/.fseventsd` binary parser, the macOS counterpart of the USN Journal) and **Spotlight** (heuristic extraction of download provenance).

**Interoperability.** New `--jsonl` flag: every report is also emitted as JSON Lines in the **Timesketch/plaso** schema, plus a single `fiuto_timeline.jsonl` per session, so FIUTO output can be loaded straight into a super-timeline or a SIEM.

**Quality.** First test suite (53 bats tests) and CI: bash syntax, ShellCheck at zero warnings, and compilation of every embedded Python parser on Python 3.9 and 3.12.

**Bug fixes.** Module 38 (PAD Offline) was **entirely non-functional**: four f-strings with nested quoting meant the embedded parser never compiled. `ci_find_file` was defined twice with incompatible semantics, silently breaking relative-path lookups (`recently-used.xbel`). Five `find | xargs` pipelines dropped files whose names contain spaces — which is the norm for Scheduled Tasks and Recent items.

**Date:** 2026-06-05 | **Version:** 2.0
**Multi-OS support**: FIUTO now auto-detects each mounted volume's operating system and proposes the relevant module set — Windows (39 modules, unchanged), **Linux (14 new modules)** and **macOS (11 new modules)**, all strictly offline. Linux coverage: system logs, systemd journal, login history (wtmp/btmp/lastlog), shell & AI CLI history, browsers, accounts, persistence, SSH, network, packages, trash, filesystem timeline, cross-module master timeline. macOS coverage: system logs, dslocal accounts, persistence (LaunchAgents/Daemons), Login Items/BTM, quarantine, TCC, KnowledgeC, browsers, shell & AI history, recent items, cross-module master timeline. New OS-aware menu/dispatch with a data-driven module registry for the Linux/macOS sets, an aggregated **Full HTML dashboard** generated when running all modules, an in-report search bar, and a cross-module **Master Timeline** that collects every timestamped finding.

**Date:** 2026-06-05 | **Version:** 1.2
New **Module 39 — AI Chat / Query History**: recovers AI assistant conversations (ChatGPT, Copilot, Claude, Cursor, Gemini, Codex, Windsurf, Continue) from offline disks. Includes a dependency-free pure-Python Snappy decompressor for ChatGPT LevelDB/IndexedDB (SSTable, WAL and external blob files), user/AI role attribution with product labelling, sensitive-string highlighting, IoC matching and Master Timeline integration.

**Date:** 2026-04-18 | **Version:** 1.1
Bug fix.

**Date:** 2026-04-15 | **Version:** 1.0
First release.

---

**Screenshot:**

<img width="1021" height="407" alt="image" src="https://github.com/user-attachments/assets/a6287b54-b09a-4653-a513-b23d9c15eaaf" />

<img width="970" height="1025" alt="image" src="https://github.com/user-attachments/assets/5e3819a3-738c-4886-a98e-70a2f71575de" />

<img width="1842" height="809" alt="image" src="https://github.com/user-attachments/assets/23ecbc8e-6ebd-4845-b9d4-20e3dce26af1" />

<img width="1861" height="750" alt="image" src="https://github.com/user-attachments/assets/727cf4dd-dd11-402b-9acd-b4fe123c154d" />
