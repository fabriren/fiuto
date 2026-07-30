
# ================================================================
#  REGISTRO MODULI PER OS (data-driven)
#
#  Formato entry:
#     "funzione|Nome|VARIABILE_COLORE|descrizione[|guardia][|flag]"
#
#  L'ordine determina la numerazione mostrata a menu e accettata da
#  --module / --modules: NON riordinare senza aggiornare il README, o si
#  rompono gli script di chi usa gia' il tool.
#
#  Nome e descrizione possono essere bilingui nella forma "italiano§english";
#  senza il separatore lo stesso testo vale per entrambe le lingue.
#
#  La guardia e' facoltativa: e' il nome di una funzione che ritorna 0 se il
#  modulo va eseguito, oppure stampa il motivo e ritorna non-zero per farlo
#  saltare in modalita' batch.
#
#  Flag riconosciuti (sesto campo):
#    defer  il modulo gira per ultimo con --all, pur mantenendo il suo numero
#           di menu. Serve alla Master Timeline, che aggrega gli altri: cosi'
#           si possono aggiungere moduli in coda senza rinumerarla.
# ================================================================

# Guardia del modulo PAD Offline: ha senso solo su un Domain Controller.
_guard_pad_offline() {
    local _ntds
    _ntds=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    [[ -n "$_ntds" ]] && return 0
    L "non è un DC" "not a DC"
    return 1
}

MODULES_WIN=(
    "module_ps_history|PowerShell History|MAGENTA|PSReadLine *_history.txt"
    "module_notepad_tabstate|Notepad TabState|MAGENTA|Tab rimasti aperti (.bin)§Open tabs (.bin)"
    "module_ifeo|IFEO Hijacking|RED|Image File Execution Options"
    "module_bam|BAM|BLUE|Background Activity Moderator"
    "module_run_keys|Run Keys & Persistenza§Run Keys & Persistence|ORANGE|Autorun nel registro§Autorun in registry"
    "module_prefetch|Prefetch|GREEN|Eseguibili tracciati (*.pf)§Tracked executables (*.pf)"
    "module_scheduled_tasks|Scheduled Tasks|YELLOW|Task pianificati (XML)§Scheduled tasks (XML)"
    "module_usb|USB Devices|BLUE|Dispositivi rimovibili (USBSTOR)§Removable devices (USBSTOR)"
    "module_lnk|LNK & JumpList|GREEN|File recenti e target path§Recent files and target path"
    "module_rdp_cache|RDP Cache|CYAN|Terminal Server Client Cache"
    "module_services|Services|RED|Servizi Windows (SYSTEM hive)§Windows Services (SYSTEM hive)"
    "module_evtx|Event Log|RED|Security/System/PS/RDP (.evtx)"
    "module_amcache|Amcache + Shimcache|YELLOW|Timeline esecuzione binari§Binary execution timeline"
    "module_recycle_bin|Recycle Bin|GREEN|File eliminati (\$Recycle.Bin)§Deleted files (\$Recycle.Bin)"
    "module_wmi|WMI Subscriptions|RED|Persistenza invisibile (T1546.003)§Fileless persistence (T1546.003)"
    "module_srum|SRUM|BLUE|Uso risorse per applicazione§Resource usage per application"
    "module_browser|Browser History|CYAN|Chrome / Edge / Firefox"
    "module_userassist|UserAssist / RunMRU|MAGENTA|Attività interattiva utente§Interactive user activity"
    "module_shellbags|ShellBags|CYAN|Navigazione cartelle (anche cancellate)§Folder navigation (including deleted)"
    "module_sam|SAM — Hash Locali§SAM — Local Hashes|RED|Hash NTLM account (impacket)§NTLM account hashes (impacket)"
    "module_mft|MFT Timeline|YELLOW|Master File Table + timestomping"
    "module_opensave|OpenSave / LastVisited MRU|GREEN|File aperti/salvati via dialogo§Files opened/saved via dialog"
    "module_usn|USN Journal|CYAN|\$UsnJrnl:\$J — change log NTFS§\$UsnJrnl:\$J — NTFS change log"
    "module_ntds|NTDS.dit|RED|Active Directory hash (DC offline)"
    "module_hiberfil|Hibernation / Pagefile|BLUE|hiberfil.sys · pagefile.sys strings"
    "module_wer_files|WER Files (Error Reports)|RED|ReportArchive · ReportQueue (.wer)"
    "module_credential_manager|Credential Manager|MAGENTA|DPAPI blob offline"
    "module_wlan|WLAN & VPN Profiles|CYAN|WiFi · NetworkList · VPN"
    "module_appx|AppX / UWP Packages|GREEN|App Store + sideload sospetti§App Store + suspicious sideloads"
    "module_browser_extra|Browser Downloads & Logins|CYAN|Download + Login Data (DPAPI)"
    "module_clipboard|Clipboard History|YELLOW|Cronologia appunti Win10+§Clipboard history Win10+"
    "module_office_mru|Office MRU|GREEN|File recenti Word/Excel/PowerPoint§Recent Word/Excel/PowerPoint files"
    "module_defender_quarantine|Defender Quarantine|RED|File in quarantena + threatname§Quarantined files + threatname"
    "module_ps_scriptblock|PS ScriptBlock Logging|MAGENTA|Event ID 4104 — PS Operational.evtx"
    "module_jumplists|JumpLists|GREEN|AutomaticDestinations · CustomDestinations"
    "module_network_artifacts|Network Artifacts|CYAN|Profili rete · Interfacce TCP/IP (registry)§Network profiles · TCP/IP interfaces (registry)"
    "module_master_timeline|Master Timeline|YELLOW|Aggregazione cross-moduli con filtri (con --all gira per ultima)§Cross-module aggregation with filters (runs last with --all)||defer"
    "module_pad_offline|PAD Offline AD Analysis|RED|NTDS.dit offline — utenti privilegiati, ACL, GPO§NTDS.dit offline — privileged users, ACL, GPO|_guard_pad_offline"
    "module_ai_chat|AI Chat History|MAGENTA|Claude · ChatGPT · Copilot · Cursor · Gemini · Codex"
    "module_setupapi|SetupAPI Device Log|BLUE|Prima installazione dispositivi (USB)§Device first install (USB)"
    "module_ps_transcripts|PowerShell Transcript|MAGENTA|Sessioni complete: comandi + output§Full sessions: commands + output"
    "module_lsa_secrets|LSA Secrets & DCC2|RED|SECURITY hive — password servizi, cache dominio§SECURITY hive — service passwords, domain cache"
    "module_vss|Volume Shadow Copies|CYAN|Snapshot precedenti del volume§Earlier volume snapshots"
    "module_pst_ost|Outlook PST / OST|YELLOW|Posta locale, allegati, item cancellati§Local mail, attachments, deleted items"
    "module_cloud_sync|Cloud Sync|BLUE|OneDrive/Dropbox/Drive — file sincronizzati§OneDrive/Dropbox/Drive — synced files"
    "module_bits|BITS Jobs|ORANGE|Download in background (T1197)§Background downloads (T1197)"
    "module_thumbcache|Thumbcache|GREEN|Miniature di file cancellati§Thumbnails of deleted files"
    "module_chat_desktop|Chat Desktop|MAGENTA|Slack/Teams/Discord — LevelDB§Slack/Teams/Discord — LevelDB"
    "module_webcache|WebCacheV01|CYAN|IE/Edge Legacy + WinINET§IE/Edge Legacy + WinINET"
    "module_search_index|Search Index|YELLOW|Windows.edb — file indicizzati§Windows.edb — indexed files"
    "module_sigma|Sigma|RED|Regole Sigma sugli EVTX (--sigma)§Sigma rules over EVTX (--sigma)|_guard_sigma"
    "module_xplat_sqlite_recovery|SQLite Recovery|MAGENTA|Record cancellati da freelist e spazio libero§Deleted records from freelist and free space"
    "module_xplat_esp_bootkit|EFI System Partition|RED|Bootkit e persistenza pre-boot§Bootkits and pre-boot persistence"
    "module_xplat_yara|YARA|RED|Scansione con regole esterne (--yara)§Scan with external rules (--yara)|_guard_yara"
)

MODULES_LINUX=(
    "module_linux_syslog|System Logs|GREEN|/var/log (syslog, auth, kern, secure...)"
    "module_linux_journal|systemd Journal|GREEN|var/log/journal/*.journal"
    "module_linux_logins|Login History|CYAN|wtmp / btmp / lastlog"
    "module_linux_shell_history|Shell History|MAGENTA|bash/zsh/sh + python/mysql/psql"
    "module_linux_ai_history|AI CLI History|MAGENTA|claude / aider / aichat / ollama"
    "module_linux_browser|Browser History|CYAN|Firefox / Chrome / Chromium / Brave"
    "module_linux_accounts|User Accounts|RED|passwd / shadow / group / sudoers"
    "module_linux_persistence|Persistence|ORANGE|cron / systemd / autostart / rc"
    "module_linux_ssh|SSH Artifacts|YELLOW|authorized_keys / known_hosts / sshd_config"
    "module_linux_network|Network Config|CYAN|hosts / NetworkManager / iptables"
    "module_linux_packages|Installed Packages|GREEN|dpkg / rpm / apt history / snap"
    "module_linux_trash|Trash & Recent|GREEN|~/.local/share/Trash + recently-used"
    "module_linux_timeline|Filesystem Timeline|YELLOW|MAC times aggregati (find/stat)"
    "module_linux_auditd|auditd|RED|/var/log/audit — syscall, auth, EXECVE"
    "module_linux_containers|Container|BLUE|Docker/Podman — inventario e fughe"
    "module_xplat_master_timeline|Master Timeline|YELLOW|aggrega le evidenze degli altri moduli (con --all gira per ultimo)§aggregates the other modules' findings (runs last with --all)||defer"
    "module_linux_pam|PAM|RED|Backdoor di autenticazione§Authentication backdoors"
    "module_linux_kernel_modules|Kernel Modules|RED|LKM rootkit, modprobe.d, initramfs§LKM rootkits, modprobe.d, initramfs"
    "module_linux_webserver_logs|Web Server Logs|ORANGE|nginx/apache — webshell, traversal, SQLi§nginx/apache — webshell, traversal, SQLi"
    "module_linux_cloud_credentials|Cloud Credentials|RED|~/.aws ~/.kube ~/.docker ~/.ssh§~/.aws ~/.kube ~/.docker ~/.ssh"
    "module_linux_suid_caps|SUID & Capabilities|ORANGE|Superficie di privilege escalation§Privilege escalation surface"
    "module_xplat_sqlite_recovery|SQLite Recovery|MAGENTA|Record cancellati da freelist e spazio libero§Deleted records from freelist and free space"
    "module_xplat_esp_bootkit|EFI System Partition|RED|Bootkit e persistenza pre-boot§Bootkits and pre-boot persistence"
    "module_xplat_yara|YARA|RED|Scansione con regole esterne (--yara)§Scan with external rules (--yara)|_guard_yara"
)

MODULES_MACOS=(
    "module_macos_logs|System Logs|GREEN|system.log / install.log / asl"
    "module_macos_accounts|User Accounts|RED|dslocal users .plist (+ hash)"
    "module_macos_persistence|Persistence|ORANGE|LaunchAgents / LaunchDaemons / cron"
    "module_macos_loginitems|Login Items (BTM)|RED|backgrounditems.btm"
    "module_macos_quarantine|Quarantine / Downloads|CYAN|QuarantineEventsV2 (download URL + origin)"
    "module_macos_tcc|TCC Privacy|RED|TCC.db (permessi cam/mic/disco)"
    "module_macos_knowledgec|KnowledgeC|BLUE|knowledgeC.db (uso app/attività)"
    "module_macos_browser|Browser History|CYAN|Safari / Chrome / Firefox"
    "module_macos_shell_ai_history|Shell & AI History|MAGENTA|zsh/bash + AI CLI"
    "module_macos_recent|Recent Items|GREEN|SFL / .Trash / recent items"
    "module_macos_fsevents|FSEvents|MAGENTA|/.fseventsd — modifiche al filesystem"
    "module_macos_spotlight|Spotlight|CYAN|store.db — provenienza download"
    "module_xplat_master_timeline|Master Timeline|YELLOW|aggrega le evidenze degli altri moduli (con --all gira per ultimo)§aggregates the other modules' findings (runs last with --all)||defer"
    "module_macos_messages|Messages|CYAN|chat.db — iMessage e SMS§chat.db — iMessage and SMS"
    "module_macos_cookies_downloads|Cookie & Download|YELLOW|Cookies.binarycookies · Downloads.plist§Cookies.binarycookies · Downloads.plist"
    "module_macos_xprotect|XProtect / Gatekeeper|RED|Difese native e autorizzazioni concesse§Native defences and granted authorisations"
    "module_macos_applications|Applications|GREEN|Inventario app, firma e posizione§App inventory, signature and location"
    "module_macos_backups|Time Machine / Snapshot|BLUE|Versioni precedenti dei file§Earlier versions of files"
    "module_macos_unified_logs|Unified Logs|MAGENTA|.tracev3 — estrazione parziale§.tracev3 — partial extraction"
    "module_xplat_sqlite_recovery|SQLite Recovery|MAGENTA|Record cancellati da freelist e spazio libero§Deleted records from freelist and free space"
    "module_xplat_esp_bootkit|EFI System Partition|RED|Bootkit e persistenza pre-boot§Bootkits and pre-boot persistence"
    "module_xplat_yara|YARA|RED|Scansione con regole esterne (--yara)§Scan with external rules (--yara)|_guard_yara"
)

# Restituisce il NOME dell'array registro per l'OS corrente (vuoto per windows/unknown)
