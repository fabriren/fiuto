
# ================================================================
#  REGISTRO MODULI PER OS NON-WINDOWS (data-driven)
#  Formato entry:  "funzione|Nome|VARIABILE_COLORE|descrizione"
#  L'ordine determina la numerazione mostrata a menu.
# ================================================================
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
    "module_xplat_master_timeline|Master Timeline|YELLOW|aggrega le evidenze degli altri moduli"
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
    "module_xplat_master_timeline|Master Timeline|YELLOW|aggrega le evidenze degli altri moduli"
)

# Restituisce il NOME dell'array registro per l'OS corrente (vuoto per windows/unknown)
