#!/usr/bin/env bash
# ================================================================
#  fiuto.sh  –  v2.6
#  Toolkit DFIR unificato per analisi di disco Windows offline
#
#  Uso:
#    ./fiuto.sh                        # menu interattivo
#    ./fiuto.sh /mnt/windows           # specifica la root
#    ./fiuto.sh /mnt/windows --all     # esegui tutto
#    ./fiuto.sh /mnt/windows --module 3  # modulo specifico
# ================================================================

set -uo pipefail

# ── Colori ───────────────────────────────────────────────────────
RED='\033[0;31m';    GREEN='\033[0;32m';  YELLOW='\033[1;33m'
CYAN='\033[0;36m';   BLUE='\033[0;34m';  MAGENTA='\033[0;35m'
BOLD='\033[1m';      DIM='\033[2m';      RESET='\033[0m'
WHITE='\033[1;37m';  ORANGE='\033[0;33m'

# ── Variabili globali ────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# Cattura la directory di invocazione risalendo l'albero dei processi finché non
# si trova una directory che non sia una directory di sistema (es. /usr/bin).
# Questo è immune a wrapper, sudo e qualsiasi meccanismo che cambia CWD.
_find_user_cwd() {
    local _pid=$$
    local _sys="|/usr/bin|/usr/local/bin|/bin|/sbin|/usr/sbin|/usr/libexec|"
    local _cwd _ppid
    for _ in 1 2 3 4 5 6 7 8 9 10; do
        _cwd=$(readlink -e /proc/$_pid/cwd 2>/dev/null) || break
        [[ "$_sys" != *"|${_cwd}|"* ]] && { echo "$_cwd"; return; }
        _ppid=$(awk '/^PPid:/{print $2}' /proc/$_pid/status 2>/dev/null) || break
        [[ -z "$_ppid" || "$_ppid" -le 1 ]] && break
        _pid=$_ppid
    done
    echo "${HOME:-/tmp}"
}
INVOCATION_DIR="$(_find_user_cwd)"
unset -f _find_user_cwd
WIN_ROOT=""          # root del volume Windows (es. /mnt/windows)
REPORT_BASE_DIR=""   # directory base dei report HTML
SCAN_DATE=""         # impostata all'avvio
declare -a GENERATED_REPORTS=()  # lista dei report HTML generati nella sessione
BATCH_MODE=false             # true in run_all_modules
declare -a SUMMARY_TABLE=()  # tabella riepilogativa batch
LOG_FILE=""                  # percorso log sessione (impostato all'avvio)
declare -a IOC_LIST=()       # lista IoC da file esterno (--ioc)
HOST_NAME=""                 # Nome macchina (da registro)
OS_VER=""                    # Sistema Operativo
IP_ADDR=""                   # Indirizzo IP
DOMAIN_NAME=""               # Dominio/Workgroup

# ── Dipendenze Python ────────────────────────────────────────────
# regipy  →  pip install regipy   (parsing hive offline)
# python-evtx → pip install python-evtx  (lettura .evtx)

# Rileva il corretto interprete Python da usare per tutto lo script.
# Il problema tipico è avere più versioni Python (es. sistema 3.12 + Homebrew 3.14):
# regipy potrebbe essere installato solo per una di esse.
# Questa routine cerca, nell'ordine, il primo python3.x che riesce ad importare regipy.
# Se nessuno ce l'ha, usa il python3 di sistema e i moduli senza regipy funzioneranno ugualmente.
_detect_python() {
    local _PY _CANDIDATES=()
    # Costruisce la lista dei candidati: prima le versioni esplicite trovate sul sistema,
    # poi il generico python3 come fallback.
    for _PY in python3.{13,12,11,10,9} python3; do
        command -v "$_PY" &>/dev/null && _CANDIDATES+=("$(command -v "$_PY")")
    done
    # Deduplicazione (risolve symlink per evitare di testare lo stesso eseguibile due volte)
    local _SEEN="" _REAL _FOUND=""
    for _PY in "${_CANDIDATES[@]}"; do
        _REAL=$(realpath "$_PY" 2>/dev/null || echo "$_PY")
        [[ "$_SEEN" == *"|${_REAL}|"* ]] && continue
        _SEEN="${_SEEN}|${_REAL}|"
        if "$_PY" -c "import regipy" 2>/dev/null; then
            _FOUND="$_PY"
            break
        fi
    done
    # Se nessuno ha regipy, prova ad aggiungere ~/.local/lib/pythonX.Y/site-packages
    # corrispondenti a ciascuna versione (pip install --user con versione diversa dal PATH)
    if [[ -z "$_FOUND" ]]; then
        for _PY in "${_CANDIDATES[@]}"; do
            _REAL=$(realpath "$_PY" 2>/dev/null || echo "$_PY")
            local _VER
            _VER=$("$_PY" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null || true)
            local _LSITE="$HOME/.local/lib/python${_VER}/site-packages"
            [[ -d "$_LSITE" ]] || continue
            if PYTHONPATH="${_LSITE}${PYTHONPATH:+:$PYTHONPATH}" "$_PY" -c "import regipy" 2>/dev/null; then
                export PYTHONPATH="${_LSITE}${PYTHONPATH:+:$PYTHONPATH}"
                _FOUND="$_PY"
                break
            fi
        done
    fi
    # Fallback finale: usa il python3 nel PATH comunque (i moduli senza regipy funzionano)
    echo "${_FOUND:-$(command -v python3 2>/dev/null || echo python3)}"
}
# PY3: variabile globale con il percorso del Python da usare in tutto lo script
PY3="$(_detect_python)"
unset -f _detect_python
# Versione Python rilevata (usata nel banner e nei messaggi di errore)
PY3_VERSION="$("$PY3" --version 2>&1 | awk '{print $2}')"

# ================================================================
#  UTILITIES
# ================================================================

print_banner() {
    clear
    echo -e "${CYAN}${BOLD}"
    echo "  ╔══════════════════════════════════════════════════════════╗"
    echo "  ║                                                          ║"
    echo "  ║      ███████╗  ██╗  ██╗   ██╗  ████████╗   ██████╗       ║"
    echo "  ║      ██╔════╝  ██║  ██║   ██║  ╚══██╔══╝  ██╔═══██╗      ║"
    echo "  ║      █████╗    ██║  ██║   ██║     ██║     ██║   ██║      ║"
    echo "  ║      ██╔══╝    ██║  ██║   ██║     ██║     ██║   ██║      ║"
    echo "  ║      ██║       ██║  ╚██████╔╝     ██║     ╚██████╔╝      ║"
    echo "  ║      ╚═╝       ╚═╝   ╚═════╝      ╚═╝      ╚═════╝       ║"
    echo "  ║                                                          ║"
    echo -e "  ║    ${CYAN}${BOLD}F${RESET}${CYAN}orensic ${BOLD}I${RESET}${CYAN}nvestigation ${BOLD}U${RESET}${CYAN}tility ${BOLD}T${RESET}${CYAN}ool for ${BOLD}O${RESET}${CYAN}ffline${RESET}       ${CYAN}${BOLD}║"
    echo -e "  ║                    ${MAGENTA}${BOLD}v2.6 - zi®iginal${RESET}${CYAN}                      ║"
    echo "  ╚══════════════════════════════════════════════════════════╝"
    echo -e "${RESET}"
    echo -e "  ${DIM}Data:   $(date '+%d/%m/%Y %H:%M:%S')${RESET}"
    echo -ne "  ${DIM}Root:   ${WIN_ROOT:-non impostata}${RESET}"
    [[ -n "$HOST_NAME" ]] && echo -ne "  ${CYAN}${BOLD}[${HOST_NAME}]${RESET}"
    echo ""
    echo -e "  ${DIM}Python: ${PY3} (${PY3_VERSION})${RESET}"
    if [[ -n "$REPORT_BASE_DIR" ]]; then
        local _RD_INFO=""
        if [[ -d "$REPORT_BASE_DIR" ]]; then
            if [[ -w "$REPORT_BASE_DIR" ]]; then
                _RD_INFO="${GREEN}[scrivibile]${RESET}"
            else
                _RD_INFO="${RED}[sola lettura!]${RESET}"
            fi
        else
            # non ancora creata: verifichiamo il parent
            local _RD_PARENT; _RD_PARENT=$(dirname "$REPORT_BASE_DIR")
            if [[ -w "$_RD_PARENT" ]]; then
                _RD_INFO="${GREEN}[creazione OK]${RESET}"
            else
                _RD_INFO="${RED}[parent non scrivibile!]${RESET}"
            fi
        fi
        echo -e "  ${DIM}Report: ${BOLD}${REPORT_BASE_DIR}${RESET}  ${_RD_INFO}"
    fi
    echo ""
}

# Stampa una sezione con titolo decorato
section_header() {
    local TITLE="$1"
    local COLOR="${2:-$CYAN}"
    echo ""
    echo -e "${COLOR}${BOLD}┌─────────────────────────────────────────────────────┐${RESET}"
    echo -e "${COLOR}${BOLD}│  ${TITLE}${RESET}"
    echo -e "${COLOR}${BOLD}└─────────────────────────────────────────────────────┘${RESET}"
    echo ""
}

# Stampa linea separatrice
separator() {
    echo -e "${DIM}  ─────────────────────────────────────────────────────${RESET}"
}

# Escape HTML — usare questa invece delle funzioni _esc_X locali nei moduli
html_esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

# Calcola SHA256 di un file (per chain of custody); ritorna stringa vuota se fallisce
sha256_file() { sha256sum "$1" 2>/dev/null | awk '{print $1}' || true; }

# Scrive un messaggio nel log di sessione (se LOG_FILE è impostato)
log_msg() {
    [[ -n "$LOG_FILE" ]] || return 0
    local _ld; _ld=$(dirname "$LOG_FILE")
    [[ -d "$_ld" ]] || mkdir -p "$_ld" 2>/dev/null || return 0
    printf '[%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$LOG_FILE" 2>/dev/null || true
}

# Timeout portabile (macOS / Linux)
portable_timeout() {
    local SECS="$1"
    shift
    if command -v timeout >/dev/null 2>&1; then
        timeout "$SECS" "$@"
    elif command -v gtimeout >/dev/null 2>&1; then
        gtimeout "$SECS" "$@"
    else
        "$@" &
        local pid=$!
        local count=0
        while kill -0 $pid 2>/dev/null; do
            sleep 1
            count=$((count+1))
            if [[ $count -ge $SECS ]]; then
                kill -9 $pid 2>/dev/null
                return 124
            fi
        done
        wait $pid
        return $?
    fi
}

# Pausa "premi un tasto per tornare al menu" — evita la ripetizione 33 volte in main()
return_to_menu() {
    echo ""
    echo -ne "  ${YELLOW}Premi qualsiasi tasto per tornare al menu...${RESET}"
    pause_key
}

# Chiede all'utente se aprire il report nel browser.
# In BATCH_MODE non apre e non chiede (nessun utente interattivo disponibile).
open_report_prompt() {
    local RPATH="$1"
    [[ "${BATCH_MODE:-false}" == "true" ]] && return 0
    local RESP
    echo -ne "  ${YELLOW}[?]${RESET} Aprire il report nel browser? [S/n]: "
    read -r RESP
    [[ "${RESP,,}" != "n" ]] && xdg-open "$RPATH" 2>/dev/null &
}

# Carica un file IoC (una entry per riga, righe # ignorate)
load_ioc_file() {
    local IOCFILE="$1"
    if [[ ! -f "$IOCFILE" ]]; then
        warn "File IoC non trovato: $IOCFILE"
        return 1
    fi
    while IFS= read -r LINE || [[ -n "$LINE" ]]; do
        [[ -z "$LINE" || "$LINE" == \#* ]] && continue
        IOC_LIST+=("$LINE")
    done < "$IOCFILE"
    ok "IoC caricati: ${#IOC_LIST[@]} da $IOCFILE"
    log_msg "[IOC] Caricati ${#IOC_LIST[@]} IoC da $IOCFILE"
}

# Controlla se una stringa contiene un IoC caricato; ritorna 0 se trovato
check_ioc() {
    local TEXT="${1,,}"
    for IOC in "${IOC_LIST[@]}"; do
        [[ "${TEXT}" == *"${IOC,,}"* ]] && return 0
    done
    return 1
}

# Info / warning / error
info()    { echo -e "  ${CYAN}[i]${RESET} $*"; log_msg "[INFO] $*"; }
ok()      { echo -e "  ${GREEN}[✓]${RESET} ${BOLD}$*${RESET}"; log_msg "[OK]   $*"; }
warn()    { echo -e "  ${YELLOW}[!]${RESET} $*"; log_msg "[WARN] $*"; }
ci_find_file() {
    local BASE="$1"
    local REL="$2"
    if [[ -f "$BASE/$REL" ]]; then
        echo "$BASE/$REL"
        return
    fi
    local DIR; DIR=$(dirname "$REL")
    local FILE; FILE=$(basename "$REL")
    local DIR_PATH; DIR_PATH=$(ci_find_dir "$BASE" "$DIR")
    [[ -z "$DIR_PATH" ]] && echo "" && return
    find "$DIR_PATH" -maxdepth 1 -iname "$FILE" -type f 2>/dev/null | head -1
}

# ================================================================
#  RECUPERO INFORMAZIONI MACCHINA
# ================================================================

gather_host_info() {
    [[ -n "$WIN_ROOT" ]] || return 1
    
    local SYSTEM_HIVE; SYSTEM_HIVE=$(get_hive "SYSTEM")
    local SOFTWARE_HIVE; SOFTWARE_HIVE=$(get_hive "SOFTWARE")
    
    if [[ -z "$SYSTEM_HIVE" && -z "$SOFTWARE_HIVE" ]]; then
        warn "Hive di registro non trovati, impossibile recuperare info macchina automaticamente."
    else
        info "Recupero informazioni macchina in corso..."
        
        local INFO_JSON
        INFO_JSON=$("$PY3" - "$SYSTEM_HIVE" "$SOFTWARE_HIVE" << 'PYEOF' 2>/dev/null
import sys, json
try:
    from regipy.registry import RegistryHive
    system_path = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] else None
    software_path = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] else None

    info = {"hostname": "", "os": "", "ip": "", "domain": ""}

    def get_key(hive, path):
        """Prova il path diretto e con prefisso ROOT\\ (regipy >= 6.x)."""
        for p in (path, 'ROOT\\' + path):
            try:
                return hive.get_key(p)
            except Exception:
                pass
        return None

    if system_path:
        try:
            sys_hive = RegistryHive(system_path)
            # Hostname
            try:
                hk = get_key(sys_hive, 'ControlSet001\\Control\\ComputerName\\ComputerName')
                if hk: info["hostname"] = hk.get_value('ComputerName') or ""
            except: pass

            # Network Info (IP / Domain)
            try:
                tk = get_key(sys_hive, 'ControlSet001\\Services\\Tcpip\\Parameters')
                if tk:
                    info["domain"] = tk.get_value('Domain') or tk.get_value('NV Domain') or ""

                # IP (cerca il primo con un IP valido)
                ik = get_key(sys_hive, 'ControlSet001\\Services\\Tcpip\\Parameters\\Interfaces')
                if ik:
                    for sub in ik.iter_subkeys():
                        ip = sub.get_value('DhcpIPAddress') or sub.get_value('IPAddress')
                        if ip and ip != '0.0.0.0':
                            if isinstance(ip, list): ip = ip[0]
                            info["ip"] = ip
                            break
            except: pass
        except: pass

    if software_path:
        try:
            soft_hive = RegistryHive(software_path)
            try:
                cvk = get_key(soft_hive, 'Microsoft\\Windows NT\\CurrentVersion')
                if cvk:
                    prod = cvk.get_value('ProductName') or ""
                    ver = cvk.get_value('DisplayVersion') or cvk.get_value('ReleaseId') or ""
                    build = cvk.get_value('CurrentBuild') or ""
                    info["os"] = f"{prod} {ver} (Build {build})".strip()
            except: pass
        except: pass

    print(json.dumps(info))
except:
    print("{}")
PYEOF
)
        HOST_NAME=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('hostname',''))" 2>/dev/null)
        OS_VER=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('os',''))" 2>/dev/null)
        IP_ADDR=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('ip',''))" 2>/dev/null)
        DOMAIN_NAME=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('domain',''))" 2>/dev/null)
    fi

    # Fallback Hostname
    if [[ -z "$HOST_NAME" ]]; then
        warn "Impossibile rilevare il nome macchina automaticamente."
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} Inserisci un nome per il caso/macchina [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi
    
    print_host_info_table
}

print_host_info_table() {
    echo -e "  ${CYAN}${BOLD}┌────────────────────────────────────────────────────────────────────────────┐${RESET}"
    echo -e "  ${CYAN}${BOLD}│                    INFORMAZIONI MACCHINA TARGET                            │${RESET}"
    echo -e "  ${CYAN}${BOLD}├────────────────────────────────────────────────────────────────────────────┤${RESET}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : ${WHITE}${BOLD}%-52s${RESET} ${CYAN}${BOLD}│${RESET}\n" "Hostname" "${HOST_NAME:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "Sistema Operativo" "${OS_VER:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "Indirizzo IP" "${IP_ADDR:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "Dominio" "${DOMAIN_NAME:-N/A}"
    echo -e "  ${CYAN}${BOLD}└────────────────────────────────────────────────────────────────────────────┘${RESET}"
    echo ""
}

err()     { echo -e "  ${RED}[✗]${RESET} $*"; log_msg "[ERR]  $*"; }
dim_msg() { echo -e "  ${DIM}[-] $*${RESET}"; log_msg "[DIM]  $*"; }

# Chiede conferma S/n, default S
ask_yn() {
    local PROMPT="$1"
    if [[ "${BATCH_MODE:-false}" == "true" ]]; then
        echo -e "  ${DIM}[auto] ${PROMPT} → S${RESET}"
        return 0
    fi
    local RESP
    echo -ne "  ${YELLOW}[?]${RESET} ${PROMPT} [S/n]: "
    read -r RESP
    [[ "${RESP,,}" != "n" ]]
}

# Attende la pressione di un singolo tasto (qualsiasi, senza aspettare INVIO).
# Usa stty raw mode per leggere 1 byte direttamente dal terminale.
# Fallback a read -r se non c'è un tty disponibile (es. pipe o batch).
pause_key() {
    local OLD_STTY
    if OLD_STTY=$(stty -g 2>/dev/null); then
        stty raw -echo min 1 time 0 2>/dev/null
        dd bs=1 count=1 >/dev/null 2>&1 </dev/tty || true
        stty "$OLD_STTY" 2>/dev/null
    else
        read -r 2>/dev/null || true
    fi
}

# Risolve un percorso case-insensitive su filesystem montato NTFS
ci_find_dir() {
    local BASE="$1"
    local REL="$2"
    # Fast path: prova il percorso diretto (case-sensitive, molto più veloce)
    if [[ -d "$BASE/$REL" ]]; then
        echo "$BASE/$REL"
        return
    fi
    local RESULT="$BASE"
    IFS='/' read -ra PARTS <<< "$REL"
    for PART in "${PARTS[@]}"; do
        [[ -z "$PART" ]] && continue
        local MATCH
        MATCH=$(find "$RESULT" -maxdepth 1 -iname "$PART" -type d 2>/dev/null | head -1)
        [[ -z "$MATCH" ]] && echo "" && return
        RESULT="$MATCH"
    done
    echo "$RESULT"
}

ci_find_file() {
    local BASE="$1"
    local NAME="$2"
    find "$BASE" -maxdepth 1 -iname "$NAME" -type f 2>/dev/null | head -1
}

# Verifica che WIN_ROOT sia impostata e contenga una struttura Windows
check_win_root() {
    if [[ -z "$WIN_ROOT" ]]; then
        err "Root Windows non impostata. Usa l'opzione [R] dal menu."
        return 1
    fi
    if [[ ! -d "$WIN_ROOT/Users" && ! -d "$WIN_ROOT/Windows" ]]; then
        # cerca case-insensitive
        local FOUND
        FOUND=$(find "$WIN_ROOT" -maxdepth 2 -type d \( -iname "Users" -o -iname "Windows" \) 2>/dev/null | head -1)
        if [[ -z "$FOUND" ]]; then
            err "Directory Users o Windows non trovata in: $WIN_ROOT"
            return 1
        fi
    fi
    return 0
}

# Torna il percorso della cartella Users (case-insensitive)
get_users_dir() {
    for C in "Users" "users" "USERS"; do
        [[ -d "$WIN_ROOT/$C" ]] && echo "$WIN_ROOT/$C" && return
    done
    find "$WIN_ROOT" -maxdepth 2 -type d -iname "Users" 2>/dev/null | head -1
}

# Elenca le home degli utenti (esclude account di sistema)
get_user_homes() {
    local USERS_DIR
    USERS_DIR=$(get_users_dir)
    [[ -z "$USERS_DIR" ]] && return
    for D in "$USERS_DIR"/*/; do
        local U
        U=$(basename "$D")
        case "${U,,}" in
            "public"|"all users"|"default"|"default user"|"defaultapppool") continue ;;
        esac
        [[ -d "$D" ]] && echo "$D"
    done
}

# Prepara la directory report e restituisce il path del file HTML.
# La directory base viene creata SOLO qui, al primo report effettivo,
# per evitare directory vuote quando l'utente non genera alcun report.
prepare_report_dir() {
    local MODULE_NAME="$1"
    [[ -n "$REPORT_BASE_DIR" && ! -d "$REPORT_BASE_DIR" ]] && mkdir -p "$REPORT_BASE_DIR"
    local DIR="${REPORT_BASE_DIR}/${MODULE_NAME}_$(date +%Y%m%d_%H%M%S)"
    mkdir -p "$DIR"
    echo "${DIR}/report.html"
}

# Registra il report nel riepilogo di sessione.
# Chiamata DOPO la scrittura effettiva del file, fuori da qualsiasi subshell.
register_report() {
    [[ -n "${1:-}" && -f "$1" ]] && GENERATED_REPORTS+=("$1")
}

# Controlla se regipy è disponibile tramite il Python rilevato all'avvio ($PY3)
check_regipy() {
    "$PY3" -c "import regipy" 2>/dev/null
}

# Torna il percorso di un hive di sistema
get_hive() {
    local NAME="$1"   # SOFTWARE, SYSTEM, SECURITY, SAM
    local PATH1="$WIN_ROOT/Windows/System32/config/${NAME}"
    local PATH2
    PATH2=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
    [[ -f "$PATH1" ]] && echo "$PATH1" && return
    [[ -n "$PATH2" ]] && {
        local F
        F=$(ci_find_file "$PATH2" "$NAME")
        [[ -n "$F" ]] && echo "$F"
    }
}

# Avvia un parser regipy su un hive e una chiave
regipy_query() {
    local HIVE="$1"
    local KEY="$2"
    "$PY3" - "$HIVE" "$KEY" << 'PYEOF'
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for v in key.get_values():
        print(f"{v.name}\t{v.value}")
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF
}

# HTML boilerplate header comune a tutti i report
html_header() {
    local TITLE="$1"
    local ACCENT="${2:-#58a6ff}"
    cat << HTMLEOF
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${TITLE} — DFIR Report</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@300;400;500;700&family=DM+Sans:wght@400;500;700&display=swap');
  :root {
    --bg:#080b0f; --bg2:#0d1117; --bg3:#131920; --bg4:#1a2332;
    --border:#1e2d3d; --accent:${ACCENT}; --accent2:#ff7b72;
    --accent3:#3fb950; --accent4:#f0883e;
    --text:#c9d1d9; --text-dim:#3d5166; --text-mid:#6e8898;
    --mono:'Fira Code',monospace; --sans:'DM Sans',sans-serif;
  }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { background:var(--bg); color:var(--text); font-family:var(--sans); min-height:100vh; }
  body::before {
    content:''; position:fixed; inset:0; pointer-events:none; z-index:0;
    background-image:linear-gradient(rgba(88,166,255,.03) 1px,transparent 1px),
      linear-gradient(90deg,rgba(88,166,255,.03) 1px,transparent 1px);
    background-size:40px 40px;
  }
  header { position:relative; z-index:1; background:var(--bg2); border-bottom:1px solid var(--border);
    padding:1.8rem 3rem; display:flex; align-items:center; gap:2rem; }
  .hicon { width:3rem; height:3rem; background:linear-gradient(135deg,var(--accent),#1f6feb);
    display:flex; align-items:center; justify-content:center; font-family:var(--mono);
    font-size:1.1rem; color:#fff; font-weight:700; flex-shrink:0;
    clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%); }
  .htxt h1 { font-size:1.35rem; font-weight:700; color:#fff; }
  .htxt h1 span { color:var(--accent); }
  .htxt .sub { font-size:.72rem; color:var(--text-dim); font-family:var(--mono); margin-top:.25rem; }
  .hmeta { margin-left:auto; font-family:var(--mono); font-size:.7rem; color:var(--text-dim); text-align:right; line-height:2; }
  .hmeta .val { color:var(--accent); }
  .statsbar { position:relative; z-index:1; background:var(--bg3); border-bottom:1px solid var(--border);
    padding:1rem 3rem; display:flex; gap:3rem; flex-wrap:wrap; }
  .stat .label { font-size:.6rem; text-transform:uppercase; letter-spacing:.15em; color:var(--text-dim); font-family:var(--mono); }
  .stat .value { font-size:1.8rem; font-weight:800; font-family:var(--mono); color:var(--accent2); }
  .stat.info .value { color:var(--accent); }
  .stat.ok .value { color:var(--accent3); }
  main { position:relative; z-index:1; padding:2rem 3rem; max-width:1400px; margin:0 auto; }
  .stitle { font-size:.65rem; text-transform:uppercase; letter-spacing:.2em; color:var(--text-dim);
    font-family:var(--mono); margin-bottom:1.5rem; padding-bottom:.5rem; border-bottom:1px solid var(--border); }
  .stitle::before { content:'> '; color:var(--accent3); }
  .cards { display:flex; flex-direction:column; gap:1.5rem; }
  .card { background:var(--bg2); border:1px solid var(--border); border-radius:6px; overflow:hidden; }
  .card:hover { border-color:rgba(88,166,255,.25); }
  .card-header { display:flex; align-items:center; gap:1rem; padding:1rem 1.5rem;
    background:var(--bg3); border-bottom:1px solid var(--border); }
  .uicon { width:2.2rem; height:2.2rem; background:linear-gradient(135deg,var(--accent),#1f6feb);
    color:#fff; display:flex; align-items:center; justify-content:center; font-size:.9rem; flex-shrink:0;
    clip-path:polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%); }
  .uname { font-weight:700; font-size:1rem; color:#fff; }
  .upath { font-family:var(--mono); font-size:.62rem; color:var(--text-dim); margin-top:.1rem; word-break:break-all; }
  .badge { margin-left:auto; background:rgba(88,166,255,.1); border:1px solid rgba(88,166,255,.3);
    color:var(--accent); font-family:var(--mono); font-size:.72rem; padding:.2rem .7rem;
    border-radius:2px; white-space:nowrap; }
  .badge.warn { background:rgba(255,123,114,.1); border-color:rgba(255,123,114,.3); color:var(--accent2); }
  table { width:100%; border-collapse:collapse; font-family:var(--mono); font-size:.78rem; }
  th { padding:.5rem 1rem; text-align:left; font-size:.6rem; text-transform:uppercase;
    letter-spacing:.15em; color:var(--text-dim); border-bottom:1px solid var(--border);
    background:rgba(0,0,0,.3); }
  td { padding:.5rem 1rem; border-bottom:1px solid rgba(30,45,61,.5); vertical-align:top; }
  tr:last-child td { border-bottom:none; }
  tr:hover { background:rgba(88,166,255,.03); }
  .mono { font-family:var(--mono); }
  .dim  { color:var(--text-dim); }
  .mid  { color:var(--text-mid); }
  .ok   { color:var(--accent3); }
  .warn { color:var(--accent4); }
  .bad  { color:var(--accent2); background:rgba(255,123,114,.07);
    border-left:2px solid rgba(255,123,114,.5); padding-left:calc(.5rem - 2px); }
  pre { white-space:pre-wrap; word-break:break-word; }
  footer { position:relative; z-index:1; border-top:1px solid var(--border); padding:1rem 3rem;
    font-family:var(--mono); font-size:.62rem; color:var(--text-dim); margin-top:3rem;
    display:flex; justify-content:space-between; }
  footer span { color:var(--accent); }
</style>
</head>
<body>
HTMLEOF
}

html_page_header() {
    local ICON="$1"
    local TITLE="$2"
    local SUBTITLE="$3"
    local SCAN="$4"
    local SRC="$5"
    local HASH="${6:-}"
    cat << HTMLEOF
<header>
  <div class="hicon">${ICON}</div>
  <div class="htxt">
    <h1>${TITLE}</h1>
    <div class="sub">${SUBTITLE}</div>
  </div>
  <div class="hmeta">
    <div>Scansione: <span class="val">${SCAN}</span></div>
    <div>Sorgente: <span class="val">${SRC}</span></div>$([ -n "$HASH" ] && echo "
    <div style='font-size:.6rem'>SHA256: <span class=\"val\" style='word-break:break-all'>${HASH}</span></div>")
  </div>
</header>
HTMLEOF
}

html_footer() {
    local SCAN="$1"
    local SRC="$2"
    cat << HTMLEOF
<footer>
  <div>fiuto.sh — generato il ${SCAN}</div>
  <div>Sorgente: <span>${SRC}</span></div>
</footer>
</body></html>
HTMLEOF
}

# ================================================================
#  MODULO 1 — PowerShell PSReadLine History
# ================================================================
module_ps_history() {
    section_header "PowerShell PSReadLine History" "$MAGENTA"
    check_win_root || return 1

    local PSREADLINE_REL="AppData/Roaming/Microsoft/Windows/PowerShell/PSReadLine"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_ROOTS=() RES_FILES=() RES_SIZES=()
    declare -a RES_CTIMES=() RES_MTIMES=() RES_PATHS=()

    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local PSRL_DIR
        PSRL_DIR=$(ci_find_dir "$USER_DIR" "$PSREADLINE_REL")
        if [[ -z "$PSRL_DIR" || ! -d "$PSRL_DIR" ]]; then
            dim_msg "$USERNAME — PSReadLine non trovata"
            continue
        fi
        mapfile -t HIST_FILES < <(find "$PSRL_DIR" -maxdepth 1 -iname "*_history.txt" -type f 2>/dev/null)
        if [[ ${#HIST_FILES[@]} -eq 0 ]]; then
            warn "$USERNAME — PSReadLine trovata ma nessun history"
            continue
        fi
        declare -a SORT_LIST=()
        for F in "${HIST_FILES[@]}"; do
            local CT
            CT=$(stat -c "%W" "$F" 2>/dev/null || echo "0")
            [[ "$CT" == "0" ]] && CT=$(stat -c "%Y" "$F" 2>/dev/null || echo "0")
            SORT_LIST+=("${CT}|${F}")
        done
        mapfile -t SORTED < <(printf '%s\n' "${SORT_LIST[@]}" | sort -t'|' -k1 -rn)
        local COUNT=${#SORTED[@]}
        ok "$USERNAME — ${BOLD}$COUNT file history"
        local FILE_NAMES="" FILE_SIZES="" FILE_CTIMES="" FILE_MTIMES="" FILE_PATHS=""
        for ENTRY in "${SORTED[@]}"; do
            local F="${ENTRY#*|}"
            local FNAME; FNAME=$(basename "$F")
            local FSIZE; FSIZE=$(stat -c "%s" "$F" 2>/dev/null || echo "?")
            local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
            local RAW_CT="${ENTRY%%|*}"
            local FCTIME
            if [[ "$RAW_CT" != "0" ]]; then
                FCTIME=$(date -d "@${RAW_CT}" "+%Y-%m-%d %H:%M:%S" 2>/dev/null || echo "?")
            else
                FCTIME=$(stat -c "%z" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
            fi
            local LINES; LINES=$(wc -l < "$F" 2>/dev/null || echo "?")
            echo -e "      ${DIM}• $FNAME  (${FSIZE} bytes, ${LINES} righe — creato: $FCTIME)${RESET}"
            # Stampa contenuto con highlight
            local DECODED
            DECODED=$("$PY3" -c "
import sys
path=sys.argv[1]
try:
    with open(path,'rb') as f: raw=f.read()
    if raw.startswith(b'\xff\xfe'):
        text=raw[2:].decode('utf-16-le',errors='replace')
    elif raw.startswith(b'\xfe\xff'):
        text=raw[2:].decode('utf-16-be',errors='replace')
    elif raw.startswith(b'\xef\xbb\xbf'):
        text=raw[3:].decode('utf-8',errors='replace')
    elif len(raw)>=2 and raw[1]==0:
        text=raw.decode('utf-16-le',errors='replace').lstrip('\ufeff')
    else:
        for enc in ('utf-8','latin-1'):
            try: text=raw.decode(enc,errors='strict'); break
            except: continue
        else: text=raw.decode('latin-1',errors='replace')
    text=text.lstrip('\ufeff').replace('\r\n','\n').replace('\r','\n')
    sys.stdout.buffer.write(text.encode('utf-8',errors='replace'))
except: pass
" "$F" 2>/dev/null || true)
            local LN=0
            while IFS= read -r LINE || [[ -n "$LINE" ]]; do
                LN=$((LN+1))
                if echo "$LINE" | grep -qi "password\|passwd\|secret\|credential\|token\|invoke-expression\|iex\|downloadstring\|bypass\|encodedcommand\|-enc\|mimikatz\|base64\|frombase64\|webclient\|net user\|net localgroup"; then
                    printf "      ${RED}%4d  %s${RESET}\n" "$LN" "$LINE"
                else
                    printf "      ${DIM}%4d${RESET}  %s\n" "$LN" "$LINE"
                fi
            done <<< "$DECODED"
            echo ""
            FILE_NAMES="${FILE_NAMES}${FNAME}|"
            FILE_SIZES="${FILE_SIZES}${FSIZE}|"
            FILE_CTIMES="${FILE_CTIMES}${FCTIME}|"
            FILE_MTIMES="${FILE_MTIMES}${FMTIME}|"
            FILE_PATHS="${FILE_PATHS}${F}|"
        done
        unset SORT_LIST SORTED
        RES_USERS+=("$USERNAME"); RES_ROOTS+=("$PSRL_DIR")
        RES_FILES+=("${FILE_NAMES%|}"); RES_SIZES+=("${FILE_SIZES%|}")
        RES_CTIMES+=("${FILE_CTIMES%|}"); RES_MTIMES+=("${FILE_MTIMES%|}")
        RES_PATHS+=("${FILE_PATHS%|}")
        TOTAL_FILES=$((TOTAL_FILES + COUNT)); TOTAL_USERS=$((TOTAL_USERS + 1))
    done < <(get_user_homes)

    separator
    info "Utenti con history: ${BOLD}$TOTAL_USERS${RESET}  |  File totali: ${BOLD}$TOTAL_FILES"

    [[ $TOTAL_FILES -eq 0 ]] && { warn "Nessun file history trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ps_history")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""

    for i in "${!RES_USERS[@]}"; do
        local USER="${RES_USERS[$i]}"
        local PSRL_PATH="${RES_ROOTS[$i]}"
        IFS='|' read -ra FNAMES  <<< "${RES_FILES[$i]}"
        IFS='|' read -ra FSIZES  <<< "${RES_SIZES[$i]}"
        IFS='|' read -ra FCTIMES <<< "${RES_CTIMES[$i]}"
        IFS='|' read -ra FMTIMES <<< "${RES_MTIMES[$i]}"
        IFS='|' read -ra FPATHS  <<< "${RES_PATHS[$i]}"
        # Combina gli array paralleli in un unico array ordinabile per FMTIME desc
        declare -a _FENTRIES=()
        for j in "${!FNAMES[@]}"; do
            _FENTRIES+=("${FMTIMES[$j]:-?}§${FCTIMES[$j]:-?}§${FNAMES[$j]}§${FSIZES[$j]:-?}§${FPATHS[$j]:-}")
        done
        mapfile -t _FENTRIES < <(printf '%s\n' "${_FENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+[""])[0],reverse=True); print("\n".join(L))')
        local FILE_BLOCKS=""
        for _FE in "${_FENTRIES[@]}"; do
            IFS='§' read -r FMTIME FCTIME FNAME FSIZE FPATH <<< "$_FE"
            local LINES_COUNT=0
            local CONTENT_HTML=""
            if [[ -f "$FPATH" && -s "$FPATH" ]]; then
                local RAW_CONTENT
                RAW_CONTENT=$("$PY3" - "$FPATH" << 'PYEOF'
import sys, html
path = sys.argv[1]
try:
    with open(path, 'rb') as f:
        raw = f.read()
    # Rileva encoding dal BOM o dalla presenza di null byte (UTF-16-LE senza BOM)
    if raw.startswith(b'\xff\xfe'):
        text = raw[2:].decode('utf-16-le', errors='replace')
    elif raw.startswith(b'\xfe\xff'):
        text = raw[2:].decode('utf-16-be', errors='replace')
    elif raw.startswith(b'\xef\xbb\xbf'):
        text = raw[3:].decode('utf-8', errors='replace')
    elif len(raw) >= 2 and raw[1] == 0:   # UTF-16-LE senza BOM (byte pari = null)
        text = raw.decode('utf-16-le', errors='replace')
    elif len(raw) >= 2 and raw[0] == 0:   # UTF-16-BE senza BOM
        text = raw.decode('utf-16-be', errors='replace')
    else:
        for enc in ('utf-8', 'latin-1'):
            try: text = raw.decode(enc); break
            except: continue
        else: text = raw.decode('latin-1', errors='replace')
    text = text.lstrip('\ufeff').replace('\r\n', '\n').replace('\r', '\n')
    lines = text.splitlines()
    SENS=['password','passwd','secret','credential','token','key','invoke-webrequest',
          'invoke-expression','iex','downloadstring','bypass','encodedcommand','-enc',
          'net user','net localgroup','whoami','mimikatz','base64','frombase64','webclient','hidden']
    out=[]
    for i,line in enumerate(lines,1):
        esc=html.escape(line)
        css='line sensitive' if any(k in line.lower() for k in SENS) else 'line'
        out.append(f'<span class="{css}"><span class="lnum">{i:4d}</span> {esc}</span>')
    print(f'LINES:{len(lines)}')
    print('\n'.join(out))
except Exception as e:
    print(f'LINES:0')
    print(f'<span class="line bad">Errore: {html.escape(str(e))}</span>')
PYEOF
)
                LINES_COUNT=$(echo "$RAW_CONTENT" | head -1 | sed 's/LINES://')
                RAW_CONTENT=$(echo "$RAW_CONTENT" | tail -n +2)
                CONTENT_HTML="<div class='hist-content'><pre class='hist-pre'>${RAW_CONTENT}</pre></div>"
            else
                CONTENT_HTML="<div style='padding:.8rem 1rem;font-family:var(--mono);font-size:.72rem;color:var(--text-dim);font-style:italic'>File vuoto</div>"
            fi
            FILE_BLOCKS+="<div class='card' style='margin-bottom:.8rem'>
              <div class='card-header'>
                <div class='uicon' style='font-size:.7rem'>PS</div>
                <div>
                  <div class='uname' style='font-size:.85rem'>${FNAME}</div>
                  <div class='upath'>${FPATH}</div>
                </div>
                <div style='margin-left:auto;text-align:right;font-family:var(--mono);font-size:.65rem;color:var(--text-dim)'>
                  <div class='ok'>creato: ${FCTIME}</div>
                  <div class='mid'>modificato: ${FMTIME}</div>
                  <div>${LINES_COUNT} righe &nbsp;·&nbsp; ${FSIZE} B</div>
                </div>
              </div>
              ${CONTENT_HTML}
            </div>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>▣</div>
            <div class='user-info'>
              <div class='uname'>${USER}</div>
              <div class='upath'>${PSRL_PATH}</div>
            </div>
            <div class='badge'>${#FNAMES[@]} file</div>
          </div>
          <div style='padding:1rem 1.5rem'>${FILE_BLOCKS}</div>
        </div>"
    done

    {
        html_header "PowerShell History"
        html_page_header "PS" "PowerShell <span>History</span> Forensics" \
            "%APPDATA%\\Microsoft\\Windows\\PowerShell\\PSReadLine\\*_history.txt" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>"
        echo "  <div class='stat'><div class='label'>File trovati</div><div class='value'>${TOTAL_FILES}</div></div>"
        echo "  <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>"
        echo "</div>"
        echo "<main><div class='stitle'>Risultati per utente</div>"
        echo "<style>
          .hist-pre{font-family:var(--mono);font-size:.75rem;line-height:1.7;padding:.8rem 1rem;
            overflow-x:auto;max-height:500px;overflow-y:auto;}
          .hist-pre::-webkit-scrollbar{width:4px;height:4px}
          .hist-pre::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
          .line{display:block;color:var(--text);white-space:pre;padding:.05rem .5rem;border-radius:2px}
          .line:hover{background:rgba(88,166,255,.05)}
          .line.sensitive{color:var(--accent2);background:rgba(255,123,114,.07);
            border-left:2px solid rgba(255,123,114,.5);padding-left:calc(.5rem - 2px)}
          .lnum{color:var(--text-dim);user-select:none;margin-right:1rem;font-size:.7rem}
        </style>"
        echo "<div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 2 — IFEO (Image File Execution Options)
# ================================================================
module_ifeo() {
    section_header "IFEO — Image File Execution Options" "$RED"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: pip install regipy  oppure  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE
    HIVE=$(get_hive "SOFTWARE")
    if [[ -z "$HIVE" || ! -f "$HIVE" ]]; then
        err "Hive SOFTWARE non trovato in $WIN_ROOT"
        return 1
    fi

    info "Parsing hive: $HIVE"

    # Chiave IFEO
    local IFEO_KEY="Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options"

    # Eseguibili sensibili tipicamente abusati
    local SENSITIVE_EXES="sethc.exe utilman.exe osk.exe magnify.exe narrator.exe displayswitch.exe atbroker.exe"

    declare -a ENTRIES_EXEC=()   # eseguibili con Debugger impostato
    declare -a ENTRIES_GFLAGS=() # GFlags sospetti

    mapfile -t IFEO_ENTRIES < <("$PY3" - "$HIVE" "$IFEO_KEY" << 'PYEOF'
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for subkey in key.iter_subkeys():
        name = subkey.name
        vals = {}
        for v in subkey.get_values():
            vals[v.name.lower()] = str(v.value)
        debugger = vals.get('debugger','')
        gflags   = vals.get('globalflag','') or vals.get('globalflags','')
        verifier = vals.get('verifierdlls','')
        if debugger or gflags or verifier:
            print(f"{name}\t{debugger}\t{gflags}\t{verifier}")
except Exception as e:
    print(f"ERROR\t{e}\t\t", file=sys.stderr)
PYEOF
)

    local FOUND_COUNT=0
    local SUSPICIOUS_COUNT=0

    echo ""
    printf "  %-35s %-45s %-15s %s\n" "Eseguibile" "Debugger" "GlobalFlag" "VerifierDLLs"
    separator

    for ENTRY in "${IFEO_ENTRIES[@]}"; do
        IFS=$'\t' read -r EXE DBG GFLAG VDLL <<< "$ENTRY"
        [[ "$EXE" == "ERROR" ]] && { err "$DBG"; continue; }
        FOUND_COUNT=$((FOUND_COUNT + 1))
        local IS_SENS=false
        for S in $SENSITIVE_EXES; do
            [[ "${EXE,,}" == "${S,,}" ]] && IS_SENS=true && break
        done
        if $IS_SENS || [[ -n "$DBG" ]]; then
            SUSPICIOUS_COUNT=$((SUSPICIOUS_COUNT + 1))
            printf "  ${RED}%-35s${RESET} ${YELLOW}%-45s${RESET} %-15s %s\n" "$EXE" "${DBG:-(nessuno)}" "${GFLAG:--}" "${VDLL:--}"
        else
            printf "  ${DIM}%-35s${RESET} %-45s %-15s %s\n" "$EXE" "${DBG:-(nessuno)}" "${GFLAG:--}" "${VDLL:--}"
        fi
        ENTRIES_EXEC+=("$EXE|${DBG}|${GFLAG}|${VDLL}")
    done

    separator
    info "Voci con Debugger/GFlags: ${BOLD}$FOUND_COUNT${RESET}  |  Sospette: ${RED}${BOLD}$SUSPICIOUS_COUNT${RESET}"

    [[ $FOUND_COUNT -eq 0 ]] && { ok "Nessuna voce IFEO con Debugger/GFlags trovata."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ifeo")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local ROWS=""
    for E in "${ENTRIES_EXEC[@]}"; do
        IFS='|' read -r EXE DBG GFLAG VDLL <<< "$E"
        local IS_SENS=false
        for S in $SENSITIVE_EXES; do
            [[ "${EXE,,}" == "${S,,}" ]] && IS_SENS=true && break
        done
        local ROW_CLASS=""
        $IS_SENS && ROW_CLASS="style='background:rgba(255,123,114,.07)'"
        local DBG_HTML="${DBG:--}"
        [[ -n "$DBG" ]] && DBG_HTML="<span class='bad'>${DBG}</span>"
        ROWS+="<tr ${ROW_CLASS}>
          <td class='mono'>${EXE}$(${IS_SENS} && echo " <span class='badge warn'>⚠ sensibile</span>" || true)</td>
          <td class='mono'>${DBG_HTML}</td>
          <td class='mono mid'>${GFLAG:--}</td>
          <td class='mono dim'>${VDLL:--}</td>
        </tr>"
    done

    {
        html_header "IFEO"
        html_page_header "IE" "IFEO <span>Hijacking</span> Forensics" \
            "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Voci con Debugger/Flags</div><div class='value'>${FOUND_COUNT}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSPICIOUS_COUNT}</div></div>
        </div>
        <main>
        <div class='stitle'>Voci IFEO con Debugger o GlobalFlags impostati</div>
        <div class='card'><table>
          <thead><tr><th>Eseguibile</th><th>Debugger</th><th>GlobalFlag</th><th>VerifierDLLs</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.82rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>ℹ Eseguibili sensibili monitorati</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>${SENSITIVE_EXES// / &nbsp;·&nbsp; }</div>
          <div style='margin-top:.8rem;font-size:.78rem;color:var(--text-mid)'>Questi eseguibili sono accessibili dalla schermata di login (prima dell&apos;autenticazione). Un Debugger su di essi consente shell SYSTEM senza credenziali. Ref: MITRE ATT&amp;CK T1546.012</div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 3 — BAM (Background Activity Moderator)
# ================================================================
module_bam() {
    section_header "BAM — Background Activity Moderator" "$BLUE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: pip install regipy  oppure  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE
    HIVE=$(get_hive "SYSTEM")
    if [[ -z "$HIVE" || ! -f "$HIVE" ]]; then
        err "Hive SYSTEM non trovato"
        return 1
    fi

    info "Parsing hive SYSTEM: $HIVE"

    # Tenta entrambi i ControlSet
    local BAM_DATA
    BAM_DATA=$("$PY3" - "$HIVE" << 'PYEOF'
import sys, datetime

try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])

    bam_key = None
    for cs in ['ControlSet001', 'ControlSet002', 'CurrentControlSet']:
        try:
            bam_key = hive.get_key(f'{cs}\\Services\\bam\\State\\UserSettings')
            break
        except:
            try:
                bam_key = hive.get_key(f'{cs}\\Services\\bam\\UserSettings')
                break
            except:
                continue

    if not bam_key:
        print("ERROR\tChiave BAM non trovata (Windows < 10 1709 o BAM disabilitato)")
        sys.exit(0)

    for sid_key in bam_key.iter_subkeys():
        sid = sid_key.name
        for v in sid_key.get_values():
            name = v.name
            if name.startswith('\\') or name.startswith('\\Device'):
                val = v.value
                ts = None
                if isinstance(val, bytes) and len(val) >= 8:
                    try:
                        ft = int.from_bytes(val[:8], 'little')
                        if ft > 0:
                            ts = datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ft//10)
                    except:
                        pass
                ts_str = ts.strftime('%Y-%m-%d %H:%M:%S UTC') if ts else '-'
                print(f"{sid}\t{name}\t{ts_str}")
except Exception as e:
    print(f"ERROR\t{e}\t")
PYEOF
)

    declare -a BAM_ENTRIES=()
    local ERROR_MSG=""
    while IFS= read -r LINE; do
        if [[ "$LINE" == ERROR* ]]; then
            ERROR_MSG="${LINE#ERROR }"
        else
            BAM_ENTRIES+=("$LINE")
        fi
    done <<< "$BAM_DATA"

    if [[ -n "$ERROR_MSG" ]]; then
        warn "$ERROR_MSG"
    fi

    local TOTAL=${#BAM_ENTRIES[@]}
    if [[ $TOTAL -eq 0 ]]; then
        warn "Nessuna voce BAM trovata."
        return 0
    fi

    # Raggruppa per SID e stampa
    local CUR_SID=""
    local SID_COUNT=0
    echo ""
    for ENTRY in "${BAM_ENTRIES[@]}"; do
        IFS=$'\t' read -r SID EXEPATH TIMESTAMP <<< "$ENTRY"
        if [[ "$SID" != "$CUR_SID" ]]; then
            [[ -n "$CUR_SID" ]] && separator
            echo -e "  ${GREEN}${BOLD}SID: $SID${RESET}"
            CUR_SID="$SID"
            SID_COUNT=$((SID_COUNT + 1))
        fi
        printf "  ${DIM}%-25s${RESET}  %s\n" "$TIMESTAMP" "$EXEPATH"
    done

    separator
    info "Totale eseguibili tracciati: ${BOLD}$TOTAL${RESET}  |  SID distinti: ${BOLD}$SID_COUNT"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "bam")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per timestamp decrescente (colonna 3) e poi per SID (colonna 1)
    mapfile -t BAM_ENTRIES < <(printf '%s\n' "${BAM_ENTRIES[@]}" | sort -t$'\t' -k3,3r -k1,1)

    # Raggruppa per SID in HTML
    local CARDS_HTML=""
    local PREV_SID="" SID_ROWS="" SID_EXEC_COUNT=0
    for ENTRY in "${BAM_ENTRIES[@]}"; do
        IFS=$'\t' read -r SID EXEPATH TIMESTAMP <<< "$ENTRY"
        if [[ "$SID" != "$PREV_SID" ]]; then
            if [[ -n "$PREV_SID" ]]; then
                CARDS_HTML+="<div class='card'>
                  <div class='card-header'><div class='uicon'>▣</div>
                    <div><div class='uname'>SID: ${PREV_SID}</div></div>
                    <div class='badge'>${SID_EXEC_COUNT} eseguibili</div>
                  </div>
                  <table><thead><tr><th>Timestamp (UTC)</th><th>Eseguibile</th></tr></thead>
                  <tbody>${SID_ROWS}</tbody></table></div>"
            fi
            PREV_SID="$SID"; SID_ROWS=""; SID_EXEC_COUNT=0
        fi
        SID_EXEC_COUNT=$((SID_EXEC_COUNT + 1))
        local EXE_HTML="<span class='mono'>${EXEPATH}</span>"
        # Evidenzia percorsi sospetti
        if echo "$EXEPATH" | grep -qi "temp\|appdata\|public\|programdata\|downloads"; then
            EXE_HTML="<span class='mono warn'>${EXEPATH}</span>"
        fi
        SID_ROWS+="<tr><td class='mono ok' style='white-space:nowrap'>${TIMESTAMP}</td><td>${EXE_HTML}</td></tr>"
    done
    # Ultima card
    if [[ -n "$PREV_SID" ]]; then
        CARDS_HTML+="<div class='card'>
          <div class='card-header'><div class='uicon'>▣</div>
            <div><div class='uname'>SID: ${PREV_SID}</div></div>
            <div class='badge'>${SID_EXEC_COUNT} eseguibili</div>
          </div>
          <table><thead><tr><th>Timestamp (UTC)</th><th>Eseguibile</th></tr></thead>
          <tbody>${SID_ROWS}</tbody></table></div>"
    fi

    {
        html_header "BAM"
        html_page_header "BA" "BAM — <span>Background Activity</span> Moderator" \
            "SYSTEM\\ControlSet001\\Services\\bam\\State\\UserSettings" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Eseguibili</div><div class='value'>${TOTAL}</div></div>
          <div class='stat info'><div class='label'>SID</div><div class='value'>${SID_COUNT}</div></div>
        </div>
        <main><div class='stitle'>Eseguibili tracciati per SID</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 4 — Cache RDP (Terminal Server Client)
# ================================================================
module_rdp_cache() {
    section_header "Cache RDP — Terminal Server Client" "$CYAN"
    check_win_root || return 1

    local RDP_REL="AppData/Local/Microsoft/Terminal Server Client/Cache"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_DIRS=() RES_FILES_LIST=()

    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local CACHE_DIR
        CACHE_DIR=$(ci_find_dir "$USER_DIR" "$RDP_REL")
        if [[ -z "$CACHE_DIR" || ! -d "$CACHE_DIR" ]]; then
            dim_msg "$USERNAME — Cache RDP non trovata"
            continue
        fi
        mapfile -t CACHE_FILES < <(find "$CACHE_DIR" -maxdepth 1 -type f \( -iname "*.bmc" -o -iname "*.bin" \) -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        if [[ ${#CACHE_FILES[@]} -eq 0 ]]; then
            mapfile -t CACHE_FILES < <(find "$CACHE_DIR" -maxdepth 1 -type f \( -iname "*.bmc" -o -iname "*.bin" \) 2>/dev/null | xargs ls -t 2>/dev/null)
        fi
        local COUNT=${#CACHE_FILES[@]}
        if [[ $COUNT -eq 0 ]]; then
            warn "$USERNAME — Directory cache trovata ma vuota"
            continue
        fi
        ok "$USERNAME — $COUNT file cache trovati in: $CACHE_DIR"
        local FILES_INFO=""
        for F in "${CACHE_FILES[@]}"; do
            local FNAME; FNAME=$(basename "$F")
            local FSIZE; FSIZE=$(stat -c "%s" "$F" 2>/dev/null || echo "?")
            local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
            echo -e "      ${DIM}• $FNAME  (${FSIZE} bytes — mod: $FMTIME)${RESET}"
            FILES_INFO="${FILES_INFO}${FNAME}:${FSIZE}:${FMTIME}|"
        done
        RES_USERS+=("$USERNAME")
        RES_DIRS+=("$CACHE_DIR")
        RES_FILES_LIST+=("${FILES_INFO%|}")
        TOTAL_FILES=$((TOTAL_FILES + COUNT))
        TOTAL_USERS=$((TOTAL_USERS + 1))
    done < <(get_user_homes)

    # Cerca anche server RDP nel registro
    echo ""
    info "Ricerca server RDP recenti nel registro (NTUSER.DAT)..."
    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local NTUSER
        NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        if check_regipy; then
            local RDP_SERVERS
            RDP_SERVERS=$("$PY3" - "$NTUSER" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    try:
        key = hive.get_key('Software\\Microsoft\\Terminal Server Client\\Servers')
        for sk in key.iter_subkeys():
            hostname = sk.name
            uname = ''
            for v in sk.get_values():
                if v.name.lower() == 'usernamehinint' or v.name.lower() == 'usernamehint':
                    uname = str(v.value)
            print(f"{hostname}\t{uname}")
    except: pass
    try:
        key2 = hive.get_key('Software\\Microsoft\\Terminal Server Client\\Default')
        for v in key2.get_values():
            if v.name.lower().startswith('mru'):
                print(f"MRU: {v.value}\t")
    except: pass
except: pass
PYEOF
)
            if [[ -n "$RDP_SERVERS" ]]; then
                echo -e "  ${GREEN}${BOLD}$USERNAME — Server RDP trovati:${RESET}"
                while IFS=$'\t' read -r HOST UNAME; do
                    printf "      ${CYAN}%-40s${RESET}  ${DIM}utente: %s${RESET}\n" "$HOST" "${UNAME:--}"
                done <<< "$RDP_SERVERS"
            fi
        fi
    done < <(get_user_homes)

    separator
    info "Utenti con cache RDP: ${BOLD}$TOTAL_USERS${RESET}  |  File totali: ${BOLD}$TOTAL_FILES"

    if [[ $TOTAL_FILES -gt 0 ]]; then
        echo ""
        info "${BOLD}Per analizzare le tile bitmap della cache usa bmc-tools:${RESET}"
        echo -e "    ${DIM}git clone https://github.com/ANSSI-FR/bmc-tools${RESET}"
        echo -e "    ${DIM}"$PY3" bmc-tools.py -s <dir_cache> -d ./output/ -b${RESET}"
    fi

    [[ $TOTAL_FILES -eq 0 ]] && { warn "Nessun file cache RDP trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "rdp_cache")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""

    for i in "${!RES_USERS[@]}"; do
        local USER="${RES_USERS[$i]}"
        local DIR="${RES_DIRS[$i]}"
        local ROWS=""
        IFS='|' read -ra FENTRIES <<< "${RES_FILES_LIST[$i]}"
        # Ordina per FMTIME (k3, resto della riga) decrescente — FMTIME può contenere ':'
        mapfile -t FENTRIES < <(printf '%s\n' "${FENTRIES[@]}" | sort -t':' -k3r)
        for FE in "${FENTRIES[@]}"; do
            IFS=':' read -r FNAME FSIZE FMTIME <<< "$FE"
            ROWS+="<tr>
              <td class='mono'>${FNAME}</td>
              <td class='mono mid' style='white-space:nowrap'>${FSIZE} B</td>
              <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
            </tr>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>RD</div>
            <div><div class='uname'>${USER}</div><div class='upath'>${DIR}</div></div>
            <div class='badge'>${#FENTRIES[@]} file</div>
          </div>
          <table><thead><tr><th>File</th><th>Dimensione</th><th>Ultima modifica</th></tr></thead>
          <tbody>${ROWS}</tbody></table>
        </div>"
    done

    {
        html_header "RDP Cache"
        html_page_header "RD" "RDP Cache <span>Forensics</span>" \
            "%LOCALAPPDATA%\\Microsoft\\Terminal Server Client\\Cache" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File cache</div><div class='value'>${TOTAL_FILES}</div></div>
          <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>
        </div>
        <main>
        <div class='stitle'>File cache per utente</div>
        <div class='cards'>${CARDS_HTML}</div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Analisi tile bitmap</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            git clone https://github.com/ANSSI-FR/bmc-tools<br>
            "$PY3" bmc-tools.py -s &lt;dir_cache&gt; -d ./output/ -b
          </div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 5 — Run Keys & Persistenza nel Registro
# ================================================================
module_run_keys() {
    section_header "Run Keys & Persistenza Registro" "$ORANGE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: pip install regipy  oppure  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SW
    HIVE_SW=$(get_hive "SOFTWARE")
    [[ -z "$HIVE_SW" ]] && { err "Hive SOFTWARE non trovato"; return 1; }

    declare -a ALL_ENTRIES=()

    # Chiavi da controllare su SOFTWARE (HKLM)
    local SW_KEYS=(
        "Microsoft\\Windows\\CurrentVersion\\Run"
        "Microsoft\\Windows\\CurrentVersion\\RunOnce"
        "Microsoft\\Windows\\CurrentVersion\\RunOnceEx"
        "Microsoft\\Windows NT\\CurrentVersion\\Winlogon"
        "Microsoft\\Windows NT\\CurrentVersion\\Windows"
        "Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run"
    )

    info "Scansione hive SOFTWARE (HKLM)..."
    for KEY in "${SW_KEYS[@]}"; do
        local VALS
        VALS=$("$PY3" - "$HIVE_SW" "$KEY" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for v in key.get_values():
        print(f"{v.name}\t{v.value}")
except Exception as e:
    pass
PYEOF
)
        if [[ -n "$VALS" ]]; then
            while IFS=$'\t' read -r VNAME VVAL; do
                ALL_ENTRIES+=("HKLM\\${KEY}|${VNAME}|${VVAL}")
            done <<< "$VALS"
        fi
    done

    # Chiavi NTUSER.DAT per ogni utente (HKCU)
    local HKCU_KEYS=(
        "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        "Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce"
    )

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        for KEY in "${HKCU_KEYS[@]}"; do
            local VALS
            VALS=$("$PY3" - "$NTUSER" "$KEY" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for v in key.get_values():
        print(f"{v.name}\t{v.value}")
except: pass
PYEOF
)
            if [[ -n "$VALS" ]]; then
                while IFS=$'\t' read -r VNAME VVAL; do
                    ALL_ENTRIES+=("HKCU[${USERNAME}]\\${KEY}|${VNAME}|${VVAL}")
                done <<< "$VALS"
            fi
        done
    done < <(get_user_homes)

    local TOTAL=${#ALL_ENTRIES[@]}
    echo ""
    printf "  ${BOLD}%-60s %-30s %s${RESET}\n" "Chiave" "Nome valore" "Dati"
    separator
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r HKEY VNAME VVAL <<< "$E"
        local VVAL_CLEAN="${VVAL}"
        local COLOR="$DIM"
        # Percorsi sospetti
        if echo "$VVAL" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|powershell\|cmd.exe\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin"; then
            COLOR="$RED"
        fi
        printf "  ${DIM}%-60s${RESET} %-30s ${COLOR}%s${RESET}\n" "$HKEY" "$VNAME" "$VVAL_CLEAN"
    done

    separator
    info "Voci totali trovate: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { ok "Nessuna voce Run trovata."; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "run_keys")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r HKEY VNAME VVAL <<< "$E"
        local CSS=""
        echo "$VVAL" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|powershell\|cmd.exe\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin" && CSS="class='bad'"
        ROWS+="<tr><td class='mono dim'>${HKEY}</td><td class='mono'>${VNAME}</td>
               <td class='mono' ${CSS}>${VVAL}</td></tr>"
    done

    {
        html_header "Run Keys"
        html_page_header "RK" "Run Keys <span>&amp; Persistenza</span>" \
            "HKLM/HKCU Run, RunOnce, Winlogon, AppInit_DLLs" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Voci trovate</div><div class='value'>${TOTAL}</div></div>
        </div>
        <main><div class='stitle'>Voci di autorun nel registro</div>
        <div class='card'><table>
          <thead><tr><th>Chiave</th><th>Nome valore</th><th>Dati (comando)</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 6 — Prefetch
# ================================================================
module_prefetch() {
    section_header "Prefetch — Eseguibili Tracciati" "$GREEN"
    check_win_root || return 1

    local PREF_DIR
    PREF_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/Prefetch")
    if [[ -z "$PREF_DIR" || ! -d "$PREF_DIR" ]]; then
        warn "Directory Prefetch non trovata (potrebbe essere disabilitato o sistema su SSD)"
        return 0
    fi

    mapfile -t PF_FILES < <(find "$PREF_DIR" -maxdepth 1 -iname "*.pf" -type f 2>/dev/null | sort)
    local TOTAL=${#PF_FILES[@]}
    if [[ $TOTAL -eq 0 ]]; then
        warn "Nessun file .pf trovato"
        return 0
    fi

    info "Trovati $TOTAL file .pf in: $PREF_DIR"

    # Parser Python per Prefetch (formato MAM/Uncompressed)
    declare -a PF_ENTRIES=()
    for PF in "${PF_FILES[@]}"; do
        local FNAME; FNAME=$(basename "$PF")
        local FSIZE; FSIZE=$(stat -c "%s" "$PF" 2>/dev/null || echo "?")
        local FMTIME; FMTIME=$(stat -c "%y" "$PF" 2>/dev/null | cut -d'.' -f1 || echo "?")
        # Estrai nome eseguibile dal nome file (es. CMD.EXE-XXXXXXXX.pf)
        local EXENAME="${FNAME%-*}"
        PF_ENTRIES+=("${EXENAME}|${FNAME}|${FSIZE}|${FMTIME}")
    done

    # Ordina per data decrescente (più recente prima) — FMTIME è campo 4 in formato YYYY-MM-DD HH:MM:SS
    mapfile -t PF_ENTRIES < <(printf '%s\n' "${PF_ENTRIES[@]}" | sort -t'|' -k4 -r)

    for E in "${PF_ENTRIES[@]}"; do
        IFS='|' read -r EXENAME FNAME FSIZE FMTIME <<< "$E"
        printf "  ${DIM}%-30s${RESET}  %-30s  %s\n" "$EXENAME" "$FMTIME" "$FSIZE B"
    done

    separator
    info "File Prefetch: ${BOLD}$TOTAL"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "prefetch")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${PF_ENTRIES[@]}"; do
        IFS='|' read -r EXENAME FNAME FSIZE FMTIME <<< "$E"
        local CSS=""
        echo "$EXENAME" | grep -qi "powershell\|cmd\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin\|mimikatz\|procdump\|psexec\|wce\|fgdump" && CSS="class='warn'"
        ROWS+="<tr><td class='mono' ${CSS}>${EXENAME}</td>
               <td class='mono dim'>${FNAME}</td>
               <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
               <td class='mono mid'>${FSIZE} B</td></tr>"
    done

    {
        html_header "Prefetch"
        html_page_header "PF" "Prefetch <span>Forensics</span>" \
            "Windows\\Prefetch\\*.pf" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File .pf</div><div class='value'>${TOTAL}</div></div>
        </div>
        <main><div class='stitle'>Eseguibili tracciati dal Prefetcher</div>
        <div class='card'><table>
          <thead><tr><th>Eseguibile</th><th>File .pf</th><th>Ultima esecuzione (mtime)</th><th>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 7 — Notepad TabState
# ================================================================
module_notepad_tabstate() {
    section_header "Notepad TabState — Tab Rimasti Aperti" "$MAGENTA"
    check_win_root || return 1

    local NOTEPAD_PKG="Microsoft.WindowsNotepad_8wekyb3d8bbwe"
    local TABSTATE_REL="AppData/Local/Packages"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_PATHS=() RES_COUNTS=()
    declare -a RES_FILES=() RES_SIZES=() RES_MTIMES=() RES_CTIMES=()

    # Parser .bin inline (stessa logica di notepad_tabstate.sh)
    local PARSER_PY=""
    PARSER_PY=$(mktemp /tmp/npad_parse_XXXXXX.py)
    trap '[[ -n "${PARSER_PY:-}" ]] && rm -f "$PARSER_PY"' RETURN
    cat > "$PARSER_PY" << 'PYEOF'
import sys, json, html as html_mod

def read_varint(data, offset):
    result = 0; shift = 0
    while offset < len(data):
        byte = data[offset]; offset += 1
        result |= (byte & 0x7F) << shift
        if not (byte & 0x80): break
        shift += 7
    return result, offset

def is_good_text(txt, thr=0.85):
    if not txt: return False
    ok = sum(1 for c in txt if c.isprintable() or c in '\n\r\t ')
    return ok / len(txt) >= thr

def parse_bin(path):
    try:
        with open(path, 'rb') as f: data = f.read()
    except Exception as e:
        return {'error': str(e), 'text': '', 'type': 'error', 'saved_path': '', 'char_count': 0}
    if len(data) < 4:
        return {'error': 'File troppo piccolo', 'text': '', 'type': 'empty', 'saved_path': '', 'char_count': 0}
    saved_path = ''; ftype = 0
    if data[0] == 0x4E and data[1] == 0x50:
        offset = 2; offset += 1
        ftype = data[offset]; offset += 1
        if ftype == 1 and offset < len(data):
            path_len = data[offset]; offset += 1
            if path_len > 0 and offset + path_len * 2 <= len(data):
                saved_path = data[offset:offset + path_len * 2].decode('utf-16-le', errors='replace')
                offset += path_len * 2
        best_text = ''; best_score = 0
        for skip in range(0, min(120, len(data) - offset)):
            o = offset + skip
            text_len, text_start = read_varint(data, o)
            if not (1 <= text_len <= 200000): continue
            end = text_start + text_len * 2
            if end > len(data): continue
            try:
                txt = data[text_start:end].decode('utf-16-le', errors='strict')
                if not is_good_text(txt): continue
                ok = sum(1 for c in txt if c.isprintable() or c in '\n\r\t ')
                score = len(txt) * (ok / len(txt))
                if score > best_score: best_text = txt; best_score = score
            except: pass
        if best_text:
            return {'text': best_text, 'saved_path': saved_path,
                    'type': 'saved' if ftype == 1 else 'unsaved',
                    'error': '', 'char_count': len(best_text)}
    # Fallback scan UTF-16LE
    start = 2 if (len(data) >= 2 and data[0] == 0x4E and data[1] == 0x50) else 0
    runs = []; i = start; run = []
    while i < len(data) - 1:
        b1, b2 = data[i], data[i+1]
        if b2 == 0 and (0x20 <= b1 <= 0x7e or b1 in (0x09, 0x0a, 0x0d)):
            run.append(chr(b1)); i += 2
        else:
            if len(run) >= 6: runs.append(''.join(run).strip())
            run = []; i += 1
    if len(run) >= 6: runs.append(''.join(run).strip())
    text = '\n'.join(r for r in runs if r)
    if text:
        return {'text': text, 'saved_path': saved_path, 'type': 'scan',
                'error': '', 'char_count': len(text)}
    return {'text': '', 'saved_path': '', 'type': 'empty', 'error': 'Nessun testo', 'char_count': 0}

if __name__ == '__main__':
    path = sys.argv[1]
    result = parse_bin(path)
    print(json.dumps({
        'text': html_mod.escape(result.get('text', '')),
        'saved_path': html_mod.escape(result.get('saved_path', '')),
        'type': result.get('type', ''),
        'error': result.get('error', ''),
        'char_count': result.get('char_count', 0)
    }))
PYEOF

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local PACKAGES_DIR
        PACKAGES_DIR=$(ci_find_dir "$USER_DIR" "AppData/Local/Packages")
        [[ -z "$PACKAGES_DIR" ]] && { dim_msg "$USERNAME — AppData\\Local\\Packages non trovata"; continue; }
        local NOTEPAD_DIR
        NOTEPAD_DIR=$(find "$PACKAGES_DIR" -maxdepth 1 -iname "${NOTEPAD_PKG}*" -type d 2>/dev/null | head -1)
        [[ -z "$NOTEPAD_DIR" ]] && { dim_msg "$USERNAME — Notepad UWP non installato"; continue; }
        local TABSTATE_DIR
        TABSTATE_DIR=$(ci_find_dir "$NOTEPAD_DIR" "LocalState/TabState")
        [[ -z "$TABSTATE_DIR" || ! -d "$TABSTATE_DIR" ]] && { warn "$USERNAME — TabState non trovata"; continue; }
        mapfile -t BIN_FILES < <(find "$TABSTATE_DIR" -maxdepth 1 -iname "*.bin" -type f -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        if [[ ${#BIN_FILES[@]} -eq 0 ]]; then
            # Fallback se printf %T@ non è supportato (BSD/macOS)
            mapfile -t BIN_FILES < <(find "$TABSTATE_DIR" -maxdepth 1 -iname "*.bin" -type f 2>/dev/null | xargs ls -t 2>/dev/null)
        fi
        local COUNT=${#BIN_FILES[@]}
        [[ $COUNT -eq 0 ]] && { warn "$USERNAME — TabState vuota"; continue; }
        ok "$USERNAME — $COUNT file .bin trovati"
        local FILE_NAMES="" FILE_SIZES="" FILE_MTIMES="" FILE_CTIMES=""
        for BIN in "${BIN_FILES[@]}"; do
            local FNAME; FNAME=$(basename "$BIN")
            local FSIZE; FSIZE=$(stat -c "%s" "$BIN" 2>/dev/null || echo "?")
            local FMTIME; FMTIME=$(stat -c "%y" "$BIN" 2>/dev/null | cut -d'.' -f1 || echo "?")
            local RAW_CT; RAW_CT=$(stat -c "%W" "$BIN" 2>/dev/null || echo "0")
            [[ "$RAW_CT" == "0" ]] && RAW_CT=$(stat -c "%Y" "$BIN" 2>/dev/null || echo "0")
            local FCTIME
            [[ "$RAW_CT" != "0" ]] && FCTIME=$(date -d "@${RAW_CT}" "+%Y-%m-%d %H:%M:%S" 2>/dev/null || echo "?") || FCTIME="?"
            if [[ "$FNAME" =~ \.[0-9]+\.bin$ ]]; then
                echo -e "      ${DIM}• $FNAME  (metadato sessione)${RESET}"
            else
                local TERM_TEXT
                TERM_TEXT=$("$PY3" "$PARSER_PY" "$BIN" 2>/dev/null \
                    | "$PY3" -c "import sys,json,html; d=json.load(sys.stdin); print(html.unescape(d.get('text','')))" 2>/dev/null || echo "")
                if [[ -n "$TERM_TEXT" ]]; then
                    echo -e "      ${GREEN}• $FNAME${RESET}  ${DIM}($FSIZE bytes)${RESET}"
                    local LN=0
                    while IFS= read -r LINE || [[ -n "$LINE" ]]; do
                        LN=$((LN+1))
                        printf "        ${DIM}%4d${RESET}  %s\n" "$LN" "$LINE"
                    done <<< "$TERM_TEXT"
                else
                    echo -e "      ${DIM}• $FNAME  ($FSIZE bytes — nessun testo)${RESET}"
                fi
            fi
            FILE_NAMES="${FILE_NAMES}${FNAME}|"
            FILE_SIZES="${FILE_SIZES}${FSIZE}|"
            FILE_MTIMES="${FILE_MTIMES}${FMTIME}|"
            FILE_CTIMES="${FILE_CTIMES}${FCTIME}|"
        done
        RES_USERS+=("$USERNAME"); RES_PATHS+=("$TABSTATE_DIR")
        RES_COUNTS+=("$COUNT")
        RES_FILES+=("${FILE_NAMES%|}"); RES_SIZES+=("${FILE_SIZES%|}")
        RES_MTIMES+=("${FILE_MTIMES%|}"); RES_CTIMES+=("${FILE_CTIMES%|}")
        TOTAL_FILES=$((TOTAL_FILES + COUNT)); TOTAL_USERS=$((TOTAL_USERS + 1))
    done < <(get_user_homes)

    separator
    info "Utenti: ${BOLD}$TOTAL_USERS${RESET}  |  File .bin: ${BOLD}$TOTAL_FILES"
    [[ $TOTAL_FILES -eq 0 ]] && { warn "Nessun tab Notepad trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "notepad_tabstate")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""
    for i in "${!RES_USERS[@]}"; do
        local USER="${RES_USERS[$i]}"
        local PATH_FULL="${RES_PATHS[$i]}"
        local COUNT="${RES_COUNTS[$i]}"
        IFS='|' read -ra FNAMES  <<< "${RES_FILES[$i]}"
        IFS='|' read -ra FSIZES  <<< "${RES_SIZES[$i]}"
        IFS='|' read -ra FMTIMES <<< "${RES_MTIMES[$i]}"
        IFS='|' read -ra FCTIMES <<< "${RES_CTIMES[$i]}"
        local ROWS=""
        for j in "${!FNAMES[@]}"; do
            local FNAME="${FNAMES[$j]}" FSIZE="${FSIZES[$j]:-?}"
            local FMTIME="${FMTIMES[$j]:-?}" FCTIME="${FCTIMES[$j]:-?}"
            local BIN_PATH="${PATH_FULL}/${FNAME}"
            local CONTENT_HTML=""
            if [[ "$FNAME" =~ \.[0-9]+\.bin$ ]]; then
                CONTENT_HTML="<span class='dim' style='font-size:.7rem'>metadato sessione</span>"
            elif [[ -f "$BIN_PATH" ]]; then
                local POUT
                POUT=$("$PY3" "$PARSER_PY" "$BIN_PATH" 2>/dev/null || echo '{}')
                local PTEXT; PTEXT=$("$PY3" -c "import sys,json; d=json.load(sys.stdin); print(d.get('text',''))" 2>/dev/null <<< "$POUT" || echo "")
                local PTYPE; PTYPE=$("$PY3" -c "import sys,json; d=json.load(sys.stdin); print(d.get('type',''))" 2>/dev/null <<< "$POUT" || echo "")
                local PPATH; PPATH=$("$PY3" -c "import sys,json; d=json.load(sys.stdin); print(d.get('saved_path',''))" 2>/dev/null <<< "$POUT" || echo "")
                local TBADGE=""
                case "$PTYPE" in
                    unsaved) TBADGE="<span style='color:var(--accent2);font-size:.7rem'>● non salvato</span>" ;;
                    saved)   TBADGE="<span style='color:var(--accent3);font-size:.7rem'>● salvato</span>" ;;
                    scan)    TBADGE="<span style='color:var(--accent);font-size:.7rem'>● scan</span>" ;;
                esac
                local SPATH_HTML=""
                [[ -n "$PPATH" ]] && SPATH_HTML="<div class='dim mono' style='font-size:.65rem;margin:.3rem 0'>📄 ${PPATH}</div>"
                if [[ -n "$PTEXT" ]]; then
                    CONTENT_HTML="${TBADGE}${SPATH_HTML}<pre style='font-family:var(--mono);font-size:.75rem;white-space:pre-wrap;max-height:250px;overflow-y:auto;margin-top:.4rem;color:var(--text)'>${PTEXT}</pre>"
                else
                    CONTENT_HTML="${TBADGE}<span class='dim' style='margin-left:.5rem;font-size:.7rem'>nessun testo leggibile</span>"
                fi
            fi
            ROWS+="<tr>
              <td class='mono ok' style='white-space:nowrap;font-size:.7rem'>${FCTIME}</td>
              <td class='mono' style='font-size:.75rem'>${FNAME}</td>
              <td class='mono mid' style='white-space:nowrap'>${FMTIME}</td>
              <td class='mono mid'>${FSIZE} B</td>
            </tr>
            <tr><td colspan='4' style='padding:.5rem 1rem 1rem'>${CONTENT_HTML}</td></tr>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>NT</div>
            <div><div class='uname'>${USER}</div><div class='upath'>${PATH_FULL}</div></div>
            <div class='badge'>${COUNT} tab</div>
          </div>
          <table><thead><tr><th>Creato</th><th>GUID / File</th><th>Modificato</th><th>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody></table>
        </div>"
    done

    {
        html_header "Notepad TabState"
        html_page_header "NT" "Notepad <span>TabState</span> Forensics" \
            "%LOCALAPPDATA%\\Packages\\Microsoft.WindowsNotepad_*\\LocalState\\TabState" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File .bin</div><div class='value'>${TOTAL_FILES}</div></div>
          <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>
        </div>
        <main><div class='stitle'>Tab Notepad per utente</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 8 — Scheduled Tasks
# ================================================================
module_scheduled_tasks() {
    section_header "Scheduled Tasks" "$YELLOW"
    check_win_root || return 1

    local TASKS_DIR
    TASKS_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/Tasks")
    if [[ -z "$TASKS_DIR" || ! -d "$TASKS_DIR" ]]; then
        warn "Directory Tasks non trovata"
        return 0
    fi

    mapfile -t TASK_FILES < <(find "$TASKS_DIR" -type f ! -iname "*.job" -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
    if [[ ${#TASK_FILES[@]} -eq 0 ]]; then
        mapfile -t TASK_FILES < <(find "$TASKS_DIR" -type f ! -iname "*.job" 2>/dev/null | xargs ls -t 2>/dev/null)
    fi
    local TOTAL=${#TASK_FILES[@]}
    info "Trovati $TOTAL task in: $TASKS_DIR"

    declare -a ALL_TASKS=()
    local SUSP_COUNT=0

    for TF in "${TASK_FILES[@]}"; do
        # Parser XML: estrae Command, Arguments, WorkingDirectory, Author, Description, Trigger
        local T_OUT
        T_OUT=$("$PY3" - "$TF" << 'PYEOF' 2>/dev/null || echo "|||||| "
import sys, xml.etree.ElementTree as ET
try:
    tree = ET.parse(sys.argv[1])
    root = tree.getroot()
    def ft(tag):
        for e in root.iter():
            if e.tag.split('}')[-1] == tag and e.text and e.text.strip():
                return e.text.strip()
        return ''
    triggers = []
    for e in root.iter():
        t = e.tag.split('}')[-1]
        if t.endswith('Trigger') and t != 'Triggers':
            sb = ''
            for c in e:
                if c.tag.split('}')[-1] == 'StartBoundary' and c.text:
                    sb = c.text.strip()
            triggers.append(f"{t}:{sb}" if sb else t)
    print('\t'.join([ft('Command'), ft('Arguments'), ft('WorkingDirectory'),
                     ft('Author'), ft('Description')[:100], ', '.join(triggers)]))
except:
    print('\t\t\t\t\t')
PYEOF
)
        IFS=$'\t' read -r T_CMD T_ARGS T_WORKDIR T_AUTHOR T_DESC T_TRIGGERS <<< "$T_OUT"

        local IS_SUSP=false
        local FULL_CMD="${T_CMD} ${T_ARGS}"
        echo "$FULL_CMD" | grep -qi "powershell\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin\|temp\|appdata\|programdata\|public\|downloads\|http\|ftp\|base64\|-enc\|-w hidden\|-nop" && IS_SUSP=true

        local RELPATH="${TF#${TASKS_DIR}/}"
        # Salva i 7 campi + flag sospetto separati da §  (carattere che non appare nei path)
        ALL_TASKS+=("${RELPATH}§${T_CMD}§${T_ARGS}§${T_WORKDIR}§${T_AUTHOR}§${T_DESC}§${T_TRIGGERS}§${IS_SUSP}")

        if $IS_SUSP; then
            echo -e "  ${RED}[!] $RELPATH${RESET}"
            [[ -n "$T_CMD"      ]] && echo -e "      ${YELLOW}Cmd:     ${RESET}${BOLD}$T_CMD${RESET}"
            [[ -n "$T_ARGS"     ]] && echo -e "      ${RED}Args:    $T_ARGS${RESET}"
            [[ -n "$T_WORKDIR"  ]] && echo -e "      ${DIM}WorkDir: $T_WORKDIR${RESET}"
            [[ -n "$T_AUTHOR"   ]] && echo -e "      ${DIM}Author:  $T_AUTHOR${RESET}"
            [[ -n "$T_TRIGGERS" ]] && echo -e "      ${DIM}Trigger: $T_TRIGGERS${RESET}"
            SUSP_COUNT=$((SUSP_COUNT + 1))
        else
            echo -e "  ${DIM}$RELPATH${RESET}"
            [[ -n "$T_CMD" ]] && echo -e "      ${DIM}↳ $T_CMD${RESET}"
        fi
    done

    separator
    info "Task totali: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "scheduled_tasks")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Funzione escape HTML — definita fuori dal loop così è disponibile anche nei subshell
    _esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

    local ROWS=""
    for E in "${ALL_TASKS[@]}"; do
        IFS='§' read -r TNAME T_CMD T_ARGS T_WORKDIR T_AUTHOR T_DESC T_TRIGGERS TISUSP <<< "$E"

        local ROW_STYLE="" ARG_EXTRA_CSS=""
        if [[ "$TISUSP" == "true" ]]; then
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
            ARG_EXTRA_CSS=" bad"
        fi

        # Ogni campo su riga propria, testo sempre wrappabile
        local DETAILS=""
        [[ -n "$T_CMD"      ]] && DETAILS+="<div class='drow'><span class='lbl'>CMD</span><span class='mono fld' style='color:var(--accent)'>$(_esc "$T_CMD")</span></div>"
        [[ -n "$T_ARGS"     ]] && DETAILS+="<div class='drow'><span class='lbl'>ARG</span><span class='mono fld${ARG_EXTRA_CSS}'>$(_esc "$T_ARGS")</span></div>"
        [[ -n "$T_WORKDIR"  ]] && DETAILS+="<div class='drow'><span class='lbl'>DIR</span><span class='mono fld mid'>$(_esc "$T_WORKDIR")</span></div>"
        [[ -n "$T_TRIGGERS" ]] && DETAILS+="<div class='drow'><span class='lbl'>TRG</span><span class='mono fld' style='color:var(--accent4)'>$(_esc "$T_TRIGGERS")</span></div>"
        [[ -n "$T_AUTHOR"   ]] && DETAILS+="<div class='drow'><span class='lbl'>AUT</span><span class='mono fld dim'>$(_esc "$T_AUTHOR")</span></div>"
        [[ -n "$T_DESC"     ]] && DETAILS+="<div class='drow'><span class='lbl'>DSC</span><span class='fld' style='font-size:.72rem;color:var(--text-mid)'>$(_esc "$T_DESC")</span></div>"
        [[ -z "$DETAILS"    ]] && DETAILS="<span class='dim' style='font-size:.72rem'>—</span>"

        ROWS+="<tr ${ROW_STYLE}>
          <td class='tname' style='vertical-align:top;padding-top:.5rem'>$(_esc "$TNAME")</td>
          <td style='padding:.35rem .9rem'>${DETAILS}</td>
        </tr>"
    done

    {
        html_header "Scheduled Tasks"
        html_page_header "ST" "Scheduled <span>Tasks</span> Forensics" \
            "Windows\\System32\\Tasks" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Task trovati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div>
        <main>
        <style>
          /* label badge fisso, testo sempre wrappabile */
          .lbl {
            display: inline-block;
            font-family: var(--mono);
            font-size: .58rem;
            color: var(--text-dim);
            width: 2.6rem;
            margin-right: .5rem;
            text-transform: uppercase;
            vertical-align: top;
            padding-top: .1rem;
            flex-shrink: 0;
          }
          .drow {
            display: flex;
            align-items: flex-start;
            margin-bottom: .22rem;
          }
          .drow:last-child { margin-bottom: 0; }
          /* campo testo: si wrappa sempre, niente overflow nascosto */
          .fld {
            word-break: break-all;
            overflow-wrap: anywhere;
            white-space: pre-wrap;
            flex: 1;
          }
          /* colonna nome task: wrappabile, font piccolo */
          .tname {
            font-family: var(--mono);
            font-size: .7rem;
            color: var(--text-mid);
            word-break: break-all;
            min-width: 160px;
            max-width: 280px;
            vertical-align: top;
            padding: .5rem .8rem .5rem 1rem;
          }
        </style>
        <div class='stitle'>Task pianificati — CMD · ARG · DIR · TRG · AUT · DSC</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:22%;min-width:160px'>Task</th>
            <th>Dettagli</th>
          </tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 9 — USB / Dispositivi Rimovibili
# ================================================================
module_usb() {
    section_header "USB — Dispositivi Rimovibili" "$BLUE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: pip install regipy  oppure  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SYS
    HIVE_SYS=$(get_hive "SYSTEM")
    [[ -z "$HIVE_SYS" ]] && { err "Hive SYSTEM non trovato"; return 1; }

    info "Parsing dispositivi USB da hive SYSTEM..."

    local USB_DATA
    USB_DATA=$("$PY3" - "$HIVE_SYS" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for cs in ['ControlSet001', 'ControlSet002']:
        try:
            usbstor = hive.get_key(f'{cs}\\Enum\\USBSTOR')
            for dtype in usbstor.iter_subkeys():
                for instance in dtype.iter_subkeys():
                    friendly = ''
                    serial = instance.name
                    for v in instance.get_values():
                        if v.name.lower() == 'friendlyname': friendly = str(v.value)
                    print(f"{dtype.name}\t{serial}\t{friendly}")
            break
        except: continue
except Exception as e:
    print(f"ERROR\t{e}\t")
PYEOF
)

    local TOTAL=0
    declare -a USB_ROWS=()
    echo ""
    printf "  ${BOLD}%-50s %-30s %s${RESET}\n" "Tipo dispositivo" "Seriale" "Nome"
    separator
    while IFS=$'\t' read -r DTYPE SERIAL FNAME; do
        [[ -z "$DTYPE" ]] && continue
        if [[ "$DTYPE" == "ERROR" ]]; then
            err "$SERIAL"
            continue
        fi
        printf "  ${CYAN}%-50s${RESET} ${DIM}%-30s${RESET} %s\n" "$DTYPE" "$SERIAL" "${FNAME:--}"
        USB_ROWS+=("${DTYPE}§${SERIAL}§${FNAME}")
        TOTAL=$((TOTAL + 1))
    done <<< "$USB_DATA"

    separator
    info "Dispositivi USB storici: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun dispositivo USB trovato nell'hive."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "usb")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for USB_E in "${USB_ROWS[@]}"; do
        IFS='§' read -r DTYPE SERIAL FNAME <<< "$USB_E"
        ROWS+="<tr>
          <td class='mono'>${DTYPE}</td>
          <td class='mono dim'>${SERIAL}</td>
          <td class='mono'>${FNAME:--}</td>
        </tr>"
    done

    {
        html_header "USB Devices"
        html_page_header "USB" "USB <span>Device History</span>" \
            "SYSTEM\\ControlSet001\\Enum\\USBSTOR" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Dispositivi</div><div class='value'>${TOTAL}</div></div>
        </div>
        <main><div class='stitle'>Dispositivi USB storici</div>
        <div class='card'><table>
          <thead><tr><th>Tipo</th><th>Seriale</th><th>Nome</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 10 — LNK & JumpList
# ================================================================
module_lnk() {
    section_header "LNK Files & JumpList" "$GREEN"
    check_win_root || return 1

    local TOTAL_LNK=0 TOTAL_USERS=0
    declare -a ALL_ENTRIES=()

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local RECENT_DIR
        RECENT_DIR=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Microsoft/Windows/Recent")
        [[ -z "$RECENT_DIR" ]] && { dim_msg "$USERNAME — Recent non trovata"; continue; }

        mapfile -t LNK_FILES < <(find "$RECENT_DIR" -maxdepth 1 -iname "*.lnk" -type f -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2- | head -50)
        if [[ ${#LNK_FILES[@]} -eq 0 ]]; then
            mapfile -t LNK_FILES < <(find "$RECENT_DIR" -maxdepth 1 -iname "*.lnk" -type f 2>/dev/null | xargs ls -t 2>/dev/null | head -50)
        fi
        local COUNT=${#LNK_FILES[@]}
        [[ $COUNT -eq 0 ]] && { dim_msg "$USERNAME — nessun .lnk trovato"; continue; }

        ok "$USERNAME — $COUNT file .lnk recenti"
        TOTAL_USERS=$((TOTAL_USERS + 1))

        for LNK in "${LNK_FILES[@]}"; do
            local FNAME; FNAME=$(basename "$LNK")
            local FMTIME; FMTIME=$(stat -c "%y" "$LNK" 2>/dev/null | cut -d'.' -f1 || echo "?")
            # Prova a estrarre il target path con python-lnk o strings
            local TARGET
            TARGET=$("$PY3" - "$LNK" << 'PYEOF' 2>/dev/null || true
import sys, struct
try:
    with open(sys.argv[1], 'rb') as f: data = f.read()
    # LNK magic: 4C 00 00 00
    if data[:4] != b'\x4c\x00\x00\x00': sys.exit()
    # Header size = 76 bytes, flags at offset 20
    flags = struct.unpack_from('<I', data, 20)[0]
    offset = 76
    # Shell Item ID List
    if flags & 0x01:
        id_list_size = struct.unpack_from('<H', data, offset)[0]
        offset += 2 + id_list_size
    # Link Info
    if flags & 0x02:
        li_size = struct.unpack_from('<I', data, offset)[0]
        li_hdr_size = struct.unpack_from('<I', data, offset + 4)[0]
        # Local base path offset
        if li_hdr_size >= 28:
            lbp_off = struct.unpack_from('<I', data, offset + 16)[0]
            if lbp_off > 0:
                abs_off = offset + lbp_off
                end = data.index(b'\x00', abs_off)
                print(data[abs_off:end].decode('latin-1'))
                sys.exit()
        offset += li_size
except: pass
PYEOF
)
            [[ -z "$TARGET" ]] && TARGET=$(strings "$LNK" 2>/dev/null | grep -E "^[A-Za-z]:\\\\" | head -1 || echo "")
            printf "  ${DIM}%-45s${RESET}  ${CYAN}%-25s${RESET}  %s\n" "$FNAME" "$FMTIME" "${TARGET:--}"
            ALL_ENTRIES+=("${USERNAME}|${FNAME}|${FMTIME}|${TARGET}")
            TOTAL_LNK=$((TOTAL_LNK + 1))
        done
        echo ""
    done < <(get_user_homes)

    separator
    info "Utenti: ${BOLD}$TOTAL_USERS${RESET}  |  .lnk trovati: ${BOLD}$TOTAL_LNK"
    [[ $TOTAL_LNK -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "lnk")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | sort -t'|' -k3,3r)

    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r UNAME FNAME FMTIME TARGET <<< "$E"
        ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>${UNAME}</td>
               <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
               <td class='mono'>${FNAME}</td>
               <td class='mono mid' style='word-break:break-all;font-size:.72rem'>${TARGET:--}</td></tr>"
    done

    {
        html_header "LNK Files"
        html_page_header "LN" "LNK Files <span>&amp; JumpList</span> Forensics" \
            "%APPDATA%\\Microsoft\\Windows\\Recent\\*.lnk" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File .lnk</div><div class='value'>${TOTAL_LNK}</div></div>
          <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>
        </div>
        <main><div class='stitle'>File recenti (LNK)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th style='width:16%;white-space:nowrap'>Data accesso</th><th style='width:24%'>File .lnk</th><th>Target path</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 11 — Services (Servizi Windows)
# ================================================================
module_services() {
    section_header "Services — Servizi Windows" "$RED"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: pip install regipy  oppure  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SYS
    HIVE_SYS=$(get_hive "SYSTEM")
    [[ -z "$HIVE_SYS" || ! -f "$HIVE_SYS" ]] && { err "Hive SYSTEM non trovato"; return 1; }

    info "Parsing servizi da hive SYSTEM: $HIVE_SYS"
    echo ""

    local SVCDATA
    SVCDATA=$("$PY3" - "$HIVE_SYS" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])

    START_MAP = {0:'Boot',1:'System',2:'Automatic',3:'Manual',4:'Disabled'}
    TYPE_MAP  = {1:'Kernel Driver',2:'FS Driver',4:'Adapter',8:'Recognizer',
                 16:'Own Process',32:'Share Process',256:'Interactive'}

    found_cs = None
    for cs in ['ControlSet001','ControlSet002','CurrentControlSet']:
        try:
            hive.get_key(f'{cs}\\Services'); found_cs = cs; break
        except: continue

    if not found_cs:
        print("ERROR\tChiave Services non trovata", file=sys.stderr); sys.exit(1)

    for svc in hive.get_key(f'{found_cs}\\Services').iter_subkeys():
        name = svc.name
        vals = {v.name.lower(): v.value for v in svc.get_values()}
        start    = START_MAP.get(vals.get('start'),    str(vals.get('start','')))
        svc_type = TYPE_MAP.get(vals.get('type'),      str(vals.get('type','')))
        image    = str(vals.get('imagepath',''))
        display  = str(vals.get('displayname',''))
        desc     = str(vals.get('description',''))[:100]
        obj_name = str(vals.get('objectname',''))

        # Flag sospetto: ImagePath fuori dalle dir di sistema standard
        img_l = image.lower()
        susp = 0
        if image and not any(p in img_l for p in [
            '\\system32\\','\\syswow64\\','\\drivers\\',
            '%systemroot%','%windir%','system32','syswow64',
            'svchost','ntosknrl'
        ]):
            susp = 1
        # ObjectName non standard
        if obj_name and obj_name.lower() not in (
            'localsystem','localservice','networkservice',''):
            susp = 1

        print(f"{name}\t{display}\t{start}\t{svc_type}\t{image}\t{obj_name}\t{desc}\t{susp}")

except Exception as e:
    print(f"ERROR\t{e}", file=sys.stderr)
PYEOF
)

    local TOTAL=0 SUSP_COUNT=0
    declare -a ALL_SVCS=()

    printf "  ${BOLD}%-30s %-12s %-18s %s${RESET}\n" "Nome" "Start" "Tipo" "ImagePath"
    separator

    while IFS=$'\t' read -r SVC_NAME SVC_DISP SVC_START SVC_TYPE SVC_IMG SVC_OBJ SVC_DESC SVC_SUSP; do
        [[ -z "$SVC_NAME" ]] && continue
        ALL_SVCS+=("${SVC_NAME}§${SVC_DISP}§${SVC_START}§${SVC_TYPE}§${SVC_IMG}§${SVC_OBJ}§${SVC_DESC}§${SVC_SUSP}")
        TOTAL=$((TOTAL + 1))
        if [[ "$SVC_SUSP" == "1" ]]; then
            SUSP_COUNT=$((SUSP_COUNT + 1))
            printf "  ${RED}%-30s${RESET} ${YELLOW}%-12s${RESET} %-18s ${RED}%s${RESET}\n" \
                "$SVC_NAME" "$SVC_START" "$SVC_TYPE" "$SVC_IMG"
            [[ -n "$SVC_OBJ" && "${SVC_OBJ,,}" != "localsystem" ]] && \
                echo -e "      ${YELLOW}ObjectName: $SVC_OBJ${RESET}"
        else
            printf "  ${DIM}%-30s${RESET} %-12s %-18s %s\n" \
                "$SVC_NAME" "$SVC_START" "$SVC_TYPE" "$SVC_IMG"
        fi
    done <<< "$SVCDATA"

    separator
    info "Servizi totali: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun servizio trovato."; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "services")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Escape HTML — usata anche per Services
    _esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

    local ROWS=""
    for E in "${ALL_SVCS[@]}"; do
        IFS='§' read -r SVC_NAME SVC_DISP SVC_START SVC_TYPE SVC_IMG SVC_OBJ SVC_DESC SVC_SUSP <<< "$E"

        local ROW_STYLE="" START_CSS="mono mid"
        [[ "$SVC_SUSP" == "1" ]] && ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        case "$SVC_START" in
            Boot|System)   START_CSS="mono bad"  ;;
            Automatic)     START_CSS="mono warn" ;;
            Disabled)      START_CSS="mono dim"  ;;
        esac

        local IMG_CSS="mono fld"
        [[ "$SVC_SUSP" == "1" ]] && IMG_CSS="mono fld bad"

        local DETAILS=""
        [[ -n "$SVC_IMG"  ]] && DETAILS+="<div class='drow'><span class='lbl'>IMG</span><span class='${IMG_CSS}'>$(_esc "$SVC_IMG")</span></div>"
        if [[ -n "$SVC_OBJ" && "${SVC_OBJ,,}" != "localsystem" ]]; then
            DETAILS+="<div class='drow'><span class='lbl'>OBJ</span><span class='mono fld warn'>$(_esc "$SVC_OBJ")</span></div>"
        fi
        [[ -n "$SVC_DESC" ]] && DETAILS+="<div class='drow'><span class='lbl'>DSC</span><span class='fld' style='font-size:.72rem;color:var(--text-mid)'>$(_esc "$SVC_DESC")</span></div>"
        [[ -z "$DETAILS"  ]] && DETAILS="<span class='dim' style='font-size:.72rem'>—</span>"

        local DISP_HTML=""
        [[ -n "$SVC_DISP" && "$SVC_DISP" != "$SVC_NAME" ]] && \
            DISP_HTML="<div class='dim' style='font-size:.65rem;margin-top:.1rem'>$(_esc "$SVC_DISP")</div>"

        ROWS+="<tr ${ROW_STYLE}>
          <td class='tname' style='vertical-align:top;padding-top:.5rem'>
            <span class='mono'>$(_esc "$SVC_NAME")</span>${DISP_HTML}
          </td>
          <td style='vertical-align:top;padding:.5rem .8rem;white-space:nowrap'>
            <span class='${START_CSS}'>${SVC_START}</span>
          </td>
          <td class='mono dim' style='vertical-align:top;padding:.5rem .8rem;white-space:nowrap;font-size:.72rem'>${SVC_TYPE}</td>
          <td style='vertical-align:top;padding:.35rem .9rem'>${DETAILS}</td>
        </tr>"
    done

    {
        html_header "Services"
        html_page_header "SV" "Windows <span>Services</span> Forensics" \
            "SYSTEM\\ControlSet001\\Services" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Servizi trovati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div>
        <main>
        <style>
          .lbl{display:inline-block;font-family:var(--mono);font-size:.58rem;color:var(--text-dim);
               width:2.6rem;margin-right:.5rem;text-transform:uppercase;vertical-align:top;padding-top:.1rem;flex-shrink:0}
          .drow{display:flex;align-items:flex-start;margin-bottom:.22rem}
          .drow:last-child{margin-bottom:0}
          .fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap;flex:1}
          .tname{font-family:var(--mono);font-size:.72rem;color:var(--text-mid);word-break:break-all;
                 min-width:140px;max-width:240px;padding:.5rem .8rem .5rem 1rem}
        </style>
        <div class='stitle'>Servizi Windows — IMG · OBJ · DSC</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:18%'>Nome / Display</th>
            <th style='width:9%'>Start</th>
            <th style='width:13%'>Tipo</th>
            <th>Dettagli</th>
          </tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>ℹ Criteri di flagging sospetto</div>
          <div style='color:var(--text-mid);line-height:1.8'>
            • ImagePath fuori da System32 / SysWOW64 / drivers<br>
            • ObjectName diverso da LocalSystem / LocalService / NetworkService
          </div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 0 — Esegui tutti i moduli
# ================================================================
# ================================================================
#  MODULO 12 — Event Log (Security / System / PowerShell / RDP)
# ================================================================
module_evtx() {
    section_header "Event Log — Analisi .evtx" "$RED"
    check_win_root || return 1

    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        warn "python-evtx non trovato."
        warn "Installalo con: ${PY3} -m pip install python-evtx"
        return 1
    fi

    local EVTX_DIR
    EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    if [[ -z "$EVTX_DIR" || ! -d "$EVTX_DIR" ]]; then
        err "Directory Event Log non trovata"; return 1
    fi
    info "Directory log: $EVTX_DIR"
    echo ""

    declare -A EVTX_MAP=(
        ["Security.evtx"]="4624 4625 4648 4688 4698 4702 4720 4726 4732 4733 4756 4757"
        ["System.evtx"]="7045 7034 7036"
        ["Microsoft-Windows-PowerShell%4Operational.evtx"]="4104"
        ["Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx"]="21 23 24 25"
        ["Microsoft-Windows-TaskScheduler%4Operational.evtx"]="106 140 141 200 201"
        ["Microsoft-Windows-Windows Defender%4Operational.evtx"]="1116 1117 1118 1119"
    )

    # Parser: legge .evtx, filtra per EID, scrive JSON su file (non su argv/stdout grande)
    local PARSER_EVTX=""
    PARSER_EVTX=$(mktemp /tmp/dfir_evtxp_XXXXXX.py)
    cat > "$PARSER_EVTX" << 'PYEOF'
import sys, json, re
try:
    import Evtx.Evtx as evtx
except ImportError:
    json.dump([], open(sys.argv[3],'w')); sys.exit(1)

path, eids_csv, outf = sys.argv[1], sys.argv[2], sys.argv[3]
target = set(eids_csv.split(','))
NS = 'http://schemas.microsoft.com/win/2004/08/events/event'

def sn(tag): return re.sub(r'\{[^}]+\}','',tag)

results = []
try:
    with evtx.Evtx(path) as log:
        for rec in log.records():
            try:
                root = rec.lxml()
                sys_el = root.find(f'{{{NS}}}System')
                if sys_el is None: continue
                eid_el = sys_el.find(f'{{{NS}}}EventID')
                if eid_el is None or eid_el.text not in target: continue
                ts_el  = sys_el.find(f'{{{NS}}}TimeCreated')
                cmp_el = sys_el.find(f'{{{NS}}}Computer')
                ts   = (ts_el.attrib.get('SystemTime','') if ts_el is not None else '')[:19].replace('T',' ')
                comp = cmp_el.text if cmp_el is not None else ''
                de   = root.find(f'.//{{{NS}}}EventData')
                data = {}
                if de is not None:
                    for it in de:
                        data[it.attrib.get('Name', sn(it.tag))] = (it.text or '').strip()
                results.append({'eid':eid_el.text,'ts':ts,'computer':comp,'data':data})
                if len(results) >= 500: break
            except: continue
except Exception as e:
    results = [{'eid':'ERROR','ts':'','computer':'','data':{'msg':str(e)}}]
json.dump(results, open(outf,'w'))
PYEOF

    # File JSONL globale — accumula tutti gli eventi senza limiti argv
    local TMP_ALL=""
    TMP_ALL=$(mktemp /tmp/dfir_allevt_XXXXXX.jsonl)
    trap '[[ -n "${PARSER_EVTX:-}" ]] && rm -f "$PARSER_EVTX"; [[ -n "${TMP_ALL:-}" ]] && rm -f "$TMP_ALL"' RETURN

    local TOTAL_FOUND=0 SUSP_COUNT=0

    for EVTX_NAME in "${!EVTX_MAP[@]}"; do
        local EVTX_FILE
        EVTX_FILE=$(ci_find_file "$EVTX_DIR" "$EVTX_NAME")
        [[ -z "$EVTX_FILE" ]] && { dim_msg "$EVTX_NAME — non trovato"; continue; }

        local EIDS_CSV="${EVTX_MAP[$EVTX_NAME]// /,}"
        local LABEL="${EVTX_NAME%.evtx}"; LABEL="${LABEL//%4/\/}"
        info "Parsing: $LABEL"

        local TMP_OUT; TMP_OUT=$(mktemp /tmp/dfir_evtxo_XXXXXX.json)
        "$PY3" "$PARSER_EVTX" "$EVTX_FILE" "$EIDS_CSV" "$TMP_OUT" 2>/dev/null \
            || echo "[]" > "$TMP_OUT"

        local COUNT
        COUNT=$("$PY3" -c "import json,sys; d=json.load(open(sys.argv[1])); print(len(d))" \
            "$TMP_OUT" 2>/dev/null || echo 0)
        [[ "$COUNT" -eq 0 ]] && { dim_msg "  Nessun evento corrispondente"; rm -f "$TMP_OUT"; continue; }

        TOTAL_FOUND=$((TOTAL_FOUND + COUNT))
        ok "  $COUNT eventi trovati"

        # Stampa terminale
        "$PY3" - "$TMP_OUT" << 'PYEOF' 2>/dev/null || true
import sys,json
EID_DESC={"4624":"Logon OK","4625":"Logon FAIL","4648":"Logon espl.","4688":"Proc. creato",
           "4698":"Task creato","4702":"Task modif.","4720":"Account creato","4726":"Account elim.",
           "4732":"Aggiunto gruppo","7045":"Servizio inst.","7034":"Svc crash",
           "4104":"PS ScriptBlock","21":"RDP Logon","23":"RDP Logoff",
           "1116":"Malware rilevato","1117":"Defender azione"}
SUSP={"4648","4698","4702","4720","7045","4104","1116","1117","4688"}
events=json.load(open(sys.argv[1]))
for ev in events[:15]:
    eid=ev['eid']; ts=ev['ts']
    mark="\033[0;31m⚠\033[0m " if eid in SUSP else "  "
    d=ev['data']
    detail=(d.get('SubjectUserName','') or d.get('TargetUserName','') or
            d.get('ImagePath','') or d.get('ProcessCommandLine','') or
            d.get('ScriptBlockText','')[:60] or d.get('TaskName','') or
            d.get('ServiceName','') or d.get('AccountName','') or '')
    print(f"    {mark}[{eid}] {ts}  {EID_DESC.get(eid,''):<22}  {detail[:60]}")
PYEOF

        # Accoda al file JSONL globale (con label sorgente)
        local LBL="$LABEL"
        "$PY3" -c "
import json, sys
events=json.load(open(sys.argv[1]))
for e in events: e['src']=sys.argv[2]
with open(sys.argv[3],'a') as f:
    for e in events: f.write(json.dumps(e)+'\n')
" "$TMP_OUT" "$LBL" "$TMP_ALL" 2>/dev/null || true

        # Conta sospetti
        local SC
        SC=$("$PY3" -c "
import json,sys
SUSP={'4648','4698','4702','4720','7045','4104','1116','1117','4688'}
print(sum(1 for e in json.load(open(sys.argv[1])) if e.get('eid') in SUSP))
" "$TMP_OUT" 2>/dev/null || echo 0)
        SUSP_COUNT=$((SUSP_COUNT + SC))
        rm -f "$TMP_OUT"
    done

    separator
    info "Totale eventi: ${BOLD}$TOTAL_FOUND${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL_FOUND -eq 0 ]] && { warn "Nessun evento trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "evtx")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Genera righe HTML leggendo dal JSONL — mai da argv
    local TMP_ROWS; TMP_ROWS=$(mktemp /tmp/dfir_evtxr_XXXXXX.html)
    "$PY3" - "$TMP_ALL" "$TMP_ROWS" << 'PYEOF' 2>/dev/null || true
import sys, json, html as H
EID_DESC={"4624":"Logon riuscito","4625":"Logon fallito","4648":"Logon credenziali espl.",
           "4688":"Processo creato","4698":"Task creato","4702":"Task modificato",
           "4720":"Account creato","4726":"Account eliminato",
           "4732":"Aggiunto a gruppo loc.","4733":"Rimosso da gruppo loc.",
           "4756":"Aggiunto a gruppo glob.","4757":"Rimosso da gruppo glob.",
           "7045":"Nuovo servizio installato","7034":"Servizio crash","7036":"Svc avv/fermato",
           "4104":"PS Script Block Logging","21":"RDP Logon","23":"RDP Logoff",
           "24":"RDP Disco.","25":"RDP Ricollega","106":"Task registrato",
           "140":"Task aggiornato","141":"Task rimosso","200":"Task eseguito","201":"Task completato",
           "1116":"Malware rilevato","1117":"Defender azione",
           "1118":"Defender scan avviato","1119":"Defender scan completo"}
SUSP={"4648","4698","4702","4720","7045","4104","1116","1117","4688"}
SKIP={'-','-1','0','%%1796','S-1-0-0','%%1833','%%1842','%%1843',''}

events=[]
with open(sys.argv[1]) as f:
    for line in f:
        line=line.strip()
        if line:
            try: events.append(json.loads(line))
            except: pass
events.sort(key=lambda e: e.get('ts',''), reverse=True)

rows=[]
for ev in events:
    eid  = H.escape(str(ev.get('eid','')))
    ts   = H.escape(str(ev.get('ts','')))
    src  = H.escape(str(ev.get('src','')))
    desc = H.escape(EID_DESC.get(ev.get('eid',''),''))
    d    = ev.get('data',{})
    parts=[]
    for k,v in d.items():
        sv=str(v).strip()
        if sv and sv not in SKIP:
            parts.append(f"<div class='drow'><span class='lbl'>{H.escape(k)}</span>"
                         f"<span class='fld mono'>{H.escape(sv[:200])}</span></div>")
    detail=''.join(parts[:8]) or "<span class='dim'>—</span>"
    susp=ev.get('eid','') in SUSP
    rs="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if susp else ""
    ec="bad" if susp else "ok"
    rows.append(f"<tr {rs}><td style='white-space:nowrap;vertical-align:top;padding:.4rem .7rem'>"
                f"<span class='mono {ec}' style='font-size:.8rem;font-weight:700'>{eid}</span><br>"
                f"<span class='dim' style='font-size:.62rem'>{desc}</span></td>"
                f"<td class='mono' style='white-space:nowrap;vertical-align:top;padding:.4rem .7rem;"
                f"font-size:.72rem;color:var(--accent4)'>{ts}</td>"
                f"<td style='vertical-align:top;padding:.35rem .8rem;font-size:.7rem;"
                f"color:var(--text-dim)'>{src}</td>"
                f"<td style='vertical-align:top;padding:.35rem .9rem'>{detail}</td></tr>")

with open(sys.argv[2],'w') as f:
    f.write('\n'.join(rows))
PYEOF

    local TABLE_ROWS
    TABLE_ROWS=$(cat "$TMP_ROWS" 2>/dev/null || echo "")
    rm -f "$TMP_ROWS"

    {
        html_header "Event Log"
        html_page_header "EV" "Event Log <span>Forensics</span>" \
            "Windows\\System32\\winevt\\Logs" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Eventi estratti</div><div class='value'>${TOTAL_FOUND}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <style>
          .lbl{display:inline-block;font-family:var(--mono);font-size:.58rem;color:var(--text-dim);
               min-width:9rem;margin-right:.4rem;flex-shrink:0;vertical-align:top;padding-top:.1rem}
          .drow{display:flex;align-items:flex-start;margin-bottom:.18rem}
          .fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap;flex:1}
        </style>
        <div class='stitle'>Eventi per timestamp — EID · Timestamp · Sorgente · Dettagli</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:11%'>EID</th>
            <th style='width:13%'>Timestamp</th>
            <th style='width:18%'>Log</th>
            <th>Dettagli</th>
          </tr></thead>
          <tbody>${TABLE_ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 13 — Amcache + Shimcache (timeline esecuzione binari)
# ================================================================
module_amcache() {
    section_header "Amcache + Shimcache — Timeline Binari" "$YELLOW"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: ${PY3} -m pip install regipy"
        return 1
    fi


    # ── Amcache ──────────────────────────────────────────────────────
    local AMC_PATH
    AMC_PATH=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "Windows/appcompat/Programs")" "Amcache.hve")
    [[ -z "$AMC_PATH" ]] && \
        AMC_PATH=$(find "$WIN_ROOT" -maxdepth 6 -iname "Amcache.hve" 2>/dev/null | head -1)

    declare -a AMC_ENTRIES=()
    local AMC_COUNT=0

    if [[ -n "$AMC_PATH" ]]; then
        info "Amcache trovato: $AMC_PATH"
        local AMC_DATA
        AMC_DATA=$("$PY3" - "$AMC_PATH" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    # Win10+: InventoryApplicationFile
    try:
        key = hive.get_key('Root\\InventoryApplicationFile')
        for sk in key.iter_subkeys():
            vals = {v.name.lower(): str(v.value) for v in sk.get_values()}
            name    = vals.get('name','')
            path    = vals.get('lowercaselong path', vals.get('filepath',''))
            sha1    = vals.get('fileid','').lstrip('0000')
            pubname = vals.get('publisher','')
            ltime   = vals.get('linkdate','')
            print(f"{name}\t{path}\t{sha1}\t{pubname}\t{ltime}\tInventoryApplicationFile")
    except Exception:
        pass
    # Fallback Win7/8: File entries
    try:
        for root_sub in ['Root\\File','Root\\Programs']:
            try:
                key = hive.get_key(root_sub)
                for vol in key.iter_subkeys():
                    for entry in vol.iter_subkeys():
                        vals = {v.name.lower(): str(v.value) for v in entry.get_values()}
                        path = vals.get('15','') or vals.get('filepath','')
                        sha1 = vals.get('101','').lstrip('0000')
                        ltime = vals.get('17','')
                        print(f"\t{path}\t{sha1}\t\t{ltime}\t{root_sub}")
            except Exception:
                pass
    except Exception:
        pass
except Exception as e:
    print(f"ERROR\t{e}\t\t\t\t")
PYEOF
)
        while IFS=$'\t' read -r AMC_NAME AMC_PATH_F AMC_SHA1 AMC_PUB AMC_TIME AMC_SRC; do
            [[ -z "$AMC_PATH_F" && -z "$AMC_NAME" ]] && continue
            [[ "$AMC_NAME" == "ERROR" ]] && { err "$AMC_PATH_F"; continue; }
            AMC_ENTRIES+=("${AMC_NAME}§${AMC_PATH_F}§${AMC_SHA1}§${AMC_PUB}§${AMC_TIME}§${AMC_SRC}")
            AMC_COUNT=$((AMC_COUNT + 1))
            printf "  ${DIM}%-40s${RESET}  ${CYAN}%-15s${RESET}  %s\n" \
                "${AMC_PATH_F: -40}" "$AMC_TIME" "${AMC_SHA1:0:16}"
        done <<< "$AMC_DATA"
        ok "Amcache: $AMC_COUNT entry"
    else
        warn "Amcache.hve non trovato"
    fi

    echo ""

    # ── Shimcache ─────────────────────────────────────────────────────
    local SYS_HIVE
    SYS_HIVE=$(get_hive "SYSTEM")
    declare -a SHIM_ENTRIES=()
    local SHIM_COUNT=0

    if [[ -n "$SYS_HIVE" ]]; then
        info "Shimcache da hive SYSTEM: $SYS_HIVE"
        local SHIM_DATA
        SHIM_DATA=$("$PY3" - "$SYS_HIVE" << 'PYEOF' 2>/dev/null || true
import sys, struct, datetime
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])

    shim_key = None
    for cs in ['ControlSet001','ControlSet002']:
        for sub in ['Session Manager\\AppCompatCache',
                    'Session Manager\\AppCompatibility']:
            try:
                shim_key = hive.get_key(f'{cs}\\{sub}'); break
            except: continue
        if shim_key: break

    if not shim_key:
        print("ERROR\tChiave AppCompatCache non trovata")
        sys.exit(0)

    for v in shim_key.get_values():
        if v.name.lower() in ('appcompatcache','cachemainkey') and isinstance(v.value, bytes):
            data = v.value
            break
    else:
        print("ERROR\tValore cache non trovato"); sys.exit(0)

    # Parse Win10 format (signature 10ts / 10ba)
    sig = data[:4]
    entries = []
    if sig in (b'\x30\x00\x00\x00', b'\x80\x00\x00\x00'):  # Win10
        offset = 52 if sig == b'\x30\x00\x00\x00' else 52
        while offset < len(data):
            try:
                esig = data[offset:offset+4]
                if esig != b'\x10\x00\x00\x00': break
                path_size = struct.unpack_from('<H', data, offset+8)[0]
                path = data[offset+12:offset+12+path_size].decode('utf-16-le', errors='replace')
                ts_raw = struct.unpack_from('<Q', data, offset+12+path_size)[0]
                ts = ''
                if ts_raw:
                    try:
                        ts = (datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
                    except: pass
                entries.append((path, ts))
                entry_size = struct.unpack_from('<I', data, offset+4)[0]
                offset += 12 + entry_size
            except: break
    else:
        # Win7/8 simple format
        try:
            num = struct.unpack_from('<I', data, 4)[0]
            off = 128
            for _ in range(min(num, 1024)):
                plen = struct.unpack_from('<H', data, off)[0]
                path = data[off+4:off+4+plen].decode('utf-16-le', errors='replace')
                ts_raw = struct.unpack_from('<Q', data, off+4+plen)[0]
                ts = ''
                if ts_raw:
                    try:
                        ts = (datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
                    except: pass
                entries.append((path, ts))
                rec_size = struct.unpack_from('<I', data, off+4+plen+8)[0]
                off += 8 + plen + rec_size
        except: pass

    for path, ts in entries:
        print(f"{path}\t{ts}")

except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
        while IFS=$'\t' read -r SHIM_PATH SHIM_TIME; do
            [[ -z "$SHIM_PATH" ]] && continue
            [[ "$SHIM_PATH" == "ERROR" ]] && { err "$SHIM_TIME"; continue; }
            SHIM_ENTRIES+=("${SHIM_PATH}§${SHIM_TIME}")
            SHIM_COUNT=$((SHIM_COUNT + 1))
            printf "  ${DIM}%-50s${RESET}  %s\n" "${SHIM_PATH: -50}" "${SHIM_TIME:--}"
        done <<< "$SHIM_DATA"
        ok "Shimcache: $SHIM_COUNT entry"
    else
        warn "Hive SYSTEM non trovato — Shimcache saltato"
    fi

    separator
    info "Amcache: ${BOLD}$AMC_COUNT${RESET}  |  Shimcache: ${BOLD}$SHIM_COUNT"
    [[ $((AMC_COUNT + SHIM_COUNT)) -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "amcache")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente (Amcache=k5 LinkDate, Shimcache=k2)
    mapfile -t AMC_ENTRIES  < <(printf '%s\n' "${AMC_ENTRIES[@]}"  | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+5*[""])[4],reverse=True); print("\n".join(L))')
    mapfile -t SHIM_ENTRIES < <(printf '%s\n' "${SHIM_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+["",""])[1],reverse=True); print("\n".join(L))')

    # Amcache rows
    local AMC_ROWS=""
    for E in "${AMC_ENTRIES[@]}"; do
        IFS='§' read -r N P S PUB T SRC <<< "$E"
        local PESC; PESC=$(html_esc "$P")
        local SUSP_S=""
        echo "$P" | grep -qi "temp\|public\|appdata\\\\local\|programdata\|downloads\|\\\\users\\\\" && SUSP_S="class='bad'"
        AMC_ROWS+="<tr>
          <td class='mono' style='font-size:.72rem;word-break:break-all'><span ${SUSP_S}>${PESC}</span></td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${T:--}</td>
          <td class='mono dim' style='font-size:.68rem;word-break:break-all'>${S:0:20}</td>
          <td class='mono dim' style='font-size:.68rem'>$(html_esc "$PUB")</td>
        </tr>"
    done

    # Shimcache rows
    local SHIM_ROWS=""
    for E in "${SHIM_ENTRIES[@]}"; do
        IFS='§' read -r P T <<< "$E"
        local PESC; PESC=$(html_esc "$P")
        local SUSP_S=""
        echo "$P" | grep -qi "temp\|public\|appdata\|programdata\|downloads" && SUSP_S="class='bad'"
        SHIM_ROWS+="<tr>
          <td class='mono' style='font-size:.72rem;word-break:break-all'><span ${SUSP_S}>${PESC}</span></td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${T:--}</td>
        </tr>"
    done

    {
        html_header "Amcache"
        html_page_header "AC" "Amcache + <span>Shimcache</span> Forensics" \
            "Windows\\appcompat\\Programs\\Amcache.hve" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Amcache entry</div><div class='value'>${AMC_COUNT}</div></div>
          <div class='stat info'><div class='label'>Shimcache entry</div><div class='value'>${SHIM_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Amcache — Eseguibili con hash SHA1</div>
        <div class='card'><table>
          <thead><tr><th>Path</th><th>LinkDate</th><th>SHA1 (troncato)</th><th>Publisher</th></tr></thead>
          <tbody>${AMC_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>Shimcache — Eseguibili visti dal sistema</div>
        <div class='card'><table>
          <thead><tr><th>Path</th><th>Timestamp</th></tr></thead>
          <tbody>${SHIM_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 14 — Recycle Bin ($Recycle.Bin)
# ================================================================
module_recycle_bin() {
    section_header "Recycle Bin — File Eliminati" "$GREEN"
    check_win_root || return 1


    # Cerca $Recycle.Bin sul volume (può stare nella root)
    local RB_DIR
    RB_DIR=$(find "$WIN_ROOT" -maxdepth 1 -iname "\$Recycle.Bin" -type d 2>/dev/null | head -1)
    if [[ -z "$RB_DIR" ]]; then
        warn "\$Recycle.Bin non trovato nella root del volume"
        return 0
    fi

    info "Directory: $RB_DIR"
    echo ""

    declare -a ALL_ITEMS=()
    local TOTAL=0 SUSP_COUNT=0

    # Parser $I files — formato binario con path originale e timestamp
    local IPARSE=""
    IPARSE=$(mktemp /tmp/dfir_rb_XXXXXX.py)
    trap '[[ -n "${IPARSE:-}" ]] && rm -f "$IPARSE"' RETURN
    cat > "$IPARSE" << 'PYEOF'
import sys, struct, datetime, os

def parse_ifile(path):
    try:
        with open(path, 'rb') as f:
            data = f.read()
        if len(data) < 24:
            return None, None, None
        ver = struct.unpack_from('<Q', data, 0)[0]
        size = struct.unpack_from('<Q', data, 8)[0]
        ts_raw = struct.unpack_from('<Q', data, 16)[0]
        ts = ''
        if ts_raw:
            try:
                ts = (datetime.datetime(1601,1,1) +
                      datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
            except: pass
        # Path: versione 2 (Win10) ha lunghezza a offset 24
        if ver == 2 and len(data) >= 28:
            plen = struct.unpack_from('<I', data, 24)[0]
            orig_path = data[28:28+plen*2].decode('utf-16-le', errors='replace').rstrip('\x00')
        else:
            orig_path = data[24:].decode('utf-16-le', errors='replace').rstrip('\x00')
        return orig_path, size, ts
    except Exception as e:
        return None, None, None

path = sys.argv[1]
orig, size, ts = parse_ifile(path)
if orig:
    print(f"{orig}\t{size or 0}\t{ts or ''}")
PYEOF

    # Scansiona ogni SID
    for SID_DIR in "$RB_DIR"/*/; do
        local SID; SID=$(basename "$SID_DIR")
        local SID_COUNT=0

        for IFILE in "$SID_DIR"/\$I*; do
            [[ -f "$IFILE" ]] || continue
            local RESULT
            RESULT=$("$PY3" "$IPARSE" "$IFILE" 2>/dev/null || true)
            [[ -z "$RESULT" ]] && continue

            IFS=$'\t' read -r ORIG_PATH FSIZE DEL_TIME <<< "$RESULT"
            [[ -z "$ORIG_PATH" ]] && continue

            local FNAME; FNAME=$(basename "${IFILE/$'\$I'/$'\$R'}")
            local IS_SUSP=false
            echo "$ORIG_PATH" | grep -qi "system32\|passwd\|shadow\|lsass\|sam\|ntds\|\.ps1\|\.bat\|\.vbs\|\.exe\|\.dll" && IS_SUSP=true

            ALL_ITEMS+=("${SID}§${ORIG_PATH}§${FSIZE}§${DEL_TIME}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            $IS_SUSP && SUSP_COUNT=$((SUSP_COUNT + 1))
            SID_COUNT=$((SID_COUNT + 1))

            if $IS_SUSP; then
                printf "  ${RED}[!] %s${RESET}\n" "$ORIG_PATH"
                printf "      ${DIM}SID: %-40s  Eliminato: %s  Size: %s B${RESET}\n" "$SID" "$DEL_TIME" "$FSIZE"
            else
                printf "  ${DIM}%-60s${RESET}  %s\n" "${ORIG_PATH: -60}" "${DEL_TIME:--}"
            fi
        done
        [[ $SID_COUNT -gt 0 ]] && ok "  SID $SID: $SID_COUNT file"
    done

    separator
    info "File nel cestino: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Cestino vuoto o nessun \$I file trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "recycle_bin")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data eliminazione decrescente globale
    mapfile -t ALL_ITEMS < <(printf '%s\n' "${ALL_ITEMS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+4*[""])[3],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_ITEMS[@]}"; do
        IFS='§' read -r SID ORIG_PATH FSIZE DEL_TIME IS_SUSP <<< "$E"
        local ROW_STYLE="" PATH_CSS=""
        [[ "$IS_SUSP" == "true" ]] && {
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
            PATH_CSS="bad"
        }
        local SIZE_HR
        if [[ "$FSIZE" -gt 1073741824 ]] 2>/dev/null; then
            SIZE_HR="$(( FSIZE / 1073741824 )) GB"
        elif [[ "$FSIZE" -gt 1048576 ]] 2>/dev/null; then
            SIZE_HR="$(( FSIZE / 1048576 )) MB"
        elif [[ "$FSIZE" -gt 1024 ]] 2>/dev/null; then
            SIZE_HR="$(( FSIZE / 1024 )) KB"
        else
            SIZE_HR="${FSIZE} B"
        fi
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono dim' style='font-size:.68rem;white-space:nowrap'>${SID}</td>
          <td class='mono ${PATH_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$ORIG_PATH")</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${DEL_TIME:--}</td>
          <td class='mono mid' style='white-space:nowrap;font-size:.72rem'>${SIZE_HR}</td>
        </tr>"
    done

    {
        html_header "Recycle Bin"
        html_page_header "RB" "Recycle Bin <span>Forensics</span>" \
            "\$Recycle.Bin\\\<SID>\\\$I*" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File eliminati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>File nel Cestino — SID · Path originale · Data eliminazione · Dimensione</div>
        <div class='card'><table>
          <thead><tr><th style='width:16%'>SID</th><th>Path originale</th><th style='width:14%'>Eliminato</th><th style='width:8%'>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 15 — WMI Subscriptions (persistenza invisibile)
# ================================================================
module_wmi() {
    section_header "WMI Subscriptions — Persistenza Invisibile" "$RED"
    check_win_root || return 1


    # Repository WMI
    local WMI_DIR
    WMI_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/wbem/Repository")
    if [[ -z "$WMI_DIR" || ! -d "$WMI_DIR" ]]; then
        warn "Repository WMI non trovato"
        return 0
    fi

    info "Repository: $WMI_DIR"
    echo ""

    # Cerca i file OBJECTS.DATA e INDEX.BTR
    local OBJECTS_FILE
    OBJECTS_FILE=$(ci_find_file "$WMI_DIR" "OBJECTS.DATA")
    [[ -z "$OBJECTS_FILE" ]] && \
        OBJECTS_FILE=$(find "$WMI_DIR" -iname "OBJECTS.DATA" 2>/dev/null | head -1)

    if [[ -z "$OBJECTS_FILE" ]]; then
        warn "OBJECTS.DATA non trovato — analisi tramite strings sul repository"
    fi

    declare -a WMI_ENTRIES=()
    local TOTAL=0 SUSP_COUNT=0

    # Parsing: usa strings per estrarre pattern WMI sospetti dal repository
    # (parsing completo del formato CIM richiede librerie specializzate)
    local WMI_DATA
    if [[ -n "$OBJECTS_FILE" ]]; then
        WMI_DATA=$("$PY3" - "$OBJECTS_FILE" << 'PYEOF' 2>/dev/null || true
import sys, re

# Legge il file come binario e cerca pattern WMI subscription
with open(sys.argv[1], 'rb') as f:
    raw = f.read()

text = raw.decode('utf-16-le', errors='replace')
text2 = raw.decode('latin-1', errors='replace')

found = set()

# Cerca CommandLineEventConsumer (esegue comandi)
for m in re.finditer(r'CommandLineTemplate\s*=\s*"([^"]{4,})"', text + text2, re.I):
    found.add(('CommandLineConsumer', m.group(1)[:300]))

# Cerca ActiveScriptEventConsumer (script VBS/PS)
for m in re.finditer(r'ScriptText\s*=\s*"([^"]{4,})"', text + text2, re.I):
    found.add(('ScriptConsumer', m.group(1)[:300]))

# Cerca __EventFilter queries
for m in re.finditer(r'Query\s*=\s*"([^"]{10,})"', text + text2, re.I):
    q = m.group(1)
    if any(k in q.lower() for k in ['win32_process','logon','startup','timer','permanentevent']):
        found.add(('EventFilter', q[:300]))

# Cerca nomi consumer e filter
for m in re.finditer(r'Name\s*=\s*"([A-Za-z0-9_\-\.]{3,64})"', text + text2, re.I):
    name = m.group(1)
    if any(k in name.lower() for k in ['persist','backdoor','evil','shell','exec','powershell','cmd','script','mof']):
        found.add(('SuspiciousName', name))

for ctype, val in sorted(found):
    print(f"{ctype}\t{val}")
PYEOF
)
    fi

    # Fallback: strings su tutto il repository
    if [[ -z "$WMI_DATA" || $(echo "$WMI_DATA" | wc -l) -lt 2 ]]; then
        info "Fallback: strings sul repository..."
        WMI_DATA=$(find "$WMI_DIR" -type f 2>/dev/null | \
            xargs strings 2>/dev/null | \
            grep -iE "CommandLineTemplate|ScriptText|ActiveScript|EventFilter|EventConsumer|FilterToConsumer" | \
            head -100 | \
            awk '{print "StringsMatch\t" $0}' || true)
    fi

    while IFS=$'\t' read -r WMI_TYPE WMI_VAL; do
        [[ -z "$WMI_TYPE" ]] && continue
        local IS_SUSP=false
        echo "$WMI_VAL" | grep -qi "powershell\|cmd\|wscript\|mshta\|certutil\|base64\|http\|download\|-enc\|bypass" && IS_SUSP=true
        [[ "$WMI_TYPE" == "CommandLineConsumer" || "$WMI_TYPE" == "ScriptConsumer" ]] && IS_SUSP=true

        WMI_ENTRIES+=("${WMI_TYPE}§${WMI_VAL}§${IS_SUSP}")
        TOTAL=$((TOTAL + 1))
        $IS_SUSP && SUSP_COUNT=$((SUSP_COUNT + 1))

        if $IS_SUSP; then
            echo -e "  ${RED}[!] ${WMI_TYPE}${RESET}"
            echo -e "      ${YELLOW}${WMI_VAL:0:120}${RESET}"
        else
            printf "  ${DIM}%-25s${RESET}  %s\n" "$WMI_TYPE" "${WMI_VAL:0:80}"
        fi
    done <<< "$WMI_DATA"

    separator
    info "Pattern WMI trovati: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessuna subscription WMI sospetta trovata."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wmi")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local ROWS=""
    for E in "${WMI_ENTRIES[@]}"; do
        IFS='§' read -r WMI_TYPE WMI_VAL IS_SUSP <<< "$E"
        local ROW_STYLE="" VAL_CSS="mono fld"
        [[ "$IS_SUSP" == "true" ]] && {
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
            VAL_CSS="mono fld bad"
        }
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono' style='white-space:nowrap;vertical-align:top;padding-top:.45rem;color:var(--accent4)'>${WMI_TYPE}</td>
          <td style='padding:.35rem .9rem'><span class='${VAL_CSS}'>$(html_esc "$WMI_VAL")</span></td>
        </tr>"
    done

    {
        html_header "WMI"
        html_page_header "WM" "WMI <span>Subscriptions</span> Forensics" \
            "Windows\\System32\\wbem\\Repository\\OBJECTS.DATA" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Pattern trovati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <style>
          .fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}
        </style>
        <div class='stitle'>WMI Event Subscriptions — Tipo · Contenuto</div>
        <div class='card'><table>
          <thead><tr><th style='width:20%'>Tipo</th><th>Valore / Query / Script</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>ℹ ATT&amp;CK T1546.003 — WMI Event Subscription</div>
          <div style='color:var(--text-mid);line-height:1.8'>
            Meccanismo di persistenza invisibile: non compare in Run keys, Task, né Services.<br>
            Classi chiave: <span class='mono'>__EventFilter</span> · <span class='mono'>CommandLineEventConsumer</span> · <span class='mono'>ActiveScriptEventConsumer</span> · <span class='mono'>__FilterToConsumerBinding</span>
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 16 — SRUM (System Resource Usage Monitor)
# ================================================================
module_srum() {
    section_header "SRUM — System Resource Usage Monitor" "$BLUE"
    check_win_root || return 1

    local SRUM_PATH
    SRUM_PATH=$(ci_find_file \
        "$(ci_find_dir "$WIN_ROOT" "Windows/System32/sru")" \
        "SRUDB.dat")
    [[ -z "$SRUM_PATH" ]] && \
        SRUM_PATH=$(find "$WIN_ROOT" -maxdepth 8 -iname "SRUDB.dat" 2>/dev/null | head -1)

    if [[ -z "$SRUM_PATH" ]]; then
        warn "SRUDB.dat non trovato"
        return 0
    fi

    info "SRUM DB: $SRUM_PATH"

    # Verifica pyesedb o libesedb-python
    local HAS_ESE=false
    "$PY3" -c "import pyesedb" 2>/dev/null && HAS_ESE=true

    if ! $HAS_ESE; then
        warn "pyesedb non disponibile — parsing tramite strings (risultati parziali)"
        warn "Per parsing completo: pip install pyesedb  (richiede libesedb)"
        echo ""
        # Fallback strings: cerca nomi applicazione e byte trasferiti
        info "Estrazione tramite strings..."
        local STRINGS_OUT
        STRINGS_OUT=$(strings "$SRUM_PATH" 2>/dev/null | \
            grep -iE '\.(exe|dll|bat|ps1|vbs|py)|\\Device\\|\\\\[A-Za-z]' | \
            grep -v '^.\{1\}$' | sort -u | head -200 || true)

        echo "$STRINGS_OUT" | while IFS= read -r LINE; do
            printf "  ${DIM}%s${RESET}\n" "$LINE"
        done | head -50

        separator
        info "Strings estratte (prime 200 uniche). Per analisi completa installa pyesedb."
        ask_yn "Generare report HTML (strings)?" || return 0

        local REPORT_HTML
        local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "srum")
        local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
        local ROWS=""
        echo "$STRINGS_OUT" | while IFS= read -r LINE; do
            local ESC_LINE; ESC_LINE=$(printf '%s' "$LINE" | sed 's/&/\&amp;/g;s/</\&lt;/g;s/>/\&gt;/g')
            ROWS+="<tr><td class='mono' style='font-size:.72rem;word-break:break-all'>${ESC_LINE}</td></tr>"
        done

        {
            html_header "SRUM"
            html_page_header "SR" "SRUM <span>Database</span> Forensics" \
                "Windows\\System32\\sru\\SRUDB.dat" "$SCAN" "$WIN_ROOT"
            echo "<div class='statsbar'>
              <div class='stat'><div class='label'>Modalità</div><div class='value' style='font-size:1rem;padding-top:.4rem;color:var(--accent4)'>strings</div></div>
            </div><main>
            <div class='stitle'>Stringhe estratte da SRUDB.dat (pyesedb non disponibile)</div>
            <div class='card'><table><thead><tr><th>Stringa</th></tr></thead><tbody>${ROWS}</tbody></table></div>
            <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px'>
              <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Per analisi completa con traffico di rete e CPU per applicazione:</div>
              <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>pip install pyesedb<br>oppure usa: srum-dump (Mark Baggett) / ese2csv</div>
            </div></main>"
            html_footer "$SCAN" "$WIN_ROOT"
        } > "$REPORT_HTML"
        register_report "$REPORT_HTML"
        ok "Report salvato: ${BOLD}$REPORT_HTML"
        open_report_prompt "$REPORT_HTML"
        return 0
    fi

    # Parser completo con pyesedb
    info "Parsing ESE con pyesedb..."
    local SRUM_JSON
    SRUM_JSON=$("$PY3" - "$SRUM_PATH" << 'PYEOF' 2>/dev/null || echo "[]"
import sys, json, datetime
try:
    import pyesedb
except ImportError:
    print("[]"); sys.exit()

try:
    db = pyesedb.open(sys.argv[1])
except Exception as e:
    print(json.dumps([{"error": str(e)}])); sys.exit()

# GUID tabelle principali
TABLE_NAMES = {
    '{5C8CF1C7-7257-4F13-B223-970EF5939312}': 'App Resource Usage',
    '{973F5D5C-1D90-4944-BE8E-24B94231A174}': 'Network Data Usage',
    '{D10CA2FE-6FCF-4F6D-848E-B2E99266FA89}': 'Push Notifications',
    '{DD6636C4-8929-4683-974E-22C046A43763}': 'Network Connectivity',
    '{FEE4E14F-02A9-4550-B5CE-5FA2DA202E37}': 'Energy Usage',
}

results = []
for i in range(db.number_of_tables):
    tbl = db.get_table(i)
    tname = tbl.name
    label = TABLE_NAMES.get(tname, tname)
    if 'SruDbIdMapTable' in tname or 'MSysObjects' in tname:
        continue
    cols = [tbl.get_column(j).name for j in range(tbl.number_of_columns)]
    rows_data = []
    rec = tbl.first_record
    count = 0
    while rec and count < 300:
        row = {}
        for j, col in enumerate(cols):
            try:
                val = rec.get_value_data_as_string(j)
                if val: row[col] = val[:200]
            except: pass
        if row:
            rows_data.append(row)
        rec = rec.next_record
        count += 1
    if rows_data:
        results.append({'table': label, 'guid': tname, 'rows': rows_data})

print(json.dumps(results))
PYEOF
)

    local TABLE_COUNT
    TABLE_COUNT=$("$PY3" -c "import sys,json; d=json.loads(sys.stdin.read()); print(len(d))" <<< "$SRUM_JSON" 2>/dev/null || echo 0)
    local ROW_COUNT
    ROW_COUNT=$("$PY3" -c "import sys,json; d=json.loads(sys.stdin.read()); print(sum(len(t.get('rows',[])) for t in d))" <<< "$SRUM_JSON" 2>/dev/null || echo 0)

    info "Tabelle SRUM: $TABLE_COUNT  |  Record: $ROW_COUNT"
    echo ""

    "$PY3" - "$SRUM_JSON" << 'PYEOF' 2>/dev/null || true
import sys, json
tables = json.loads(sys.argv[1])
for tbl in tables[:4]:
    print(f"  [{tbl['table']}]  {len(tbl.get('rows',[]))} record")
    for row in tbl.get('rows',[])[:5]:
        interesting = {k:v for k,v in row.items() if any(
            x in k.lower() for x in ['app','exe','byte','sent','recv','connect','user','time']
        )}
        if interesting:
            parts = [f"{k}={v[:40]}" for k,v in list(interesting.items())[:4]]
            print(f"    {', '.join(parts)}")
    print()
PYEOF

    separator
    info "Record totali: ${BOLD}$ROW_COUNT"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "srum")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local SECTIONS_HTML
    SECTIONS_HTML=$("$PY3" - "$SRUM_JSON" << 'PYEOF' 2>/dev/null || echo ""
import sys, json, html as html_mod
tables = json.loads(sys.argv[1])
out = []
for tbl in tables:
    label = html_mod.escape(tbl.get('table',''))
    rows  = tbl.get('rows', [])
    if not rows: continue
    all_cols = list(dict.fromkeys(k for r in rows for k in r.keys()))
    show_cols = all_cols[:10]
    th = ''.join(f"<th>{html_mod.escape(c)}</th>" for c in show_cols)
    trs = []
    for row in rows[:200]:
        tds = ''.join(f"<td class='mono' style='font-size:.68rem;word-break:break-all;max-width:200px'>{html_mod.escape(str(row.get(c,''))[:100])}</td>" for c in show_cols)
        trs.append(f"<tr>{tds}</tr>")
    out.append(f"""
<div class='stitle' style='margin-top:2rem'>{label} ({len(rows)} record)</div>
<div class='card' style='overflow-x:auto'><table style='min-width:600px'>
  <thead><tr>{th}</tr></thead>
  <tbody>{''.join(trs)}</tbody>
</table></div>""")
print('\n'.join(out))
PYEOF
)

    {
        html_header "SRUM"
        html_page_header "SR" "SRUM <span>Database</span> Forensics" \
            "Windows\\System32\\sru\\SRUDB.dat" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Tabelle</div><div class='value'>${TABLE_COUNT}</div></div>
          <div class='stat info'><div class='label'>Record</div><div class='value'>${ROW_COUNT}</div></div>
        </div><main>${SECTIONS_HTML}</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 17 — Browser History (Chrome / Edge / Firefox)
# ================================================================
module_browser() {
    section_header "Browser History — Chrome / Edge / Firefox" "$CYAN"
    check_win_root || return 1


    # Verifica sqlite3 disponibile
    if ! "$PY3" -c "import sqlite3" 2>/dev/null; then
        err "sqlite3 non disponibile nel Python rilevato"
        return 1
    fi

    declare -a PROFILES=(
        # [label]  [path relativo nella home utente]  [tipo]
        "Chrome|AppData/Local/Google/Chrome/User Data/Default/History|chromium"
        "Edge|AppData/Local/Microsoft/Edge/User Data/Default/History|chromium"
        "Chromium|AppData/Local/Chromium/User Data/Default/History|chromium"
        "Brave|AppData/Local/BraveSoftware/Brave-Browser/User Data/Default/History|chromium"
        "Firefox|AppData/Roaming/Mozilla/Firefox/Profiles|firefox"
    )

    ALL_VISITS=()
    TOTAL=0
    SUSP_COUNT=0

    local TMP_ROWS
    TMP_ROWS=$(mktemp /tmp/dfir_br_rows_XXXXXX.html 2>/dev/null) || {
        err "Impossibile creare file temporaneo per le righe HTML"
        return 1
    }

    local HIST_PARSER=""
    HIST_PARSER=$(mktemp /tmp/dfir_brhist_XXXXXX.py 2>/dev/null) || {
        err "Impossibile creare file temporaneo per il parser browser"
        rm -f "$TMP_ROWS" 2>/dev/null
        return 1
    }

    trap 'rm -f "${HIST_PARSER:-}" "${TMP_ROWS:-}" 2>/dev/null' RETURN
    cat > "$HIST_PARSER" << 'PYEOF'
import sys, sqlite3, shutil, os, tempfile, datetime, json

mode     = sys.argv[1]   # chromium | firefox
db_path  = sys.argv[2]
browser  = sys.argv[3]
user     = sys.argv[4]

SUSP_DOMAINS = ['pastebin','transfer.sh','mega.nz','anonfiles','ufile.io',
                'ngrok','serveo','temp-mail','guerrillamail',
                '10minutemail','hastebin','ghostbin']
SUSP_PATTERNS = ['.exe','.ps1','.bat','.vbs','.dll','.hta','.msi',
                 'powershell','base64','invoke','download','rat','trojan',
                 'metasploit','cobalt','mimikatz','payload']

def is_susp(url):
    ul = url.lower()
    return any(p in ul for p in SUSP_DOMAINS + SUSP_PATTERNS)

def chrome_time(t):
    if not t: return ''
    try:
        return (datetime.datetime(1601,1,1) +
                datetime.timedelta(microseconds=t)).strftime('%Y-%m-%d %H:%M:%S')
    except: return ''

# Copia DB in tmp per evitare lock
tmp = tempfile.mktemp(suffix='.db')
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    results = []
    if mode == 'chromium':
        cur.execute("""
            SELECT v.visit_time, u.url, u.title, u.visit_count
            FROM visits v JOIN urls u ON v.url=u.id
            ORDER BY v.visit_time DESC LIMIT 8000
        """)
        for row in cur.fetchall():
            ts  = chrome_time(row[0])
            url = row[1] or ''
            title = row[2] or ''
            cnt   = row[3] or 0
            susp  = is_susp(url)
            results.append({'ts':ts,'url':url,'title':title[:100],
                            'visits':cnt,'browser':browser,'user':user,'susp':susp})
    elif mode == 'firefox':
        cur.execute("""
            SELECT h.last_visit_date, p.url, p.title, p.visit_count
            FROM moz_historyvisits h JOIN moz_places p ON h.place_id=p.id
            ORDER BY h.last_visit_date DESC LIMIT 8000
        """)
        for row in cur.fetchall():
            ts_raw = row[0]
            ts = ''
            if ts_raw:
                try:
                    ts = (datetime.datetime(1970,1,1) +
                          datetime.timedelta(microseconds=ts_raw)).strftime('%Y-%m-%d %H:%M:%S')
                except: pass
            url   = row[1] or ''
            title = row[2] or ''
            cnt   = row[3] or 0
            susp  = is_susp(url)
            results.append({'ts':ts,'url':url,'title':title[:100],
                            'visits':cnt,'browser':browser,'user':user,'susp':susp})
    conn.close()
    print(json.dumps(results))
except Exception as e:
    print(json.dumps([{"error": str(e), "browser": browser, "user": user}]))
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        for PROFILE_SPEC in "${PROFILES[@]}"; do
            IFS='|' read -r BLABEL BPATH BTYPE <<< "$PROFILE_SPEC"

            if [[ "$BTYPE" == "firefox" ]]; then
                # Firefox: cerca in tutti i profili
                local FF_BASE
                FF_BASE=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Mozilla/Firefox/Profiles")
                [[ -z "$FF_BASE" ]] && continue
                for FF_PROF in "$FF_BASE"/*/; do
                    local FF_DB="${FF_PROF}places.sqlite"
                    [[ -f "$FF_DB" ]] || continue
                    local JSON_OUT
                    JSON_OUT=$("$PY3" "$HIST_PARSER" "firefox" "$FF_DB" "Firefox" "$USERNAME" 2>/dev/null || echo "[]")
                    _process_browser_json "$JSON_OUT" "$USERNAME" "$BLABEL"
                done
            else
                # Chromium-based
                local HIST_FILE
                HIST_FILE=$(ci_find_dir "$USER_DIR" "$(dirname "$BPATH")")
                [[ -z "$HIST_FILE" ]] && continue
                HIST_FILE="${HIST_FILE}/$(basename "$BPATH")"
                [[ -f "$HIST_FILE" ]] || continue
                local JSON_OUT
                JSON_OUT=$("$PY3" "$HIST_PARSER" "chromium" "$HIST_FILE" "$BLABEL" "$USERNAME" 2>/dev/null || echo "[]")
                _process_browser_json "$JSON_OUT" "$USERNAME" "$BLABEL"
            fi
        done
    done < <(get_user_homes)

    separator
    info "Visite totali: ${BOLD}$TOTAL${RESET}  |  Sospette: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessuna history browser trovata."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "browser")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Chiamata del parser su tutte le home (questo aggiungerà l'HTML a TMP_ROWS e modificherà TOTAL/SUSP_COUNT)


    {
        html_header "Browser History"
        html_page_header "BR" "Browser <span>History</span> Forensics" \
            "Chrome · Edge · Firefox · Brave" "$SCAN" "$WIN_ROOT"
        cat << 'FILTERHTML'
<style>
  .fbtn{background:var(--bg3);border:1px solid var(--border);color:var(--text-mid);
    font-family:var(--mono);font-size:.72rem;padding:.35rem .9rem;border-radius:4px;
    cursor:pointer;transition:all .15s}
  .fbtn:hover{border-color:var(--accent);color:var(--accent)}
  .fbtn-on{border-color:var(--accent2)!important;color:var(--accent2)!important;
    background:rgba(255,123,114,.12)!important}
  .user-sep td{padding:.4rem 1rem;background:var(--bg4);border-top:2px solid var(--border);
    border-bottom:1px solid var(--border)}
  .user-sep-first td{border-top:none}
  .sep-label{font-family:var(--mono);font-size:.68rem;text-transform:uppercase;
    letter-spacing:.15em;color:var(--accent);font-weight:700}
</style>
<div class='statsbar' style='align-items:center;flex-wrap:wrap'>
  <div class='stat'><div class='label'>Visite totali</div>
    <div class='value' id='cnt-all' style='cursor:pointer' onclick="setFilter('all')" title='Mostra tutto'>0</div></div>
  <div class='stat'><div class='label'>⚠ Sospette — clicca per filtrare</div>
    <div class='value' id='cnt-susp' style='color:var(--accent2);cursor:pointer' onclick="setFilter('susp')" title='Filtra solo sospette'>0</div></div>
  <div style='margin-left:auto;display:flex;gap:.6rem;align-items:center'>
    <button id='btn-all'  class='fbtn fbtn-on' onclick="setFilter('all')">Tutte</button>
    <button id='btn-susp' class='fbtn'          onclick="setFilter('susp')">⚠ Solo sospette</button>
    <span id='finfo' style='font-family:var(--mono);font-size:.68rem;color:var(--text-dim);margin-left:.6rem'></span>
  </div>
</div>
FILTERHTML
        echo "<main>
        <style>.fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}</style>
        <div class='stitle'>Cronologia navigazione — Timestamp · Utente/Browser · URL · Titolo · Visite</div>
        <div class='card'><table id='brtable'>
          <thead><tr>
            <th style='width:12%'>Timestamp</th>
            <th style='width:12%'>Utente/Browser</th>
            <th>URL</th>
            <th style='width:18%'>Titolo</th>
            <th style='width:5%'>N.</th>
          </tr></thead>
          <tbody id='brtbody'>"
        cat "$TMP_ROWS"
        echo "</tbody>
        </table></div></main>"
        cat << 'FILTERJS'
<script>
(function(){
  var allRows = Array.from(document.querySelectorAll('#brtbody tr'));
  var dataRows = allRows.filter(function(r){ return r.dataset.susp !== 'sep'; });
  var sepRows  = allRows.filter(function(r){ return r.dataset.susp === 'sep'; });
  var total = dataRows.length;
  var susp  = dataRows.filter(function(r){ return r.dataset.susp === 'true'; }).length;
  document.getElementById('cnt-all').textContent  = total;
  document.getElementById('cnt-susp').textContent = susp;
  function setFilter(mode){
    window._brFilter = mode;
    var vis = 0;
    // Prima nascondi/mostra le righe dati
    dataRows.forEach(function(r){
      var show = mode === 'all' || r.dataset.susp === 'true';
      r.style.display = show ? '' : 'none';
      if(show) vis++;
    });
    // Separatori: visibili solo se almeno una riga del loro blocco è visibile
    sepRows.forEach(function(sep){
      var next = sep.nextElementSibling;
      var hasVisible = false;
      while(next && next.dataset.susp !== 'sep'){
        if(next.style.display !== 'none'){ hasVisible = true; break; }
        next = next.nextElementSibling;
      }
      sep.style.display = hasVisible ? '' : 'none';
    });
    document.getElementById('btn-all').className  = 'fbtn' + (mode === 'all'  ? ' fbtn-on' : '');
    document.getElementById('btn-susp').className = 'fbtn' + (mode === 'susp' ? ' fbtn-on' : '');
    document.getElementById('finfo').textContent  =
      mode === 'susp' ? 'Mostrando ' + vis + ' sospette su ' + total + ' totali'
                      : 'Mostrando tutte le ' + total + ' voci';
  }
  window.setFilter = setFilter;
  setFilter('all');
})();
</script>
FILTERJS
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# Funzione helper usata da module_browser
# Funzione helper usata da module_browser
# Aggiorna variabili totali e accoda l'HTML per il report in TMP_ROWS
_process_browser_json() {
    local JSON_OUT="$1"
    local USERNAME="$2"
    local BLABEL="$3"

    # Facciamo fare un job completo a Python: restituisce COUNT su stdout,
    # e scrive l'HTML convertito su file. Ritorna pure il count sospetti formattato come COUNT|SUSP
    local COUNTS
    COUNTS=$(printf '%s' "$JSON_OUT" | "$PY3" -c "
import sys, json, html
try:
    data = json.loads(sys.stdin.read())
except:
    sys.exit(0)

d = [x for x in data if 'error' not in x][:8000]
if not d:
    print('0|0')
    sys.exit(0)

susp_count = sum(1 for e in d if e.get('susp', False))
print(f'{len(d)}|{susp_count}')

tmp_file = sys.argv[1]
with open(tmp_file, 'a', encoding='utf-8') as f:
    f.write(\"<tr data-susp='sep' class='user-sep'><td colspan='5'><span class='sep-label'>▶ \" + html.escape(sys.argv[2]) + \" / \" + html.escape(sys.argv[3]) + \"</span></td></tr>\\n\")
    for v in d:
        is_susp = v.get('susp', False)
        susp_str = 'true' if is_susp else 'false'
        row_style = \"style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'\" if is_susp else ''
        url_css = \"mono fld bad\" if is_susp else \"mono fld\"
        u = html.escape(v.get('url','')[:300])
        t = html.escape(v.get('title','')[:100])
        ts = html.escape(v.get('ts',''))
        cnt = str(v.get('visits',0))
        f.write(f\"<tr data-susp='{susp_str}' {row_style}><td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{ts}</td><td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{html.escape(sys.argv[2])} / {html.escape(sys.argv[3])}</td><td style='word-break:break-all;font-size:.72rem'><span class='{url_css}'>{u}</span></td><td style='font-size:.7rem;color:var(--text-mid);max-width:200px;word-break:break-word'>{t}</td><td class='mono mid' style='white-space:nowrap;font-size:.7rem;text-align:center'>{cnt}</td></tr>\\n\")
" "$TMP_ROWS" "$USERNAME" "$BLABEL" 2>/dev/null)

    [[ -z "$COUNTS" || "$COUNTS" == "0|0" ]] && return

    local P_COUNT="${COUNTS%|*}"
    local P_SUSP="${COUNTS#*|}"

    ok "$USERNAME / $BLABEL — $P_COUNT visite"
    TOTAL=$((TOTAL + P_COUNT))
    SUSP_COUNT=$((SUSP_COUNT + P_SUSP))

    # Stampa in console un riepilogo rapido delle sospette
    printf '%s' "$JSON_OUT" | "$PY3" -c "
import sys, json
data = json.loads(sys.stdin.read())
susp = [x for x in data if x.get('susp') and 'error' not in x][:5]
normal = [x for x in data if not x.get('susp') and 'error' not in x][:5]
for v in susp:
    print(f\"  \033[0;31m[!]\033[0m {v.get('ts','')}  {v.get('url','')[:80]}\")
for v in normal:
    print(f\"      {v.get('ts','')}  {v.get('url','')[:80]}\")
" 2>/dev/null || true
}

# ================================================================
#  MODULO 18 — UserAssist + RunMRU + TypedPaths + WordWheelQuery
# ================================================================
module_userassist() {
    section_header "UserAssist / RunMRU / TypedPaths — Attività Utente" "$MAGENTA"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: ${PY3} -m pip install regipy"
        return 1
    fi


    local TOTAL_UA=0 TOTAL_RUN=0 TOTAL_TP=0 TOTAL_WW=0
    declare -a UA_ENTRIES=()   # UserAssist
    declare -a RUN_ENTRIES=()  # RunMRU
    declare -a TP_ENTRIES=()   # TypedPaths
    declare -a WW_ENTRIES=()   # WordWheelQuery

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && { dim_msg "$USERNAME — NTUSER.DAT non trovato"; continue; }

        info "Parsing NTUSER.DAT: $USERNAME"

        # ── UserAssist ───────────────────────────────────────────────
        local UA_DATA
        UA_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, json, codecs, struct, datetime
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    UA_KEY = 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\UserAssist'
    try:
        ua_root = hive.get_key(UA_KEY)
    except:
        sys.exit(0)
    for guid_key in ua_root.iter_subkeys():
        try:
            count_key = hive.get_key(f'{UA_KEY}\\{guid_key.name}\\Count')
        except:
            continue
        for v in count_key.get_values():
            # Decodifica ROT13 del nome
            name = codecs.decode(v.name, 'rot_13')
            val  = v.value
            ts   = ''
            count = 0
            if isinstance(val, bytes) and len(val) >= 16:
                try:
                    count = struct.unpack_from('<I', val, 4)[0]
                    ft    = struct.unpack_from('<Q', val, 8)[0]
                    if ft > 0:
                        ts = (datetime.datetime(1601,1,1) +
                              datetime.timedelta(microseconds=ft//10)).strftime('%Y-%m-%d %H:%M:%S')
                except: pass
            if name and '{' not in name and count > 0:
                print(f"{user}\t{name}\t{count}\t{ts}")
except Exception as e:
    pass
PYEOF
)
        while IFS=$'\t' read -r U N C T; do
            [[ -z "$N" ]] && continue
            UA_ENTRIES+=("${U}§${N}§${C}§${T}")
            TOTAL_UA=$((TOTAL_UA + 1))
        done <<< "$UA_DATA"

        # ── RunMRU ───────────────────────────────────────────────────
        local RUN_DATA
        RUN_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    try:
        key = hive.get_key('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU')
        mru_order = ''
        entries = {}
        for v in key.get_values():
            if v.name.lower() == 'mrulist':
                mru_order = str(v.value)
            else:
                entries[v.name] = str(v.value).rstrip('\\1')
        # Stampa in ordine MRU
        for ch in mru_order:
            if ch in entries:
                print(f"{user}\t{entries[ch]}")
    except: pass
except: pass
PYEOF
)
        while IFS=$'\t' read -r U CMD; do
            [[ -z "$CMD" ]] && continue
            RUN_ENTRIES+=("${U}§${CMD}")
            TOTAL_RUN=$((TOTAL_RUN + 1))
        done <<< "$RUN_DATA"

        # ── TypedPaths ───────────────────────────────────────────────
        local TP_DATA
        TP_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    try:
        key = hive.get_key('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\TypedPaths')
        for v in key.get_values():
            print(f"{user}\t{v.name}\t{v.value}")
    except: pass
except: pass
PYEOF
)
        while IFS=$'\t' read -r U K V; do
            [[ -z "$V" ]] && continue
            TP_ENTRIES+=("${U}§${K}§${V}")
            TOTAL_TP=$((TOTAL_TP + 1))
        done <<< "$TP_DATA"

        # ── WordWheelQuery (ricerche Start) ──────────────────────────
        local WW_DATA
        WW_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    try:
        key = hive.get_key('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\WordWheelQuery')
        mru_order = ''
        entries = {}
        for v in key.get_values():
            if v.name.lower() == 'mrulistex':
                continue
            val = v.value
            if isinstance(val, bytes):
                try:
                    val = val.decode('utf-16-le', errors='replace').rstrip('\x00')
                except:
                    val = repr(val)
            entries[v.name] = str(val)
        for k,v in sorted(entries.items()):
            if v:
                print(f"{user}\t{v}")
    except: pass
except: pass
PYEOF
)
        while IFS=$'\t' read -r U TERM; do
            [[ -z "$TERM" ]] && continue
            WW_ENTRIES+=("${U}§${TERM}")
            TOTAL_WW=$((TOTAL_WW + 1))
        done <<< "$WW_DATA"

    done < <(get_user_homes)

    # ── Stampa terminale ─────────────────────────────────────────────
    echo ""
    echo -e "  ${MAGENTA}${BOLD}UserAssist${RESET} (${BOLD}$TOTAL_UA${RESET} voci)"
    for E in "${UA_ENTRIES[@]}"; do
        IFS='§' read -r U N C T <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  cnt=${CYAN}%-3s${RESET}  ${DIM}%s${RESET}  %s\n" \
            "$U" "$C" "$T" "${N: -70}"
    done | head -30
    [[ $TOTAL_UA -gt 30 ]] && echo -e "  ${DIM}... e altri $((TOTAL_UA-30)) (vedi report HTML)${RESET}"

    echo ""
    echo -e "  ${MAGENTA}${BOLD}RunMRU${RESET} (${BOLD}$TOTAL_RUN${RESET} voci)"
    for E in "${RUN_ENTRIES[@]}"; do
        IFS='§' read -r U CMD <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  %s\n" "$U" "$CMD"
    done

    echo ""
    echo -e "  ${MAGENTA}${BOLD}TypedPaths${RESET} (${BOLD}$TOTAL_TP${RESET} voci)"
    for E in "${TP_ENTRIES[@]}"; do
        IFS='§' read -r U K V <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  %s\n" "$U" "$V"
    done

    echo ""
    echo -e "  ${MAGENTA}${BOLD}WordWheelQuery${RESET} (${BOLD}$TOTAL_WW${RESET} ricerche Start)"
    for E in "${WW_ENTRIES[@]}"; do
        IFS='§' read -r U TERM <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  ${CYAN}%s${RESET}\n" "$U" "$TERM"
    done

    separator
    info "UserAssist: ${BOLD}$TOTAL_UA${RESET}  RunMRU: ${BOLD}$TOTAL_RUN${RESET}  TypedPaths: ${BOLD}$TOTAL_TP${RESET}  WordWheel: ${BOLD}$TOTAL_WW"
    local GRAND=$((TOTAL_UA + TOTAL_RUN + TOTAL_TP + TOTAL_WW))
    [[ $GRAND -eq 0 ]] && { warn "Nessuna voce trovata."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "userassist")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina UserAssist per timestamp decrescente
    mapfile -t UA_ENTRIES < <(printf '%s\n' "${UA_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+4*[""])[3],reverse=True); print("\n".join(L))')

    # UserAssist rows
    local UA_ROWS=""
    for E in "${UA_ENTRIES[@]}"; do
        IFS='§' read -r U N C T <<< "$E"
        local SUSP_CSS=""
        echo "$N" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|\\\\temp\\\\" && SUSP_CSS="class='bad'"
        UA_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono' style='word-break:break-all;font-size:.72rem'><span ${SUSP_CSS}>$(html_esc "$N")</span></td>
          <td class='mono ok' style='white-space:nowrap;text-align:center'>${C}</td>
          <td class='mono mid' style='white-space:nowrap'>${T:--}</td></tr>"
    done

    # RunMRU rows
    local RUN_ROWS=""
    for E in "${RUN_ENTRIES[@]}"; do
        IFS='§' read -r U CMD <<< "$E"
        local SUSP_CSS=""
        echo "$CMD" | grep -qi "powershell\|cmd\|wscript\|mshta\|certutil\|bitsadmin\|-enc\|base64" && SUSP_CSS="class='bad'"
        RUN_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono' style='word-break:break-all'><span ${SUSP_CSS}>$(html_esc "$CMD")</span></td></tr>"
    done

    # TypedPaths rows
    local TP_ROWS=""
    for E in "${TP_ENTRIES[@]}"; do
        IFS='§' read -r U K V <<< "$E"
        TP_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono mid' style='white-space:nowrap'>${K}</td>
          <td class='mono' style='word-break:break-all'>$(html_esc "$V")</td></tr>"
    done

    # WordWheelQuery rows
    local WW_ROWS=""
    for E in "${WW_ENTRIES[@]}"; do
        IFS='§' read -r U TERM <<< "$E"
        WW_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono' style='color:var(--accent)'>$(html_esc "$TERM")</td></tr>"
    done

    {
        html_header "UserAssist"
        html_page_header "UA" "UserAssist / RunMRU / <span>TypedPaths</span>" \
            "NTUSER.DAT — Attività interattiva utente" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>UserAssist</div><div class='value'>${TOTAL_UA}</div></div>
          <div class='stat info'><div class='label'>RunMRU</div><div class='value'>${TOTAL_RUN}</div></div>
          <div class='stat ok'><div class='label'>TypedPaths</div><div class='value'>${TOTAL_TP}</div></div>
          <div class='stat'><div class='label'>WordWheel</div><div class='value' style='color:var(--accent4)'>${TOTAL_WW}</div></div>
        </div><main>
        <div class='stitle'>UserAssist — Programmi avviati dalla GUI (ROT13 decodificato)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th>Eseguibile</th><th style='width:6%'>Avvii</th><th style='width:16%'>Ultimo uso</th></tr></thead>
          <tbody>${UA_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>RunMRU — Comandi digitati in Esegui (Win+R)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th>Comando</th></tr></thead>
          <tbody>${RUN_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>TypedPaths — Percorsi digitati in Explorer</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th style='width:8%'>Slot</th><th>Percorso</th></tr></thead>
          <tbody>${TP_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>WordWheelQuery — Ricerche in Start Menu</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th>Termine cercato</th></tr></thead>
          <tbody>${WW_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 19 — ShellBags (navigazione cartelle, anche cancellate)
# ================================================================
module_shellbags() {
    section_header "ShellBags — Navigazione Cartelle" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: ${PY3} -m pip install regipy"
        return 1
    fi


    declare -a ALL_BAGS=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        # UsrClass.dat contiene le ShellBags principali (Win7+)
        local USRCLASS; USRCLASS=$(ci_find_file \
            "$(ci_find_dir "$USER_DIR" "AppData/Local/Microsoft/Windows")" \
            "UsrClass.dat")
        # Fallback: anche in NTUSER.DAT ci sono bag per desktop/drive locali
        local NTUSER; NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")

        local FOUND_ANY=false

        for HIVE_PATH in "$USRCLASS" "$NTUSER"; do
            [[ -z "$HIVE_PATH" || ! -f "$HIVE_PATH" ]] && continue
            local HIVE_LABEL; HIVE_LABEL=$(basename "$HIVE_PATH")

            local BAG_DATA
            BAG_DATA=$("$PY3" - "$HIVE_PATH" "$USERNAME" "$HIVE_LABEL" << 'PYEOF' 2>/dev/null || true
import sys, struct, datetime
try:
    from regipy.registry import RegistryHive
    hive  = RegistryHive(sys.argv[1])
    user  = sys.argv[2]
    label = sys.argv[3]

    # Chiavi ShellBags in UsrClass.dat e NTUSER.DAT
    BAGMRU_KEYS = [
        'Local Settings\\Software\\Microsoft\\Windows\\Shell\\BagMRU',
        'Software\\Microsoft\\Windows\\Shell\\BagMRU',
        'Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\BagMRU',
    ]

    def parse_shell_item(data):
        """
        Estrae il nome leggibile da uno ShellItem binario (SHITEMID).
        Struttura Microsoft: offset 14=short name ASCII, ext block con
        wVersion>=7: long name UTF-16 a ext+4+24. Network items: ASCII a offset 4.
        """
        import struct as _struct
        if not data or len(data) < 4:
            return ''
        item_type = data[2]

        if item_type & 0x70 == 0x10:
            return ''

        if item_type in (0x2F, 0x23, 0x25, 0x29):
            try:
                ch = chr(data[3])
                if ch.isalpha():
                    return ch + ':'
            except Exception:
                pass
            return ''

        if item_type in (0x41, 0x42, 0x46, 0x47, 0x4C, 0x61):
            try:
                end = data.index(0, 4)
                s = data[4:end].decode('ascii', errors='replace').strip()
                if len(s) >= 1:
                    return s
            except Exception:
                pass
            return ''

        if item_type in (0x30, 0x31, 0x32, 0x35, 0x36, 0xB1, 0x74):
            short_name = ''
            ext_start = 14
            try:
                sn_end = data.index(0, 14)
                short_name = data[14:sn_end].decode('ascii', errors='replace').strip()
                ext_start = sn_end + 1
                if ext_start % 2 != 0:
                    ext_start += 1
            except (ValueError, Exception):
                return ''

            if ext_start + 4 > len(data):
                return short_name

            cb_size   = _struct.unpack_from('<H', data, ext_start)[0]
            w_version = _struct.unpack_from('<H', data, ext_start + 2)[0]

            if w_version >= 7 and cb_size >= 32:
                ln_off = ext_start + 4 + 24
                if ln_off < len(data) - 2:
                    try:
                        end = ln_off
                        while end + 1 < len(data) and not (data[end] == 0 and data[end+1] == 0):
                            end += 2
                        s = data[ln_off:end].decode('utf-16-le', errors='strict').rstrip('\x00')
                        if (1 <= len(s) <= 260 and
                                not any('\u2E80' <= c <= '\u9FFF' for c in s) and
                                any(c.isalnum() or c in r'._-~$! ()[]' for c in s)):
                            return s
                    except Exception:
                        pass

            if short_name and not short_name.startswith('1SPS'):
                return short_name

        off = 4
        while off < min(len(data) - 1, 256):
            if 0x20 <= data[off] <= 0x7e:
                end = off
                while end < len(data) and 0x20 <= data[end] <= 0x7e:
                    end += 1
                s = data[off:end].decode('ascii', errors='replace')
                if len(s) >= 2 and not s.startswith(('1SPS', 'SPS', '{', '\x00')):
                    return s
                off = end
            else:
                off += 1
        return ''


    def walk_bagmru(key_path, path_prefix=''):
        try:
            key = hive.get_key(key_path)
        except:
            return
        for v in key.get_values():
            if not isinstance(v.value, bytes) or v.name == 'MRUListEx':
                continue
            name = parse_shell_item(v.value)
            full_path = (path_prefix + '\\' + name).lstrip('\\') if name else path_prefix
            if full_path:
                print(f"{user}\t{full_path}\t{label}")
            try:
                subkey_path = f'{key_path}\\{v.name}'
                walk_bagmru(subkey_path, full_path)
            except:
                pass

    for key_name in BAGMRU_KEYS:
        walk_bagmru(key_name)

except Exception as e:
    pass
PYEOF
)
            if [[ -n "$BAG_DATA" ]]; then
                FOUND_ANY=true
                while IFS=$'\t' read -r U P L; do
                    [[ -z "$P" ]] && continue
                    ALL_BAGS+=("${U}§${P}§${L}")
                    TOTAL=$((TOTAL + 1))
                    local IS_SUSP=false
                    echo "$P" | grep -qi "\\\\temp\\\\\|appdata\\\\local\\\\\|programdata\\\\\|\\\\public\\\\\|network\|\\\\c\$\|admin\$\|ipc\$" \
                        && IS_SUSP=true
                    $IS_SUSP && SUSP_COUNT=$((SUSP_COUNT + 1))
                done <<< "$BAG_DATA"
            fi
        done

        if $FOUND_ANY; then
            ok "$USERNAME — $TOTAL shellbags totali"
        else
            dim_msg "$USERNAME — nessuna ShellBag trovata"
        fi

    done < <(get_user_homes)

    # Stampa terminale (prime 40)
    echo ""
    local CNT=0
    for E in "${ALL_BAGS[@]}"; do
        IFS='§' read -r U P L <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  %s\n" "$U" "$P"
        CNT=$((CNT+1))
        [[ $CNT -ge 40 ]] && break
    done
    [[ $TOTAL -gt 40 ]] && echo -e "  ${DIM}... e altri $((TOTAL-40)) (vedi report HTML)${RESET}"

    separator
    info "ShellBags totali: ${BOLD}$TOTAL${RESET}  |  Sospette: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessuna ShellBag trovata."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "shellbags")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_BAGS[@]}"; do
        IFS='§' read -r U P L <<< "$E"
        local SUSP_CSS="" ROW_STYLE=""
        echo "$P" | grep -qi "\\\\temp\|programdata\|\\\\public\|network\|\\\\c\$\|admin\$" && {
            SUSP_CSS="bad"; ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        }
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono ${SUSP_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
          <td class='mono dim' style='white-space:nowrap;font-size:.68rem'>${L}</td>
        </tr>"
    done

    {
        html_header "ShellBags"
        html_page_header "SB" "ShellBags — <span>Navigazione Cartelle</span>" \
            "UsrClass.dat → BagMRU" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>ShellBags</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Cartelle visitate — anche se poi cancellate o smontate</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th>Percorso ricostruito</th><th style='width:12%'>Hive</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 20 — SAM: Hash password account locali
# ================================================================
module_sam() {
    section_header "SAM — Hash Account Locali" "$RED"
    check_win_root || return 1

    local SAM_HIVE;  SAM_HIVE=$(get_hive "SAM")
    local SYS_HIVE;  SYS_HIVE=$(get_hive "SYSTEM")

    if [[ -z "$SAM_HIVE" || ! -f "$SAM_HIVE" ]]; then
        err "Hive SAM non trovato"
        return 1
    fi
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "Hive SYSTEM non trovato (necessario per la SysKey)"
        return 1
    fi

    # Verifica impacket
    if ! "$PY3" -c "from impacket.examples.secretsdump import LocalOperations" 2>/dev/null; then
        warn "impacket non trovato."
        warn "Installalo con: ${PY3} -m pip install impacket"
        info "Estrazione parziale tramite regipy (senza hash decifrati)..."

        if ! check_regipy; then
            return 1
        fi
        # Fallback: solo lista account senza hash
        local ACC_DATA
        ACC_DATA=$("$PY3" - "$SAM_HIVE" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    try:
        users_key = hive.get_key('SAM\\Domains\\Account\\Users\\Names')
        for sk in users_key.iter_subkeys():
            print(f"{sk.name}\t(hash non disponibile senza impacket)")
    except Exception as e:
        print(f"ERROR\t{e}")
except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
        echo ""
        while IFS=$'\t' read -r NAME NOTE; do
            [[ "$NAME" == "ERROR" ]] && { err "$NOTE"; continue; }
            printf "  ${CYAN}%-25s${RESET}  ${DIM}%s${RESET}\n" "$NAME" "$NOTE"
        done <<< "$ACC_DATA"
        warn "Per gli hash installa impacket e rilancia il modulo."
        return 0
    fi

    info "Estrazione hash con impacket secretsdump..."
    echo ""

    local TMP_DIR; TMP_DIR=$(mktemp -d /tmp/dfir_sam_XXXXXX)
    trap '[[ -n "${TMP_DIR:-}" ]] && rm -rf "$TMP_DIR"' RETURN
    cp "$SAM_HIVE" "$TMP_DIR/SAM" 2>/dev/null || true
    cp "$SYS_HIVE" "$TMP_DIR/SYSTEM" 2>/dev/null || true
    chmod 600 "$TMP_DIR/SAM" "$TMP_DIR/SYSTEM" 2>/dev/null || true

    local DUMP_OUT
    DUMP_OUT=$("$PY3" - "$TMP_DIR/SAM" "$TMP_DIR/SYSTEM" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from impacket.examples.secretsdump import LocalOperations, SAMHashes
    from impacket.smbconnection import SessionError

    sys_ops = LocalOperations(sys.argv[2])
    boot_key = sys_ops.getBootKey()

    sam_hashes = SAMHashes(sys.argv[1], boot_key, isRemote=False)
    sam_hashes.dump()
    sam_hashes.export()
except Exception as e:
    # Fallback: usa secretsdump direttamente
    import subprocess, os
    try:
        result = subprocess.run(
            [sys.executable, '-c',
             f'''
from impacket.examples.secretsdump import LocalOperations, SAMHashes
import sys
sys_ops = LocalOperations("{sys.argv[2]}")
boot_key = sys_ops.getBootKey()
sam = SAMHashes("{sys.argv[1]}", boot_key, isRemote=False)
sam.dump()
'''],
            capture_output=True, text=True, timeout=30
        )
        if result.stdout:
            print(result.stdout)
        else:
            print(f"ERROR\t{str(e)}")
    except:
        print(f"ERROR\t{str(e)}")
PYEOF
)

    declare -a ACC_ENTRIES=()
    local TOTAL=0

    local KNOWN_HASHES=(
        "31d6cfe0d16ae931b73c59d7e0c089c0"   # Password vuota
        "aad3b435b51404eeaad3b435b51404ee"   # LM vuoto
    )

    while IFS= read -r LINE; do
        [[ -z "$LINE" ]] && continue
        [[ "$LINE" == "ERROR"* ]] && { err "${LINE#ERROR    }"; continue; }
        # Formato impacket: username:RID:LMhash:NThash:::
        if echo "$LINE" | grep -qE '^[^:]+:[0-9]+:[a-fA-F0-9]{32}:[a-fA-F0-9]{32}'; then
            IFS=':' read -r UNAME RID LM NT REST <<< "$LINE"
            local IS_EMPTY=false IS_SUSP=false
            [[ "${NT,,}" == "31d6cfe0d16ae931b73c59d7e0c089c0" ]] && IS_EMPTY=true
            # RID 500=Admin, 501=Guest, altri potrebbero essere account nascosti
            [[ "$RID" -gt 1000 ]] 2>/dev/null && IS_SUSP=true
            printf "  ${CYAN}%-20s${RESET}  RID:${DIM}%-6s${RESET}  NT:${YELLOW}%s${RESET}" \
                "$UNAME" "$RID" "$NT"
            $IS_EMPTY && printf "  ${DIM}(password vuota)${RESET}"
            $IS_SUSP  && printf "  ${RED}[utente non default]${RESET}"
            echo ""
            ACC_ENTRIES+=("${UNAME}§${RID}§${LM}§${NT}§${IS_EMPTY}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
        else
            echo -e "  ${DIM}$LINE${RESET}"
        fi
    done <<< "$DUMP_OUT"

    separator
    info "Account trovati: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun account estratto."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "sam")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ACC_ENTRIES[@]}"; do
        IFS='§' read -r UNAME RID LM NT IS_EMPTY IS_SUSP <<< "$E"
        local ROW_STYLE="" NT_CSS="mono"
        [[ "$IS_SUSP" == "true" ]] && ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        [[ "$IS_EMPTY" == "true" ]] && NT_CSS="mono warn"
        local NOTES=""
        [[ "$IS_EMPTY" == "true" ]] && NOTES+="<span class='dim' style='font-size:.68rem'>password vuota</span> "
        [[ "$IS_SUSP"  == "true" ]] && NOTES+="<span style='color:var(--accent2);font-size:.68rem'>account non default</span>"
        [[ "$RID" == "500" ]]       && NOTES+="<span style='color:var(--accent4);font-size:.68rem'>Administrator</span>"
        [[ "$RID" == "501" ]]       && NOTES+="<span class='dim' style='font-size:.68rem'>Guest</span>"
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono'>${UNAME}</td>
          <td class='mono mid' style='white-space:nowrap'>${RID}</td>
          <td class='mono dim' style='font-size:.65rem;word-break:break-all'>${LM}</td>
          <td class='${NT_CSS}' style='font-size:.72rem;word-break:break-all'>${NT}</td>
          <td style='font-size:.72rem'>${NOTES}</td>
        </tr>"
    done

    {
        html_header "SAM"
        html_page_header "SM" "SAM — <span>Hash Account Locali</span>" \
            "Windows\\System32\\config\\SAM" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Account</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>Hash NTLM account locali — craccabili offline con hashcat/john</div>
        <div class='card'><table>
          <thead><tr><th>Username</th><th style='width:7%'>RID</th><th style='width:27%'>LM Hash</th><th style='width:27%'>NT Hash</th><th>Note</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Cracking offline</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            hashcat -m 1000 hashes.txt wordlist.txt<br>
            john --format=NT hashes.txt
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 21 — MFT Timeline (Master File Table)
# ================================================================
module_mft() {
    section_header "MFT — Master File Table Timeline" "$YELLOW"
    check_win_root || return 1

    local MFT_FILE=""
    MFT_FILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "\$mft" -type f 2>/dev/null | head -1)
    [[ -z "$MFT_FILE" ]] && \
        MFT_FILE=$(find "$WIN_ROOT" -maxdepth 1 \( -iname "mft" -o -iname "\$mft" \) 2>/dev/null | head -1)

    # Se non trovato, prova a estrarlo con ntfscat (non richiede rimontaggio)
    if [[ -z "$MFT_FILE" ]]; then
        # Recupera il device block associato a WIN_ROOT
        local DEV=""
        while IFS=' ' read -r RD RM _ _; do
            local DM; DM=$(printf '%b' "$RM")
            [[ "$DM" == "$WIN_ROOT" ]] && DEV=$(printf '%b' "$RD") && break
        done < /proc/mounts

        local TMP_MFT=""
        if [[ -n "$DEV" ]] && command -v ntfscat &>/dev/null; then
            info "Estrazione \$MFT tramite ntfscat dal device ${DEV}..."
            TMP_MFT=$(mktemp /tmp/dfir_MFT_XXXXXX)
            if ntfscat -f "$DEV" '$MFT' > "$TMP_MFT" 2>/dev/null && [[ -s "$TMP_MFT" ]]; then
                MFT_FILE="$TMP_MFT"
                ok "\$MFT estratto con ntfscat ($(du -h "$MFT_FILE" | cut -f1))"
                trap '[[ -n "${TMP_MFT:-}" ]] && rm -f "$TMP_MFT"' RETURN
            else
                rm -f "$TMP_MFT"
                TMP_MFT=""
            fi
        fi

        if [[ -z "$MFT_FILE" ]]; then
            warn "\$MFT non accessibile — il volume è montato con ntfs3 (driver kernel)"
            echo ""
            info "Il driver ntfs3 non espone i file di sistema (\$MFT, \$LogFile, ecc.)."
            info "Per accedere all'\$MFT hai due opzioni:"
            echo ""
            if [[ -n "$DEV" ]]; then
                echo -e "  ${CYAN}Opzione A${RESET} — rimonta con ntfs-3g (supporta file di sistema):"
                echo -e "  ${DIM}sudo umount \"$WIN_ROOT\"${RESET}"
                echo -e "  ${DIM}sudo mount -t ntfs-3g -o ro,noload,show_sys_files \"$DEV\" \"$WIN_ROOT\"${RESET}"
                echo ""
                echo -e "  ${CYAN}Opzione B${RESET} — estrai \$MFT senza rimontare (richiede ntfs-3g-tools):"
                echo -e "  ${DIM}sudo apt install ntfs-3g${RESET}"
                echo -e "  ${DIM}ntfscat -f \"$DEV\" '\$MFT' > /tmp/MFT && sudo chmod a+r /tmp/MFT${RESET}"
                echo -e "  ${DIM}# poi imposta WIN_ROOT e riavvia il modulo${RESET}"
            else
                echo -e "  ${CYAN}Opzione A${RESET} — rimonta con ntfs-3g:"
                echo -e "  ${DIM}sudo umount \"$WIN_ROOT\"${RESET}"
                echo -e "  ${DIM}sudo mount -t ntfs-3g -o ro,noload,show_sys_files /dev/DEVICE \"$WIN_ROOT\"${RESET}"
                echo ""
                echo -e "  ${CYAN}Opzione B${RESET} — estrai direttamente:"
                echo -e "  ${DIM}ntfscat -f /dev/DEVICE '\$MFT' > /tmp/MFT${RESET}"
            fi
            echo ""
            return 0
        fi
    fi

    info "MFT trovato: $MFT_FILE  ($(du -h "$MFT_FILE" 2>/dev/null | cut -f1 || echo '?'))"

    # Verifica python-mft
    local HAS_PYMFT=false
    "$PY3" -c "import mft" 2>/dev/null && HAS_PYMFT=true

    if ! $HAS_PYMFT; then
        warn "python-mft non trovato. Installalo con: ${PY3} -m pip install mft"
        info "Analisi rapida tramite strings (parziale)..."
        echo ""
        # Estrai nomi file e timestamp con strings come fallback
        local STRINGS_OUT
        STRINGS_OUT=$(strings "$MFT_FILE" 2>/dev/null | \
            grep -iE '\.(exe|dll|bat|ps1|vbs|py|zip|rar|7z|iso)$' | \
            sort -u | head -200 || true)
        echo "$STRINGS_OUT" | while IFS= read -r L; do
            printf "  ${DIM}%s${RESET}\n" "$L"
        done | head -50
        warn "Per l'analisi completa installa python-mft."
        return 0
    fi

    # Calcola il numero approssimativo di record (ogni record MFT è 1KB)
    local MFT_SIZE; MFT_SIZE=$(stat -c "%s" "$MFT_FILE" 2>/dev/null || echo 0)
    local MFT_RECORDS=$((MFT_SIZE / 1024))
    
    info "Parsing MFT con python-mft (~${MFT_RECORDS} record presenti)..."
    if [[ $MFT_SIZE -gt 500000000 ]]; then
        warn "⏳ L'MFT supera i 500MB! Il parsing completo può richiedere vari minuti."
        info "   (Specialmente su volumi montati con BitLocker/FUSE). Attendi senza interrompere..."
    fi

    local TMP_MFT_CSV; TMP_MFT_CSV=$(mktemp /tmp/dfir_mft_XXXXXX.csv)
    local TMP_MFT_ROWS; TMP_MFT_ROWS=$(mktemp /tmp/dfir_mft_rows_XXXXXX.html)
    trap '[[ -n "${TMP_MFT_CSV:-}" ]] && rm -f "$TMP_MFT_CSV"; [[ -n "${TMP_MFT_ROWS:-}" ]] && rm -f "$TMP_MFT_ROWS"' RETURN

    "$PY3" - "$MFT_FILE" "$TMP_MFT_CSV" << 'PYEOF' 2>/dev/null
import sys, json, datetime
try:
    from mft import PyMftParser
except ImportError:
    sys.exit(1)

mft_file = sys.argv[1]
outf     = sys.argv[2]

SKIP_EXTS = {'.lnk','.log','.dat','.mui','.dll','.sys','.manifest',''}
SUSP_EXTS = {'.exe','.ps1','.bat','.vbs','.hta','.js','.cmd',
             '.zip','.rar','.7z','.iso','.img',
             '.py','.rb','.php','.aspx','.jsp'}

try:
    parser = PyMftParser(mft_file)
    with open(outf,'w', encoding='utf-8') as out:
        count = 0
        for chunk_bytes in parser.entries_csv():
            for line in chunk_bytes.decode('utf-8', 'replace').split('\n'):
                if not line or line.startswith('Signature'):
                    continue
                parts = line.split(',', 21)
                if len(parts) < 22:
                    continue
                
                if parts[10] == 'true': # IsADirectory
                    continue
                
                fpath = parts[21].strip('"\r')
                if '\\' in fpath:
                    name = fpath.rsplit('\\', 1)[-1]
                else:
                    name = fpath
                
                ext = ''
                if '.' in name:
                    ext = '.' + name.rsplit('.', 1)[-1].lower()
                if ext in SKIP_EXTS:
                    continue

                si_cre = parts[16][:19]
                si_mod = parts[14][:19]
                fn_cre = parts[20][:19]

                stomped = ''
                if si_cre and fn_cre and si_cre != fn_cre:
                    try:
                        d_si = datetime.datetime.fromisoformat(si_cre)
                        d_fn = datetime.datetime.fromisoformat(fn_cre)
                        if abs((d_si - d_fn).total_seconds()) > 3600:
                            stomped = 'TIMESTOMP'
                    except:
                        pass
                
                susp = '1' if ext in SUSP_EXTS or stomped else '0'
                si_cre_s = si_cre.replace('T', ' ')
                si_mod_s = si_mod.replace('T', ' ')
                fn_cre_s = fn_cre.replace('T', ' ')
                
                out.write(f"{name}\t{ext}\t{si_cre_s}\t{si_mod_s}\t\t{fn_cre_s}\t{stomped}\t{susp}\n")
                count += 1
    print(f"Records processati: {count}", file=sys.stderr)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF

    local TOTAL_LINES
    TOTAL_LINES=$(wc -l < "$TMP_MFT_CSV" 2>/dev/null || echo 0)
    local STOMP_COUNT
    STOMP_COUNT=$(grep -c "TIMESTOMP" "$TMP_MFT_CSV" 2>/dev/null || echo 0)
    local SUSP_COUNT
    SUSP_COUNT=$(awk -F'\t' '$8=="1"' "$TMP_MFT_CSV" 2>/dev/null | wc -l || echo 0)

    info "File processati: ${BOLD}$TOTAL_LINES${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT${RESET}  |  Timestomping: ${RED}${BOLD}$STOMP_COUNT"

    # Stampa terminale: solo sospetti e timestomped
    echo ""
    awk -F'\t' '$8=="1" || $7=="TIMESTOMP" {print}' "$TMP_MFT_CSV" 2>/dev/null | head -30 | \
    while IFS=$'\t' read -r NAME EXT CREATED MODIFIED MFTMOD FNCREATED STOMPED SUSP; do
        local MARK=""
        [[ "$STOMPED" == "TIMESTOMP" ]] && MARK="${RED}[STOMP]${RESET} "
        printf "  %s${DIM}%-40s${RESET}  ${CYAN}%s${RESET}  %s\n" \
            "$MARK" "${NAME: -40}" "$CREATED" "$EXT"
    done

    [[ $TOTAL_LINES -eq 0 ]] && { warn "Nessun record MFT processato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "mft")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Genera righe HTML dal CSV (solo sospetti + primi 2000 per non gonfiare)
    "$PY3" - "$TMP_MFT_CSV" "$TMP_MFT_ROWS" << 'PYEOF' 2>/dev/null || true
import sys, html as H

rows_data = []
with open(sys.argv[1]) as f:
    for line in f:
        parts = line.rstrip('\n').split('\t')
        if len(parts) < 8:
            continue
        susp = parts[7]
        stomped = parts[6]
        if susp != '1' and stomped != 'TIMESTOMP':
            continue
        rows_data.append(parts)

# Ordina per data di modifica (SI Modified - colonna indice 3) decrescente
rows_data.sort(key=lambda x: x[3], reverse=True)

rows = []
for parts in rows_data[:2000]:
    name,ext,created,modified,mftmod,fncreated,stomped,susp = parts[:8]
    stomp_badge = ("<span style='color:var(--accent2);font-weight:700;font-size:.68rem;"
                   "margin-right:.3rem'>⚠ TIMESTOMP</span>") if stomped == 'TIMESTOMP' else ''
    row_style = ("style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
                 if stomped == 'TIMESTOMP' else "")
    rows.append(
        f"<tr {row_style}>"
        f"<td class='mono' style='word-break:break-all;font-size:.72rem'>{stomp_badge}{H.escape(name)}</td>"
        f"<td class='mono mid' style='white-space:nowrap'>{H.escape(ext)}</td>"
        f"<td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{H.escape(created)}</td>"
        f"<td class='mono' style='white-space:nowrap;font-size:.7rem;color:var(--accent4)'>{H.escape(modified)}</td>"
        f"<td class='mono dim' style='white-space:nowrap;font-size:.7rem'>{H.escape(fncreated)}</td>"
        f"</tr>"
    )

with open(sys.argv[2],'w') as f:
    f.write('\n'.join(rows))
PYEOF

    local TABLE_ROWS; TABLE_ROWS=$(cat "$TMP_MFT_ROWS" 2>/dev/null || echo "")

    {
        html_header "MFT"
        html_page_header "MF" "MFT — <span>Master File Table</span> Timeline" \
            "\$MFT — radice volume NTFS" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File analizzati</div><div class='value'>${TOTAL_LINES}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
          <div class='stat'><div class='label'>Timestomping</div><div class='value' style='color:var(--accent2)'>${STOMP_COUNT}</div></div>
        </div>
        <div class='notice' style='margin:1rem;padding:1rem;background:rgba(255,255,255,0.05);border-radius:6px;font-size:0.85rem'>
          <b>Nota:</b> Per prevenire blocchi del browser, vengono mostrati a schermo solo i primi <b>2000</b> record sospetti/timestomped ordinati per data decrescente. <br>
          Il report completo (TUTTI i record analizzati) è salvato in formato CSV nella stessa cartella: <br><code style='color:var(--accent4)'>$(basename "${REPORT_HTML%.html}").csv</code>
        </div>
        <main>
        <div class='stitle'>File sospetti e con possibile timestomping (STANDARD_INFO ≠ FILENAME &gt;1h)</div>
        <div class='card'><table style='table-layout: fixed; width: 100%;'>
          <thead><tr>
            <th style='width:57%'>Nome file</th>
            <th style='width:7%'>Ext</th>
            <th style='width:12%'>SI Created</th>
            <th style='width:12%'>SI Modified</th>
            <th style='width:12%'>FN Created</th>
          </tr></thead>
          <tbody>${TABLE_ROWS:-<tr><td colspan='5' class='dim' style='padding:1rem'>Nessun file sospetto trovato</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    # Salva il CSV completo accanto al report
    cp "$TMP_MFT_CSV" "${REPORT_HTML%.html}.csv" 2>/dev/null || echo ""

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    ok "Dati completi salvati: ${BOLD}${REPORT_HTML%.html}.csv"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 22 — OpenSaveMRU / LastVisitedMRU (dialoghi file)
# ================================================================
module_opensave() {
    section_header "OpenSaveMRU / LastVisitedMRU — Dialoghi File" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato per ${PY3} (${PY3_VERSION})"
        warn "Installalo con: ${PY3} -m pip install regipy"
        return 1
    fi


    declare -a OS_ENTRIES=()   # OpenSaveMRU
    declare -a LV_ENTRIES=()   # LastVisitedMRU
    local TOTAL_OS=0 TOTAL_LV=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue

        # ── OpenSavePidlMRU / OpenSaveMRU ───────────────────────────
        local OS_DATA
        OS_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    # Prova entrambe le chiavi (Win7 usa PidlMRU, versioni precedenti MRU)
    for key_path in [
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\OpenSavePidlMRU',
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\OpenSaveMRU',
    ]:
        try:
            root = hive.get_key(key_path)
            for ext_key in root.iter_subkeys():
                ext = ext_key.name  # es. "exe", "ps1", "*"
                for v in ext_key.get_values():
                    if v.name.lower() in ('mrulistex','mrulist'):
                        continue
                    val = v.value
                    # PidlMRU: valore binario con path embedded
                    if isinstance(val, bytes):
                        # Cerca stringa UTF-16 nel blob
                        try:
                            decoded = val.decode('utf-16-le', errors='replace').rstrip('\x00')
                            # Prendi solo la parte che sembra un path
                            parts = [p for p in decoded.split('\x00') if '\\' in p and len(p) > 3]
                            for p in parts[:1]:
                                print(f"{user}\t{ext}\t{p.strip()}")
                        except:
                            pass
                    elif isinstance(val, str) and '\\' in val:
                        print(f"{user}\t{ext}\t{val}")
        except:
            continue
except:
    pass
PYEOF
)
        while IFS=$'\t' read -r U EXT PATH_V; do
            [[ -z "$PATH_V" ]] && continue
            OS_ENTRIES+=("${U}§${EXT}§${PATH_V}")
            TOTAL_OS=$((TOTAL_OS + 1))
        done <<< "$OS_DATA"

        # ── LastVisitedPidlMRU / LastVisitedMRU ─────────────────────
        local LV_DATA
        LV_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    for key_path in [
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\LastVisitedPidlMRU',
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\LastVisitedMRU',
    ]:
        try:
            key = hive.get_key(key_path)
            for v in key.get_values():
                if v.name.lower() in ('mrulistex','mrulist'):
                    continue
                val = v.value
                if isinstance(val, bytes):
                    try:
                        # Prima stringa terminata da null = nome app
                        decoded = val.decode('utf-16-le', errors='replace')
                        parts = decoded.split('\x00')
                        app = parts[0].strip() if parts else ''
                        # Cerca path nelle parti successive
                        path = next((p for p in parts[1:] if '\\' in p and len(p) > 3), '')
                        if app:
                            print(f"{user}\t{app}\t{path}")
                    except:
                        pass
                elif isinstance(val, str):
                    print(f"{user}\t{val}\t")
        except:
            continue
except:
    pass
PYEOF
)
        while IFS=$'\t' read -r U APP PATH_V; do
            [[ -z "$APP" ]] && continue
            LV_ENTRIES+=("${U}§${APP}§${PATH_V}")
            TOTAL_LV=$((TOTAL_LV + 1))
        done <<< "$LV_DATA"

    done < <(get_user_homes)

    # Stampa terminale
    echo ""
    echo -e "  ${GREEN}${BOLD}OpenSaveMRU${RESET} (${BOLD}$TOTAL_OS${RESET} file aperti/salvati)"
    for E in "${OS_ENTRIES[@]}"; do
        IFS='§' read -r U EXT P <<< "$E"
        local COLOR="$DIM"
        echo "$P" | grep -qi "\.ps1\|\.bat\|\.exe\|\.vbs\|\.hta\|temp\|public" && COLOR="$RED"
        printf "  [${DIM}%-12s${RESET}]  ${CYAN}%-8s${RESET}  ${COLOR}%s${RESET}\n" "$U" ".$EXT" "$P"
    done | head -30

    echo ""
    echo -e "  ${GREEN}${BOLD}LastVisitedMRU${RESET} (${BOLD}$TOTAL_LV${RESET} applicazioni/cartelle)"
    for E in "${LV_ENTRIES[@]}"; do
        IFS='§' read -r U APP P <<< "$E"
        printf "  [${DIM}%-12s${RESET}]  ${YELLOW}%-25s${RESET}  ${DIM}%s${RESET}\n" "$U" "$APP" "$P"
    done | head -20

    separator
    info "OpenSave: ${BOLD}$TOTAL_OS${RESET}  |  LastVisited: ${BOLD}$TOTAL_LV"
    local GRAND=$((TOTAL_OS + TOTAL_LV))
    [[ $GRAND -eq 0 ]] && { warn "Nessun dato trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "opensave")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local OS_ROWS=""
    for E in "${OS_ENTRIES[@]}"; do
        IFS='§' read -r U EXT P <<< "$E"
        local SUSP_CSS="" ROW_STYLE=""
        echo "$P" | grep -qi "\.ps1\|\.bat\|\.exe\|\.vbs\|\.hta\|temp\|public\|programdata" && {
            SUSP_CSS="bad"
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        }
        OS_ROWS+="<tr ${ROW_STYLE}>
          <td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono' style='white-space:nowrap;color:var(--accent4)'>.${EXT}</td>
          <td class='mono ${SUSP_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
        </tr>"
    done

    local LV_ROWS=""
    for E in "${LV_ENTRIES[@]}"; do
        IFS='§' read -r U APP P <<< "$E"
        LV_ROWS+="<tr>
          <td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$APP")</td>
          <td class='mono mid' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
        </tr>"
    done

    {
        html_header "OpenSaveMRU"
        html_page_header "OS" "OpenSave / <span>LastVisited</span> MRU" \
            "NTUSER.DAT → ComDlg32" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File aperti/salvati</div><div class='value'>${TOTAL_OS}</div></div>
          <div class='stat info'><div class='label'>App + cartelle</div><div class='value'>${TOTAL_LV}</div></div>
        </div><main>
        <div class='stitle'>OpenSaveMRU — File aperti o salvati tramite dialogo di Windows</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th style='width:8%'>Estensione</th><th>Path</th></tr></thead>
          <tbody>${OS_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>LastVisitedMRU — Ultima cartella visitata per applicazione</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th style='width:20%'>Applicazione</th><th>Ultima cartella</th></tr></thead>
          <tbody>${LV_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>Nessun dato</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 23 — USN Journal ($UsnJrnl:$J)
# ================================================================
module_usn() {
    section_header "USN Journal — Change Log NTFS" "$CYAN"
    check_win_root || return 1


    # Cerca $UsnJrnl nella root del volume (ntfs-3g show_sys_files) o via ntfscat
    local UJFILE=""
    UJFILE=$(find "$WIN_ROOT" -maxdepth 3 \( -iname "\$j" -o -iname "usnjrnl" \) 2>/dev/null | head -1)
    [[ -z "$UJFILE" ]] && \
        UJFILE=$(find "$WIN_ROOT" -maxdepth 2 \( -iname "\$usnjrnl" -o -iname "\$j" \) 2>/dev/null | head -1)

    if [[ -z "$UJFILE" ]]; then
        local DEV=""
        while IFS=' ' read -r RD RM _ _; do
            local DM; DM=$(printf '%b' "$RM")
            [[ "$DM" == "$WIN_ROOT" ]] && DEV=$(printf '%b' "$RD") && break
        done < /proc/mounts
        if [[ -n "$DEV" ]] && command -v ntfscat &>/dev/null; then
            info "Estrazione \$UsnJrnl via ntfscat da $DEV..."
            local TMP_USN; TMP_USN=$(mktemp /tmp/dfir_usn_XXXXXX)
            if ntfscat -f "$DEV" '$Extend/$UsnJrnl:$J' > "$TMP_USN" 2>/dev/null && [[ -s "$TMP_USN" ]]; then
                UJFILE="$TMP_USN"
                ok "\$UsnJrnl estratto ($(du -h "$UJFILE" | cut -f1))"
                trap '[[ -n "${TMP_USN:-}" ]] && rm -f "$TMP_USN"' RETURN
            else
                rm -f "$TMP_USN"
                warn "\$UsnJrnl non accessibile. Monta con: mount -t ntfs-3g -o ro,show_sys_files"
                return 0
            fi
        else
            warn "\$UsnJrnl non accessibile — volume montato con ntfs3 (driver kernel)"
            [[ -n "$DEV" ]] && info "Prova: ntfscat -f \"$DEV\" '\$Extend/\$UsnJrnl:\$J' > /tmp/J"
            return 0
        fi
    fi

    info "USN Journal: $UJFILE  ($(du -h "$UJFILE" 2>/dev/null | cut -f1 || echo '?'))"

    local TMP_USN_CSV; TMP_USN_CSV=$(mktemp /tmp/dfir_usn_XXXXXX.csv)
    local TMP_USN_ROWS; TMP_USN_ROWS=$(mktemp /tmp/dfir_usn_rows_XXXXXX.html)
    trap '[[ -n "${TMP_USN_CSV:-}" ]] && rm -f "$TMP_USN_CSV"; [[ -n "${TMP_USN_ROWS:-}" ]] && rm -f "$TMP_USN_ROWS"' RETURN

    "$PY3" - "$UJFILE" "$TMP_USN_CSV" << 'PYEOF' 2>/dev/null || true
import sys, struct, datetime, os

REASONS = {
    0x00000001:'DATA_OVERWRITE', 0x00000002:'DATA_EXTEND', 0x00000004:'DATA_TRUNCATION',
    0x00000010:'NAMED_DATA_OVERWRITE', 0x00000020:'NAMED_DATA_EXTEND', 0x00000040:'NAMED_DATA_TRUNCATION',
    0x00000100:'FILE_CREATE', 0x00000200:'FILE_DELETE', 0x00000400:'EA_CHANGE',
    0x00000800:'SECURITY_CHANGE', 0x00001000:'RENAME_OLD_NAME', 0x00002000:'RENAME_NEW_NAME',
    0x00004000:'INDEXABLE_CHANGE', 0x00008000:'BASIC_INFO_CHANGE', 0x00010000:'HARD_LINK_CHANGE',
    0x00020000:'COMPRESSION_CHANGE', 0x00040000:'ENCRYPTION_CHANGE', 0x00080000:'OBJECT_ID_CHANGE',
    0x00100000:'REPARSE_POINT_CHANGE', 0x00200000:'STREAM_CHANGE', 0x80000000:'CLOSE',
}
SUSP_EXTS = {'.exe','.dll','.ps1','.bat','.vbs','.hta','.cmd','.js','.py','.msi','.iso','.zip','.rar','.7z'}
SUSP_PATHS_KWORDS = ['temp','tmp','appdata','public','programdata','downloads','users\\']
SUSP_REASONS = {0x00000100,0x00000200,0x00001000,0x00002000}  # CREATE DELETE RENAME

path = sys.argv[1]
outf = sys.argv[2]
count = 0

def decode_reason(r):
    parts = [n for mask,n in REASONS.items() if r & mask]
    return '|'.join(parts) if parts else hex(r)

try:
    with open(path, 'rb') as f, open(outf, 'w') as out:
        raw = f.read()
    offset = 0
    while offset < len(raw) - 60:
        # Skip nulls (padding between records)
        if raw[offset:offset+4] == b'\x00\x00\x00\x00':
            offset += 8; continue
        try:
            rec_len = struct.unpack_from('<I', raw, offset)[0]
            if rec_len < 60 or rec_len > 65536:
                offset += 8; continue
            major = struct.unpack_from('<H', raw, offset+4)[0]
            if major not in (2, 3):
                offset += 8; continue
            reason = struct.unpack_from('<I', raw, offset+40)[0]
            ts_raw = struct.unpack_from('<q', raw, offset+24)[0]
            fname_off = struct.unpack_from('<H', raw, offset+58)[0]
            fname_len = struct.unpack_from('<H', raw, offset+56)[0]
            ts = ''
            if ts_raw > 0:
                try:
                    ts = (datetime.datetime(1601,1,1) +
                          datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
                except: pass
            abs_off = offset + fname_off
            name = raw[abs_off:abs_off+fname_len].decode('utf-16-le', errors='replace')
            ext = ''
            if '.' in name:
                ext = '.' + name.rsplit('.',1)[-1].lower()
            reason_str = decode_reason(reason)
            susp = '1' if (ext in SUSP_EXTS or any(reason & m for m in SUSP_REASONS)) else '0'
            with open(outf, 'a') as out:
                out.write(f"{ts}\t{name}\t{ext}\t{reason_str}\t{susp}\n")
            count += 1
            if count >= 200000:
                break
            offset += rec_len
            if offset % 8 != 0:
                offset += 8 - (offset % 8)
        except Exception:
            offset += 8
    print(f"Records: {count}", file=sys.stderr)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF

    local TOTAL_LINES; TOTAL_LINES=$(wc -l < "$TMP_USN_CSV" 2>/dev/null || echo 0)
    local SUSP_COUNT; SUSP_COUNT=$(awk -F'\t' '$5=="1"' "$TMP_USN_CSV" 2>/dev/null | wc -l || echo 0)

    info "Record USN: ${BOLD}$TOTAL_LINES${RESET}  |  Sospetti (ext+azione): ${RED}${BOLD}$SUSP_COUNT"
    echo ""
    awk -F'\t' '$5=="1"' "$TMP_USN_CSV" 2>/dev/null | head -25 | \
    while IFS=$'\t' read -r TS NAME EXT REASON SUSP; do
        printf "  ${RED}%-20s${RESET}  ${CYAN}%-30s${RESET}  %s\n" "$TS" "$NAME" "$REASON"
    done

    [[ $TOTAL_LINES -eq 0 ]] && { warn "Nessun record USN processato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "usn")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    "$PY3" - "$TMP_USN_CSV" "$TMP_USN_ROWS" << 'PYEOF' 2>/dev/null || true
import sys, html as H
rows = []
with open(sys.argv[1]) as f:
    for line in f:
        parts = line.rstrip('\n').split('\t')
        if len(parts) < 5: continue
        ts, name, ext, reason, susp = parts[0], parts[1], parts[2], parts[3], parts[4]
        is_susp = susp == '1'
        rs = "style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if is_susp else ""
        ec = "bad" if is_susp else "ok"
        rows.append(f"<tr {rs}><td class='mono {ec}' style='white-space:nowrap;font-size:.7rem'>{H.escape(ts)}</td>"
                    f"<td class='mono' style='word-break:break-all;font-size:.72rem'>{H.escape(name)}</td>"
                    f"<td class='mono mid' style='white-space:nowrap'>{H.escape(ext)}</td>"
                    f"<td class='mono dim' style='font-size:.68rem;word-break:break-all'>{H.escape(reason)}</td></tr>")
        if len(rows) >= 5000: break
with open(sys.argv[2], 'w') as f:
    f.write('\n'.join(rows))
PYEOF

    local TABLE_ROWS; TABLE_ROWS=$(cat "$TMP_USN_ROWS" 2>/dev/null || echo "")
    {
        html_header "USN Journal"
        html_page_header "UJ" "USN Journal — <span>Change Log</span> NTFS" \
            "\$Extend\\\$UsnJrnl:\$J" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Record totali</div><div class='value'>${TOTAL_LINES}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Change Journal NTFS — Timestamp · File · Ext · Azione</div>
        <div class='card'><table>
          <thead><tr><th style='width:14%'>Timestamp</th><th>Nome file</th><th style='width:6%'>Ext</th><th>Motivo</th></tr></thead>
          <tbody>${TABLE_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun record</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 24 — NTDS.dit (Active Directory hashes)
# ================================================================
module_ntds() {
    section_header "NTDS.dit — Active Directory Hashes" "$RED"
    check_win_root || return 1


    local NTDS_FILE
    NTDS_FILE=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f \
        -not -path "*/WinSxS/*" -not -path "*/WinSXS/*" \
        2>/dev/null | head -1)
    # Fallback: try canonical path via ci_find_dir if find returned nothing
    if [[ -z "$NTDS_FILE" ]]; then
        local _NTDS_DIR; _NTDS_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/NTDS")
        [[ -n "$_NTDS_DIR" ]] && NTDS_FILE=$(find "$_NTDS_DIR" -maxdepth 1 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    fi
    if [[ -z "$NTDS_FILE" ]]; then
        warn "ntds.dit non trovato — questo non è un Domain Controller (o la struttura NTDS/ è assente)"
        return 0
    fi
    local SYS_HIVE; SYS_HIVE=$(get_hive "SYSTEM")
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "Hive SYSTEM non trovato (necessario per il BootKey)"; return 1
    fi
    info "NTDS.dit: $NTDS_FILE"
    info "SYSTEM:   $SYS_HIVE"
    echo ""

    if ! "$PY3" -c "from impacket.examples.secretsdump import NTDSHashes" 2>/dev/null; then
        warn "impacket non trovato — installalo con: ${PY3} -m pip install impacket"
        info "Fallback: estrazione account senza hash tramite strings..."
        local ACC_STRINGS
        ACC_STRINGS=$(strings "$NTDS_FILE" 2>/dev/null | grep -E '^[A-Za-z0-9_.-]{3,20}\$?$' | sort -u | head -100 || true)
        echo "$ACC_STRINGS" | while IFS= read -r L; do
            printf "  ${DIM}%s${RESET}\n" "$L"
        done
        warn "Installa impacket per l'estrazione degli hash NTLM."
        return 0
    fi

    info "Estrazione hash con impacket NTDSHashes..."
    local TMP_DIR; TMP_DIR=$(mktemp -d /tmp/dfir_ntds_XXXXXX)
    trap '[[ -n "${TMP_DIR:-}" ]] && rm -rf "$TMP_DIR"' RETURN

    # Copia con timeout: su share SMB da DC live la copia può bloccarsi indefinitamente
    info "Copia ntds.dit in /tmp (timeout 180s — file di rete, attendere)..."
    if ! portable_timeout 180 cp "$NTDS_FILE" "$TMP_DIR/ntds.dit" 2>/dev/null; then
        err "Timeout o errore nella copia di ntds.dit (share lenta o file bloccato da Windows)"
        info "Suggerimento: esegui il modulo su un'immagine montata offline oppure tramite VSS snapshot"
        return 1
    fi
    info "Copia SYSTEM hive..."
    if ! portable_timeout 60 cp "$SYS_HIVE" "$TMP_DIR/SYSTEM" 2>/dev/null; then
        err "Timeout o errore nella copia del SYSTEM hive"
        return 1
    fi
    chmod 600 "$TMP_DIR/ntds.dit" "$TMP_DIR/SYSTEM" 2>/dev/null || true

    # Patch dirty state ESE: NTDSHashes usa ESENT_DB internamente, che rifiuta
    # database in stato DirtyShutdown (3) — condizione normale su DC live.
    # Offset 52 (0x34): campo dbstate nello standard ESE/JET Blue.
    info "Verifica/patch dirty state ESE (necessario per DC live)..."
    "$PY3" - "$TMP_DIR/ntds.dit" << 'PYEOF_PATCH' 2>/dev/null || true
import sys, struct
path = sys.argv[1]
try:
    with open(path, 'r+b') as f:
        f.seek(4)
        if struct.unpack('<I', f.read(4))[0] == 0x89ABCDEF:
            f.seek(52)
            state = struct.unpack('<I', f.read(4))[0]
            if state in (3, 4, 5):
                f.seek(52)
                f.write(struct.pack('<I', 2))
                sys.stderr.write(f"INFO: dirty state {state}→2 (CleanShutdown) patchato\n")
except Exception as ex:
    sys.stderr.write(f"WARN: patch dirty state fallita: {ex}\n")
PYEOF_PATCH

    info "Dump hash NTLM (può richiedere diversi minuti per domini grandi)..."
    local DUMP_OUT
    DUMP_OUT=$(portable_timeout 600 "$PY3" - "$TMP_DIR/ntds.dit" "$TMP_DIR/SYSTEM" << 'PYEOF' 2>/dev/null || true
import sys, io
try:
    from impacket.examples.secretsdump import LocalOperations, NTDSHashes
    sys_ops = LocalOperations(sys.argv[2])
    boot_key = sys_ops.getBootKey()
    class _CB:
        def __init__(self): self.lines = []
        def __call__(self, secret): self.lines.append(secret); print(secret)
    cb = _CB()
    ntds = NTDSHashes(sys.argv[1], boot_key, isRemote=False, history=False,
                      noLMHash=True, remoteOps=None, useVSSMethod=False,
                      justNTLM=True, pwdLastSet=False, resumeSession=None,
                      outputFileName=None, justUser=None, ldapFilter=None,
                      printUserStatus=True)
    ntds.dump()
    ntds.finish()
except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
    if [[ -z "$DUMP_OUT" ]]; then
        if grep -qiE 'Errno 35|EAGAIN|temporarily unavailable' "$TMP_DIR/ntds_err.log" 2>/dev/null \
        || [[ "$NTDS_FILE" == /Volumes/* || "$NTDS_FILE" == /mnt/* || "$NTDS_FILE" == /media/* ]]; then
            err "ntds.dit è bloccato dal servizio Active Directory (DC live, Errno 35 / EAGAIN)."
            info "Soluzioni: 1) ntdsutil IFM sul DC  2) impacket-secretsdump remoto  3) immagine disco offline"
        else
            err "Dump vuoto — ntds.dit potrebbe essere in stato dirty o corrotto."
            info "Prova il modulo 38 (PAD Offline) che usa analisi ESE con recovery avanzato."
        fi
        return 1
    fi

    declare -a ACC_ENTRIES=()
    local TOTAL=0
    local KNOWN_EMPTY="31d6cfe0d16ae931b73c59d7e0c089c0"

    while IFS= read -r LINE; do
        [[ -z "$LINE" ]] && continue
        [[ "$LINE" == ERROR* ]] && { err "${LINE#ERROR	}"; continue; }
        if echo "$LINE" | grep -qE '^[^:]+:[0-9]+:[a-fA-F0-9]{32}:[a-fA-F0-9]{32}'; then
            IFS=':' read -r UNAME RID LM NT REST <<< "$LINE"
            local IS_EMPTY=false; [[ "${NT,,}" == "$KNOWN_EMPTY" ]] && IS_EMPTY=true
            local IS_DISABLED=false; echo "$LINE" | grep -qi "disabled\|account disabled\|(disabled)" && IS_DISABLED=true
            printf "  ${CYAN}%-30s${RESET}  RID:${DIM}%-6s${RESET}  NT:${YELLOW}%s${RESET}" "$UNAME" "$RID" "$NT"
            $IS_EMPTY    && printf "  ${DIM}(password vuota)${RESET}"
            $IS_DISABLED && printf "  ${RED}[DISABLED]${RESET}"
            echo ""
            ACC_ENTRIES+=("${UNAME}§${RID}§${NT}§${IS_EMPTY}§${IS_DISABLED}")
            TOTAL=$((TOTAL + 1))
        else
            echo -e "  ${DIM}$LINE${RESET}"
        fi
    done <<< "$DUMP_OUT"

    separator
    info "Account AD estratti: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun account estratto."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ntds")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ACC_ENTRIES[@]}"; do
        IFS='§' read -r UNAME RID NT IS_EMPTY IS_DISABLED <<< "$E"
        local ROW_STYLE="" NT_CSS="mono"
        [[ "$IS_DISABLED" == "true" ]] && ROW_STYLE="style='opacity:.5'"
        [[ "$IS_EMPTY"    == "true" ]] && NT_CSS="mono warn"
        local NOTES=""
        [[ "$IS_EMPTY"    == "true" ]] && NOTES+="<span class='dim' style='font-size:.68rem'>password vuota</span> "
        [[ "$IS_DISABLED" == "true" ]] && NOTES+="<span style='color:var(--text-dim);font-size:.68rem'>DISABLED</span>"
        [[ "$RID" == "500" ]]          && NOTES+="<span style='color:var(--accent2);font-size:.68rem'>Administrator</span>"
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono'>${UNAME}</td>
          <td class='mono mid' style='white-space:nowrap'>${RID}</td>
          <td class='${NT_CSS}' style='font-size:.72rem;word-break:break-all'>${NT}</td>
          <td style='font-size:.72rem'>${NOTES}</td>
        </tr>"
    done
    {
        html_header "NTDS.dit"
        html_page_header "AD" "NTDS.dit — <span>Active Directory</span> Hashes" \
            "Windows\\NTDS\\ntds.dit" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Account</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>Hash NTLM account AD — craccabili offline con hashcat -m 1000</div>
        <div class='card'><table>
          <thead><tr><th>Username</th><th style='width:8%'>RID</th><th style='width:30%'>NT Hash</th><th>Note</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Cracking offline</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            hashcat -m 1000 ntds_hashes.txt wordlist.txt<br>
            john --format=NT ntds_hashes.txt
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 25 — Hibernation / Pagefile (artefatti memoria)
# ================================================================
module_hiberfil() {
    section_header "Hibernation / Pagefile — Artefatti Memoria" "$BLUE"
    check_win_root || return 1


    local HFILE PFILE
    HFILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "hiberfil.sys" -type f 2>/dev/null | head -1)
    PFILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "pagefile.sys" -type f 2>/dev/null | head -1)

    [[ -z "$HFILE" && -z "$PFILE" ]] && {
        warn "hiberfil.sys e pagefile.sys non trovati nella root del volume"
        info "Il volume potrebbe essere montato senza visibilità sui file di sistema"
        return 0
    }

    declare -a FILE_ENTRIES=()
    local TOTAL_STRINGS=0

    for MFILE in "$HFILE" "$PFILE"; do
        [[ -z "$MFILE" || ! -f "$MFILE" ]] && continue
        local MNAME; MNAME=$(basename "$MFILE")
        local MSIZE; MSIZE=$(du -h "$MFILE" 2>/dev/null | cut -f1 || echo "?")
        ok "$MNAME trovato: $MFILE  ($MSIZE)"

        # Verifica magic bytes per hiberfil.sys
        local MAGIC=""
        if [[ "$MNAME" =~ hiberfil ]]; then
            MAGIC=$(dd if="$MFILE" bs=4 count=1 2>/dev/null | od -An -tx1 | tr -d ' \n' | cut -c1-8 || echo "")
            case "${MAGIC^^}" in
                "5041474D") info "  Formato: HIBR (hibernated)" ;;
                "52535452") info "  Formato: RSTR (restored)" ;;
                *) info "  Magic: 0x${MAGIC} (sconosciuto)" ;;
            esac
        fi

        info "  Estrazione stringhe significative..."
        local STR_OUT
        STR_OUT=$(python3 - "$MFILE" << 'PYEOF' 2>/dev/null || true
import sys, re
path = sys.argv[1]
PATTERNS = [
    (r'https?://[a-zA-Z0-9_/\-\.\?&=%~+#]+', 'URL'),
    (r'\b(?:\d{1,3}\.){3}\d{1,3}\b', 'IP'),
    (r'[A-Za-z]:\\[a-zA-Z0-9_\\\-\.\s\(\)~!@#%&\+,;=\[\]\{\}\^]{5,200}', 'PATH'),
    (r'(?i)(?:password|passwd|pwd|secret|token|apikey|authorization)[=:\s]+[a-zA-Z0-9_!@#\$\^&\*\-\.\?]{4,50}', 'CRED'),
    (r'(?i)(?:cmd\.exe|powershell|wscript|mshta|certutil|bitsadmin)[ -~]{0,100}', 'EXEC'),
    (r'[A-Za-z0-9+/]{40,}={0,2}', 'B64_CANDIDATE'),
]
results = []
try:
    CHUNK = 1024*1024*4  # 4MB at a time
    with open(path,'rb') as f:
        while True:
            chunk = f.read(CHUNK)
            if not chunk: break
            text = chunk.decode('utf-16-le',errors='replace') + chunk.decode('latin-1',errors='replace')
            for pat, label in PATTERNS:
                for m in re.findall(pat, text):
                    v = m.strip(' "''<>\\r\\n')[:200]
                    # Filter out purely non-ascii / bad decoding artifacts
                    if len(v) > 6 and v not in results and v.isascii():
                        results.append(v)
                        print(f"{label}\t{v}")
            if len(results) > 500: break
except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
        local COUNT; COUNT=$(echo "$STR_OUT" | wc -l)
        TOTAL_STRINGS=$((TOTAL_STRINGS + COUNT))
        echo "$STR_OUT" | grep -v "^ERROR" | head -20 | while IFS=$'\t' read -r LBL VAL; do
            local COLOR="$DIM"
            [[ "$LBL" == "CRED" || "$LBL" == "URL" ]] && COLOR="$RED"
            printf "  ${CYAN}%-16s${RESET} ${COLOR}%s${RESET}\n" "$LBL" "${VAL:0:100}"
        done
        echo ""
        FILE_ENTRIES+=("${MNAME}§${MFILE}§${MSIZE}§${MAGIC}§$(echo "$STR_OUT" | grep -v "^ERROR" | head -100 | base64 -w0 2>/dev/null || true)")
    done

    separator
    info "Stringhe estratte totali: ${BOLD}$TOTAL_STRINGS"
    info "${BOLD}Per analisi completa della memoria usa Volatility3:${RESET}"
    echo -e "    ${DIM}vol -f \"${HFILE:-<hiberfil.sys>}\" windows.pstree${RESET}"
    echo -e "    ${DIM}vol -f \"${HFILE:-<hiberfil.sys>}\" windows.netscan${RESET}"
    echo -e "    ${DIM}vol -f \"${HFILE:-<hiberfil.sys>}\" windows.cmdline${RESET}"

    [[ ${#FILE_ENTRIES[@]} -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "hiberfil")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""
    for E in "${FILE_ENTRIES[@]}"; do
        IFS='§' read -r MNAME MFILE MSIZE MAGIC B64_STRINGS <<< "$E"
        local STR_ROWS=""
        local DECODED; DECODED=$(echo "$B64_STRINGS" | base64 -d 2>/dev/null || true)
        while IFS=$'\t' read -r LBL VAL; do
            [[ -z "$LBL" ]] && continue
            local CSS="mono fld"; [[ "$LBL" == "CRED" || "$LBL" == "URL" ]] && CSS="mono fld bad"
            STR_ROWS+="<tr><td class='mono' style='white-space:nowrap;color:var(--accent4)'>${LBL}</td>
              <td><span class='${CSS}' style='font-size:.72rem'>$(html_esc "${VAL:0:200}")</span></td></tr>"
        done <<< "$DECODED"
        CARDS_HTML+="<div class='card'>
          <div class='card-header'><div class='uicon'>MEM</div>
            <div><div class='uname'>${MNAME}</div><div class='upath'>${MFILE}</div></div>
            <div class='badge'>${MSIZE}</div>
          </div>
          <table><thead><tr><th style='width:14%'>Tipo</th><th>Valore estratto</th></tr></thead>
          <tbody>${STR_ROWS:-<tr><td colspan='2' class='dim' style='padding:.8rem'>Nessuna stringa significativa estratta</td></tr>}</tbody></table>
          <div style='padding:.8rem 1.5rem;font-family:var(--mono);font-size:.7rem;color:var(--text-dim)'>
            Magic: <span style='color:var(--accent)'>0x${MAGIC}</span> &nbsp;·&nbsp;
            Per analisi completa: <span style='color:var(--accent4)'>vol -f \"${MFILE}\" windows.pstree</span>
          </div></div>"
    done
    {
        html_header "Hibernation"
        html_page_header "HB" "Hibernation / <span>Pagefile</span> Forensics" \
            "hiberfil.sys · pagefile.sys" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File trovati</div><div class='value'>${#FILE_ENTRIES[@]}</div></div>
          <div class='stat info'><div class='label'>Stringhe estratte</div><div class='value'>${TOTAL_STRINGS}</div></div>
        </div><main>
        <style>.fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}</style>
        <div class='stitle'>Artefatti in memoria — analisi tramite strings (parziale)</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 26 — WER Files (Windows Error Reporting)
# ================================================================
module_wer_files() {
    section_header "WER Files — Windows Error Reporting" "$MAGENTA"
    check_win_root || return 1

    local TOTAL_FILES=0
    declare -a WER_FILES=()

    # Per-system WER (ProgramData) - ordina per data decrescente
    local SYSPATHS=(
        "ProgramData/Microsoft/Windows/WER/ReportArchive"
        "ProgramData/Microsoft/Windows/WER/ReportQueue"
        "Windows/System32/WER/ReportArchive"
        "Windows/System32/WER/ReportQueue"
    )
    for REL in "${SYSPATHS[@]}"; do
        local DIR
        DIR=$(ci_find_dir "$WIN_ROOT" "$REL")
        [[ -n "$DIR" && -d "$DIR" ]] || continue
        while IFS= read -r FDIR; do
            [[ -d "$FDIR" ]] || continue
            local WERF
            WERF=$(find "$FDIR" -maxdepth 1 -iname "Report.wer" -type f 2>/dev/null | head -1)
            [[ -n "$WERF" ]] && WER_FILES+=( "$WERF|System" )
        done < <(find "$DIR" -maxdepth 1 -mindepth 1 -type d -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
    done

    # Per-user WER (AppData/Local)
    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local USERPATHS=(
            "AppData/Local/Microsoft/Windows/WER/ReportArchive"
            "AppData/Local/Microsoft/Windows/WER/ReportQueue"
        )
        for REL in "${USERPATHS[@]}"; do
            local DIR
            DIR=$(ci_find_dir "$USER_DIR" "$REL")
            [[ -n "$DIR" && -d "$DIR" ]] || continue
            while IFS= read -r FDIR; do
                [[ -d "$FDIR" ]] || continue
                local WERF
                WERF=$(find "$FDIR" -maxdepth 1 -iname "Report.wer" -type f 2>/dev/null | head -1)
                [[ -n "$WERF" ]] && WER_FILES+=( "$WERF|$USERNAME" )
            done < <(find "$DIR" -maxdepth 1 -mindepth 1 -type d -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        done
    done < <(get_user_homes)

    TOTAL_FILES=${#WER_FILES[@]}
    info "File WER trovati: ${BOLD}$TOTAL_FILES"

    if [[ $TOTAL_FILES -eq 0 ]]; then
        ok "Nessun file WER trovato."
        return 0
    fi

    local SCAN_COUNT=0
    for ENTRY in "${WER_FILES[@]}"; do
        SCAN_COUNT=$((SCAN_COUNT+1))
        local F="${ENTRY%|*}"
        local OWNER="${ENTRY#*|}"
        local FNAME; FNAME=$(basename "$(dirname "$F")")
        local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
        echo -e "      ${DIM}[$(printf "%2d" $SCAN_COUNT)] • $FNAME  (${OWNER}) — $FMTIME${RESET}"
    done

    separator
    info "Processati ${BOLD}$TOTAL_FILES${RESET} report di errore."

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wer_files")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""

    for ENTRY in "${WER_FILES[@]}"; do
        local F="${ENTRY%|*}"
        local OWNER="${ENTRY#*|}"
        local FDIR_NAME; FDIR_NAME=$(basename "$(dirname "$F")")
        local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
        
        # Caricamento contenuto file .wer (solitamente UTF-16LE con BOM)
        local CONTENT
        CONTENT=$("$PY3" - "$F" << 'PYEOF'
import sys, html
try:
    with open(sys.argv[1], 'rb') as f:
        raw = f.read()
    # Rilevamento encoding
    if raw.startswith(b'\xff\xfe'):
        text = raw[2:].decode('utf-16-le', errors='replace')
    elif raw.startswith(b'\xfe\xff'):
        text = raw[2:].decode('utf-16-be', errors='replace')
    elif raw.startswith(b'\xef\xbb\xbf'):
        text = raw[3:].decode('utf-8', errors='replace')
    else:
        # Tenta UTF-16-LE comunque se ci sono molti null
        if len(raw) >= 2 and raw[1] == 0:
            text = raw.decode('utf-16-le', errors='replace')
        else:
            text = raw.decode('utf-8', errors='replace')
    print(html.escape(text.strip()))
except Exception as e:
    print(f"Errore caricamento: {html.escape(str(e))}")
PYEOF
)
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon' style='background:linear-gradient(135deg,var(--accent2),#c0392b)'>!</div>
            <div>
              <div class='uname'>${FDIR_NAME}</div>
              <div class='upath'>Sorgente: ${OWNER} &nbsp;·&nbsp; Data: ${FMTIME}</div>
            </div>
            <div class='badge warn'>Report.wer</div>
          </div>
          <div style='padding:0.8rem 1.2rem; background:var(--bg3); border-top:1px solid var(--border)'>
            <pre style='font-family:var(--mono); font-size:0.75rem; color:var(--text); line-height:1.5; white-space:pre-wrap; max-height:350px; overflow-y:auto'>${CONTENT}</pre>
          </div>
        </div>"
    done

    {
        html_header "Windows Error Reporting (WER)"
        html_page_header "WER" "Windows <span>Error Reporting</span> Analysis" \
            "WER\\ReportArchive | WER\\ReportQueue" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Totale Report</div><div class='value'>${TOTAL_FILES}</div></div>
        </div><main>
        <div class='stitle'>Dati rilevati dalle code di errore di sistema</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 27 — Credential Manager (DPAPI blobs)
# ================================================================
module_credential_manager() {
    section_header "Credential Manager — DPAPI Blobs" "$MAGENTA"
    check_win_root || return 1


    declare -a ALL_CREDS=()
    local TOTAL=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        for REL in "AppData/Local/Microsoft/Credentials" "AppData/Roaming/Microsoft/Credentials" \
                   "AppData/Local/Microsoft/Vault" "AppData/Roaming/Microsoft/Vault"; do
            local CRED_DIR
            CRED_DIR=$(ci_find_dir "$USER_DIR" "$REL")
            [[ -z "$CRED_DIR" || ! -d "$CRED_DIR" ]] && continue
            while IFS= read -r CFILE; do
                [[ -f "$CFILE" ]] || continue
                local FNAME; FNAME=$(basename "$CFILE")
                local FSIZE; FSIZE=$(stat -c "%s" "$CFILE" 2>/dev/null || echo "?")
                local FMTIME; FMTIME=$(stat -c "%y" "$CFILE" 2>/dev/null | cut -d'.' -f1 || echo "?")
                # Parse DPAPI blob header
                local DPAPI_INFO
                DPAPI_INFO=$("$PY3" - "$CFILE" << 'PYEOF' 2>/dev/null || true
import sys, struct, binascii
path = sys.argv[1]
try:
    with open(path,'rb') as f: data = f.read()
    # DPAPI blob: version(4) + provider_guid(16) + mk_guid(16) + ...
    if len(data) < 36: sys.exit()
    # Cerca DPAPI magic 01 00 00 00 44 50 41 50 49
    idx = data.find(b'\x01\x00\x00\x00\x44\x50\x41\x50\x49')
    if idx == -1: idx = 0
    mk_off = idx + 20  # after version + provider guid
    if mk_off + 16 <= len(data):
        mk = data[mk_off:mk_off+16]
        mk_guid = '{%s-%s-%s-%s-%s}' % (
            binascii.hexlify(mk[0:4]).decode(), binascii.hexlify(mk[4:6]).decode(),
            binascii.hexlify(mk[6:8]).decode(), binascii.hexlify(mk[8:10]).decode(),
            binascii.hexlify(mk[10:16]).decode())
        print(f"DPAPI\t{mk_guid}\t{len(data)}")
    else:
        print(f"BLOB\t-\t{len(data)}")
except Exception as e:
    print(f"ERROR\t{e}\t0")
PYEOF
)
                local DTYPE MKGUID DSIZE
                IFS=$'\t' read -r DTYPE MKGUID DSIZE <<< "$DPAPI_INFO"
                [[ "$DTYPE" == "ERROR" ]] && continue
                local LABEL; LABEL=$(basename "$(dirname "$CFILE")")
                printf "  ${MAGENTA}[%-15s]${RESET}  ${CYAN}%-30s${RESET}  ${DIM}MK: %s${RESET}\n" \
                    "$USERNAME" "$FNAME" "${MKGUID:--}"
                ALL_CREDS+=("${USERNAME}§${FNAME}§${CFILE}§${REL}§${MKGUID}§${FSIZE}§${FMTIME}")
                TOTAL=$((TOTAL + 1))
            done < <(find "$CRED_DIR" -maxdepth 2 -type f 2>/dev/null)
        done
    done < <(get_user_homes)

    separator
    info "Blob DPAPI trovati: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun blob DPAPI trovato."; return 0; }
    warn "I blob sono cifrati con DPAPI — richiedono la master key utente per la decifratura"
    info "Tool utili: dpapick3, mimikatz (sekurlsa::dpapi), impacket dpapi"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "credential_manager")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente
    mapfile -t ALL_CREDS < <(printf '%s\n' "${ALL_CREDS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+7*[""])[6],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_CREDS[@]}"; do
        IFS='§' read -r UNAME FNAME CPATH REL MKGUID FSIZE FMTIME <<< "$E"
        ROWS+="<tr>
          <td class='mono dim' style='white-space:nowrap'>${UNAME}</td>
          <td class='mono' style='font-size:.72rem;word-break:break-all'>$(html_esc "$FNAME")</td>
          <td class='mono mid' style='font-size:.68rem;word-break:break-all'>$(html_esc "$REL")</td>
          <td class='mono dim' style='font-size:.65rem;word-break:break-all'>${MKGUID:--}</td>
          <td class='mono mid' style='white-space:nowrap'>${FSIZE} B</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${FMTIME}</td>
        </tr>"
    done
    {
        html_header "Credential Manager"
        html_page_header "CM" "Credential Manager — <span>DPAPI Blobs</span>" \
            "AppData\\Local\\Microsoft\\Credentials" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Blob trovati</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>Blob DPAPI — Utente · File · Path · MasterKey GUID · Dim · Timestamp</div>
        <div class='card'><table>
          <thead><tr><th>Utente</th><th>File</th><th>Store</th><th>MK GUID</th><th>Dim.</th><th>Modificato</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Decifratura offline</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            impacket: dpapi.py masterkey -file &lt;MK&gt; -password &lt;pwd&gt;<br>
            mimikatz: sekurlsa::dpapi (su sistema live)
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 28 — WLAN & VPN Profiles
# ================================================================
module_wlan() {
    section_header "WLAN & VPN — Profili di Rete" "$CYAN"
    check_win_root || return 1


    declare -a ALL_WLAN=()
    local TOTAL_WLAN=0

    # WLAN Profiles: ProgramData\Microsoft\Wlansvc\Profiles\Interfaces\*\*.xml
    local WLANSVC_DIR
    WLANSVC_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Wlansvc/Profiles/Interfaces")
    if [[ -n "$WLANSVC_DIR" && -d "$WLANSVC_DIR" ]]; then
        info "WLAN Profiles: $WLANSVC_DIR"
        while IFS= read -r WXML; do
            [[ -f "$WXML" ]] || continue
            local WOUT
            WOUT=$("$PY3" - "$WXML" << 'PYEOF' 2>/dev/null || true
import sys, xml.etree.ElementTree as ET
try:
    tree = ET.parse(sys.argv[1])
    root = tree.getroot()
    ns = {'w': 'http://www.microsoft.com/networking/WLAN/profile/v1'}
    def gt(tag):
        el = root.find(f'.//{{http://www.microsoft.com/networking/WLAN/profile/v1}}{tag}')
        return el.text.strip() if el is not None and el.text else ''
    ssid   = gt('name') or gt('SSID') or gt('SSIDConfig')
    auth   = gt('authentication')
    enc    = gt('encryption')
    key    = gt('keyMaterial')  # WEP/WPA PSK in chiaro
    print(f"{ssid}\t{auth}\t{enc}\t{key}")
except Exception as e:
    print(f"ERROR\t{e}\t\t")
PYEOF
)
            IFS=$'\t' read -r SSID AUTH ENC KEYMAT <<< "$WOUT"
            [[ "$SSID" == "ERROR" || -z "$SSID" ]] && continue
            local HAS_KEY=""; [[ -n "$KEYMAT" ]] && HAS_KEY="${RED}[KEY IN CHIARO]${RESET}"
            printf "  ${CYAN}%-35s${RESET}  %-15s  %-10s  %b\n" "$SSID" "$AUTH" "$ENC" "$HAS_KEY"
            ALL_WLAN+=("WLAN§${SSID}§${AUTH}§${ENC}§${KEYMAT}§${WXML}")
            TOTAL_WLAN=$((TOTAL_WLAN + 1))
        done < <(find "$WLANSVC_DIR" -name "*.xml" -type f 2>/dev/null)
    else
        warn "WLAN Profiles non trovati (ProgramData/Microsoft/Wlansvc/Profiles)"
    fi

    # VPN / Network profiles dal registro SOFTWARE
    echo ""
    info "Profili di rete (NetworkList) dal registro..."
    if check_regipy; then
        local SW_HIVE; SW_HIVE=$(get_hive "SOFTWARE")
        if [[ -n "$SW_HIVE" ]]; then
            local NET_DATA
            NET_DATA=$("$PY3" - "$SW_HIVE" << 'PYEOF' 2>/dev/null || true
import sys, datetime
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for key_path in ['Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles',
                     'Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Signatures\\Unmanaged',
                     'Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Signatures\\Managed']:
        try:
            key = hive.get_key(key_path)
            for sk in key.iter_subkeys():
                vals = {v.name.lower(): v.value for v in sk.get_values()}
                name = str(vals.get('profilename', vals.get('firstname', sk.name)))
                desc = str(vals.get('description',''))
                cat  = str(vals.get('category',''))
                gw   = str(vals.get('defaultgatewayipv4', vals.get('defaultgatewayipv6','')))
                dns  = str(vals.get('dnsserver',''))
                print(f"NET\t{name}\t{desc}\t{cat}\t{gw}\t{dns}")
        except: continue
except Exception as e:
    pass
PYEOF
)
            while IFS=$'\t' read -r TYPE NAME DESC CAT GW DNS; do
                [[ -z "$NAME" ]] && continue
                printf "  ${GREEN}%-40s${RESET}  ${DIM}cat:%s  gw:%s${RESET}\n" "$NAME" "$CAT" "$GW"
                ALL_WLAN+=("VPN§${NAME}§${CAT}§${GW}§${DNS}§-")
                TOTAL_WLAN=$((TOTAL_WLAN + 1))
            done <<< "$NET_DATA"
        fi
    fi

    separator
    info "Profili di rete totali: ${BOLD}$TOTAL_WLAN"
    [[ $TOTAL_WLAN -eq 0 ]] && { warn "Nessun profilo trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wlan")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_WLAN[@]}"; do
        IFS='§' read -r TYPE NAME F2 F3 F4 F5 <<< "$E"
        local CSS=""; [[ "$TYPE" == "WLAN" && -n "$F4" ]] && CSS="style='background:rgba(255,123,114,.07)'"
        local KEYHTML=""; [[ "$TYPE" == "WLAN" && -n "$F4" ]] && KEYHTML="<span class='bad' style='font-size:.68rem'>KEY: $(html_esc "${F4:0:40}")</span>"
        ROWS+="<tr ${CSS}>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>${TYPE}</td>
          <td class='mono'>${NAME} ${KEYHTML}</td>
          <td class='mono dim' style='white-space:nowrap'>${F2}</td>
          <td class='mono mid' style='white-space:nowrap'>${F3}</td>
        </tr>"
    done
    {
        html_header "WLAN VPN"
        html_page_header "WL" "WLAN / VPN — <span>Profili di Rete</span>" \
            "ProgramData\\Microsoft\\Wlansvc + NetworkList" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Profili totali</div><div class='value'>${TOTAL_WLAN}</div></div>
        </div><main>
        <div class='stitle'>Reti WiFi e profili VPN/network memorizzati</div>
        <div class='card'><table>
          <thead><tr><th style='width:8%'>Tipo</th><th>Nome / SSID</th><th style='width:12%'>Auth/Cat</th><th style='width:15%'>Gateway/Enc</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 29 — AppX / UWP Packages
# ================================================================
module_appx() {
    section_header "AppX / UWP — Pacchetti Store Installati" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. Installalo con: ${PY3} -m pip install regipy"; return 1
    fi


    declare -a ALL_PKGS=()
    local TOTAL=0 SUSP_COUNT=0

    # SOFTWARE hive (sistema)
    local SW_HIVE; SW_HIVE=$(get_hive "SOFTWARE")
    if [[ -n "$SW_HIVE" ]]; then
        info "Parsing AppX packages da SOFTWARE hive..."
        local PKG_DATA
        PKG_DATA=$("$PY3" - "$SW_HIVE" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for key_path in [
        'Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages',
        'Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages',
    ]:
        try:
            root = hive.get_key(key_path)
            for sk in root.iter_subkeys():
                vals = {v.name.lower(): str(v.value) for v in sk.get_values()}
                name = sk.name
                path = vals.get('packagerootfolder', vals.get('installlocation',''))
                pub  = vals.get('packageid.publisher', vals.get('publisher',''))
                ver  = vals.get('packageid.version','')
                arch = vals.get('packageid.processorarchitecture','')
                print(f"SYS\t{name}\t{path}\t{pub}\t{ver}\t{arch}")
        except: continue
except: pass
PYEOF
)
        while IFS=$'\t' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER PKG_ARCH; do
            [[ -z "$PKG_NAME" ]] && continue
            local IS_SUSP="false"
            echo "$PKG_PATH" | grep -qiv "program files\|windows apps\|system32" && \
                [[ -n "$PKG_PATH" ]] && IS_SUSP="true"
            [[ "${PKG_PUB,,}" == *"cn="* ]] || echo "$PKG_NAME" | grep -qi "sideload\|test\|debug" && IS_SUSP="true"
            ALL_PKGS+=("${SCOPE}§${PKG_NAME}§${PKG_PATH}§${PKG_PUB}§${PKG_VER}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
            if [[ "$IS_SUSP" == "true" ]]; then
                printf "  ${RED}[!] %-50s${RESET}  %s\n" "$PKG_NAME" "$PKG_PATH"
            else
                printf "  ${DIM}%-50s${RESET}\n" "${PKG_NAME:0:60}"
            fi
        done <<< "$PKG_DATA"
    fi

    # Per utente: NTUSER.DAT
    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        local UPKG_DATA
        UPKG_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    for key_path in ['Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages']:
        try:
            root = hive.get_key(key_path)
            for sk in root.iter_subkeys():
                vals = {v.name.lower(): str(v.value) for v in sk.get_values()}
                path = vals.get('packagerootfolder','')
                pub  = vals.get('publisher','')
                print(f"USR:{user}\t{sk.name}\t{path}\t{pub}\t\t")
        except: continue
except: pass
PYEOF
)
        while IFS=$'\t' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER PKG_ARCH; do
            [[ -z "$PKG_NAME" ]] && continue
            local IS_SUSP="false"
            echo "$PKG_PATH" | grep -qiv "program files\|windows apps" && [[ -n "$PKG_PATH" ]] && IS_SUSP="true"
            ALL_PKGS+=("${SCOPE}§${PKG_NAME}§${PKG_PATH}§${PKG_PUB}§${PKG_VER}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
        done <<< "$UPKG_DATA"
    done < <(get_user_homes)

    separator
    info "Pacchetti AppX: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun pacchetto AppX trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "appx")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_PKGS[@]}"; do
        IFS='§' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER IS_SUSP <<< "$E"
        local RS=""; [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        local NC="mono"; [[ "$IS_SUSP" == "true" ]] && NC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap;font-size:.68rem'>${SCOPE}</td>
          <td class='${NC}' style='word-break:break-all;font-size:.72rem'>$(html_esc "${PKG_NAME:0:80}")</td>
          <td class='mono dim' style='word-break:break-all;font-size:.68rem'>$(html_esc "$PKG_PUB")</td>
          <td class='mono mid' style='word-break:break-all;font-size:.68rem'>$(html_esc "$PKG_PATH")</td>
        </tr>"
    done
    {
        html_header "AppX"
        html_page_header "AX" "AppX / UWP — <span>Pacchetti Store</span>" \
            "AppModel\\Repository\\Packages" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Pacchetti</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Pacchetti UWP installati — path fuori Windows Apps = potenziale sideload</div>
        <div class='card'><table>
          <thead><tr><th style='width:10%'>Scope</th><th>Package Name</th><th style='width:20%'>Publisher</th><th>Install Path</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 30 — Browser Downloads & Login Data
# ================================================================
module_browser_extra() {
    section_header "Browser Downloads & Login Data" "$CYAN"
    check_win_root || return 1

    if ! "$PY3" -c "import sqlite3" 2>/dev/null; then
        err "sqlite3 non disponibile"; return 1
    fi


    declare -a DL_PROFILES=(
        "Chrome|AppData/Local/Google/Chrome/User Data/Default/History|chromium"
        "Edge|AppData/Local/Microsoft/Edge/User Data/Default/History|chromium"
        "Brave|AppData/Local/BraveSoftware/Brave-Browser/User Data/Default/History|chromium"
        "Firefox|AppData/Roaming/Mozilla/Firefox/Profiles|firefox"
    )
    declare -a LOGIN_PROFILES=(
        "Chrome|AppData/Local/Google/Chrome/User Data/Default/Login Data|chromium"
        "Edge|AppData/Local/Microsoft/Edge/User Data/Default/Login Data|chromium"
        "Brave|AppData/Local/BraveSoftware/Brave-Browser/User Data/Default/Login Data|chromium"
    )

    local TMP_DL_ROWS; TMP_DL_ROWS=$(mktemp /tmp/dfir_dlrows_XXXXXX.html)
    local TMP_LG_ROWS; TMP_LG_ROWS=$(mktemp /tmp/dfir_lgrows_XXXXXX.html)
    trap '[[ -n "${TMP_DL_ROWS:-}" ]] && rm -f "$TMP_DL_ROWS"; [[ -n "${TMP_LG_ROWS:-}" ]] && rm -f "$TMP_LG_ROWS"' RETURN
    local TOTAL_DL=0 TOTAL_LG=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        # === Downloads ===
        for SPEC in "${DL_PROFILES[@]}"; do
            IFS='|' read -r BLABEL BPATH BTYPE <<< "$SPEC"
            local DB_PATH
            if [[ "$BTYPE" == "firefox" ]]; then
                local FF_BASE
                FF_BASE=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Mozilla/Firefox/Profiles")
                [[ -z "$FF_BASE" ]] && continue
                for FF_PROF in "$FF_BASE"/*/; do
                    local FF_DB="${FF_PROF}places.sqlite"
                    [[ -f "$FF_DB" ]] || continue
                    local DL_OUT
                    DL_OUT=$("$PY3" - "$FF_DB" "firefox" "$BLABEL" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, sqlite3, shutil, os, tempfile, datetime, html as H
db_path, mode, browser, user = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
tmp = tempfile.mktemp(suffix='.db')
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("""
        SELECT p.url, a.content, h.visit_date
        FROM moz_annos a
        JOIN moz_places p ON a.place_id=p.id
        LEFT JOIN moz_historyvisits h ON h.place_id=p.id
        WHERE a.anno_attribute_id IN (
            SELECT id FROM moz_anno_attributes WHERE name='downloads/destinationFileName'
        )
        ORDER BY h.visit_date DESC LIMIT 500
    """)
    for row in cur.fetchall():
        url = (row[0] or '')[:300]
        dest = (row[1] or '')[:200]
        ts_raw = row[2]
        ts = ''
        if ts_raw:
            try: ts = (datetime.datetime(1970,1,1)+datetime.timedelta(microseconds=ts_raw)).strftime('%Y-%m-%d %H:%M:%S')
            except: pass
        susp = any(p in url.lower() for p in ['.exe','.ps1','.bat','.dll','.hta','.msi','transfer.sh','mega.nz','pastebin'])
        rs = "style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if susp else ""
        uc = "mono fld bad" if susp else "mono fld"
        print(f"<tr {rs}><td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{H.escape(ts)}</td>"
              f"<td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{H.escape(user)} / {H.escape(browser)}</td>"
              f"<td><span class='{uc}' style='font-size:.72rem'>{H.escape(url)}</span></td>"
              f"<td class='mono mid' style='font-size:.7rem;word-break:break-word'>{H.escape(dest)}</td></tr>")
    conn.close()
except Exception as e:
    pass
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF
)
                    if [[ -n "$DL_OUT" ]]; then
                        echo "$DL_OUT" >> "$TMP_DL_ROWS"
                        TOTAL_DL=$((TOTAL_DL + $(echo "$DL_OUT" | wc -l)))
                        ok "$USERNAME / $BLABEL — download trovati"
                    fi
                done
            else
                DB_PATH=$(ci_find_dir "$USER_DIR" "$(dirname "$BPATH")")
                [[ -z "$DB_PATH" ]] && continue
                DB_PATH="${DB_PATH}/$(basename "$BPATH")"
                [[ -f "$DB_PATH" ]] || continue
                local DL_OUT
                DL_OUT=$("$PY3" - "$DB_PATH" "chromium" "$BLABEL" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, sqlite3, shutil, os, tempfile, datetime, html as H
db_path, mode, browser, user = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
tmp = tempfile.mktemp(suffix='.db')
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("""
        SELECT tab_url, target_path, start_time, total_bytes
        FROM downloads ORDER BY start_time DESC LIMIT 500
    """)
    def ct(t):
        if not t: return ''
        try: return (datetime.datetime(1601,1,1)+datetime.timedelta(microseconds=t)).strftime('%Y-%m-%d %H:%M:%S')
        except: return ''
    for row in cur.fetchall():
        url  = (row[0] or '')[:300]
        dest = (row[1] or '')[:200]
        ts   = ct(row[2])
        susp = any(p in url.lower() for p in ['.exe','.ps1','.bat','.dll','.hta','.msi','transfer.sh','mega.nz','pastebin'])
        rs = "style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if susp else ""
        uc = "mono fld bad" if susp else "mono fld"
        print(f"<tr {rs}><td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{H.escape(ts)}</td>"
              f"<td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{H.escape(user)} / {H.escape(browser)}</td>"
              f"<td><span class='{uc}' style='font-size:.72rem'>{H.escape(url)}</span></td>"
              f"<td class='mono mid' style='font-size:.7rem;word-break:break-word'>{H.escape(dest)}</td></tr>")
    conn.close()
except Exception as e:
    pass
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF
)
                if [[ -n "$DL_OUT" ]]; then
                    echo "$DL_OUT" >> "$TMP_DL_ROWS"
                    TOTAL_DL=$((TOTAL_DL + $(echo "$DL_OUT" | wc -l)))
                    ok "$USERNAME / $BLABEL — download trovati"
                fi
            fi
        done

        # === Login Data ===
        for SPEC in "${LOGIN_PROFILES[@]}"; do
            IFS='|' read -r BLABEL BPATH BTYPE <<< "$SPEC"
            local LDB_PATH
            LDB_PATH=$(ci_find_dir "$USER_DIR" "$(dirname "$BPATH")")
            [[ -z "$LDB_PATH" ]] && continue
            LDB_PATH="${LDB_PATH}/$(basename "$BPATH")"
            [[ -f "$LDB_PATH" ]] || continue
            local LG_OUT
            LG_OUT=$("$PY3" - "$LDB_PATH" "$BLABEL" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, sqlite3, shutil, os, tempfile, html as H
db_path, browser, user = sys.argv[1], sys.argv[2], sys.argv[3]
tmp = tempfile.mktemp(suffix='.db')
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("SELECT origin_url, username_value, length(password_value), date_password_changed FROM logins ORDER BY date_password_changed DESC LIMIT 500")
    for row in cur.fetchall():
        url  = (row[0] or '')[:200]
        uname= (row[1] or '???')[:80]
        pwlen= row[2] or 0
        print(f"<tr><td class='mono ok' style='white-space:nowrap;font-size:.7rem'></td>"
              f"<td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{H.escape(user)} / {H.escape(browser)}</td>"
              f"<td class='mono fld' style='font-size:.72rem'>{H.escape(url)}</td>"
              f"<td class='mono' style='white-space:nowrap'>{H.escape(uname)}</td>"
              f"<td class='mono dim' style='white-space:nowrap'>{pwlen}B (DPAPI)</td></tr>")
    conn.close()
except Exception as e:
    pass
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF
)
            if [[ -n "$LG_OUT" ]]; then
                echo "$LG_OUT" >> "$TMP_LG_ROWS"
                TOTAL_LG=$((TOTAL_LG + $(echo "$LG_OUT" | wc -l)))
                ok "$USERNAME / $BLABEL — credenziali salvate trovate"
            fi
        done
    done < <(get_user_homes)

    separator
    info "Download: ${BOLD}$TOTAL_DL${RESET}  |  Credenziali salvate: ${BOLD}$TOTAL_LG"
    [[ $((TOTAL_DL + TOTAL_LG)) -eq 0 ]] && { warn "Nessun dato trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "browser_extra")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local DL_ROWS; DL_ROWS=$(cat "$TMP_DL_ROWS" 2>/dev/null || echo "")
    local LG_ROWS; LG_ROWS=$(cat "$TMP_LG_ROWS" 2>/dev/null || echo "")
    {
        html_header "Browser Downloads"
        html_page_header "BD" "Browser <span>Downloads &amp; Login Data</span>" \
            "Chrome · Edge · Firefox · Brave" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Download</div><div class='value'>${TOTAL_DL}</div></div>
          <div class='stat info'><div class='label'>Login salvati</div><div class='value'>${TOTAL_LG}</div></div>
        </div><main>
        <style>.fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}</style>
        <div class='stitle'>Download — Timestamp · Utente/Browser · URL · Destinazione</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Timestamp</th><th style='width:14%'>Utente/Browser</th><th>URL</th><th style='width:20%'>Destinazione</th></tr></thead>
          <tbody>${DL_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun download trovato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>Login Data — Credenziali salvate (password cifrate DPAPI)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Timestamp</th><th style='width:14%'>Utente/Browser</th><th>URL</th><th style='width:15%'>Username</th><th style='width:10%'>Password</th></tr></thead>
          <tbody>${LG_ROWS:-<tr><td colspan='5' class='dim' style='padding:1rem'>Nessuna credenziale trovata</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 31 — Clipboard History (Win10+)
# ================================================================
module_clipboard() {
    section_header "Clipboard History — Cronologia Appunti" "$YELLOW"
    check_win_root || return 1


    declare -a ALL_CLIPS=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local CB_DIR
        CB_DIR=$(ci_find_dir "$USER_DIR" "AppData/Local/Microsoft/Windows/Clipboard")
        [[ -z "$CB_DIR" || ! -d "$CB_DIR" ]] && continue
        info "Clipboard dir trovata: $CB_DIR"

        while IFS= read -r CFILE; do
            [[ -f "$CFILE" ]] || continue
            local FSIZE; FSIZE=$(stat -c "%s" "$CFILE" 2>/dev/null || echo "0")
            [[ "$FSIZE" -lt 4 ]] && continue
            local CLIP_TEXT
            CLIP_TEXT=$("$PY3" - "$CFILE" << 'PYEOF' 2>/dev/null || true
import sys, re
path = sys.argv[1]
try:
    with open(path,'rb') as f:
        data = f.read()
    # Try UTF-16-LE first (common for clipboard .dat)
    for enc in ('utf-16-le','utf-8','latin-1'):
        try:
            text = data.decode(enc, errors='ignore')
            text = text.replace('\x00','').strip()
            if len(text) > 3:
                print(text[:500])
                break
        except: continue
except Exception as e:
    pass
PYEOF
)
            [[ -z "$CLIP_TEXT" ]] && continue
            local FNAME; FNAME=$(basename "$CFILE")
            local FMTIME; FMTIME=$(stat -c "%y" "$CFILE" 2>/dev/null | cut -d'.' -f1 || echo "")
            local IS_SUSP="false"
            echo "$CLIP_TEXT" | grep -qiE '([0-9]{1,3}\.){3}[0-9]{1,3}|https?://|[A-Za-z0-9+/]{40,}={0,2}|password|token|secret' \
                && IS_SUSP="true"
            [[ "$IS_SUSP" == "true" ]] && printf "  ${RED}[!] [%-15s]${RESET}  %s\n" "$USERNAME" "${CLIP_TEXT:0:100}"
            ALL_CLIPS+=("${USERNAME}§${FMTIME}§${FNAME}§$(echo "$CLIP_TEXT" | head -1)§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
        done < <(find "$CB_DIR" -maxdepth 3 -name "*.dat" -type f 2>/dev/null | sort)
    done < <(get_user_homes)

    separator
    info "Voci clipboard: ${BOLD}$TOTAL${RESET}  |  Sospette: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessuna voce clipboard trovata (feature Win10 1809+, va abilitata)."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "clipboard")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per timestamp decrescente globale
    mapfile -t ALL_CLIPS < <(printf '%s\n' "${ALL_CLIPS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+["",""])[1],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_CLIPS[@]}"; do
        IFS='§' read -r UNAME FMTIME FNAME CLIP_FIRST IS_SUSP <<< "$E"
        local RS="" TC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && TC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>${UNAME}</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${FMTIME}</td>
          <td class='${TC}' style='word-break:break-all;font-size:.72rem'>$(html_esc "${CLIP_FIRST:0:300}")</td>
        </tr>"
    done
    {
        html_header "Clipboard History"
        html_page_header "CB" "Clipboard History — <span>Cronologia Appunti</span>" \
            "AppData\\Local\\Microsoft\\Windows\\Clipboard" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Voci</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Clipboard History — Utente · Timestamp · Contenuto (troncato a 300 char)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th style='width:14%'>Timestamp</th><th>Contenuto</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 32 — Office MRU & Recent Docs
# ================================================================
module_office_mru() {
    section_header "Office MRU — File Recenti Office" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. Installalo con: ${PY3} -m pip install regipy"; return 1
    fi


    declare -a ALL_MRU=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(ci_find_file "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        info "Parsing Office MRU: $USERNAME"

        local MRU_DATA
        MRU_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    APPS = ['Word', 'Excel', 'PowerPoint', 'Access', 'Publisher', 'Visio', 'Project', 'OneNote']
    # Office versioni comuni: 14.0 (2010), 15.0 (2013), 16.0 (2016/2019/365)
    VERSIONS = ['14.0', '15.0', '16.0']
    for ver in VERSIONS:
        for app in APPS:
            for key_path in [
                f'Software\\Microsoft\\Office\\{ver}\\{app}\\User MRU',
                f'Software\\Microsoft\\Office\\{ver}\\{app}\\File MRU',
            ]:
                try:
                    root = hive.get_key(key_path)
                    # Cerca sub-chiavi AD_* (per User MRU) o valori diretti
                    try:
                        for sk in root.iter_subkeys():
                            if not sk.name.startswith('AD_'): continue
                            for mru_key_path in [
                                f'{key_path}\\{sk.name}\\File MRU',
                                f'{key_path}\\{sk.name}',
                            ]:
                                try:
                                    mru_key = hive.get_key(mru_key_path)
                                    for v in mru_key.get_values():
                                        if v.name.lower() in ('mruliste','mrulist','max display'): continue
                                        val = str(v.value)
                                        if '\\' in val or '/' in val:
                                            print(f"{user}\t{app}\t{ver}\t{val[:300]}")
                                except: pass
                    except:
                        for v in root.get_values():
                            if v.name.lower() in ('mrulist','mrulistex','max display'): continue
                            val = str(v.value)
                            if '\\' in val or '/' in val:
                                print(f"{user}\t{app}\t{ver}\t{val[:300]}")
                except: continue
except Exception as e:
    pass
PYEOF
)
        while IFS=$'\t' read -r U APP VER FPATH; do
            [[ -z "$FPATH" ]] && continue
            local IS_SUSP="false"
            echo "$FPATH" | grep -qi "temp\|appdata\|public\|programdata\|\\\\temp\\\\" && IS_SUSP="true"
            printf "  ${DIM}[%-15s]${RESET}  ${GREEN}%-12s${RESET}  [%s]  %s\n" "$U" "$APP" "$VER" "${FPATH:0:80}"
            ALL_MRU+=("${U}§${APP}§${VER}§${FPATH}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
        done <<< "$MRU_DATA"
    done < <(get_user_homes)

    separator
    info "File Office recenti: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun file Office recente trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "office_mru")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_MRU[@]}"; do
        IFS='§' read -r U APP VER FPATH IS_SUSP <<< "$E"
        local RS="" NC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && NC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>${U}</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>${APP}</td>
          <td class='mono dim' style='white-space:nowrap'>${VER}</td>
          <td class='${NC}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$FPATH")</td>
        </tr>"
    done
    {
        html_header "Office MRU"
        html_page_header "OF" "Office MRU — <span>File Recenti</span>" \
            "NTUSER.DAT → Software\\Microsoft\\Office" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File recenti</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>File aperti con Word · Excel · PowerPoint · Access · OneNote</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Utente</th><th style='width:12%'>App</th><th style='width:7%'>Versione</th><th>Path file</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 33 — Defender Quarantine
# ================================================================
module_defender_quarantine() {
    section_header "Defender Quarantine — File in Quarantena" "$RED"
    check_win_root || return 1


    local QENTRIES_DIR
    QENTRIES_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Windows Defender/Quarantine/Entries")
    local QRESOURCE_DIR
    QRESOURCE_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Windows Defender/Quarantine/ResourceData")

    if [[ -z "$QENTRIES_DIR" || ! -d "$QENTRIES_DIR" ]]; then
        warn "Directory quarantena Defender non trovata"
        info "Percorso atteso: ProgramData/Microsoft/Windows Defender/Quarantine/"
        return 0
    fi

    info "Quarantine Entries: $QENTRIES_DIR"
    [[ -n "$QRESOURCE_DIR" ]] && info "ResourceData:       $QRESOURCE_DIR"
    echo ""

    declare -a ALL_ENTRIES=()
    local TOTAL=0

    while IFS= read -r EFILE; do
        [[ -f "$EFILE" ]] || continue
        local FNAME; FNAME=$(basename "$EFILE")
        local FSIZE; FSIZE=$(stat -c "%s" "$EFILE" 2>/dev/null || echo "0")
        local FMTIME; FMTIME=$(stat -c "%y" "$EFILE" 2>/dev/null | cut -d'.' -f1 || echo "")

        local ENTRY_INFO
        ENTRY_INFO=$("$PY3" - "$EFILE" << 'PYEOF' 2>/dev/null || true
import sys, struct, binascii, re

path = sys.argv[1]
try:
    with open(path,'rb') as f:
        data = f.read()
    # Cerca stringhe UTF-16LE leggibili: path file, nome minaccia, SHA256-like
    strings_utf16 = data.decode('utf-16-le', errors='replace').replace('\x00','')
    strings_ascii = data.decode('latin-1', errors='replace')

    # Cerca path Windows
    paths = re.findall(r'[A-Za-z]:\\[^\x00\n\r"<>|]{5,260}', strings_utf16)
    # Cerca nome minaccia (pattern Defender: Trojan:Win32/...)
    threats = re.findall(r'(?:Trojan|Ransom|Worm|Backdoor|Exploit|HackTool|Riskware|Generic|Behavior)[:\.\w\/\\]+', strings_utf16)
    # Cerca SHA256 hex
    hashes = re.findall(r'[0-9a-fA-F]{64}', strings_ascii)

    orig_path = paths[0] if paths else ''
    threat    = threats[0] if threats else ''
    sha256    = hashes[0] if hashes else ''
    print(f"{orig_path}\t{threat}\t{sha256}")
except Exception as e:
    print(f"\t\t")
PYEOF
)
        local ORIG_PATH THREAT HASH
        IFS=$'\t' read -r ORIG_PATH THREAT HASH <<< "$ENTRY_INFO"

        printf "  ${RED}%-45s${RESET}  ${YELLOW}%s${RESET}\n" "${THREAT:-sconosciuta}" "${ORIG_PATH:0:60}"
        [[ -n "$HASH" ]] && printf "  ${DIM}SHA256: %s${RESET}\n" "$HASH"
        ALL_ENTRIES+=("${FNAME}§${ORIG_PATH}§${THREAT}§${HASH}§${FSIZE}§${FMTIME}")
        TOTAL=$((TOTAL + 1))
    done < <(find "$QENTRIES_DIR" -maxdepth 2 -type f 2>/dev/null | sort)

    separator
    info "Voci in quarantena trovate: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "Nessuna voce trovata nella quarantena."; return 0; }

    if [[ -n "$QRESOURCE_DIR" ]]; then
        info "${BOLD}Payload cifrati disponibili in:${RESET} $QRESOURCE_DIR"
        info "Per ripristino (su VM isolata): MpCmdRun.exe -Restore -Name <ThreatName>"
    fi

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "defender_quarantine")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    # Ordina per timestamp decrescente (campo 6: FMTIME)
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+6*[""])[5],reverse=True); print("\n".join(L))')
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='§' read -r FNAME ORIG_PATH THREAT HASH FSIZE FMTIME <<< "$E"
        ROWS+="<tr style='background:rgba(255,123,114,.05);border-left:3px solid var(--accent2)'>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${FMTIME}</td>
          <td class='mono bad' style='word-break:break-all;font-size:.72rem'>$(html_esc "${THREAT:-Sconosciuta}")</td>
          <td class='mono' style='word-break:break-all;font-size:.72rem'>$(html_esc "$ORIG_PATH")</td>
          <td class='mono dim' style='word-break:break-all;font-size:.65rem'>$(html_esc "${HASH:0:64}")</td>
          <td class='mono mid' style='white-space:nowrap'>${FSIZE}B</td>
        </tr>"
    done
    {
        html_header "Defender Quarantine"
        html_page_header "DQ" "Defender — <span>Quarantine</span>" \
            "ProgramData\\Microsoft\\Windows Defender\\Quarantine" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>In quarantena</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>File in quarantena — recuperabili offline da ResourceData/ (cifrati XOR)</div>
        <div class='card'><table>
          <thead><tr><th style='width:14%'>Rilevato</th><th style='width:22%'>Minaccia</th><th>Path originale</th><th style='width:18%'>SHA256</th><th style='width:6%'>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Analisi payload (ambiente isolato)</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            I file in ResourceData/ sono cifrati con RC4 (key: hardcoded Defender) — tool: defender-un-quarantine.py<br>
            Ripristino: MpCmdRun.exe -Restore -FilePath &lt;path&gt;
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 34 — PowerShell Script Block Logging (EID 4104)
# ================================================================
module_ps_scriptblock() {
    section_header "PowerShell Script Block Logging — EID 4104" "$MAGENTA"
    check_win_root || return 1

    local EVTX_DIR
    EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    if [[ -z "$EVTX_DIR" ]]; then
        warn "Directory EVTX non trovata"
        return 1
    fi

    local PS_EVTX
    PS_EVTX=$(find "$EVTX_DIR" -maxdepth 1 -iname "Microsoft-Windows-PowerShell%4Operational.evtx" -type f 2>/dev/null | head -1)
    if [[ -z "$PS_EVTX" || ! -f "$PS_EVTX" ]]; then
        warn "Log PowerShell Operational non trovato"
        info "Atteso: Windows/System32/winevt/Logs/Microsoft-Windows-PowerShell%4Operational.evtx"
        return 0
    fi

    info "Parsing: $PS_EVTX"

    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        warn "python-evtx non trovato. Installalo con: ${PY3} -m pip install python-evtx"
        return 1
    fi

    local BLOCK_DATA
    BLOCK_DATA=$("$PY3" - "$PS_EVTX" << 'PYEOF' 2>/dev/null || true
import sys, xml.etree.ElementTree as ET
try:
    from Evtx.Evtx import Evtx
    SENS = ['invoke-expression','iex','downloadstring','encodedcommand','-enc','bypass',
            'frombase64string','webclient','net user','net localgroup','mimikatz',
            'password','token','secret','hidden','noprofile','windowstyle']

    # Accumula frammenti per ScriptBlockId (EID 4104 suddivide script grandi
    # in N eventi consecutivi con MessageNumber/MessageTotal)
    blocks = {}  # block_id -> {'total':N, 'parts':{num:text}, 'ts':'', 'path':''}

    with Evtx(sys.argv[1]) as log:
        for record in log.records():
            try:
                xml_str = record.xml()
                root = ET.fromstring(xml_str)
                ns = {'e': 'http://schemas.microsoft.com/win/2004/08/events/event'}
                eid_el = root.find('.//e:EventID', ns)
                if eid_el is None or eid_el.text != '4104': continue

                vals = {}
                for ed in root.findall('.//e:EventData/e:Data', ns):
                    vals[ed.get('Name', '')] = ed.text or ''

                sbt = vals.get('ScriptBlockText', '')
                if not sbt: continue

                block_id = vals.get('ScriptBlockId', '') or id(vals)
                msg_num  = int(vals.get('MessageNumber', '1') or '1')
                msg_tot  = int(vals.get('MessageTotal',  '1') or '1')
                path     = vals.get('Path', '')

                ts_el = root.find('.//e:TimeCreated', ns)
                ts = ts_el.get('SystemTime', '') if ts_el is not None else ''

                if block_id not in blocks:
                    blocks[block_id] = {'total': msg_tot, 'parts': {}, 'ts': ts, 'path': path}
                blocks[block_id]['parts'][msg_num] = sbt
                if ts and not blocks[block_id]['ts']:
                    blocks[block_id]['ts'] = ts
            except: continue

    # Riassembla ogni script e stampa come singola riga:
    # TYPE \t TIMESTAMP \t PATH \t SCRIPT_FLAT
    # I newline nello script sono codificati come \x1f (unit separator)
    # per non rompere il parsing bash riga-per-riga.
    for info in blocks.values():
        parts = info['parts']
        total = info['total']
        full  = ''.join(parts.get(i, '') for i in range(1, total + 1))
        if not full:
            full = ''.join(parts[k] for k in sorted(parts))
        ts    = info['ts']
        path  = info['path']
        is_sens = any(k in full.lower() for k in SENS)
        flat  = full.replace('\r\n', '\x1f').replace('\n', '\x1f').replace('\r', '\x1f')
        sys.stdout.buffer.write(
            f"{'SENS' if is_sens else 'NORM'}\x1e{ts}\x1e{path}\x1e{flat[:1000000]}\n".encode('utf-8', errors='replace')
        )

except Exception as e:
    sys.stderr.write(f"ERROR\x1e\x1e\x1e{e}\n")
PYEOF
)

    local TOTAL=0 SUSP=0
    declare -a BLOCKS=()
    while IFS=$'\x1e' read -r TYPE TS FPATH SCRIPT; do
        [[ "$TYPE" == "ERROR" ]] && { err "Errore parser: $SCRIPT"; break; }
        # Salta righe orfane (non devono esistere dopo il fix, ma per sicurezza)
        [[ -z "$TS" && -z "$FPATH" ]] && continue
        TOTAL=$((TOTAL+1))
        [[ "$TYPE" == "SENS" ]] && SUSP=$((SUSP+1))
        # Ripristina i newline (\x1f → \n) per il display terminale
        local DISP; DISP="${SCRIPT:0:120}"
        DISP="${DISP//$'\x1f'/$'\n'      }"
        if [[ "$TYPE" == "SENS" ]]; then
            printf "  ${RED}[!]${RESET} ${DIM}%s${RESET}  ${DIM}%s${RESET}\n  ${RED}%s${RESET}\n\n" "$TS" "$FPATH" "$DISP"
        else
            printf "  ${DIM}[i] %s  %s${RESET}\n" "$TS" "$DISP"
        fi
        BLOCKS+=("${TYPE}§${TS}§${FPATH}§${SCRIPT}")
    done <<< "$BLOCK_DATA"

    separator
    info "Script block unici: ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && { warn "Nessun EID 4104 trovato (Script Block Logging potrebbe non essere abilitato)."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ps_scriptblock")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local HIVE_HASH; HIVE_HASH=$(sha256_file "$PS_EVTX")
    # Ordina per timestamp decrescente (campo 2: TS)
    mapfile -t BLOCKS < <(printf '%s\n' "${BLOCKS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+2*[""])[1],reverse=True); print("\n".join(L))')
    local ROWS=""
    for E in "${BLOCKS[@]}"; do
        IFS='§' read -r TYPE TS FPATH SCRIPT <<< "$E"
        local RS="" TC="mono"
        [[ "$TYPE" == "SENS" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && TC="mono bad"
        # Ripristina i newline (\x1f → LF) e applica HTML escape
        local SCRIPT_ESC
        SCRIPT_ESC=$(printf '%s' "${SCRIPT:0:1000000}" \
            | sed $'s/\x1f/\\n/g' \
            | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g')
        local PATH_ESC; PATH_ESC=$(html_esc "$FPATH")
        ROWS+="<tr ${RS}>
          <td class='mono ok' style='white-space:nowrap;font-size:.68rem'>${TS}</td>
          <td class='mono dim' style='white-space:nowrap;font-size:.65rem;max-width:220px;overflow:hidden;text-overflow:ellipsis' title='${PATH_ESC}'>${PATH_ESC}</td>
          <td class='${TC}' style='word-break:break-all;font-size:.72rem;white-space:pre-wrap'>${SCRIPT_ESC}</td>
        </tr>"
    done
    {
        html_header "PS ScriptBlock Log"
        html_page_header "PS4" "PS Script Block Logging — <span>EID 4104</span>" \
            "Microsoft-Windows-PowerShell%4Operational.evtx" "$SCAN" "$WIN_ROOT" "$HIVE_HASH"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Script Block unici</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP}</div></div>
        </div><main>
        <div class='stitle'>Event ID 4104 — Frammenti riassemblati per ScriptBlockId — ScriptBlockText (max 1000000 char)</div>
        <div class='card'><table>
          <thead><tr><th style='width:16%'>Timestamp</th><th style='width:18%'>Path script</th><th>ScriptBlockText</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 35 — JumpLists (AutomaticDestinations + CustomDestinations)
# ================================================================
module_jumplists() {
    section_header "JumpLists — AutomaticDestinations & CustomDestinations" "$GREEN"
    check_win_root || return 1

    declare -a ALL_ENTRIES=()
    local TOTAL=0 SUSP=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local AUTO_DIR
        AUTO_DIR=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Microsoft/Windows/Recent/AutomaticDestinations")
        local CUST_DIR
        CUST_DIR=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Microsoft/Windows/Recent/CustomDestinations")

        for JL_DIR in "$AUTO_DIR" "$CUST_DIR"; do
            [[ -z "$JL_DIR" || ! -d "$JL_DIR" ]] && continue
            local JL_TYPE="Auto"
            [[ "$JL_DIR" == *"CustomDestinations"* ]] && JL_TYPE="Custom"

            while IFS= read -r JFILE; do
                [[ -f "$JFILE" ]] || continue
                local FNAME; FNAME=$(basename "$JFILE")
                local FMTIME; FMTIME=$(stat -c "%y" "$JFILE" 2>/dev/null | cut -d'.' -f1 || echo "")
                local STRINGS_OUT
                STRINGS_OUT=$(strings -n 8 "$JFILE" 2>/dev/null | grep -E '\\\\|https?://|\.exe|\.dll|\.ps1|\.bat|\.vbs|\.lnk' | head -20 || true)
                [[ -z "$STRINGS_OUT" ]] && continue
                local FIRST_PATH; FIRST_PATH=$(echo "$STRINGS_OUT" | head -1)
                local IS_SUSP="false"
                echo "$STRINGS_OUT" | grep -qiE 'temp|appdata\\local\\temp|\.ps1|\.bat|\.vbs|programdata|public\\' && IS_SUSP="true"
                if [[ ${#IOC_LIST[@]} -gt 0 ]]; then
                    check_ioc "$STRINGS_OUT" && IS_SUSP="true"
                fi
                [[ "$IS_SUSP" == "true" ]] && printf "  ${RED}[!]${RESET} ${DIM}[%-10s]${RESET} [${JL_TYPE}] %s\n" "$USERNAME" "$FIRST_PATH"
                ALL_ENTRIES+=("${USERNAME}§${JL_TYPE}§${FMTIME}§${FNAME}§$(echo "$STRINGS_OUT" | head -3 | tr '\n' '|')§${IS_SUSP}")
                TOTAL=$((TOTAL+1))
                [[ "$IS_SUSP" == "true" ]] && SUSP=$((SUSP+1))
            done < <(find "$JL_DIR" -maxdepth 1 -type f 2>/dev/null | sort)
        done
    done < <(get_user_homes)

    separator
    info "JumpList entries: ${BOLD}$TOTAL${RESET}  |  Sospette: ${RED}${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && { warn "Nessuna JumpList trovata."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "jumplists")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    # Ordina per timestamp decrescente (campo 3: FMTIME)
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+3*[""])[2],reverse=True); print("\n".join(L))')
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='§' read -r UNAME JT FMTIME FNAME PATHS IS_SUSP <<< "$E"
        local RS="" TC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && TC="mono bad"
        local PATHS_DISP; PATHS_DISP=$(html_esc "${PATHS//|/<br>}")
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>${UNAME}</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>${JT}</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.68rem'>${FMTIME}</td>
          <td class='mono dim' style='font-size:.68rem'>${FNAME}</td>
          <td class='${TC}' style='word-break:break-all;font-size:.72rem'>${PATHS_DISP}</td>
        </tr>"
    done
    {
        html_header "JumpLists"
        html_page_header "JL" "JumpLists — <span>AutomaticDestinations · CustomDestinations</span>" \
            "AppData\\Roaming\\Microsoft\\Windows\\Recent\\*Destinations" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Entries</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSP}</div></div>
        </div><main>
        <div class='stitle'>JumpLists — Utente · Tipo · Timestamp · File · Percorsi estratti (strings)</div>
        <div class='card'><table>
          <thead><tr><th style='width:10%'>Utente</th><th style='width:7%'>Tipo</th><th style='width:14%'>Modificato</th><th style='width:14%'>File JL</th><th>Percorsi estratti</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 36 — Network Artifacts (DNS Cache, Interfacce, NetworkList)
# ================================================================
module_network_artifacts() {
    section_header "Network Artifacts — DNS · Interfacce · NetworkList" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. Installalo con: ${PY3} -m pip install regipy"; return 1
    fi

    local SYS_HIVE; SYS_HIVE=$(get_hive "SOFTWARE")
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "Hive SOFTWARE non trovato"; return 1
    fi

    info "Parsing NetworkList da hive SOFTWARE: $SYS_HIVE"

    local NET_DATA
    NET_DATA=$("$PY3" - "$SYS_HIVE" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    try:
        nl = hive.get_key('Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles')
        for sk in nl.iter_subkeys():
            vals = {v.name: str(v.value) for v in sk.get_values()}
            name       = vals.get('ProfileName', '')
            desc       = vals.get('Description', '')
            category   = vals.get('Category', '')
            date_first = vals.get('DateFirstConnected', '')
            date_last  = vals.get('DateLastConnected', '')
            cat_str = {'0':'Public','1':'Private','2':'Domain'}.get(str(category), str(category))
            print(f"PROFILE\t{name}\t{desc}\t{cat_str}\t{date_first}\t{date_last}")
    except: pass
    try:
        sigs = hive.get_key('Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Signatures\\Unmanaged')
        for sk in sigs.iter_subkeys():
            vals = {v.name: str(v.value) for v in sk.get_values()}
            desc     = vals.get('Description', '')
            dns      = vals.get('DnsSuffix', '')
            mac      = vals.get('DefaultGatewayMac', '')
            ssid     = vals.get('FirstNetwork', '')
            print(f"SIGNATURE\t{desc}\t{dns}\t{mac}\t{ssid}\t")
    except: pass
except Exception as e:
    print(f"ERROR\t{e}", file=sys.stderr)
PYEOF
)

    local SYS_HIVE2; SYS_HIVE2=$(get_hive "SYSTEM")
    local IFACE_DATA=""
    if [[ -n "$SYS_HIVE2" && -f "$SYS_HIVE2" ]]; then
        info "Parsing interfacce TCP/IP da SYSTEM hive: $SYS_HIVE2"
        IFACE_DATA=$("$PY3" - "$SYS_HIVE2" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for cs in ['ControlSet001','ControlSet002','CurrentControlSet']:
        try:
            ifaces = hive.get_key(f'{cs}\\Services\\Tcpip\\Parameters\\Interfaces')
            for sk in ifaces.iter_subkeys():
                vals = {v.name: str(v.value) for v in sk.get_values()}
                ip    = vals.get('IPAddress','') or vals.get('DhcpIPAddress','')
                gw    = vals.get('DefaultGateway','') or vals.get('DhcpDefaultGateway','')
                dns   = vals.get('NameServer','') or vals.get('DhcpNameServer','')
                dhcp  = vals.get('EnableDHCP','')
                if ip and ip not in ('0.0.0.0',''):
                    print(f"IFACE\t{cs}\t{sk.name}\t{ip}\t{gw}\t{dns}\t{dhcp}")
            break
        except: continue
except Exception as e:
    pass
PYEOF
)
    fi

    declare -a PROFILES=() SIGNATURES=() IFACES=()
    local TOTAL_P=0 TOTAL_S=0 TOTAL_I=0

    while IFS=$'\t' read -r TYPE F1 F2 F3 F4 F5; do
        case "$TYPE" in
            PROFILE)
                printf "  ${GREEN}[P]${RESET}  ${BOLD}%-30s${RESET}  ${DIM}%s${RESET}  [%s]\n" "$F1" "$F3" "$F4"
                PROFILES+=("${F1}§${F2}§${F3}§${F4}§${F5}")
                TOTAL_P=$((TOTAL_P+1))
                ;;
            SIGNATURE)
                printf "  ${BLUE}[S]${RESET}  %-30s  DNS:${DIM}%s${RESET}  MAC:${DIM}%s${RESET}\n" "$F1" "$F2" "$F3"
                SIGNATURES+=("${F1}§${F2}§${F3}§${F4}")
                TOTAL_S=$((TOTAL_S+1))
                ;;
        esac
    done <<< "$NET_DATA"

    while IFS=$'\t' read -r TYPE CS IFACE IP GW DNS DHCP; do
        [[ "$TYPE" != "IFACE" ]] && continue
        printf "  ${CYAN}[I]${RESET}  ${DIM}%s${RESET}  IP:${BOLD}%s${RESET}  GW:%s  DNS:%s\n" "$IFACE" "$IP" "$GW" "$DNS"
        IFACES+=("${CS}§${IFACE}§${IP}§${GW}§${DNS}§${DHCP}")
        TOTAL_I=$((TOTAL_I+1))
    done <<< "$IFACE_DATA"

    separator
    info "Profili rete: ${BOLD}$TOTAL_P${RESET}  |  Firme: ${BOLD}$TOTAL_S${RESET}  |  Interfacce TCP/IP: ${BOLD}$TOTAL_I"
    [[ $((TOTAL_P + TOTAL_S + TOTAL_I)) -eq 0 ]] && { warn "Nessun dato di rete trovato."; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "network_artifacts")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local HIVE_HASH; HIVE_HASH=$(sha256_file "$SYS_HIVE")

    # Ordina profili rete per ultima connessione decrescente (campo 5: DLAST)
    mapfile -t PROFILES < <(printf '%s\n' "${PROFILES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+5*[""])[4],reverse=True); print("\n".join(L))')
    local ROWS_P="" ROWS_S="" ROWS_I=""
    for E in "${PROFILES[@]}"; do
        IFS='§' read -r NAME DESC CAT DFIRST DLAST <<< "$E"
        local CAT_COLOR="ok"
        [[ "$CAT" == "Public" ]] && CAT_COLOR="warn"
        ROWS_P+="<tr><td class='mono'>${NAME}</td><td class='mono dim'>${DESC}</td>
          <td class='mono ${CAT_COLOR}'>${CAT}</td>
          <td class='mono dim' style='font-size:.68rem'>${DFIRST}</td>
          <td class='mono ok'  style='font-size:.68rem'>${DLAST}</td></tr>"
    done
    for E in "${SIGNATURES[@]}"; do
        IFS='§' read -r NAME DNS_S MAC SSID <<< "$E"
        ROWS_S+="<tr><td class='mono'>${NAME}</td><td class='mono dim'>${DNS_S}</td>
          <td class='mono warn'>${MAC}</td><td class='mono'>${SSID}</td></tr>"
    done
    for E in "${IFACES[@]}"; do
        IFS='§' read -r CS IFACE IP GW DNS_I DHCP <<< "$E"
        ROWS_I+="<tr><td class='mono dim'>${CS}</td><td class='mono dim' style='font-size:.65rem'>${IFACE}</td>
          <td class='mono ok'>${IP}</td><td class='mono'>${GW}</td>
          <td class='mono dim'>${DNS_I}</td><td class='mono'>${DHCP}</td></tr>"
    done
    {
        html_header "Network Artifacts"
        html_page_header "NET" "Network Artifacts — <span>Profili · Firme · Interfacce</span>" \
            "SOFTWARE · SYSTEM hive" "$SCAN" "$WIN_ROOT" "$HIVE_HASH"
        echo "<div class='statsbar'>
          <div class='stat info'><div class='label'>Profili rete</div><div class='value'>${TOTAL_P}</div></div>
          <div class='stat'><div class='label'>Firme</div><div class='value'>${TOTAL_S}</div></div>
          <div class='stat ok'><div class='label'>Interfacce TCP/IP</div><div class='value'>${TOTAL_I}</div></div>
        </div><main>"
        if [[ -n "$ROWS_P" ]]; then
            echo "<div class='stitle'>Profili di rete noti (NetworkList\\Profiles)</div>
            <div class='card'><table>
              <thead><tr><th>Nome</th><th>Descrizione</th><th>Categoria</th><th>Prima connessione</th><th>Ultima connessione</th></tr></thead>
              <tbody>${ROWS_P}</tbody></table></div>"
        fi
        if [[ -n "$ROWS_S" ]]; then
            echo "<div class='stitle' style='margin-top:2rem'>Firme di rete (Signatures\\Unmanaged)</div>
            <div class='card'><table>
              <thead><tr><th>Descrizione</th><th>DNS Suffix</th><th>MAC Gateway</th><th>SSID</th></tr></thead>
              <tbody>${ROWS_S}</tbody></table></div>"
        fi
        if [[ -n "$ROWS_I" ]]; then
            echo "<div class='stitle' style='margin-top:2rem'>Interfacce TCP/IP (SYSTEM\\Tcpip\\Parameters)</div>
            <div class='card'><table>
              <thead><tr><th>ControlSet</th><th>GUID interfaccia</th><th>IP</th><th>Gateway</th><th>DNS</th><th>DHCP</th></tr></thead>
              <tbody>${ROWS_I}</tbody></table></div>"
        fi
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 37 — Master Timeline (aggregazione cross-moduli)
# ================================================================
module_master_timeline() {
    section_header "Master Timeline — Aggregazione Cross-Moduli" "$YELLOW"

    if [[ ${#GENERATED_REPORTS[@]} -eq 0 ]]; then
        warn "Nessun report generato in questa sessione."
        info "Esegui prima i moduli che ti interessano, poi richiama la Master Timeline."
        return 0
    fi

    info "Report disponibili in sessione: ${#GENERATED_REPORTS[@]}"
    echo ""

    # Python estrae le righe con timestamp da tutti i report HTML.
    # Usa re.DOTALL per matchare <tr>...</tr> multiriga (grep -E non può farlo).
    # Output: YYYY-MM-DD HH:MM:SS \t modulo \t descrizione (una riga per evento)
    local TL_RAW
    TL_RAW=$("$PY3" - "${GENERATED_REPORTS[@]}" << 'PYEOF' 2>/dev/null || true
import sys, re, html as html_mod

TS_PAT  = re.compile(r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}')
TD_PAT  = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL | re.IGNORECASE)
TAG_PAT = re.compile(r'<[^>]+>')

def strip_tags(s):
    return TAG_PAT.sub('', html_mod.unescape(s)).strip()

seen = set()
for rpath in sys.argv[1:]:
    mod = rpath.rstrip('/').split('/')[-2]
    mod = re.sub(r'_\d{8}_\d{6}$', '', mod)
    try:
        with open(rpath, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        for tr_m in re.finditer(r'<tr[^>]*>(.*?)</tr>', content, re.DOTALL | re.IGNORECASE):
            cells = [strip_tags(m.group(1)) for m in TD_PAT.finditer(tr_m.group(1))]
            if not cells:
                continue
            ts = ''
            for cell in cells:
                m = TS_PAT.search(cell)
                if m:
                    ts = m.group(0).replace('T', ' ')[:19]
                    break
            if not ts:
                continue
            # Descrizione: celle che non contengono solo il timestamp
            desc_parts = [c[:80] for c in cells if c and not TS_PAT.fullmatch(c.strip())]
            desc = ' | '.join(desc_parts)[:200]
            desc = ' '.join(desc.split())   # normalizza spazi
            key = (ts, mod, desc[:40])
            if key in seen:
                continue
            seen.add(key)
            sys.stdout.buffer.write(f"{ts}\t{mod}\t{desc}\n".encode('utf-8', errors='replace'))
    except Exception:
        pass
PYEOF
)

    local TOTAL_EVENTS; TOTAL_EVENTS=$(printf '%s' "$TL_RAW" | grep -c $'\t' 2>/dev/null || echo 0)

    # Ordina per timestamp decrescente e tronca a 2000 (mostreremo i primi 1000 nel report)
    local TL_SORTED
    TL_SORTED=$(printf '%s\n' "$TL_RAW" | sort -t$'\t' -k1 -r | head -2000)

    separator
    info "Eventi con timestamp estratti: ${BOLD}$TOTAL_EVENTS${RESET} (mostrati max 1000 nel report)"
    [[ $TOTAL_EVENTS -eq 0 ]] && { warn "Nessun timestamp trovato nei report. Verifica che i moduli abbiano generato dati."; return 0; }
    ask_yn "Generare report HTML Master Timeline?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "master_timeline")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local NR="${#GENERATED_REPORTS[@]}"

    {
        html_header "Master Timeline"
        html_page_header "TL" "Master Timeline — <span>Aggregazione Cross-Moduli</span>" \
            "Aggregazione di ${NR} report" "$SCAN" "$WIN_ROOT"
        cat << STATSEOF
<div class='statsbar'>
  <div class='stat'><div class='label'>Report aggregati</div><div class='value'>${NR}</div></div>
  <div class='stat info'><div class='label'>Eventi con timestamp</div><div class='value'>${TOTAL_EVENTS}</div></div>
</div>
<main>
<div class='stitle'>Timeline cronologica (decrescente) — max 1000 eventi · ordinati per data</div>
<div class='card'>
<style>
  .tl-filter { padding:.8rem 1rem; background:var(--bg3); border-bottom:1px solid var(--border);
    display:flex; gap:1rem; align-items:center; flex-wrap:wrap; }
  .tl-filter input { background:var(--bg2); border:1px solid var(--border); color:var(--text);
    padding:.3rem .7rem; font-family:var(--mono); font-size:.75rem; border-radius:3px; width:250px; }
  .tl-filter input:focus { outline:none; border-color:var(--accent); }
  .tl-filter label { font-family:var(--mono); font-size:.65rem; color:var(--text-dim); }
</style>
<div class='tl-filter'>
  <label>Filtra:</label>
  <input type='text' id='tl-filter-ts'   placeholder='Timestamp (es. 2024-03)' oninput='filterTimeline()'>
  <input type='text' id='tl-filter-mod'  placeholder='Modulo (es. evtx)'       oninput='filterTimeline()'>
  <input type='text' id='tl-filter-desc' placeholder='Testo'                   oninput='filterTimeline()'>
</div>
<table id='tl-table'>
  <thead><tr>
    <th style='width:16%'>Timestamp</th>
    <th style='width:14%'>Modulo</th>
    <th>Descrizione</th>
  </tr></thead>
  <tbody id='tl-body'>
STATSEOF

        # Scrivi le righe direttamente nel file — evita la variabile ROWS gonfia
        local _COUNT=0
        while IFS=$'\t' read -r _TS _MOD _DESC; do
            [[ -z "$_TS" ]] && continue
            (( _COUNT++ > 1000 )) && break
            local _BC="var(--accent)"
            case "${_MOD,,}" in
                *evtx*|*service*|*ifeo*|*wmi*|*sam*|*ntds*) _BC="var(--accent2)" ;;
                *ps_hist*|*script*|*prefetch*|*amcache*)     _BC="var(--accent4)" ;;
                *browser*|*lnk*|*shell*|*userassist*)        _BC="var(--accent3)" ;;
            esac
            local _TE; _TE=$(html_esc "$_TS")
            local _ME; _ME=$(html_esc "$_MOD")
            local _DE; _DE=$(html_esc "${_DESC:0:150}")
            printf '<tr><td class="mono ok" style="white-space:nowrap;font-size:.72rem">%s</td>' "$_TE"
            printf '<td><span style="background:rgba(88,166,255,.1);border:1px solid rgba(88,166,255,.2);'
            printf 'color:%s;font-family:var(--mono);font-size:.65rem;padding:.15rem .5rem;border-radius:2px">%s</span></td>' "$_BC" "$_ME"
            printf '<td class="mono" style="font-size:.72rem;color:var(--text)">%s</td></tr>\n' "$_DE"
        done <<< "$TL_SORTED"

        cat << 'SCRIPTEOF'
  </tbody>
</table>
</div>
<script>
function filterTimeline() {
  var ft  = document.getElementById('tl-filter-ts').value.toLowerCase();
  var fm  = document.getElementById('tl-filter-mod').value.toLowerCase();
  var fd  = document.getElementById('tl-filter-desc').value.toLowerCase();
  document.querySelectorAll('#tl-body tr').forEach(function(row) {
    var c = row.querySelectorAll('td');
    if (!c.length) return;
    var ok = (!ft || c[0].textContent.toLowerCase().includes(ft)) &&
             (!fm || c[1].textContent.toLowerCase().includes(fm)) &&
             (!fd || c[2].textContent.toLowerCase().includes(fd));
    row.style.display = ok ? '' : 'none';
  });
}
</script>
</main>
SCRIPTEOF

        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "Report salvato: ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# Esegue un modulo in modalità batch e registra il risultato in SUMMARY_TABLE
run_batch_module() {
    local mod_num="$1"
    local mod_func="$2"
    local mod_name="$3"
    local total_mods="${4:-38}"

    echo -ne "  ${CYAN}[*]${RESET} [${mod_num}/${total_mods}] Esecuzione modulo $mod_num ($mod_name)... \r"
    log_msg "[BATCH] Modulo $mod_num: $mod_name"

    local before_reports=${#GENERATED_REPORTS[@]}
    $mod_func >/dev/null 2>&1
    local after_reports=${#GENERATED_REPORTS[@]}

    if (( after_reports > before_reports )); then
        local rep_path="${GENERATED_REPORTS[-1]}"
        echo -e "  ${GREEN}[✓]${RESET} [${mod_num}/${total_mods}] $mod_name — report: ${DIM}${rep_path}${RESET}"
        SUMMARY_TABLE+=("$mod_num|$mod_name|SI|$rep_path")
    else
        echo -e "  ${DIM}[i] [${mod_num}/${total_mods}] $mod_name — nessun risultato${RESET}                     "
        SUMMARY_TABLE+=("$mod_num|$mod_name|NO|-")
    fi
}

# ================================================================
#  MODULO 38 — PAD Offline (Active Directory Analysis)
# ================================================================
module_pad_offline() {
    section_header "PAD Offline — Active Directory Analysis (NTDS.dit)" "$CYAN"
    check_win_root || return 1

    # ── Trova NTDS.dit ────────────────────────────────────────────
    local NTDS_DIR="" NTDS_PATH="" SYSTEM_HIVE=""
    NTDS_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/NTDS")
    if [[ -n "$NTDS_DIR" ]]; then
        NTDS_PATH=$(find "$NTDS_DIR" -maxdepth 1 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    fi
    if [[ -z "$NTDS_PATH" ]]; then
        NTDS_PATH=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f \
            -not -path "*/WinSxS/*" -not -path "*/WinSXS/*" \
            2>/dev/null | head -1)
    fi
    if [[ -z "$NTDS_PATH" || ! -f "$NTDS_PATH" ]]; then
        warn "ntds.dit non trovato sotto $WIN_ROOT"
        info "Percorso atteso: Windows/NTDS/ntds.dit"
        info "Questo modulo richiede un Domain Controller (non una workstation)"
        return 0
    fi
    info "NTDS.dit: $NTDS_PATH"

    # ── Trova SYSTEM hive (opzionale) ────────────────────────────
    local CFG_DIR
    CFG_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
    if [[ -n "$CFG_DIR" ]]; then
        SYSTEM_HIVE=$(find "$CFG_DIR" -maxdepth 1 -iname "SYSTEM" -type f 2>/dev/null | head -1)
    fi
    if [[ -n "$SYSTEM_HIVE" ]]; then
        info "SYSTEM hive: $SYSTEM_HIVE"
    else
        warn "SYSTEM hive non trovato — attributi cifrati (hash) non disponibili; tutti gli altri OK"
        SYSTEM_HIVE=""
    fi

    # ── Avviso dirty database ─────────────────────────────────────
    warn "Se ntds.dit proviene da un DC attivo, alcune transazioni potrebbero essere nei log (.jrs)."
    info "Per massima integrità usa una VSS snapshot o esegui 'esentutl /r edb /l <logdir>' prima."

    # ── Verifica dipendenza impacket ──────────────────────────────
    if ! "$PY3" -c "from impacket.ese import ESENT_DB" 2>/dev/null; then
        warn "impacket non trovato. Installalo con: ${PY3} -m pip install impacket"
        return 1
    fi

    info "Avvio analisi NTDS.dit (può richiedere qualche minuto su database grandi)..."
    separator

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "pad_offline")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local NTDS_HASH; NTDS_HASH=$(sha256_file "$NTDS_PATH")

    # Copia ntds.dit in /tmp (preferita: velocizza accesso ESE e abilita patch binaria).
    # Se fallisce (file bloccato da Windows su DC live, o rete lenta) si usa il path
    # originale — il recovery in-memory del dirty state non richiede write access al file.
    local TMP_PAD_DIR; TMP_PAD_DIR=$(mktemp -d /tmp/dfir_pad_XXXXXX)
    local TMP_NTDS; TMP_NTDS="${TMP_PAD_DIR}/ntds.dit"
    local NTDS_TO_USE="$NTDS_PATH"
    info "Tentativo copia ntds.dit in /tmp (timeout 180s)..."
    if portable_timeout 180 cp "$NTDS_PATH" "$TMP_NTDS" 2>/dev/null; then
        chmod 600 "$TMP_NTDS" 2>/dev/null || true
        NTDS_TO_USE="$TMP_NTDS"
        info "Copia completata — uso copia locale."
    else
        warn "Copia non riuscita (file bloccato da Windows o rete lenta) — uso path originale."
        info "Il dirty state verrà corretto in memoria (patch in-memory, senza write access)."
    fi

    TMP_NTDS_SIZE=$(wc -c < "$TMP_NTDS" 2>/dev/null | tr -d ' ' || echo "0")
    ORIG_NTDS_SIZE=$(wc -c < "$NTDS_TO_USE" 2>/dev/null | tr -d ' ' || echo "0")
    info "Dimensione NTDS.dit prima del parsing: ${TMP_NTDS_SIZE} bytes (originale: ${ORIG_NTDS_SIZE})"
    
    # ── Script Python: analisi completa + generazione HTML ───────
    local PYERR_FILE; PYERR_FILE="${TMP_PAD_DIR}/pyerr.log"
    "$PY3" - "$NTDS_TO_USE" "${SYSTEM_HIVE}" "$REPORT_HTML" "$SCAN" "$WIN_ROOT" "$NTDS_HASH" "$(dirname "$NTDS_PATH")" "$TMP_PAD_DIR" << 'PYEOF' 2>"$PYERR_FILE"
import sys, os, struct, datetime, traceback, html as html_mod

ntds_path   = sys.argv[1]
sys_hive    = sys.argv[2] if len(sys.argv) > 2 else ''
report_path = sys.argv[3]
scan_date   = sys.argv[4]
source_root = sys.argv[5]
ntds_hash   = sys.argv[6]
ntds_log_dir= sys.argv[7] if len(sys.argv) > 7 else ''
tmp_work_dir= sys.argv[8] if len(sys.argv) > 8 else os.path.dirname(ntds_path)

# ── Copia log ESE in /tmp per permettere soft-recovery automatico ─
# I file .jrs devono essere nella stessa dir del ntds.dit aperto da impacket
# PRIMA che impacket lo apra, e il dirty state NON deve essere patchato
# altrimenti impacket non tenta il recovery.
_has_ese_logs = False
import shutil, glob as _glob
_log_search_dirs = []
if ntds_log_dir and os.path.isdir(ntds_log_dir):
    _log_search_dirs.append(ntds_log_dir)
# Cerca anche nella dir del file originale se diversa
_orig_ntds_dir = os.path.dirname(os.path.realpath(ntds_path))
if _orig_ntds_dir not in _log_search_dirs:
    _log_search_dirs.append(_orig_ntds_dir)

for _log_src_dir in _log_search_dirs:
    _log_files = (
        _glob.glob(os.path.join(_log_src_dir, 'edb*.log')) +
        _glob.glob(os.path.join(_log_src_dir, 'edb*.jrs')) +
        _glob.glob(os.path.join(_log_src_dir, 'edb.chk'))
    )
    for _lf in _log_files:
        try:
            _dst = os.path.join(tmp_work_dir, os.path.basename(_lf))
            if not os.path.exists(_dst):
                shutil.copy2(_lf, _dst)
                sys.stderr.write(f"INFO: copiato log ESE: {os.path.basename(_lf)}\n")
            _has_ese_logs = True
        except Exception as _le:
            sys.stderr.write(f"INFO: log copy skip {_lf}: {_le}\n")
    if _has_ese_logs:
        break

if _has_ese_logs:
    sys.stderr.write("INFO: log ESE trovati e copiati in tmp — impacket tenterà soft-recovery\n")
else:
    sys.stderr.write("INFO: nessun log ESE trovato — verrà usata solo patch dirty state\n")

try:
    from impacket.ese import ESENT_DB as _ESENT_DB_CLS
    ESENT_DB = _ESENT_DB_CLS   # alias usato in _open_esent_db
    _HAVE_IMPACKET = True
except ImportError:
    _HAVE_IMPACKET = False
    ESENT_DB = None

try:
    import pyesedb as _pyesedb
    _HAVE_PYESEDB = True
except ImportError:
    _HAVE_PYESEDB = False

if not _HAVE_IMPACKET and not _HAVE_PYESEDB:
    sys.stderr.write("FATAL: nessuna libreria ESE disponibile. Installa impacket o libesedb-python.\n")
    sys.exit(1)

# ── Wrapper pyesedb compatibile con l'API impacket ───────────────
# Produce righe come dict {col_name: value}, compatibile col resto del codice.
class _PyesedbCursor:
    def __init__(self, table):
        self._table = table
        self._ncols = table.get_number_of_columns()
        self._nrec  = table.get_number_of_records()
        self._idx   = 0
        # Mappa indice colonna → nome
        self._colnames = [table.get_column(i).name for i in range(self._ncols)]
        # Mappa indice colonna → tipo (0=null,1=bool,2=uint8,3=int16,4=int32,5=currency,
        #  6=float,7=double,8=datetime,9=guid,10=uint16,11=int32,12=int64,13=float,
        #  14=double,15=datetime,16=binary,17=text,18=blob,19=text)
        self._coltypes = [table.get_column(i).type for i in range(self._ncols)]

    def _decode_val(self, record, idx):
        ct = self._coltypes[idx]
        try:
            raw = record.get_value_data(idx)
            if raw is None: return None
            # Interi
            if ct in (2, 3, 4, 10, 11):
                return record.get_value_data_as_integer(idx)
            if ct == 12:   # int64
                return record.get_value_data_as_integer(idx)
            if ct == 5:    # currency (int64)
                return record.get_value_data_as_integer(idx)
            if ct in (17, 19):  # testo Unicode (UTF-16-LE)
                return raw.decode('utf-16-le', errors='replace').rstrip('\x00')
            # Binario / GUID / tutto il resto → restituisce bytes
            return bytes(raw)
        except Exception:
            return None

    def next_row(self):
        if self._idx >= self._nrec:
            return None
        rec = self._table.get_record(self._idx)
        self._idx += 1
        row = {}
        for i, name in enumerate(self._colnames):
            row[name] = self._decode_val(rec, i)
        return row

class _PyesedbDB:
    def __init__(self, path):
        self._db = _pyesedb.file()
        self._db.open(path)
    def openTable(self, name):
        t = self._db.get_table_by_name(name)
        if t is None:
            raise KeyError(f"Tabella '{name}' non trovata")
        return _PyesedbCursor(t)
    def moveHead(self, cur):
        cur._idx = 0
    def getNextRow(self, cur):
        return cur.next_row()
    def closeTable(self, cur):
        pass

PRIV_GROUPS = {
    "Account Operators","Administrators","Backup Operators","Cert Publishers",
    "Cloneable Domain Controllers","DnsAdmins","DnsUpdateProxy","Domain Admins",
    "Domain Controllers","Debugger Users","Distributed COM Users","Enterprise Admins",
    "Enterprise Key Admins","Enterprise Read-only Domain Controllers",
    "Group Policy Creator Owners","Hyper-V Administrators",
    "Incoming Forest Trust Builders","Key Admins","Performance Log Users",
    "Print Operators","Remote Desktop Users","Remote Management Users",
    "Replicator","Schema Admins","Server Operators",
    "Storage Replica Administrators","System Managed Accounts Group"
}

UAC_DISABLED    = 0x00000002
UAC_WORKSTATION = 0x00001000
UAC_DC          = 0x00002000
UAC_NO_EXPIRE   = 0x00010000
UAC_SMARTCARD   = 0x00040000

SAMTYPE_USER     = 805306368
SAMTYPE_COMPUTER = 805306369
SAMTYPE_GROUP    = 268435456

FL_MAP = {0:'Windows 2000',1:'Windows Server 2003 Interim',2:'Windows Server 2003',
          3:'Windows Server 2008',4:'Windows Server 2008 R2',5:'Windows Server 2012',
          6:'Windows Server 2012 R2',7:'Windows Server 2016',10:'Windows Server 2025'}

WELLKNOWN_SIDS = {
    'S-1-1-0':'Everyone','S-1-5-11':'Authenticated Users','S-1-5-18':'SYSTEM',
    'S-1-3-0':'Creator Owner','S-1-5-9':'Enterprise Domain Controllers',
    'S-1-5-10':'Self','S-1-5-7':'Anonymous','S-1-5-15':'This Organization',
    'S-1-5-32-544':'BUILTIN\\Administrators','S-1-5-32-545':'BUILTIN\\Users',
    'S-1-5-32-546':'BUILTIN\\Guests','S-1-5-32-547':'BUILTIN\\Power Users',
    'S-1-5-32-548':'BUILTIN\\Account Operators',
    'S-1-5-32-549':'BUILTIN\\Server Operators',
    'S-1-5-32-550':'BUILTIN\\Print Operators',
    'S-1-5-32-551':'BUILTIN\\Backup Operators',
    'S-1-5-32-552':'BUILTIN\\Replicators',
}

# SID skip per ACL (non interessanti per IR)
ACL_SKIP_SIDS = {'S-1-5-18','S-1-3-0','S-1-5-10','S-1-5-9'}

DANGEROUS_RIGHTS = [
    (0x10000000,'GenericAll'),(0xF01FF,'FullControl'),
    (0x00040000,'WriteDACL'),(0x00080000,'WriteOwner'),
    (0x00000100,'ExtendedRight (DCSync/etc)'),
    (0x00000020,'WriteProperty'),(0x00000001,'CreateChild'),
    (0x00000002,'DeleteChild'),
]

# ── Helper functions ─────────────────────────────────────────────
def find_col(candidates, avail):
    for c in candidates:
        if c in avail:
            return c
    return None

def filetime_to_dt(ft):
    if not ft or ft <= 0 or ft >= 0x7FFFFFFFFFFFFFFF:
        return None
    try:
        return datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ft//10)
    except:
        return None

def ft_str(ft):
    dt = filetime_to_dt(ft)
    return dt.strftime('%d/%m/%Y %H:%M') if dt else 'N/A'

def ft_days_ago(ft):
    dt = filetime_to_dt(ft)
    if dt is None: return None
    return (datetime.datetime.utcnow() - dt).days

def gt_str(s):
    if not s: return 'N/A'
    try:
        if isinstance(s, bytes): s = s.decode('utf-8','replace')
        return datetime.datetime.strptime(s[:14],'%Y%m%d%H%M%S').strftime('%d/%m/%Y %H:%M')
    except:
        return str(s)[:16]

def sid_to_str(b):
    if not b or len(b) < 8: return ''
    try:
        rev = b[0]; n = b[1]
        auth = int.from_bytes(b[2:8],'big')
        subs = struct.unpack_from(f'<{n}I', b, 8)
        return f'S-{rev}-{auth}-' + '-'.join(str(s) for s in subs)
    except:
        return ''

def decode_str(v):
    if v is None: return ''
    if isinstance(v, bytes): return v.decode('utf-8','replace')
    return str(v)

def h(s):
    return html_mod.escape(str(s))

# ── Apri NTDS.dit (con recovery per dirty state da DC live) ─────
# PROBLEMA: alcune versioni di impacket chiamano mountDB() direttamente nel
# costruttore __init__ di ESENT_DB. Se ntds.dit ha dirty state (3,4,5),
# il crash avviene PRIMA che si possa intercettare con openDatabase().
#
# SOLUZIONE: patch binaria PREVENTIVA dell'offset 52 (dbstate ESE/JET Blue)
# sul file copiato in /tmp. Solo dopo, istanziamo ESENT_DB.
# Il file è già una copia locale → scrivere su di esso è sicuro e legale.
def _patch_ese_state(path):
    """
    Scansiona l'header ESE cercando il magic 0x89ABCDEF e,
    se lo trova, cerca dbstate nei byte successivi e lo patcha a 2 (CleanShutdown).
    Stampa hex dump diagnostico dei primi 80 byte.
    Restituisce (patched: bool, state_found: int).
    """
    ESE_MAGIC = 0x89ABCDEF
    try:
        with open(path, 'r+b') as f:
            header = f.read(80)
        # Hex dump diagnostico (visibile nei log con Debug ESE)
        hexdump = ' '.join(f'{b:02X}' for b in header[:80])
        sys.stderr.write(f"INFO: ESE header[0:80] = {hexdump}\n")

        # Cerca il magic nei primi 64 byte (offset tipici: 0 o 4)
        magic_off = -1
        for off in range(0, 64, 4):
            if len(header) < off + 4:
                break
            val = struct.unpack_from('<I', header, off)[0]
            if val == ESE_MAGIC:
                magic_off = off
                sys.stderr.write(f"INFO: ESE magic trovato a offset {off}\n")
                break

        if magic_off == -1:
            sys.stderr.write("INFO: ESE magic NON trovato nei primi 64 byte — patch impossibile\n")
            return False, -1

        # dbstate si trova tipicamente a +28 o +48 bytes dal magic
        # Prova entrambi gli offset e cerca un valore plausibile (0-6)
        state_off = -1
        state_val = -1
        for delta in (28, 48, 52, 44, 32, 36, 40, 56, 24, 20):
            cand_off = magic_off + delta
            if cand_off + 4 > len(header):
                continue
            v = struct.unpack_from('<I', header, cand_off)[0]
            if 1 <= v <= 6:   # valori ESE dbstate validi
                state_off = cand_off
                state_val = v
                sys.stderr.write(f"INFO: dbstate candidato: offset {cand_off} = {v}\n")
                if v in (3, 4, 5):   # dirty → prendi subito questo
                    break

        if state_off == -1:
            sys.stderr.write("INFO: dbstate non trovato con offset standard — skip patch\n")
            return False, 0

        sys.stderr.write(f"INFO: ESE dbstate = {state_val} (2=Clean, 3=DirtyShutdown) @ offset {state_off}\n")
        if state_val in (3, 4, 5):
            with open(path, 'r+b') as f:
                f.seek(state_off)
                f.write(struct.pack('<I', 2))   # CleanShutdown
            sys.stderr.write(f"INFO: dirty state patchato a CleanShutdown (2) @ offset {state_off}\n")
            return True, state_val

        return False, state_val

    except OSError as e:
        sys.stderr.write(f"INFO: impossibile patchare {path}: {e}\n")
        return False, -1

def _open_esent_db(path, skip_patch=False):
    import traceback

    # Step 1: patch binaria preventiva PRIMA di istanziare ESENT_DB.
    # SKIP se i log ESE sono disponibili: impacket usa il dirty state per capire
    # che deve applicare il soft-recovery. Se patchiamo a CleanShutdown, non lo fa.
    if skip_patch:
        patched, state = False, -1
        sys.stderr.write("INFO: skip patch dirty state — log ESE disponibili per soft-recovery\n")
    else:
        patched, state = _patch_ese_state(path)
        if state == -1:
            sys.stderr.write("WARNING: patch ESE non possibile, tentativo apertura diretta\n")
        else:
            sys.stderr.write(f"INFO: state={state}, patched={patched}\n")

    _last_err = "unknown"
    _last_tb  = ""

    # Step 2: apertura standard (funziona con impacket che richiede openDatabase())
    try:
        db = ESENT_DB(path)
        db.openDatabase()
        return db
    except Exception as e:
        _last_err = str(e)
        _last_tb  = traceback.format_exc()

    # Step 3: alcune versioni di impacket chiamano mountDB() in __init__ e caricano
    # già le tabelle senza richiedere openDatabase(). Proviamo a usare db direttamente.
    try:
        db3 = ESENT_DB(path)
        # Verifichiamo che le tabelle siano state caricate (indica mountDB ok)
        _ = db3.openTable('datatable')
        sys.stderr.write("INFO: aperto senza openDatabase() (mountDB in __init__)\n")
        return db3
    except Exception as e3:
        _last_err = f"{_last_err} | step3: {e3}"
        _last_tb  = traceback.format_exc()

    sys.stderr.write(f"FATAL cannot open NTDS.dit: {_last_err}\n")
    sys.stderr.write(_last_tb)
    sys.exit(1)

# ── Apertura DB (pyesedb → impacket fallback) ────────────────────
db = None
_db_backend = 'none'

if _HAVE_PYESEDB:
    try:
        # pyesedb handles logs automatically if they are in the same dir
        db = _PyesedbDB(ntds_path)
        _db_backend = 'pyesedb'
        sys.stderr.write("INFO: backend=pyesedb\n")
    except Exception as e2:
        sys.stderr.write(f"INFO: pyesedb fallito ({e2}), provo impacket\n")

if db is None and _HAVE_IMPACKET:
    try:
        import impacket.ese as _ese_mod
        
        # Monkey patch
        _orig_pp = _ese_mod.ESENT_DB.parsePage
        _ENTRIES_ATTR = '_ESENT_DB__pageEntries'
        def _safe_parsePage(self, page):
            try:
                _orig_pp(self, page)
            except Exception as _e:
                sys.stderr.write(f"INFO: parsePage tolerant skip pg: {_e}\n")
                try: self.__dict__[_ENTRIES_ATTR] = []
                except Exception: pass
        _ese_mod.ESENT_DB.parsePage = _safe_parsePage

        _orig_pc = _ese_mod.ESENT_DB.parseCatalog
        def _safe_parseCatalog(self, pageNum):
            try: _orig_pc(self, pageNum)
            except Exception as _ep: sys.stderr.write(f"INFO: parseCatalog tolerant skip pageNum={pageNum}: {_ep}\n")
        _ese_mod.ESENT_DB.parseCatalog = _safe_parseCatalog

        _orig_gnr = _ese_mod.ESENT_DB.getNextRow
        def _safe_getNextRow(self, cursor):
            # Loop di recupero tag/pagine corrotti.
            # NOTA: _orig_gnr chiama self.getNextRow() (= questa funzione) per le
            # transizioni di pagina normali → non richiamare mai _orig_gnr nel ramo
            # di errore, altrimenti si crea double-traversal esponenziale.
            # Strategia:
            #   - successo  → restituiamo la riga
            #   - eccezione → _orig_gnr ha già fatto CurrentTag += 1 prima di fallire
            #       * se CurrentTag >= FirstAvailablePageTag: pagina esaurita,
            #         avanziamo manualmente a NextPageNumber (reset tag=0)
            #       * altrimenti: tag corrotto nel mezzo della pagina;
            #         il prossimo giro incrementerà di nuovo e salterà il tag rotto
            for _attempts in range(500000):
                try:
                    return _orig_gnr(self, cursor)
                except Exception:
                    try:
                        page = cursor['CurrentPageData']
                        fap  = page.record['FirstAvailablePageTag']
                        if cursor['CurrentTag'] >= fap:
                            # Pagina esaurita con errore: segui NextPageNumber
                            next_pg = page.record['NextPageNumber']
                            if not next_pg:
                                return None
                            cursor['CurrentPageData'] = self.getPage(next_pg)
                            cursor['CurrentTag'] = 0
                        # else: tag corrotto nel mezzo — il prossimo _orig_gnr salta
                    except Exception:
                        return None
            return None
        _ese_mod.ESENT_DB.getNextRow = _safe_getNextRow
        ESENT_DB = _ese_mod.ESENT_DB
        if not hasattr(ESENT_DB, 'closeTable'): ESENT_DB.closeTable = lambda self, cur: None
        if not hasattr(ESENT_DB, 'moveHead'): ESENT_DB.moveHead = lambda self, cur: None

        # skip_patch=False: patchiamo sempre a CleanShutdown per sicurezza nella lettura.
        # I log ESE sono già stati copiati in tmp e verranno usati da openDatabase().
        db_imp = _open_esent_db(ntds_path, skip_patch=False)
        db = db_imp
        _db_backend = 'impacket'
        sys.stderr.write(f"INFO: backend=impacket has_logs={_has_ese_logs}\n")
    except SystemExit:
        raise
    except Exception as e:
        sys.stderr.write(f"INFO: impacket fallito ({e})\n")

if db is None:
    sys.stderr.write("FATAL: impossibile aprire NTDS.dit con nessun backend disponibile.\n")
    sys.stderr.write("INFO: installa libesedb-python: pip install libesedb-python\n")
    sys.exit(1)

# ── Scoperta colonne ─────────────────────────────────────────────
# Con impacket, getNextRow restituisce solo le colonne con dati nella riga corrente
# e le chiavi possono essere bytes. Normalizziamo tutto a str e campionamo più righe.
def _norm_key(k):
    return k.decode('utf-8', 'replace') if isinstance(k, bytes) else str(k)

def _norm_row(r):
    """Normalizza le chiavi del dict da bytes a str (impacket può restituire bytes)."""
    if r is None: return None
    return {_norm_key(k): v for k, v in r.items()}

_KEY_COLS = {'ATTm590045', 'ATTj590014', 'ATTi590014', 'ATTi590126', 'ATTq589920'}
try:
    cur = db.openTable('datatable')
    avail = set()
    for _ in range(500):
        _r = _norm_row(db.getNextRow(cur))
        if _r is None: break
        avail.update(_r.keys())
        if _KEY_COLS.issubset(avail): break   # colonne chiave trovate, stop anticipato
    db.closeTable(cur)
    sys.stderr.write(f"INFO: schema discovery — {len(avail)} colonne trovate in datatable\n")
    sys.stderr.write(f"INFO: COL_SAM={'ATTm590045' in avail} COL_STYPE={'ATTj590014' in avail or 'ATTi590014' in avail} COL_UAC={'ATTi590126' in avail}\n")
except Exception as e:
    sys.stderr.write(f"FATAL cannot read datatable schema: {e}\n"); sys.exit(1)

COL_SAM   = find_col(['ATTm590045'], avail)
COL_NAME  = find_col(['ATTm590480'], avail)
COL_UAC   = find_col(['ATTi590126'], avail)
COL_PWD   = find_col(['ATTq589920'], avail)
COL_LOGON = find_col(['ATTq589876'], avail)
COL_CRE   = find_col(['ATTm131074'], avail)
COL_CHG   = find_col(['ATTl131075','ATTm131075'], avail)
COL_SID   = find_col(['ATTb590597'], avail)
COL_ADMC  = find_col(['ATTj590420'], avail)
COL_EXP   = find_col(['ATTq591520','ATTq589984'], avail)
COL_NTSEC = find_col(['ATTp131353','ATTb131353'], avail)
COL_OS    = find_col(['ATTm590474'], avail)
COL_DISP  = find_col(['ATTm131085','ATTm590042'], avail)
COL_GPC   = find_col(['ATTm590258'], avail)
COL_BEHAV = find_col(['ATTj590182','ATTl590182'], avail)
COL_DEL   = find_col(['ATTi590829'], avail)
COL_STYPE = find_col(['ATTj590014','ATTi590014'], avail)
COL_SIDH  = find_col(['ATTb589971'], avail)
COL_OPTF  = find_col(['ATTb591520','ATTm591520'], avail)

# ── Diagnosi catalogo corrotto ────────────────────────────────────
_catalog_corrupt = (not COL_SAM or not COL_STYPE or not COL_UAC)
if _catalog_corrupt:
    _sample = sorted(c for c in list(avail)[:20] if c.startswith('ATT'))
    sys.stderr.write(f"WARNING: catalogo ESE parzialmente corrotto — colonne chiave non trovate\n")
    sys.stderr.write(f"WARNING: COL_SAM={COL_SAM} COL_STYPE={COL_STYPE} COL_UAC={COL_UAC}\n")
    sys.stderr.write(f"WARNING: campione colonne ATT trovate: {_sample[:10]}\n")
    sys.stderr.write(f"CATALOG_CORRUPT:1\n")
    # Scrivi comunque un report parziale, ma avvisa l'utente
    print("\n  [!] NTDS.dit con catalogo ESE parzialmente corrotto.")
    print("      I log ESE (.jrs) sono necessari per il soft-recovery.")
    print("      Vedi sezione 'Recupero NTDS.dit' nel report HTML.")

# ── Scansione datatable ──────────────────────────────────────────
objects      = {}   # dnt -> obj dict
names_by_dnt = {}   # dnt -> name (RDN)
pdnt_by_dnt  = {}   # dnt -> parent dnt
_scan_count = 0

try:
    sys.stdout.write("  [i] Lettura datatable in corso (0 oggetti)...\r")
    sys.stdout.flush()
    cur = db.openTable('datatable')
    db.moveHead(cur)
    while True:
        row = _norm_row(db.getNextRow(cur))
        if row is None: break

        _scan_count += 1
        if _scan_count % 500 == 0:
            sys.stdout.write(f"  [i] Lettura datatable in corso ({_scan_count} oggetti)... \r")
            sys.stdout.flush()

        try:
            dnt  = row.get('DNT_col')
            pdnt = row.get('PDNT_col')
            if dnt is None: continue
            if COL_DEL and row.get(COL_DEL): continue   # skip deleted

            rdn = decode_str(row.get(COL_NAME)) if COL_NAME else ''
            if rdn: names_by_dnt[dnt] = rdn
            if pdnt: pdnt_by_dnt[dnt] = pdnt

            sam     = decode_str(row.get(COL_SAM))   if COL_SAM  else ''
            ntsec   = row.get(COL_NTSEC)              if COL_NTSEC else None
            behav   = row.get(COL_BEHAV)              if COL_BEHAV else None
            gpc     = decode_str(row.get(COL_GPC))   if COL_GPC  else ''
            optf    = row.get(COL_OPTF)               if COL_OPTF  else None

            if not sam and not ntsec and behav is None and not gpc and not optf:
                continue

            sid_raw = row.get(COL_SID) if COL_SID else None
            objects[dnt] = {
                'dnt':       dnt,
                'pdnt':      pdnt,
                'name':      rdn,
                'sam':       sam,
                'uac':       (row.get(COL_UAC)   if COL_UAC  else None) or 0,
                'pwdLastSet':(row.get(COL_PWD)   if COL_PWD  else None),
                'lastLogon': (row.get(COL_LOGON) if COL_LOGON else None),
                'whenCreated':decode_str(row.get(COL_CRE)) if COL_CRE else '',
                'whenChanged':(row.get(COL_CHG)  if COL_CHG  else None),
                'sid':       sid_to_str(sid_raw) if sid_raw else '',
                'adminCount':(row.get(COL_ADMC)  if COL_ADMC else None) or 0,
                'expires':   (row.get(COL_EXP)   if COL_EXP  else None),
                'ntsec':     ntsec,
                'os':        decode_str(row.get(COL_OS))   if COL_OS  else '',
                'disp':      decode_str(row.get(COL_DISP)) if COL_DISP else '',
                'gpc':       gpc,
                'behavior':  behav,
                'samType':   (row.get(COL_STYPE) if COL_STYPE else None),
                'hasSidHist':bool(row.get(COL_SIDH)) if COL_SIDH else False,
                'optf':      optf,
            }
        except:
            continue
    db.closeTable(cur)
except Exception as e:
    sys.stderr.write(f"ERROR datatable scan: {e}\n")

sys.stderr.write(f"INFO: scan_count={_scan_count} objects={len(objects)} names={len(names_by_dnt)}\n")

# ── Classifica oggetti ────────────────────────────────────────────
def obj_type(o):
    st = o.get('samType')
    if st == SAMTYPE_COMPUTER: return 'computer'
    if st in (SAMTYPE_GROUP, 268435457): return 'group'
    if st == SAMTYPE_USER: return 'user'
    uac = o.get('uac', 0) or 0
    if uac & UAC_DC:          return 'dc'
    if uac & UAC_WORKSTATION: return 'computer'
    sam = o.get('sam', '') or ''
    if sam:
        # Fallback quando samAccountType non disponibile (catalogo corrotto):
        # account computer terminano con '$', gli altri sono utenti o gruppi.
        # I nomi gruppo coincidono spesso con o['name'] — distinguiamo per assenza di '$'.
        if sam.endswith('$'): return 'computer'
        return 'user'   # assume utente se ha SAM e non è un computer
    return 'other'

# ── Mappa SID → nome (dai dati estratti) ─────────────────────────
sid_map = dict(WELLKNOWN_SIDS)
for o in objects.values():
    if o['sid'] and o['sam']:
        sid_map[o['sid']] = o['sam']

# ── Costruzione DN ────────────────────────────────────────────────
def build_dn(dnt):
    parts = []; seen = set(); cur = dnt
    for _ in range(25):
        if cur in seen: break
        seen.add(cur)
        n = names_by_dnt.get(cur,'')
        if n: parts.append(f'CN={n}')
        p = pdnt_by_dnt.get(cur)
        if p is None or p == cur: break
        cur = p
    return ','.join(parts)

# ── Lettura link_table (group membership) ────────────────────────
group_members = {}   # group_dnt -> set(member_dnt)
member_of     = {}   # user_dnt  -> set(group_dnt)

try:
    cur = db.openTable('link_table')
    first_lrow = _norm_row(db.getNextRow(cur))
    lcols = set(first_lrow.keys()) if first_lrow else set()
    db.closeTable(cur)

    LCOL_GRP  = find_col(['link_DNT'],     lcols)
    LCOL_MBR  = find_col(['backlink_DNT'], lcols)
    LCOL_DEL  = find_col(['link_deltime'], lcols)
    LCOL_BASE = find_col(['link_base'],    lcols)
    # Filtriamo per link_base=2 (forward link dell'attributo `member` AD standard).
    # Se il catalogo è corrotto e link_base non è mappato correttamente,
    # accettiamo qualsiasi valore di link_base noto per l'attributo `member`:
    # linkID=2 nel default Windows AD schema.
    # Se LCOL_BASE è None (colonna non trovata), non filtriamo.
    MEMBER_LINK_BASE = 2  # forward linkID del `member` attribute

    if LCOL_GRP and LCOL_MBR:
        sys.stdout.write("\n  [i] Lettura link_table in corso (0 link)...\r")
        sys.stdout.flush()
        _link_count = 0
        cur = db.openTable('link_table')
        db.moveHead(cur)
        while True:
            row = _norm_row(db.getNextRow(cur))
            if row is None: break

            _link_count += 1
            if _link_count % 500 == 0:
                sys.stdout.write(f"  [i] Lettura link_table in corso ({_link_count} link)... \r")
                sys.stdout.flush()

            try:
                deltime = row.get(LCOL_DEL) if LCOL_DEL else None
                if deltime and deltime != 0: continue
                if LCOL_BASE:
                    lb = row.get(LCOL_BASE)
                    if lb is not None and lb != MEMBER_LINK_BASE: continue
                gd  = row.get(LCOL_GRP)
                md  = row.get(LCOL_MBR)
                if gd is None or md is None: continue
                group_members.setdefault(gd, set()).add(md)
                member_of.setdefault(md, set()).add(gd)
            except:
                continue
        db.closeTable(cur)
except Exception as e:
    sys.stderr.write(f"WARNING link_table: {e}\n")

# ── Indice gruppi per nome ────────────────────────────────────────
groups_by_name = {}  # name -> dnt
for dnt, o in objects.items():
    if _catalog_corrupt:
        # COL_NAME (ATTm590480) spesso corrotto → o['name'] quasi sempre vuoto.
        # I gruppi AD hanno sAMAccountName uguale al loro nome (es. "Domain Admins").
        # Indexiamo sia per name che per sam così PRIV_GROUPS viene trovato.
        if o['name']:
            groups_by_name[o['name']] = dnt
        if o['sam']:
            groups_by_name[o['sam']] = dnt
    else:
        if obj_type(o) in ('group','other') and o['name']:
            groups_by_name[o['name']] = dnt

priv_group_dnts = {dnt for name, dnt in groups_by_name.items() if name in PRIV_GROUPS}

# ── Membership transitiva (BFS) ────────────────────────────────────
def transitive_groups(user_dnt):
    visited = set(); queue = list(member_of.get(user_dnt, []))
    while queue:
        g = queue.pop()
        if g in visited: continue
        visited.add(g)
        for pg in member_of.get(g, []):
            if pg not in visited: queue.append(pg)
    return visited

# ── Analisi utenti privilegiati ────────────────────────────────────
priv_users = []
all_users  = [(dnt, o) for dnt, o in objects.items() if obj_type(o) in ('user','dc')]

for dnt, o in all_users:
    if not o['sam']: continue
    all_grp_dnts = transitive_groups(dnt)
    matched = all_grp_dnts & priv_group_dnts
    if not matched: continue

    memberships = []
    for gd in matched:
        go = objects.get(gd, {})
        gname = go.get('name', str(gd))
        is_direct = dnt in group_members.get(gd, set())
        memberships.append(f"{h(gname)} ({'Direct' if is_direct else 'Nested'})")

    uac = o['uac'] or 0
    pwd_dt  = filetime_to_dt(o['pwdLastSet'])
    logon_dt= filetime_to_dt(o['lastLogon'])
    expires_dt = filetime_to_dt(o.get('expires'))
    
    pwd_days= (datetime.datetime.utcnow()-pwd_dt).days if pwd_dt else None
    logon_days= (datetime.datetime.utcnow()-logon_dt).days if logon_dt else None
    is_expired= (datetime.datetime.utcnow() > expires_dt) if expires_dt else False
    is_inactive90 = (logon_days is not None and logon_days > 90)

    priv_users.append({
        'sam':        o['sam'],
        'groups':     '; '.join(memberships),
        'pwdLastSet': ft_str(o['pwdLastSet']),
        'pwdDays':    pwd_days,
        'lastLogon':  ft_str(o['lastLogon']),
        'disabled':   bool(uac & UAC_DISABLED),
        'noExpire':   bool(uac & UAC_NO_EXPIRE),
        'adminCount': o['adminCount'],
        'sid':        o['sid'],
        'hasSidHist': o['hasSidHist'],
        'stale':      (pwd_days is not None and pwd_days > 1),
        'expired':    is_expired,
        'inactive90': is_inactive90,
    })

priv_users.sort(key=lambda u: u['sam'].lower())

# ── KRBTGT ─────────────────────────────────────────────────────────
krbtgt_pwd = 'N/A'; krbtgt_days = None
for dnt, o in objects.items():
    if o['sam'].lower() == 'krbtgt':
        krbtgt_pwd  = ft_str(o['pwdLastSet'])
        krbtgt_days = ft_days_ago(o['pwdLastSet'])
        break

# ── Functional level ───────────────────────────────────────────────
domain_fl = 'N/A'; domain_root_dnt = None
for dnt, o in objects.items():
    if o.get('behavior') is not None:
        fv = o['behavior']
        domain_fl = f"{fv} — {FL_MAP.get(fv, 'Unknown')}"
        domain_root_dnt = dnt
        break

# ── Recycle Bin ────────────────────────────────────────────────────
recycle_bin = 'Disabled'
for dnt, o in objects.items():
    if o['name'] == 'Recycle Bin Feature':
        recycle_bin = 'Enabled'
        break
    if o.get('optf') and o['name']:
        recycle_bin = 'Enabled'
        break

# ── GPO ────────────────────────────────────────────────────────────
gpos = []
for dnt, o in objects.items():
    if not o['gpc']: continue
    disp = o['disp'] or o['name'] or '(no name)'
    when_chg = ft_str(o['whenChanged']) if isinstance(o['whenChanged'], int) else gt_str(o.get('whenCreated',''))
    days_ago = ft_days_ago(o['whenChanged']) if isinstance(o['whenChanged'], int) else None
    gpos.append({'name': disp, 'changed': when_chg, 'path': o['gpc'], 'days': days_ago})
gpos.sort(key=lambda g: g['changed'], reverse=True)

# ── Computer ────────────────────────────────────────────────────────
computers = []
for dnt, o in objects.items():
    if obj_type(o) not in ('computer','dc'): continue
    if not o['name']: continue
    computers.append({'name': o['name'], 'os': o['os'] or 'N/A', 'dn': build_dn(dnt)})
computers.sort(key=lambda c: c['name'].lower())

# ── Parser Security Descriptor ─────────────────────────────────────
def parse_sd(sd_bytes):
    aces = []
    if not sd_bytes or len(sd_bytes) < 20: return aces
    try:
        off_dacl = struct.unpack_from('<I', sd_bytes, 16)[0]
        if off_dacl == 0 or off_dacl >= len(sd_bytes): return aces
        ace_count = struct.unpack_from('<H', sd_bytes, off_dacl + 4)[0]
        ace_off   = off_dacl + 8
        for _ in range(min(ace_count, 200)):
            if ace_off + 8 > len(sd_bytes): break
            ace_type  = sd_bytes[ace_off]
            ace_flags = sd_bytes[ace_off + 1]
            ace_size  = struct.unpack_from('<H', sd_bytes, ace_off + 2)[0]
            if ace_size < 8: break
            access_mask = struct.unpack_from('<I', sd_bytes, ace_off + 4)[0]
            # SID offset: standard ACE=8, object ACE (type 5/6)=8+16[+16]
            sid_off = ace_off + 8
            if ace_type in (5, 6):
                flags_oe = struct.unpack_from('<I', sd_bytes, ace_off + 8)[0]
                sid_off = ace_off + 8 + 4
                if flags_oe & 1: sid_off += 16
                if flags_oe & 2: sid_off += 16
            sid_str = sid_to_str(sd_bytes[sid_off:sid_off + 68])
            if sid_str and sid_str not in ACL_SKIP_SIDS and ace_type in (0, 5):
                flagged = [nm for mask, nm in DANGEROUS_RIGHTS if access_mask & mask]
                if flagged:
                    aces.append({
                        'sid':       sid_str,
                        'name':      sid_map.get(sid_str, sid_str),
                        'rights':    ', '.join(flagged),
                        'mask':      hex(access_mask),
                        'inherited': bool(ace_flags & 0x10),
                        'everyone':  sid_str in ('S-1-1-0','S-1-5-11','S-1-5-7'),
                    })
            ace_off += ace_size
    except Exception as e:
        pass
    return aces

# ── ACL domain root + DC OU ────────────────────────────────────────
root_aces = []; dc_ou_aces = []

if domain_root_dnt is not None:
    ntsec = objects[domain_root_dnt].get('ntsec')
    if ntsec: root_aces = parse_sd(ntsec)

dc_ou_dnt = None
for dnt, o in objects.items():
    if o['name'] == 'Domain Controllers':
        p = pdnt_by_dnt.get(dnt)
        if p == domain_root_dnt or True:
            dc_ou_dnt = dnt; break

if dc_ou_dnt is not None:
    ntsec = objects[dc_ou_dnt].get('ntsec')
    if ntsec: dc_ou_aces = parse_sd(ntsec)

# ── Statistiche ────────────────────────────────────────────────────
total_priv    = len(priv_users)
total_stale   = sum(1 for u in priv_users if u['stale'])
total_disabled= sum(1 for u in priv_users if u['disabled'])
total_expired = sum(1 for u in priv_users if u['expired'])
total_inactive= sum(1 for u in priv_users if u['inactive90'])
total_gpo     = len(gpos)
gpo_recent    = sum(1 for g in gpos if g['days'] is not None and g['days'] <= 30)
total_comp    = len(computers)

# ── CSS / HTML ────────────────────────────────────────────────────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@300;400;500;700&family=DM+Sans:wght@400;500;700&display=swap');
:root{--bg:#080b0f;--bg2:#0d1117;--bg3:#131920;--bg4:#1a2332;
  --border:#1e2d3d;--accent:#58a6ff;--accent2:#ff7b72;
  --accent3:#3fb950;--accent4:#f0883e;
  --text:#c9d1d9;--text-dim:#3d5166;--text-mid:#6e8898;
  --mono:'Fira Code',monospace;--sans:'DM Sans',sans-serif;}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:var(--text);font-family:var(--sans);min-height:100vh;}
body::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background-image:linear-gradient(rgba(88,166,255,.03) 1px,transparent 1px),
  linear-gradient(90deg,rgba(88,166,255,.03) 1px,transparent 1px);
  background-size:40px 40px;}
header{position:relative;z-index:1;background:var(--bg2);border-bottom:1px solid var(--border);
  padding:1.8rem 3rem;display:flex;align-items:center;gap:2rem;}
.hicon{width:3rem;height:3rem;background:linear-gradient(135deg,var(--accent),#1f6feb);
  display:flex;align-items:center;justify-content:center;font-family:var(--mono);
  font-size:1.1rem;color:#fff;font-weight:700;flex-shrink:0;
  clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%);}
.htxt h1{font-size:1.35rem;font-weight:700;color:#fff;}
.htxt h1 span{color:var(--accent);}
.htxt .sub{font-size:.72rem;color:var(--text-dim);font-family:var(--mono);margin-top:.25rem;}
.hmeta{margin-left:auto;font-family:var(--mono);font-size:.7rem;color:var(--text-dim);text-align:right;line-height:2;}
.hmeta .val{color:var(--accent);}
.statsbar{position:relative;z-index:1;background:var(--bg3);border-bottom:1px solid var(--border);
  padding:1rem 3rem;display:flex;gap:3rem;flex-wrap:wrap;}
.stat .label{font-size:.6rem;text-transform:uppercase;letter-spacing:.15em;color:var(--text-dim);font-family:var(--mono);}
.stat .value{font-size:1.8rem;font-weight:800;font-family:var(--mono);color:var(--accent2);}
.stat.ok .value{color:var(--accent3);}
.stat.info .value{color:var(--accent);}
main{position:relative;z-index:1;padding:2rem 3rem;max-width:1600px;margin:0 auto;}
.stitle{font-size:.65rem;text-transform:uppercase;letter-spacing:.2em;color:var(--text-dim);
  font-family:var(--mono);margin-bottom:1.5rem;padding-bottom:.5rem;border-bottom:1px solid var(--border);}
.stitle::before{content:'> ';color:var(--accent3);}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:6px;overflow:hidden;margin-bottom:2rem;}
table{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:.78rem;}
th{padding:.5rem 1rem;text-align:left;font-size:.6rem;text-transform:uppercase;
  letter-spacing:.15em;color:var(--text-dim);border-bottom:1px solid var(--border);
  background:rgba(0,0,0,.3);}
td{padding:.5rem 1rem;border-bottom:1px solid rgba(30,45,61,.5);vertical-align:top;}
tr:last-child td{border-bottom:none;}
tr:hover{background:rgba(88,166,255,.03);}
.mono{font-family:var(--mono);}
.dim{color:var(--text-dim);}
.ok{color:var(--accent3);}
.bad{color:var(--accent2);}
.warn{color:var(--accent4);}
.tag{display:inline-block;padding:.1rem .4rem;border-radius:2px;font-size:.65rem;margin:.1rem;font-family:var(--mono);}
.tag-bad{background:rgba(255,123,114,.15);color:var(--accent2);border:1px solid rgba(255,123,114,.3);}
.tag-ok{background:rgba(63,185,80,.15);color:var(--accent3);border:1px solid rgba(63,185,80,.3);}
.tag-info{background:rgba(88,166,255,.15);color:var(--accent);border:1px solid rgba(88,166,255,.3);}
.tag-warn{background:rgba(240,136,62,.15);color:var(--accent4);border:1px solid rgba(240,136,62,.3);}
.row-bad{background:rgba(255,123,114,.06);border-left:3px solid var(--accent2);}
.row-warn{background:rgba(240,136,62,.06);}
footer{position:relative;z-index:1;border-top:1px solid var(--border);padding:1rem 3rem;
  font-family:var(--mono);font-size:.62rem;color:var(--text-dim);margin-top:3rem;
  display:flex;justify-content:space-between;}
footer span{color:var(--accent);}
"""

def tag(label, cls='info'):
    return f"<span class='tag tag-{cls}'>{h(label)}</span>"

def uac_tags(u):
    out = ''
    if u['disabled']:  out += tag('DISABLED','bad')
    if u['expired']:   out += tag('EXPIRED','bad')
    if u['inactive90']:out += tag('INACTIVE >90d','warn')
    if u['noExpire']:  out += tag('PWD-NO-EXPIRE','warn')
    if u['hasSidHist']:out += tag('SID-HISTORY','bad')
    if u['stale']:     out += tag('PWD-STALE','warn')
    return out

# ── Sezione utenti privilegiati ───────────────────────────────────
rows_priv = ''
for u in priv_users:
    cls = 'row-bad' if (u['disabled'] or u['hasSidHist'] or u['expired']) else ('row-warn' if (u['stale'] or u['inactive90']) else '')
    rows_priv += f"""<tr class='{cls}'>
<td class='mono ok' style='white-space:nowrap'>{h(u['sam'])}</td>
<td style='font-size:.68rem'>{u['groups']}</td>
<td class='mono {"bad" if u["stale"] else "ok"}' style='white-space:nowrap'>{h(u['pwdLastSet'])}</td>
<td class='mono dim' style='white-space:nowrap'>{h(u['lastLogon'])}</td>
<td>{uac_tags(u)}</td>
<td class='mono dim' style='font-size:.65rem;word-break:break-all'>{h(u['sid'])}</td>
</tr>"""

# ── Sezione GPO ────────────────────────────────────────────────────
rows_gpo = ''
for g in gpos:
    cls = 'row-warn' if (g['days'] is not None and g['days'] <= 30) else ''
    badge = tag(f"{g['days']}d ago",'warn') if (g['days'] is not None and g['days'] <= 30) else ''
    rows_gpo += f"""<tr class='{cls}'>
<td class='mono'>{h(g['name'])}{badge}</td>
<td class='mono {"warn" if (g["days"] is not None and g["days"]<=30) else "dim"}' style='white-space:nowrap'>{h(g['changed'])}</td>
<td class='mono dim' style='font-size:.65rem;word-break:break-all'>{h(g['path'])}</td>
</tr>"""

# ── Sezione Computer ────────────────────────────────────────────────
rows_comp = ''
for c in computers:
    rows_comp += f"""<tr>
<td class='mono ok'>{h(c['name'])}</td>
<td class='mono dim'>{h(c['os'])}</td>
<td class='mono dim' style='font-size:.65rem'>{h(c['dn'])}</td>
</tr>"""

# ── Sezione ACL ────────────────────────────────────────────────────
def ace_rows(aces):
    if not aces:
        return "<tr><td colspan='5' class='dim' style='padding:1rem;text-align:center'>Nessun ACE pericoloso rilevato (o nTSecurityDescriptor non trovato)</td></tr>"
    out = ''
    for a in aces:
        cls = 'row-bad' if a['everyone'] else 'row-warn'
        ev  = tag('EVERYONE-LIKE','bad') if a['everyone'] else ''
        inh = tag('inherited','info') if a['inherited'] else ''
        out += f"""<tr class='{cls}'>
<td class='mono {"bad" if a["everyone"] else "warn"}'>{h(a['name'])}{ev}</td>
<td class='mono dim' style='font-size:.65rem;word-break:break-all'>{h(a['sid'])}</td>
<td class='mono bad'>{h(a['rights'])}</td>
<td class='mono dim'>{h(a['mask'])}</td>
<td>{inh}</td>
</tr>"""
    return out

# ── Functional level badge ─────────────────────────────────────────
fl_badge_cls = 'ok' if domain_fl != 'N/A' else 'dim'
krbtgt_cls = 'bad' if (krbtgt_days is not None and krbtgt_days > 180) else 'ok'
krbtgt_badge = tag(f'{krbtgt_days}d ago','bad') if (krbtgt_days and krbtgt_days > 180) else (tag(f'{krbtgt_days}d ago','ok') if krbtgt_days else '')

# ── Blocco HTML per catalogo corrotto (pre-generato per evitare f-string annidati) ──
_catalog_corrupt_html = ""
if _catalog_corrupt:
    _objects_count = len(objects)
    _catalog_corrupt_html = f"""
<div class="stitle" style="color:var(--accent2)">&#9888; NTDS.dit &mdash; Catalogo ESE corrotto: analisi privilegiata non disponibile</div>
<div class="card" style="border-color:rgba(255,123,114,.5);background:rgba(255,123,114,.06)">
<div style="padding:1.4rem 1.6rem;font-family:var(--mono);font-size:.82rem;line-height:1.9">
<div style="color:var(--accent2);font-weight:700;font-size:.95rem;margin-bottom:.9rem">Stato: DirtyShutdown &mdash; Pagine ESE corrotte nel catalogo di sistema</div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem">
  <div style="background:rgba(0,0,0,.25);border-radius:6px;padding:.9rem 1rem">
    <div style="color:var(--accent3);font-weight:700;margin-bottom:.5rem">&#10003; Dati estratti con successo</div>
    <div style="color:var(--text-mid)">&#x2022; Password KRBTGT (pwdLastSet)</div>
    <div style="color:var(--text-mid)">&#x2022; Livello funzionale dominio</div>
    <div style="color:var(--text-mid)">&#x2022; Oggetti GPO e percorsi</div>
    <div style="color:var(--text-mid)">&#x2022; Elenco computer (parziale)</div>
    <div style="color:var(--text-mid)">&#x2022; {_objects_count} oggetti datatable recuperati</div>
  </div>
  <div style="background:rgba(0,0,0,.25);border-radius:6px;padding:.9rem 1rem">
    <div style="color:var(--accent2);font-weight:700;margin-bottom:.5rem">&#10007; Dati non disponibili</div>
    <div style="color:var(--text-mid)">&#x2022; sAMAccountName (colonna non mappata)</div>
    <div style="color:var(--text-mid)">&#x2022; samAccountType / userAccountControl</div>
    <div style="color:var(--text-mid)">&#x2022; Membership gruppi privilegiati</div>
    <div style="color:var(--text-mid)">&#x2022; Link-table group membership</div>
    <div style="color:var(--text-mid)">&#x2022; Hash delle password</div>
  </div>
</div>
<div style="color:var(--text-dim);margin-bottom:1rem">
  <b style="color:var(--text)">Causa tecnica:</b> Le pagine ESE del catalogo di sistema (pagine 14, 19&ndash;22) sono corrotte.
  Il catalogo mappa i nomi delle colonne AD (es. <code style="color:var(--accent)">ATTm590045</code> &rarr; sAMAccountName) agli ID interni:
  senza di esso, impacket e libesedb non possono identificare correttamente i campi.
  Le entry di group membership nella link_table risultano su pagine non leggibili.
</div>
<div style="margin-top:.8rem;color:var(--text);font-weight:700;border-top:1px solid rgba(255,255,255,.08);padding-top:.8rem">Procedura di recupero</div>
<div style="margin:.6rem 0">
  <div style="color:var(--accent3);font-weight:600">Opzione 1 &mdash; esentutl su Windows <span style="color:var(--text-dim)">(recovery completo, consigliato)</span></div>
  <div style="color:var(--text-dim);margin:.2rem 0 .2rem 1rem">Sul DC sorgente o una VM Windows con ntds.dit + log ESE nella stessa cartella:</div>
  <div style="background:rgba(0,0,0,.3);border-radius:4px;padding:.4rem .8rem;margin:.2rem 0 .4rem 1rem;color:var(--accent)">esentutl /r edb /l "C:\\Windows\\NTDS" /s "C:\\Windows\\NTDS" /!32768</div>
  <div style="color:var(--text-dim);margin-left:1rem">Poi copia ntds.dit recuperato e rilancia il modulo 38.</div>
</div>
<div style="margin:.6rem 0">
  <div style="color:var(--accent3);font-weight:600">Opzione 2 &mdash; ntdsutil IFM <span style="color:var(--text-dim)">(se il DC &egrave; ancora operativo)</span></div>
  <div style="background:rgba(0,0,0,.3);border-radius:4px;padding:.4rem .8rem;margin:.2rem 0 .4rem 1rem;color:var(--accent)">ntdsutil "activate instance ntds" "ifm" "create full C:\\ifm" quit quit</div>
  <div style="color:var(--text-dim);margin-left:1rem">Il file <code>C:\\ifm\\Active Directory\\ntds.dit</code> sar&agrave; gi&agrave; in CleanShutdown.</div>
</div>
<div style="margin:.6rem 0">
  <div style="color:var(--accent3);font-weight:600">Opzione 3 &mdash; esentutl /p (hard repair) <span style="color:var(--text-dim)">(richiede Windows &mdash; usa solo se Opzione 1 fallisce)</span></div>
  <div style="color:var(--text-dim);margin:.2rem 0 .2rem 1rem">
    Ripara le pagine fisicamente danneggiate azzerandole. <b style="color:var(--accent2)">Possibile perdita dati parziale.</b> Esegui prima un backup.
  </div>
  <div style="background:rgba(0,0,0,.3);border-radius:4px;padding:.4rem .8rem;margin:.2rem 0 .4rem 1rem;color:var(--accent)">
    esentutl /p ntds.dit /!32768 /8 /o
  </div>
  <div style="color:var(--text-dim);margin-left:1rem">
    Dopo /p: ri-esegui <code>/r edb</code> per il soft-recovery, poi rilancia il modulo 38.<br>
    <b>Nota:</b> il soft-recovery (<code>impacket</code> o <code>esentutl /r</code>) non ripara pagine fisicamente corrotte &mdash; serve /p prima.
  </div>
</div>
<div style="margin:.6rem 0;padding:.6rem .8rem;background:rgba(255,123,114,.08);border-radius:4px;border-left:3px solid rgba(255,123,114,.5)">
  <b style="color:var(--accent2)">Perch&eacute; il soft-recovery non basta:</b>
  <span style="color:var(--text-dim)"> Il replay dei log ESE corregge solo lo stato <i>DirtyShutdown</i> (transazioni non committed).
  Le pagine di catalogo 14, 19&ndash;22 di questo NTDS.dit sono fisicamente danneggiate (byte corrotti nel payload ESE):
  nessun log pu&ograve; ripristinare dati che non esistono pi&ugrave; sul disco.</span>
</div>
</div></div>"""

# ── HTML output ────────────────────────────────────────────────────
html_out = f"""<!DOCTYPE html>
<html lang="it"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>PAD Offline — DFIR Report</title>
<style>{CSS}</style></head><body>
<header>
  <div class="hicon">AD</div>
  <div class="htxt">
    <h1>PAD Offline — <span>Panoramica Active Directory</span></h1>
    <div class="sub">NTDS.dit offline forensic analysis</div>
  </div>
  <div class="hmeta">
    <div>Scansione: <span class="val">{h(scan_date)}</span></div>
    <div>Sorgente: <span class="val">{h(source_root)}</span></div>
    <div style="font-size:.6rem">SHA256: <span class="val" style="word-break:break-all">{h(ntds_hash)}</span></div>
  </div>
</header>
<div class="statsbar">
  <div class="stat"><div class="label">Utenti Privilegiati</div><div class="value">{total_priv}</div></div>
  <div class="stat"><div class="label">Gruppi Monitorati</div><div class="value info" style="color:var(--accent)">{len(PRIV_GROUPS)}</div></div>
  <div class="stat"><div class="label">KRBTGT (giorni)</div><div class="value {krbtgt_cls}">{krbtgt_days if krbtgt_days is not None else 'N/A'}</div></div>
  <div class="stat"><div class="label">Password Stale (&gt;24h)</div><div class="value">{total_stale}</div></div>
  <div class="stat"><div class="label">GPO Modificate (30gg)</div><div class="value {"" if gpo_recent==0 else ""}">{gpo_recent}/{total_gpo}</div></div>
  <div class="stat ok"><div class="label">Computer nel Dominio</div><div class="value">{total_comp}</div></div>
</div>
<main>

<div class="stitle">Configurazione Dominio</div>
<div class="card"><table>
<thead><tr><th>Parametro</th><th>Valore</th><th>Note</th></tr></thead>
<tbody>
<tr><td class="mono dim">Domain Functional Level</td>
    <td class="mono {fl_badge_cls}">{h(domain_fl)}</td><td></td></tr>
<tr><td class="mono dim">KRBTGT pwdLastSet</td>
    <td class="mono {krbtgt_cls}">{h(krbtgt_pwd)}</td>
    <td>{krbtgt_badge}{'<span class="tag tag-bad">GOLDEN TICKET RISK: rotazione &gt;180gg</span>' if (krbtgt_days and krbtgt_days>180) else ''}</td></tr>
<tr><td class="mono dim">AD Recycle Bin</td>
    <td class="mono {"ok" if recycle_bin=="Enabled" else "bad"}">{h(recycle_bin)}</td>
    <td>{"" if recycle_bin=="Enabled" else "<span class='tag tag-warn'>Oggetti eliminati non recuperabili</span>"}</td></tr>
<tr><td class="mono dim">Utenti Privilegiati Disabilitati</td>
    <td class="mono {"warn" if total_disabled>0 else "ok"}">{total_disabled}</td><td></td></tr>
</tbody></table></div>

{_catalog_corrupt_html}
<div class="stitle">Utenti Privilegiati — {total_priv} account ({len(PRIV_GROUPS)} gruppi monitorati)</div>
<div class="card"><table>
<thead><tr><th>sAMAccountName</th><th>Gruppi (Direct/Nested)</th><th>pwdLastSet</th><th>lastLogon</th><th>Flags</th><th>SID</th></tr></thead>
<tbody>{rows_priv if rows_priv else "<tr><td colspan='6' class='dim' style='padding:1rem;text-align:center'>Nessun utente privilegiato trovato</td></tr>"}</tbody>
</table></div>

<div class="stitle">ACL Domain Root — ACE con diritti pericolosi</div>
<div class="card"><table>
<thead><tr><th>Trustee</th><th>SID</th><th>Diritti</th><th>Access Mask</th><th>Flags</th></tr></thead>
<tbody>{ace_rows(root_aces)}</tbody>
</table></div>

<div class="stitle">ACL OU Domain Controllers — ACE con diritti pericolosi</div>
<div class="card"><table>
<thead><tr><th>Trustee</th><th>SID</th><th>Diritti</th><th>Access Mask</th><th>Flags</th></tr></thead>
<tbody>{ace_rows(dc_ou_aces)}</tbody>
</table></div>

<div class="stitle">Group Policy Objects — {total_gpo} GPO totali</div>
<div class="card"><table>
<thead><tr><th>Display Name</th><th>Ultima Modifica</th><th>Path SYSVOL</th></tr></thead>
<tbody>{rows_gpo if rows_gpo else "<tr><td colspan='3' class='dim' style='padding:1rem;text-align:center'>Nessuna GPO trovata</td></tr>"}</tbody>
</table></div>

<div class="stitle">Computer nel Dominio — {total_comp} oggetti</div>
<div class="card"><table>
<thead><tr><th>Computer Name</th><th>Operating System</th><th>Distinguished Name</th></tr></thead>
<tbody>{rows_comp if rows_comp else "<tr><td colspan='3' class='dim' style='padding:1rem;text-align:center'>Nessun computer trovato</td></tr>"}</tbody>
</table></div>

</main>
<footer>
  <span>PAD Offline</span>
  <span>fiuto.sh — DFIR Toolkit</span>
</footer>
</body></html>"""

try:
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html_out)
except Exception as e:
    sys.stderr.write(f"ERROR writing report: {e}\n"); sys.exit(1)

# Output stats per bash (su stderr per non interferire)
sys.stderr.write(f"STATS:{total_priv}:{total_stale}:{krbtgt_days or 'N/A'}:{gpo_recent}:{total_comp}\n")
PYEOF

    if [[ -f "$REPORT_HTML" ]]; then
        # Copia il pyerr.log nel report dir per debug postumo
        if [[ -f "$PYERR_FILE" ]]; then
            cp "$PYERR_FILE" "$(dirname "$REPORT_HTML")/impacket_debug.log" 2>/dev/null || true
        fi

        # Avviso catalogo corrotto
        if grep -q '^CATALOG_CORRUPT:1' "$PYERR_FILE" 2>/dev/null; then
            echo ""
            warn "NTDS.dit con catalogo ESE parzialmente corrotto — dati utente non disponibili."
            info "Le colonne sAMAccountName/samAccountType sono su pagine di catalogo non leggibili."
            echo ""
            echo -e "  ${BOLD}${YELLOW}Recupero necessario — scegli una delle opzioni:${RESET}"
            echo -e "  ${CYAN}1) Soft-recovery Linux${RESET} (richiede i log .jrs nella stessa dir di ntds.dit)"
            echo -e "     Copia ntds.dit + edb*.jrs + edb.chk in una cartella, poi rilancia."
            echo ""
            echo -e "  ${CYAN}2) esentutl su Windows:${RESET}"
            echo -e "     ${DIM}esentutl /r edb /l C:\\Windows\\NTDS /s C:\\Windows\\NTDS${RESET}"
            echo ""
            echo -e "  ${CYAN}3) IFM snapshot (DC attivo):${RESET}"
            echo -e "     ${DIM}ntdsutil \"activate instance ntds\" \"ifm\" \"create full C:\\ifm\" quit quit${RESET}"
            echo ""
            info "Il report HTML contiene la guida completa al recupero."
        fi

        info "Report generato: ${BOLD}$REPORT_HTML${RESET}"
        open_report_prompt "$REPORT_HTML"
    else
        local _PYERR _PYINFO
        _PYERR=$(grep -v '^STATS:' "$PYERR_FILE" 2>/dev/null | grep -iv '^INFO:' | head -10 || true)
        _PYINFO=$(grep -i '^INFO:' "$PYERR_FILE" 2>/dev/null | head -10 || true)
        [[ -n "$_PYINFO" ]] && info "Debug ESE: $_PYINFO"
        [[ -n "$_PYERR" ]] && warn "Errore Python: $_PYERR"

        # Diagnosi specifica per errori noti
        if echo "${_PYERR}" | grep -qiE 'Errno 35|EAGAIN|temporarily unavailable|Resource temporarily'; then
            err "ntds.dit è bloccato dal servizio Active Directory (Errno 35 / EAGAIN)."
            echo ""
            echo -e "  ${YELLOW}Il client SMB rispetta il lock esclusivo tenuto da lsass/ntds.exe.${RESET}"
            echo -e "  ${YELLOW}Il file non è apribile nemmeno in lettura finché il DC è attivo.${RESET}"
            echo ""
            echo -e "  ${BOLD}Soluzioni per DC live:${RESET}"
            echo -e "  ${CYAN}1) VSS snapshot (consigliato):${RESET}"
            echo -e "     Sul DC (PowerShell admin):"
            echo -e "     ${DIM}ntdsutil \"activate instance ntds\" \"ifm\" \"create full C:\\\\ifm_snapshot\" quit quit${RESET}"
            echo -e "     Poi monta/copia la cartella C:\\\\ifm_snapshot e rilancia su quel path."
            echo ""
            echo -e "  ${CYAN}2) impacket secretsdump remoto (solo hash NTLM, senza analisi ACL/GPO):${RESET}"
            echo -e "     ${DIM}impacket-secretsdump -just-dc-ntlm DOMINIO/utente:password@IP_DC${RESET}"
            echo ""
            echo -e "  ${CYAN}3) Se hai già un'immagine disco (VHDX/E01) del DC:${RESET}"
            echo -e "     Monta offline con ${BOLD}./mount_image.sh${RESET} e rilancia il modulo su quel mount point."
        elif echo "${_PYERR}" | grep -qiE 'dirty|state|DirtyShutdown|Unknown state'; then
            err "Generazione report fallita — database ESE in stato dirty non recuperabile."
            info "Suggerimento: copia il file e prova 'esentutl /r edb /l .' in un ambiente Windows"
        else
            err "Generazione report fallita. Verifica che ntds.dit sia leggibile e non corrotto."
            info "Suggerimento: copia il file e prova 'esentutl /r edb' in un ambiente Windows"
        fi
    fi
    rm -rf "$TMP_PAD_DIR" 2>/dev/null || true
}

run_all_modules() {
    clear
    print_banner
    info "Esecuzione di tutti i moduli in modalità batch..."
    echo ""
    if [[ -z "$REPORT_BASE_DIR" ]]; then
        REPORT_BASE_DIR="${INVOCATION_DIR}/fiuto_reports_$(date +%Y%m%d_%H%M%S)"
        LOG_FILE="${REPORT_BASE_DIR}/fiuto_session_$(date +%Y%m%d_%H%M%S).log"
    fi
    info "Report base dir: ${BOLD}$REPORT_BASE_DIR${RESET}"
    log_msg "=== FIUTO avviato in modalità batch — WIN_ROOT=$WIN_ROOT ==="
    sleep 1

    BATCH_MODE=true
    SUMMARY_TABLE=()

    echo ""
    run_batch_module 1 module_ps_history "PowerShell History" 38
    run_batch_module 2 module_notepad_tabstate "Notepad TabState" 38
    run_batch_module 3 module_ifeo "IFEO" 38
    run_batch_module 4 module_bam "BAM" 38
    run_batch_module 5 module_run_keys "Run Keys" 38
    run_batch_module 6 module_prefetch "Prefetch" 38
    run_batch_module 7 module_scheduled_tasks "Scheduled Tasks" 38
    run_batch_module 8 module_usb "USB Devices" 38
    run_batch_module 9 module_lnk "LNK Files" 38
    run_batch_module 10 module_rdp_cache "RDP Cache" 38
    run_batch_module 11 module_services "Services" 38
    run_batch_module 12 module_evtx "EVTX" 38
    run_batch_module 13 module_amcache "Amcache" 38
    run_batch_module 14 module_recycle_bin "Recycle Bin" 38
    run_batch_module 15 module_wmi "WMI" 38
    run_batch_module 16 module_srum "SRUM" 38
    run_batch_module 17 module_browser "Browser History" 38
    run_batch_module 18 module_userassist "UserAssist" 38
    run_batch_module 19 module_shellbags "Shellbags" 38
    run_batch_module 20 module_sam "SAM" 38
    run_batch_module 21 module_mft "MFT" 38
    run_batch_module 22 module_opensave "Open/Save MRU" 38
    run_batch_module 23 module_usn "USN Journal" 38
    run_batch_module 24 module_ntds "NTDS.dit" 38
    run_batch_module 25 module_hiberfil "Hibernation / Pagefile" 38
    run_batch_module 26 module_evtx_tamper "EVTX Tampering" 38
    run_batch_module 27 module_credential_manager "Credential Manager" 38
    run_batch_module 28 module_wlan "WLAN Profiles" 38
    run_batch_module 29 module_appx "AppX / UWP" 38
    run_batch_module 30 module_browser_extra "Browser Logins/Downloads" 38
    run_batch_module 31 module_clipboard "Clipboard History" 38
    run_batch_module 32 module_office_mru "Office MRU" 38
    run_batch_module 33 module_defender_quarantine "Defender Quarantine" 38
    run_batch_module 34 module_ps_scriptblock "PS ScriptBlock Log" 38
    run_batch_module 35 module_jumplists "JumpLists" 38
    run_batch_module 36 module_network_artifacts "Network Artifacts" 38
    run_batch_module 37 module_master_timeline "Master Timeline" 38
    # Modulo 38: eseguito solo se il disco è un Domain Controller (ntds.dit presente)
    local _ntds_check
    _ntds_check=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    if [[ -n "$_ntds_check" ]]; then
        run_batch_module 38 module_pad_offline "PAD Offline AD" 38
    else
        echo -e "  ${DIM}[i] [38/38] PAD Offline AD — saltato (non è un Domain Controller)${RESET}"
        SUMMARY_TABLE+=("38|PAD Offline AD|SKIP|non è un DC")
    fi

    BATCH_MODE=false

    echo ""
    section_header "Riepilogo Scansione Globale" "$GREEN"
    echo -e "  ${BOLD}MOD  NOME MODULO                        EVIDENZE     FILE GENERATI${RESET}"
    echo "  ─────────────────────────────────────────────────────────────────────────────────────────"
    for row in "${SUMMARY_TABLE[@]}"; do
        IFS='|' read -r mnum mname msy mpath <<< "$row"
        if [[ "$msy" == "SI" ]]; then
            local rel_path="${mpath#$REPORT_BASE_DIR/}"
            printf "  ${CYAN}%02d${RESET}   ${BOLD}%-32s${RESET} ${GREEN}%-10s${RESET}   ${DIM}%s${RESET}\n" "$mnum" "$mname" "TROVATE" "$rel_path"
        else
            printf "  ${CYAN}%02d${RESET}   %-32s ${DIM}%-10s${RESET}   ${DIM}-${RESET}\n" "$mnum" "$mname" "NESSUNA"
        fi
    done
    echo ""
    ok "Report salvati integralmente in: ${BOLD}$REPORT_BASE_DIR"
}

# ================================================================
#  AUTODETECT ROOT WINDOWS
# ================================================================

debug_mounts() {
    echo ""
    section_header "DEBUG — Mount attivi su questo sistema" "$YELLOW"
    echo -e "  ${DIM}── /proc/mounts (non di sistema) ───────────────────${RESET}"
    echo ""
    local SKIP_FS='tmpfs|sysfs|proc|devtmpfs|cgroup2?|fusectl|tracefs|securityfs|pstore|bpf|hugetlbfs|mqueue|debugfs|configfs|overlay|squashfs|nsfs|efivarfs|autofs|ramfs|rpc_pipefs'
    local SKIP_MNT='^/(proc|sys|dev|run|snap)(/|$)'
    while IFS=' ' read -r RAW_DEV RAW_MNT FSTYPE _; do
        local DEV MNT
        DEV=$(printf '%b' "$RAW_DEV")
        MNT=$(printf '%b' "$RAW_MNT")
        [[ "$MNT" =~ $SKIP_MNT || "$MNT" == "/" || "$MNT" == /tmp/* ]] && continue
        if [[ "$FSTYPE" =~ ^($SKIP_FS)$ ]]; then
            printf "  ${DIM}  %-38s %-30s %s${RESET}\n" "$DEV" "$MNT" "$FSTYPE"
        else
            printf "  ${CYAN}→ %-38s ${GREEN}%-30s${RESET} ${YELLOW}%s${RESET}\n" "$DEV" "$MNT" "$FSTYPE"
        fi
    done < /proc/mounts
    echo ""
    echo -e "  ${DIM}(${CYAN}→${DIM} = candidati; grigio = esclusi)${RESET}"
    echo ""
    echo -ne "  ${YELLOW}Premi qualsiasi tasto per tornare al menu...${RESET}"
    pause_key
}

_find_windows_mounts() {
    local -a CANDIDATES=()
    while IFS=' ' read -r RAW_DEV RAW_MNT FSTYPE _REST; do
        local MNT
        MNT=$(printf '%b' "$RAW_MNT")
        [[ -z "$MNT" || "$MNT" == "/" ]] && continue
        [[ "$MNT" == /proc* || "$MNT" == /sys* || "$MNT" == /dev* ]] && continue
        [[ "$MNT" == /snap/* || "$MNT" == /run/* || "$MNT" == /tmp/* ]] && continue
        [[ "$RAW_DEV" == *.AppImage ]] && continue
        [[ "$FSTYPE" == "fuse.ewfmount" || "$FSTYPE" == "fuse.xmount" ]] && continue
        if [[ -d "$MNT/Users" || -d "$MNT/Windows" ||
              -d "$MNT/users" || -d "$MNT/windows" ]]; then
            CANDIDATES+=("$MNT"); continue
        fi
        if find "$MNT" -maxdepth 1 -type d \( -iname "Users" -o -iname "Windows" \) \
               2>/dev/null | grep -q .; then
            CANDIDATES+=("$MNT")
        fi
    done < /proc/mounts
    local -A SEEN=()
    for MNT in "${CANDIDATES[@]}"; do
        [[ -n "${SEEN[$MNT]+x}" ]] && continue
        SEEN["$MNT"]=1; echo "$MNT"
    done
}

autodetect_win_root() {
    local SILENT="${1:-}"
    local -a FOUND=()
    mapfile -t FOUND < <(_find_windows_mounts)
    if [[ ${#FOUND[@]} -eq 0 ]]; then
        [[ "$SILENT" != "silent" ]] && warn "Nessun volume Windows rilevato tra i filesystem montati."
        return 1
    fi

    # Per ogni volume trovato, aggiungi Windows.old se presente e valido
    local -a EXTRA_PATHS=() EXTRA_BASES=()
    for MNT in "${FOUND[@]}"; do
        local WOLD
        WOLD=$(find "$MNT" -maxdepth 1 -type d -iname "Windows.old" 2>/dev/null | head -1)
        if [[ -n "$WOLD" ]]; then
            # Verifica che Windows.old contenga una struttura Windows
            if find "$WOLD" -maxdepth 1 -type d \( -iname "Users" -o -iname "Windows" \) \
                    2>/dev/null | grep -q .; then
                EXTRA_PATHS+=("$WOLD")
                EXTRA_BASES+=("$MNT")
            fi
        fi
    done

    # Costruisci la lista finale: volumi base + Windows.old
    local -a ALL_PATHS=("${FOUND[@]}" "${EXTRA_PATHS[@]}")

    echo ""
    echo -e "  ${CYAN}${BOLD}Volumi Windows rilevati:${RESET}"
    echo ""
    local IDX=1
    for MNT in "${FOUND[@]}"; do
        local LABEL; LABEL=$(basename "$MNT")
        local USERS_DIR
        USERS_DIR=$(find "$MNT" -maxdepth 1 -type d -iname "Users" 2>/dev/null | head -1)
        local USER_COUNT=0
        [[ -n "$USERS_DIR" ]] && \
            USER_COUNT=$(find "$USERS_DIR" -maxdepth 1 -mindepth 1 -type d 2>/dev/null \
                | grep -ciEv '/(Public|Default|Default User|All Users)$' || true)
        echo -e "  ${GREEN}[${IDX}]${RESET}  ${BOLD}${MNT}${RESET} ${MAGENTA}(${LABEL})${RESET}  ${CYAN}${USER_COUNT} utenti${RESET}"
        IDX=$((IDX + 1))
    done
    # Mostra le opzioni Windows.old con indicatore visivo
    for i in "${!EXTRA_PATHS[@]}"; do
        local WOLD="${EXTRA_PATHS[$i]}"
        local BASE="${EXTRA_BASES[$i]}"
        local USERS_DIR
        USERS_DIR=$(find "$WOLD" -maxdepth 1 -type d -iname "Users" 2>/dev/null | head -1)
        local USER_COUNT=0
        [[ -n "$USERS_DIR" ]] && \
            USER_COUNT=$(find "$USERS_DIR" -maxdepth 1 -mindepth 1 -type d 2>/dev/null \
                | grep -ciEv '/(Public|Default|Default User|All Users)$' || true)
        echo -e "  ${YELLOW}[${IDX}]${RESET}  ${BOLD}${WOLD}${RESET}  ${YELLOW}★ Windows.old${RESET} ${DIM}(da ${BASE})${RESET}  ${CYAN}${USER_COUNT} utenti${RESET}"
        IDX=$((IDX + 1))
    done
    echo ""

    local CHOICE
    if [[ ${#ALL_PATHS[@]} -eq 1 ]]; then
        echo -ne "  ${YELLOW}[?]${RESET} Usare ${BOLD}${ALL_PATHS[0]}${RESET} come root Windows? [S/n]: "
        read -r CHOICE || true
        [[ "${CHOICE,,}" == "n" ]] && return 1
        _apply_win_root "${ALL_PATHS[0]}"; return 0
    fi
    echo -ne "  ${YELLOW}[?]${RESET} Seleziona numero, inserisci path manuale, o [N] per saltare: "
    read -r CHOICE || true
    case "${CHOICE,,}" in
        n|"") return 1 ;;
        [0-9]*)
            local SEL=$((CHOICE - 1))
            if [[ $SEL -ge 0 && $SEL -lt ${#ALL_PATHS[@]} ]]; then
                _apply_win_root "${ALL_PATHS[$SEL]}"; return 0
            else
                err "Selezione non valida"; return 1
            fi ;;
        *)
            local MP; MP=$(realpath -m "$CHOICE" 2>/dev/null || echo "$CHOICE")
            [[ ! -d "$MP" ]] && err "Directory non trovata: $MP" && return 1
            _apply_win_root "$MP"; return 0 ;;
    esac
}

# Imposta WIN_ROOT e innesca la raccolta informazioni
_apply_win_root() {
    local ROOT="$1"
    WIN_ROOT="$ROOT"
    ok "Root impostata: ${BOLD}$WIN_ROOT"

    # Recupera info macchina (hostname, OS, IP, dominio)
    gather_host_info
    # Resetta REPORT_BASE_DIR per ricalcolarla con il nuovo hostname
    REPORT_BASE_DIR=""
    setup_report_dir || true
}

# ================================================================
#  IMPOSTAZIONE MANUALE ROOT WINDOWS  (voce R del menu)
# ================================================================
set_win_root() {
    echo ""
    # Prima prova autodetect
    echo -e "  ${CYAN}[*]${RESET} Ricerca volumi Windows montati..."
    if autodetect_win_root; then
        return 0
    fi
    # Fallback: input manuale
    echo ""
    echo -ne "  ${YELLOW}[?]${RESET} Inserisci il path della root Windows (es. /mnt/windows): "
    read -r INPUT_ROOT
    [[ -z "$INPUT_ROOT" ]] && return 1
    INPUT_ROOT=$(realpath -m "$INPUT_ROOT" 2>/dev/null || echo "$INPUT_ROOT")
    if [[ ! -d "$INPUT_ROOT" ]]; then
        err "Directory non trovata: $INPUT_ROOT"
        return 1
    fi
    _apply_win_root "$INPUT_ROOT"
}

# ================================================================
#  SETUP DIRECTORY REPORT
# ================================================================
# Chiesta interattivamente la prima volta (REPORT_BASE_DIR vuota).
# Propone una directory di default, verifica i permessi di scrittura
# e aggiorna REPORT_BASE_DIR.
setup_report_dir() {
    local TS; TS=$(date +%Y%m%d_%H%M)
    local SUGGESTED_DEFAULT="${INVOCATION_DIR}/${HOST_NAME:-CASE}_fiuto_${TS}"
    echo ""
    echo -e "  ${CYAN}${BOLD}╔══════════════════════════════════════════════════════╗${RESET}"
    echo -e "  ${CYAN}${BOLD}║     Configurazione cartella di output dei report     ║${RESET}"
    echo -e "  ${CYAN}${BOLD}╚══════════════════════════════════════════════════════╝${RESET}"
    echo ""
    # Verifica se la directory suggerita è scrivibile
    local _PARENT_OK=false
    if [[ -d "$SUGGESTED_DEFAULT" ]]; then
        [[ -w "$SUGGESTED_DEFAULT" ]] && _PARENT_OK=true
    elif [[ -w "$(dirname "$SUGGESTED_DEFAULT")" ]]; then
        _PARENT_OK=true
    fi
    if $_PARENT_OK; then
        echo -e "  ${GREEN}[✓]${RESET} Directory suggerita: ${BOLD}${SUGGESTED_DEFAULT}${RESET}  ${GREEN}[scrivibile]${RESET}"
    else
        echo -e "  ${RED}[!]${RESET} Directory suggerita: ${BOLD}${SUGGESTED_DEFAULT}${RESET}  ${RED}[non scrivibile o parent protetto]${RESET}"
    fi
    echo ""
    echo -e "  ${DIM}I report di ogni modulo verranno salvati in sottocartelle con data/ora.${RESET}"
    echo -e "  ${DIM}Puoi inserire un percorso diverso oppure premere INVIO per usare quello suggerito.${RESET}"
    echo ""
    echo -ne "  ${YELLOW}[?]${RESET} Cartella report [${BOLD}${SUGGESTED_DEFAULT}${RESET}]: "
    local _INPUT
    read -r _INPUT
    local _CHOSEN
    if [[ -z "$_INPUT" ]]; then
        _CHOSEN="$SUGGESTED_DEFAULT"
    else
        _CHOSEN=$(realpath -m "$_INPUT" 2>/dev/null || echo "$_INPUT")
    fi
    # Verifica i permessi sulla scelta
    local _RW_OK=false _RW_MSG=""
    if [[ -d "$_CHOSEN" ]]; then
        if [[ -w "$_CHOSEN" ]]; then
            _RW_OK=true
            _RW_MSG="${GREEN}[scrivibile]${RESET}"
        else
            _RW_MSG="${RED}[SOLA LETTURA — i report NON potranno essere salvati!]${RESET}"
        fi
    else
        local _P; _P=$(dirname "$_CHOSEN")
        if [[ -w "$_P" ]]; then
            _RW_OK=true
            _RW_MSG="${GREEN}[verrà creata — parent scrivibile]${RESET}"
        else
            _RW_MSG="${RED}[parent '${_P}' NON scrivibile — i report NON potranno essere salvati!]${RESET}"
        fi
    fi
    echo ""
    echo -e "  ${CYAN}[→]${RESET} Cartella scelta: ${BOLD}${_CHOSEN}${RESET}  ${_RW_MSG}"
    echo ""
    if ! $_RW_OK; then
        warn "Attenzione: la directory selezionata non è scrivibile."
        echo -ne "  ${YELLOW}[?]${RESET} Vuoi comunque usarla? [s/N]: "
        local _CONF; read -r _CONF
        [[ "${_CONF,,}" != "s" ]] && { warn "Setup annullato. Ripeti con [P] dal menu."; return 1; }
    fi
    REPORT_BASE_DIR="$_CHOSEN"
    LOG_FILE="${REPORT_BASE_DIR}/fiuto_session_$(date +%Y%m%d_%H%M%S).log"
    log_msg "=== Log sessione inizializzato ==="
    ok "Log sessione: ${BOLD}$LOG_FILE"
    ok "Report dir impostata: ${BOLD}$REPORT_BASE_DIR"
    sleep 1
}

# ================================================================
#  MENU PRINCIPALE
# ================================================================
print_menu() {
    echo -e "  ${CYAN}${BOLD}╔══════════════════════════════════════════════════╗${RESET}"
    echo -e "  ${CYAN}${BOLD}║           F I U T O  —  SELEZIONA UN MODULO      ║${RESET}"
    echo -e "  ${CYAN}${BOLD}╚══════════════════════════════════════════════════╝${RESET}"
    echo ""
    # ── Cartella report (sempre visibile in cima) ──────────────────
    if [[ -n "$REPORT_BASE_DIR" ]]; then
        local _RW_LABEL _RW_COLOR
        if [[ -d "$REPORT_BASE_DIR" ]]; then
            if [[ -w "$REPORT_BASE_DIR" ]]; then
                _RW_LABEL="scrivibile"; _RW_COLOR="$GREEN"
            else
                _RW_LABEL="SOLA LETTURA"; _RW_COLOR="$RED"
            fi
        else
            local _RD_PARENT; _RD_PARENT=$(dirname "$REPORT_BASE_DIR")
            if [[ -w "$_RD_PARENT" ]]; then
                _RW_LABEL="OK (non ancora creata)"; _RW_COLOR="$GREEN"
            else
                _RW_LABEL="PARENT NON SCRIVIBILE"; _RW_COLOR="$RED"
            fi
        fi
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${DIM}${REPORT_BASE_DIR}${RESET}  ${_RW_COLOR}[${_RW_LABEL}]${RESET}"
    else
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${RED}non impostata — premi [P] per configurare${RESET}"
    fi
    echo -e "  ${WHITE}[R]${RESET}  ${BOLD}Imposta root Windows${RESET}      ${DIM}${WIN_ROOT:-(non impostata)}${RESET}"
    echo -e "  ${YELLOW}[D]${RESET}  ${BOLD}Debug mount attivi${RESET}        ${DIM}Diagnostica volumi montati${RESET}"
    echo ""
    echo -e "  ${MAGENTA}[1]${RESET}  PowerShell History            ${DIM}PSReadLine *_history.txt${RESET}"
    echo -e "  ${MAGENTA}[2]${RESET}  Notepad TabState              ${DIM}Tab rimasti aperti (.bin)${RESET}"
    echo -e "  ${RED}[3]${RESET}  IFEO Hijacking                ${DIM}Image File Execution Options${RESET}"
    echo -e "  ${BLUE}[4]${RESET}  BAM                           ${DIM}Background Activity Moderator${RESET}"
    echo -e "  ${ORANGE}[5]${RESET}  Run Keys & Persistenza        ${DIM}Autorun nel registro${RESET}"
    echo -e "  ${GREEN}[6]${RESET}  Prefetch                      ${DIM}Eseguibili tracciati (*.pf)${RESET}"
    echo -e "  ${YELLOW}[7]${RESET}  Scheduled Tasks               ${DIM}Task pianificati (XML)${RESET}"
    echo -e "  ${BLUE}[8]${RESET}  USB Devices                   ${DIM}Dispositivi rimovibili (USBSTOR)${RESET}"
    echo -e "  ${GREEN}[9]${RESET}  LNK & JumpList                ${DIM}File recenti e target path${RESET}"
    echo -e "  ${CYAN}[10]${RESET} Cache RDP                     ${DIM}Terminal Server Client Cache${RESET}"
  echo -e "  ${RED}[11]${RESET} Services                      ${DIM}Servizi Windows (SYSTEM hive)${RESET}"
  echo -e "  ${RED}[12]${RESET} Event Log                     ${DIM}Security/System/PS/RDP (.evtx)${RESET}"
  echo -e "  ${YELLOW}[13]${RESET} Amcache + Shimcache           ${DIM}Timeline esecuzione binari${RESET}"
  echo -e "  ${GREEN}[14]${RESET} Recycle Bin                   ${DIM}File eliminati (\$Recycle.Bin)${RESET}"
  echo -e "  ${RED}[15]${RESET} WMI Subscriptions             ${DIM}Persistenza invisibile (T1546.003)${RESET}"
  echo -e "  ${BLUE}[16]${RESET} SRUM                          ${DIM}Uso risorse per applicazione${RESET}"
  echo -e "  ${CYAN}[17]${RESET} Browser History               ${DIM}Chrome / Edge / Firefox${RESET}"
  echo -e "  ${MAGENTA}[18]${RESET} UserAssist / RunMRU           ${DIM}Attività interattiva utente${RESET}"
  echo -e "  ${CYAN}[19]${RESET} ShellBags                     ${DIM}Navigazione cartelle (anche cancellate)${RESET}"
  echo -e "  ${RED}[20]${RESET} SAM — Hash Locali             ${DIM}Hash NTLM account (impacket)${RESET}"
  echo -e "  ${YELLOW}[21]${RESET} MFT Timeline                  ${DIM}Master File Table + timestomping${RESET}"
  echo -e "  ${GREEN}[22]${RESET} OpenSave / LastVisited MRU    ${DIM}File aperti/salvati via dialogo${RESET}"
  echo -e "  ${CYAN}[23]${RESET} USN Journal                   ${DIM}\$UsnJrnl:\$J — change log NTFS${RESET}"
  echo -e "  ${RED}[24]${RESET} NTDS.dit                      ${DIM}Active Directory hash (DC offline)${RESET}"
  echo -e "  ${BLUE}[25]${RESET} Hibernation / Pagefile        ${DIM}hiberfil.sys · pagefile.sys strings${RESET}"
  echo -e "  ${RED}[26]${RESET} WER Files (Error Reports)     ${DIM}ReportArchive · ReportQueue (.wer)${RESET}"
  echo -e "  ${MAGENTA}[27]${RESET} Credential Manager           ${DIM}DPAPI blob offline${RESET}"
  echo -e "  ${CYAN}[28]${RESET} WLAN & VPN Profiles           ${DIM}WiFi · NetworkList · VPN${RESET}"
  echo -e "  ${GREEN}[29]${RESET} AppX / UWP Packages           ${DIM}App Store + sideload sospetti${RESET}"
  echo -e "  ${CYAN}[30]${RESET} Browser Downloads & Logins   ${DIM}Download + Login Data (DPAPI)${RESET}"
  echo -e "  ${YELLOW}[31]${RESET} Clipboard History             ${DIM}Cronologia appunti Win10+${RESET}"
  echo -e "  ${GREEN}[32]${RESET} Office MRU                   ${DIM}File recenti Word/Excel/PowerPoint${RESET}"
  echo -e "  ${RED}[33]${RESET} Defender Quarantine          ${DIM}File in quarantena + threatname${RESET}"
  echo -e "  ${MAGENTA}[34]${RESET} PS ScriptBlock Logging        ${DIM}Event ID 4104 — PS Operational.evtx${RESET}"
  echo -e "  ${GREEN}[35]${RESET} JumpLists                     ${DIM}AutomaticDestinations · CustomDestinations${RESET}"
  echo -e "  ${CYAN}[36]${RESET} Network Artifacts             ${DIM}Profili rete · Interfacce TCP/IP (registry)${RESET}"
  echo -e "  ${YELLOW}[37]${RESET} Master Timeline               ${DIM}Aggregazione cross-moduli con filtri${RESET}"
  echo -e "  ${RED}[38]${RESET} PAD Offline AD Analysis      ${DIM}NTDS.dit offline — utenti privilegiati, ACL, GPO${RESET}"
    echo ""
    echo -e "  ${WHITE}${BOLD}[0]${RESET}  ${BOLD}Esegui TUTTI i moduli${RESET}"
    echo ""
    # Report generati nella sessione corrente
    if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
        echo -e "  ${DIM}── Report generati (${#GENERATED_REPORTS[@]}) ──────────────────────────${RESET}"
        for _R in "${GENERATED_REPORTS[@]}"; do
            echo -e "  ${CYAN}↳${RESET} ${DIM}${_R}${RESET}"
        done
        echo ""
    fi
    echo -e "  ${RED}[Q]  Esci${RESET}"
    echo ""
    echo -ne "  ${YELLOW}Scelta:${RESET} "
}

# ================================================================
#  HELPER FUNCTIONS PER MODALITÀ NON INTERATTIVA
# ================================================================

# Espande una stringa tipo "1,3,5-8,12" in una lista di numeri
expand_module_list() {
    local spec="$1"
    local result=()
    IFS=',' read -ra PARTS <<< "$spec"
    for P in "${PARTS[@]}"; do
        if [[ "$P" =~ ^([0-9]+)-([0-9]+)$ ]]; then
            local from="${BASH_REMATCH[1]}" to="${BASH_REMATCH[2]}"
            for (( n=from; n<=to; n++ )); do result+=("$n"); done
        elif [[ "$P" =~ ^[0-9]+$ ]]; then
            result+=("$P")
        fi
    done
    printf '%s\n' "${result[@]}" | sort -n -u
}

run_module_by_number() {
    case "$1" in
        1)  module_ps_history ;;
        2)  module_notepad_tabstate ;;
        3)  module_ifeo ;;
        4)  module_bam ;;
        5)  module_run_keys ;;
        6)  module_prefetch ;;
        7)  module_scheduled_tasks ;;
        8)  module_usb ;;
        9)  module_lnk ;;
        10) module_rdp_cache ;;
        11) module_services ;;
        12) module_evtx ;;
        13) module_amcache ;;
        14) module_recycle_bin ;;
        15) module_wmi ;;
        16) module_srum ;;
        17) module_browser ;;
        18) module_userassist ;;
        19) module_shellbags ;;
        20) module_sam ;;
        21) module_mft ;;
        22) module_opensave ;;
        23) module_usn ;;
        24) module_ntds ;;
        25) module_hiberfil ;;
        26) module_wer_files ;;
        27) module_credential_manager ;;
        28) module_wlan ;;
        29) module_appx ;;
        30) module_browser_extra ;;
        31) module_clipboard ;;
        32) module_office_mru ;;
        33) module_defender_quarantine ;;
        34) module_ps_scriptblock ;;
        35) module_jumplists ;;
        36) module_network_artifacts ;;
        37) module_master_timeline ;;
        38) module_pad_offline ;;
        *)  err "Modulo sconosciuto: $1" ;;
    esac
}

# ================================================================
#  MAIN
# ================================================================
main() {
    SCAN_DATE=$(date "+%d/%m/%Y %H:%M:%S")

    # Parsing argomenti
    local ARG_ROOT=""
    local ARG_MODULE=""
    local ARG_ALL=false
    local ARG_MODULES=""
    local ARG_IOC=""

    while [[ $# -gt 0 ]]; do
        case "$1" in
            -h|--help)
                echo ""
                echo -e "${CYAN}${BOLD}fiuto.sh${RESET} — Toolkit DFIR per analisi di disco Windows offline"
                echo ""
                echo -e "  ${BOLD}Uso:${RESET}"
                echo -e "    ./fiuto.sh                            # menu interattivo"
                echo -e "    ./fiuto.sh /mnt/windows               # imposta root e apre il menu"
                echo -e "    ./fiuto.sh /mnt/windows --all         # esegui tutti i moduli"
                echo -e "    ./fiuto.sh /mnt/windows --module 3    # esegui modulo specifico"
                echo -e "    ./fiuto.sh /mnt/windows --modules 1,3,5-8  # esegui selezione"
                echo -e "    ./fiuto.sh /mnt/windows --all --ioc /path/to/ioc.txt  # con IoC"
                echo ""
                echo -e "  ${BOLD}Moduli disponibili (1-38):${RESET}"
                echo -e "    1  PowerShell History        2  Notepad TabState"
                echo -e "    3  IFEO Hijacking            4  BAM"
                echo -e "    5  Run Keys                  6  Prefetch"
                echo -e "    7  Scheduled Tasks           8  USB Devices"
                echo -e "    9  LNK Files                10  RDP Cache"
                echo -e "   11  Services                 12  Event Log (EVTX)"
                echo -e "   13  Amcache                  14  Recycle Bin"
                echo -e "   15  WMI Subscriptions        16  SRUM"
                echo -e "   17  Browser History          18  UserAssist"
                echo -e "   19  ShellBags                20  SAM"
                echo -e "   21  MFT Timeline             22  OpenSave MRU"
                echo -e "   23  USN Journal              24  NTDS.dit"
                echo -e "   25  Hibernation/Pagefile     26  WER Files (Error Reports)"
                echo -e "   27  Credential Manager       28  WLAN Profiles"
                echo -e "   29  AppX / UWP               30  Browser Downloads"
                echo -e "   31  Clipboard History        32  Office MRU"
                echo -e "   33  Defender Quarantine      34  PS ScriptBlock Log"
                echo -e "   35  JumpLists                36  Network Artifacts"
                echo -e "   37  Master Timeline          38  PAD Offline AD"
                echo ""
                exit 0
                ;;
            --all)       ARG_ALL=true ;;
            --module)    ARG_MODULE="$2"; shift ;;
            --modules)   ARG_MODULES="$2"; shift ;;
            --ioc)       ARG_IOC="$2"; shift ;;
            -*)          warn "Opzione sconosciuta: $1" ;;
            *)           [[ -z "$ARG_ROOT" ]] && ARG_ROOT="$1" ;;
        esac
        shift
    done

    if [[ -n "$ARG_ROOT" ]]; then
        if [[ ! -d "$ARG_ROOT" ]]; then
            err "Directory non trovata: $ARG_ROOT"; exit 1
        fi
        _apply_win_root "$ARG_ROOT"
    fi

    # Modalità non interattiva
    if $ARG_ALL; then
        print_banner
        [[ -z "$WIN_ROOT" ]] && { err "Specifica la root: $0 /mnt/windows --all"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        run_all_modules
        exit 0
    fi
    if [[ -n "$ARG_MODULE" ]]; then
        [[ -z "$WIN_ROOT" ]] && { err "Specifica la root: $0 /mnt/windows --module N"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        case "$ARG_MODULE" in
            1)  module_ps_history ;;
            2)  module_notepad_tabstate ;;
            3)  module_ifeo ;;
            4)  module_bam ;;
            5)  module_run_keys ;;
            6)  module_prefetch ;;
            7)  module_scheduled_tasks ;;
            8)  module_usb ;;
            9)  module_lnk ;;
            10) module_rdp_cache ;;
            11) module_services ;;
            12) module_evtx ;;
            13) module_amcache ;;
            14) module_recycle_bin ;;
            15) module_wmi ;;
            16) module_srum ;;
            17) module_browser ;;
            18) module_userassist ;;
            19) module_shellbags ;;
            20) module_sam ;;
            21) module_mft ;;
            22) module_opensave ;;
            23) module_usn ;;
            24) module_ntds ;;
            25) module_hiberfil ;;
            26) module_wer_files ;;
            27) module_credential_manager ;;
            28) module_wlan ;;
            29) module_appx ;;
            30) module_browser_extra ;;
            31) module_clipboard ;;
            32) module_office_mru ;;
            33) module_defender_quarantine ;;
            34) module_ps_scriptblock ;;
            35) module_jumplists ;;
            36) module_network_artifacts ;;
            37) module_master_timeline ;;
            38) module_pad_offline ;;
            *)  err "Modulo sconosciuto: $ARG_MODULE" ;;
        esac
        exit 0
    fi
    if [[ -n "$ARG_MODULES" ]]; then
        [[ -z "$WIN_ROOT" ]] && { err "Specifica la root: $0 /mnt/windows --modules $ARG_MODULES"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        local MOD_NUMS
        mapfile -t MOD_NUMS < <(expand_module_list "$ARG_MODULES")
        for N in "${MOD_NUMS[@]}"; do
            run_module_by_number "$N"
        done
        exit 0
    fi

    # Carica IoC se specificato (per --all o modalità interattiva)
    [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"

    # Modalità interattiva — chiedi prima ROOT, poi REPORT dir
    if [[ -z "$WIN_ROOT" ]]; then
        print_banner
        echo -e "  ${CYAN}[*]${RESET} Ricerca automatica di volumi Windows montati..."
        if ! autodetect_win_root silent; then
            echo ""
            warn "Nessun volume Windows rilevato automaticamente."
            echo ""
            echo -ne "  ${YELLOW}[?]${RESET} Inserisci il path della root Windows (o INVIO per saltare): "
            local _MR; read -r _MR || true
            if [[ -n "$_MR" ]]; then
                _MR=$(realpath -m "$_MR" 2>/dev/null || echo "$_MR")
                if [[ -d "$_MR" ]]; then
                    _apply_win_root "$_MR"
                else
                    err "Directory non trovata: $_MR"
                    info "Usa [R] dal menu per impostare la root."
                    sleep 2
                fi
            else
                info "Root non impostata. Usa [R] dal menu."
                sleep 1
            fi
        else
            echo ""
            sleep 1
        fi
    fi

    # Setup directory report (fallback se ROOT non è stata impostata)
    if [[ -z "$REPORT_BASE_DIR" ]]; then
        print_banner
        setup_report_dir || true
    fi

    # Modalità interattiva
    while true; do
        print_banner
        print_menu
        read -r CHOICE
        echo ""

        case "${CHOICE^^}" in
            P)  setup_report_dir || true; sleep 1 ;;
            R)  set_win_root; sleep 1 ;;
            D)  debug_mounts ;;
            1)  module_ps_history; return_to_menu ;;
            2)  module_notepad_tabstate; return_to_menu ;;
            3)  module_ifeo; return_to_menu ;;
            4)  module_bam; return_to_menu ;;
            5)  module_run_keys; return_to_menu ;;
            6)  module_prefetch; return_to_menu ;;
            7)  module_scheduled_tasks; return_to_menu ;;
            8)  module_usb; return_to_menu ;;
            9)  module_lnk; return_to_menu ;;
            10) module_rdp_cache; return_to_menu ;;
            11) module_services; return_to_menu ;;
            12) module_evtx; return_to_menu ;;
            13) module_amcache; return_to_menu ;;
            14) module_recycle_bin; return_to_menu ;;
            15) module_wmi; return_to_menu ;;
            16) module_srum; return_to_menu ;;
            17) module_browser; return_to_menu ;;
            18) module_userassist; return_to_menu ;;
            19) module_shellbags; return_to_menu ;;
            20) module_sam; return_to_menu ;;
            21) module_mft; return_to_menu ;;
            22) module_opensave; return_to_menu ;;
            23) module_usn; return_to_menu ;;
            24) module_ntds; return_to_menu ;;
            25) module_hiberfil; return_to_menu ;;
            26) module_wer_files; return_to_menu ;;
            27) module_credential_manager; return_to_menu ;;
            28) module_wlan; return_to_menu ;;
            29) module_appx; return_to_menu ;;
            30) module_browser_extra; return_to_menu ;;
            31) module_clipboard; return_to_menu ;;
            32) module_office_mru; return_to_menu ;;
            33) module_defender_quarantine; return_to_menu ;;
            34) module_ps_scriptblock; return_to_menu ;;
            35) module_jumplists; return_to_menu ;;
            36) module_network_artifacts; return_to_menu ;;
            37) module_master_timeline; return_to_menu ;;
            38) module_pad_offline; return_to_menu ;;
            0)  run_all_modules
                return_to_menu ;;
            Q)  echo ""
                if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
                    clear
                    echo -e "${CYAN}${BOLD}"
                    echo "  ╔══════════════════════════════════════════════════════════╗"
                    echo "  ║         FIUTO — Report generati in questa sessione       ║"
                    echo "  ╚══════════════════════════════════════════════════════════╝"
                    echo -e "${RESET}"
                    _IDX=0
                    for _R in "${GENERATED_REPORTS[@]}"; do
                        _IDX=$(( _IDX + 1 ))
                        if [[ -f "$_R" ]]; then
                            _RSIZE=$(du -h "$_R" 2>/dev/null | cut -f1 || echo "?")
                            _RMOD=$(basename "$(dirname "$_R")")
                            printf "  ${CYAN}%2d${RESET}  ${BOLD}%-30s${RESET}\n" "$_IDX" "$_RMOD"
                            printf "      ${DIM}%s${RESET}  ${DIM}(%s)${RESET}\n" "$_R" "$_RSIZE"
                            echo ""
                        fi
                    done
                    echo -e "  ${DIM}Apri con: xdg-open \"<percorso>\"${RESET}"
                    echo ""
                fi
                echo -e "  ${DIM}Uscita.${RESET}"; echo ""; exit 0 ;;
            *)  warn "Scelta non valida: '$CHOICE'"; sleep 1 ;;
        esac
    done
}

main "$@"
