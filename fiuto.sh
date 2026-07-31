#!/usr/bin/env bash
# ================================================================
#  fiuto.sh  –  v2.3
#  Toolkit DFIR unificato per analisi offline di dischi Windows/Linux/macOS
#
#  Uso:
#    ./fiuto.sh                        # menu interattivo
#    ./fiuto.sh /mnt/disk              # specifica la root (Windows/Linux/macOS)
#    ./fiuto.sh /mnt/disk --all        # esegui tutto
#    ./fiuto.sh /mnt/disk --module 3   # modulo specifico
#    ./fiuto.sh --image disco.E01 --partition 2 --all   # da immagine forense
# ================================================================
#
#  ATTENZIONE — QUESTO FILE E' GENERATO. NON MODIFICARLO A MANO.
#
#  Il sorgente sta in src/, diviso per sistema operativo:
#    src/lib/                funzioni condivise (core, fs, report, registro...)
#    src/modules/win/        moduli Windows
#    src/modules/linux/      moduli Linux
#    src/modules/macos/      moduli macOS
#    src/modules/xplat/      moduli cross-OS
#
#  Dopo aver modificato un sorgente rigenera questo file con:
#    ./build.sh
#
#  La CI verifica con ./build.sh --check che i due siano allineati: una
#  modifica fatta qui e non nei sorgenti verrebbe persa al build successivo.
# ================================================================

set -uo pipefail

# Versione del tool. Unica fonte: la usano il banner e il manifesto di catena
# di custodia, che non devono poter divergere.
FIUTO_VERSION="2.3"

# -- Cleanup globale file temporanei (signal-safe) ----------------
# La trap EXIT viene invocata su ritorno normale, SIGINT (Ctrl+C),
# SIGTERM, SIGHUP e errori.
# SIGKILL (kill -9) non e' intercettabile per design del kernel.
declare -a _GLOBAL_TMP_FILES=()
_global_cleanup() {
    # Il manifesto di catena di custodia va scritto prima di rimuovere i file
    # temporanei: lo stato della sessione vive li'.
    declare -F write_evidence_manifest > /dev/null && write_evidence_manifest 2>/dev/null || true
    # Poi si smonta: un rm -rf su una directory che e' ancora un punto di mount
    # cancellerebbe dentro l'immagine invece che attorno.
    declare -F image_cleanup > /dev/null && image_cleanup 2>/dev/null || true
    local _f
    for _f in "${_GLOBAL_TMP_FILES[@]:-}"; do
        [[ -e "$_f" ]] && rm -rf "$_f" 2>/dev/null || true
    done
}
trap _global_cleanup EXIT

# Helper: registra un file/dir temporaneo per cleanup automatico
register_tmp() { _GLOBAL_TMP_FILES+=("$1"); }

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
WIN_ROOT=""          # root del volume selezionato (Windows/Linux/macOS — es. /mnt/disk)
OS_TYPE=""           # tipo OS del volume: windows | linux | macos | unknown
REPORT_BASE_DIR=""   # directory base dei report HTML
SCAN_DATE=""         # impostata all'avvio
declare -a GENERATED_REPORTS=()  # lista dei report HTML generati nella sessione
BATCH_MODE=false             # true in run_all_modules
declare -a SUMMARY_TABLE=()  # tabella riepilogativa batch
LOG_FILE=""                  # percorso log sessione (impostato all'avvio)
declare -a IOC_LIST=()       # lista IoC da file esterno (--ioc)
HIVE_REPLAY=true             # replay dei transaction log del registro (--no-log-replay per disattivarlo)
EXPORT_JSONL=false           # export JSONL schema Timesketch accanto a ogni report (--jsonl)
YARA_RULES=""                # file o directory di regole YARA (--yara)
YARA_SCAN_PATH=""            # ambito alternativo da scansionare (--yara-scan)
YARA_MAX_MB=64               # tetto per file: oltre, il file viene dichiarato saltato
YARA_MAX_FILES=200000        # tetto complessivo: oltre, la scansione si dichiara parziale
SIGMA_RULES=""               # file o directory di regole Sigma (--sigma)
SIGMA_MAX_RECORDS=300000     # tetto sui record EVTX letti: oltre, valutazione parziale dichiarata
JOBS=1                       # moduli eseguiti in parallelo con --all (--jobs N)
REPORT_DIR_FIXED=""          # --report-dir: cartella dei report scelta da riga di comando
# Lo stato del replay (cache, esiti, avvisi gia' emessi) vive su disco in
# ${TMPDIR:-/tmp}/fiuto_hives_$$ e non in variabili: recover_hive gira quasi
# sempre dentro una command substitution, quindi in subshell.
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
#  LANGUAGE SUPPORT / SUPPORTO LINGUE
# ================================================================

export LANG="en"  # Default: English / Default: Italiano (EXPORTED for persistence)

# L "testo italiano" "english text" — restituisce il testo nella lingua corrente
L() { [[ "${LANG:-en}" == "it" ]] && printf '%s' "$1" || printf '%s' "$2"; }

# Funzione per selezionare la lingua all'avvio
select_language() {
    echo ""
    echo -e "${CYAN}${BOLD}Select Language / Selezionare Lingua${RESET}"
    echo -e "${CYAN}${BOLD}═══════════════════════════════════════${RESET}"
    echo ""
    echo "  1) English"
    echo "  2) Italiano"
    echo ""
    echo -ne "${YELLOW}[?]${RESET} Choose / Scegli [1-2, default=1]: "
    read -r LANG_CHOICE
    if [[ "$LANG_CHOICE" == "2" ]]; then
        export LANG="it"
    else
        export LANG="en"
    fi
    echo ""
}

# Funzione di traduzione per stringhe
t() {
    local KEY="$1"
    case "$LANG:$KEY" in
        # === BANNER & SETUP ===
        "en:root_not_set") echo "Windows root not set. Use option [R] from menu." ;;
        "it:root_not_set") echo "Root Windows non impostata. Usa l'opzione [R] dal menu." ;;

        "en:root_invalid") echo "Directory Users or Windows not found in:" ;;
        "it:root_invalid") echo "Directory Users o Windows non trovata in:" ;;

        "en:hostname_prompt") echo "Enter a name for the case/machine" ;;
        "it:hostname_prompt") echo "Inserisci un nome per il caso/macchina" ;;

        "en:machine_info") echo "TARGET MACHINE INFORMATION" ;;
        "it:machine_info") echo "INFORMAZIONI MACCHINA TARGET" ;;

        "en:hostname") echo "Hostname" ;;
        "it:hostname") echo "Hostname" ;;

        "en:os") echo "Operating System" ;;
        "it:os") echo "Sistema Operativo" ;;

        "en:ip") echo "IP Address" ;;
        "it:ip") echo "Indirizzo IP" ;;

        "en:domain") echo "Domain" ;;
        "it:domain") echo "Dominio" ;;

        # === MESSAGES ===
        "en:press_key") echo "Press any key to return to menu..." ;;
        "it:press_key") echo "Premi qualsiasi tasto per tornare al menu..." ;;

        "en:open_browser") echo "Open report in browser?" ;;
        "it:open_browser") echo "Aprire il report nel browser?" ;;

        "en:generate_html") echo "Generate HTML report?" ;;
        "it:generate_html") echo "Generare report HTML?" ;;

        "en:ioc_not_found") echo "IoC file not found:" ;;
        "it:ioc_not_found") echo "File IoC non trovato:" ;;

        "en:ioc_loaded") echo "IoC loaded:" ;;
        "it:ioc_loaded") echo "IoC caricati:" ;;

        "en:hive_not_found") echo "Registry hives not found, unable to retrieve machine info automatically." ;;
        "it:hive_not_found") echo "Hive di registro non trovati, impossibile recuperare info macchina automaticamente." ;;

        "en:retrieving_info") echo "Retrieving machine information..." ;;
        "it:retrieving_info") echo "Recupero informazioni macchina in corso..." ;;

        # === MAIN MENU & BATCH ===
        "en:batch_running") echo "Running all modules in batch mode..." ;;
        "it:batch_running") echo "Esecuzione di tutti i moduli in modalità batch..." ;;

        "en:batch_report_dir") echo "Report base directory:" ;;
        "it:batch_report_dir") echo "Report base dir:" ;;

        "en:batch_started") echo "=== FIUTO started in batch mode — WIN_ROOT=" ;;
        "it:batch_started") echo "=== FIUTO avviato in modalità batch — WIN_ROOT=" ;;

        "en:specify_root_all") echo "Specify the root: $0 /mnt/windows --all" ;;
        "it:specify_root_all") echo "Specifica la root: $0 /mnt/windows --all" ;;

        "en:specify_root_module") echo "Specify the root: $0 /mnt/windows --module N" ;;
        "it:specify_root_module") echo "Specifica la root: $0 /mnt/windows --module N" ;;

        "en:dir_not_found") echo "Directory not found:" ;;
        "it:dir_not_found") echo "Directory non trovata:" ;;

        "en:searching_mounts") echo "Automatically searching for mounted Windows volumes..." ;;
        "it:searching_mounts") echo "Ricerca automatica di volumi Windows montati..." ;;

        "en:no_windows_found") echo "No Windows volume detected automatically." ;;
        "it:no_windows_found") echo "Nessun volume Windows rilevato automaticamente." ;;

        "en:enter_windows_root") echo "Enter the Windows root path (or ENTER to skip):" ;;
        "it:enter_windows_root") echo "Inserisci il path della root Windows (o INVIO per saltare):" ;;

        "en:unable_detect_hostname") echo "Unable to detect machine name automatically." ;;
        "it:unable_detect_hostname") echo "Impossibile rilevare il nome macchina automaticamente." ;;

        *) echo "$KEY" ;;
    esac
}

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
    echo -e "  ║    ${CYAN}${BOLD}F${RESET}${CYAN}orensic ${BOLD}I${RESET}${CYAN}nvestigation ${BOLD}U${RESET}${CYAN}tility ${BOLD}T${RESET}${CYAN}ool ${BOLD}O${RESET}${CYAN}ffline${RESET}           ${CYAN}${BOLD}║"
    echo -e "  ║                    ${MAGENTA}${BOLD}v${FIUTO_VERSION} - zi®iginal${RESET}${CYAN}                      ║"
    echo "  ╚══════════════════════════════════════════════════════════╝"
    echo -e "${RESET}"
    local DATE_LABEL="$([ "$LANG" = "it" ] && echo "Data" || echo "Date")"
    local ROOT_LABEL="$([ "$LANG" = "it" ] && echo "Root" || echo "Root")"
    local PYTHON_LABEL="$([ "$LANG" = "it" ] && echo "Python" || echo "Python")"
    local REPORT_LABEL="$([ "$LANG" = "it" ] && echo "Report" || echo "Report")"
    local NOT_SET_LABEL="$([ "$LANG" = "it" ] && echo "non impostata" || echo "not set")"
    local WRITABLE_LABEL="$([ "$LANG" = "it" ] && echo "scrivibile" || echo "writable")"
    local READONLY_LABEL="$([ "$LANG" = "it" ] && echo "sola lettura!" || echo "read-only!")"
    local CREATE_OK_LABEL="$([ "$LANG" = "it" ] && echo "creazione OK" || echo "creation OK")"
    local PARENT_NOT_OK_LABEL="$([ "$LANG" = "it" ] && echo "parent non scrivibile!" || echo "parent not writable!")"

    echo -e "  ${DIM}${DATE_LABEL}:   $(date '+%d/%m/%Y %H:%M:%S')${RESET}"
    echo -ne "  ${DIM}${ROOT_LABEL}:   ${WIN_ROOT:-$NOT_SET_LABEL}${RESET}"
    [[ -n "$HOST_NAME" ]] && echo -ne "  ${CYAN}${BOLD}[${HOST_NAME}]${RESET}"
    echo ""
    echo -e "  ${DIM}${PYTHON_LABEL}: ${PY3} (${PY3_VERSION})${RESET}"
    if [[ -n "$REPORT_BASE_DIR" ]]; then
        local _RD_INFO=""
        if [[ -d "$REPORT_BASE_DIR" ]]; then
            if [[ -w "$REPORT_BASE_DIR" ]]; then
                _RD_INFO="${GREEN}[${WRITABLE_LABEL}]${RESET}"
            else
                _RD_INFO="${RED}[${READONLY_LABEL}]${RESET}"
            fi
        else
            # non ancora creata: verifichiamo il parent
            local _RD_PARENT; _RD_PARENT=$(dirname "$REPORT_BASE_DIR")
            if [[ -w "$_RD_PARENT" ]]; then
                _RD_INFO="${GREEN}[${CREATE_OK_LABEL}]${RESET}"
            else
                _RD_INFO="${RED}[${PARENT_NOT_OK_LABEL}]${RESET}"
            fi
        fi
        echo -e "  ${DIM}${REPORT_LABEL}: ${BOLD}${REPORT_BASE_DIR}${RESET}  ${_RD_INFO}"
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
    echo -ne "  ${YELLOW}$(t press_key)${RESET}"
    pause_key
}

# Comando per aprire un file con l'applicazione predefinita.
#
# xdg-open e' di freedesktop e su macOS non esiste: li' si chiama open. Il
# rilevamento e' a runtime e non nello strato di compatibilita' perche' cosi'
# vale anche per la build Linux eseguita su un Mac, e per i desktop che hanno
# gio ma non xdg-open. Se non c'e' niente, si dice: aprire un file e' un
# comodo, non un requisito, ma un comando che fallisce in silenzio lascia
# l'utente a chiedersi perche' non succede nulla.
report_opener() {
    if [[ "$(uname -s 2>/dev/null)" == "Darwin" ]] && command -v open > /dev/null 2>&1; then
        echo "open"; return 0
    fi
    local C
    for C in xdg-open gio open; do
        command -v "$C" > /dev/null 2>&1 && { echo "$C"; return 0; }
    done
    return 1
}

# Chiede all'utente se aprire il report nel browser.
# In BATCH_MODE non apre e non chiede (nessun utente interattivo disponibile).
open_report_prompt() {
    local RPATH="$1"
    [[ "${BATCH_MODE:-false}" == "true" ]] && return 0
    local RESP
    local YES_LABEL="$([ "$LANG" = "it" ] && echo "S" || echo "Y")"
    local NO_LABEL="$([ "$LANG" = "it" ] && echo "n" || echo "n")"
    echo -ne "  ${YELLOW}[?]${RESET} $(t open_browser) [${YES_LABEL}/${NO_LABEL}]: "
    read -r RESP
    [[ "${RESP,,}" == "n" ]] && return 0
    local OPENER
    if OPENER=$(report_opener); then
        "$OPENER" "$RPATH" > /dev/null 2>&1 &
    else
        warn "$(L "Nessun comando per aprire i file (xdg-open, open). Apri a mano:" \
                 "No command available to open files (xdg-open, open). Open it manually:") $RPATH"
    fi
}

# Info / warning / error
info()    { echo -e "  ${CYAN}[i]${RESET} $*"; log_msg "[INFO] $*"; }
ok()      { echo -e "  ${GREEN}[✓]${RESET} ${BOLD}$*${RESET}"; log_msg "[OK]   $*"; }
warn()    { echo -e "  ${YELLOW}[!]${RESET} $*"; log_msg "[WARN] $*"; }

err()     { echo -e "  ${RED}[✗]${RESET} $*"; log_msg "[ERR]  $*"; }
dim_msg() { echo -e "  ${DIM}[-] $*${RESET}"; log_msg "[DIM]  $*"; }

# Chiede conferma S/n, default S
ask_yn() {
    local PROMPT="$1"
    local YES_LABEL="$([ "$LANG" = "it" ] && echo "S" || echo "Y")"
    local NO_LABEL="$([ "$LANG" = "it" ] && echo "n" || echo "n")"
    local AUTO_LABEL="$([ "$LANG" = "it" ] && echo "auto" || echo "auto")"
    if [[ "${BATCH_MODE:-false}" == "true" ]]; then
        echo -e "  ${DIM}[${AUTO_LABEL}] ${PROMPT} → ${YES_LABEL}${RESET}"
        return 0
    fi
    local RESP
    echo -ne "  ${YELLOW}[?]${RESET} ${PROMPT} [${YES_LABEL}/${NO_LABEL}]: "
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

# Escape HTML — usare questa invece delle funzioni _esc_X locali nei moduli
# (per contenuto testuale tra i tag: basta neutralizzare & < >)
html_esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

# Escape HTML per valori dentro un ATTRIBUTO (title='...', data-*='...'):
# oltre a & < > neutralizza anche gli apici, che altrimenti chiuderebbero l'attributo.
html_attr() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g; s/"/\&quot;/g; s/'"'"'/\&#39;/g'; }

# Calcola SHA256 di un file (per chain of custody); ritorna stringa vuota se fallisce
sha256_file() { sha256sum "$1" 2>/dev/null | awk '{print $1}' || true; }

# Scrive un messaggio nel log di sessione (se LOG_FILE è impostato)
log_msg() {
    [[ -n "$LOG_FILE" ]] || return 0
    local _ld; _ld=$(dirname "$LOG_FILE")
    [[ -d "$_ld" ]] || mkdir -p "$_ld" 2>/dev/null || return 0
    printf '[%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$LOG_FILE" 2>/dev/null || true
}

# ================================================================
#  MOTORE IoC
#
#  Fino alla 2.2 il match era una sottostringa case-insensitive. Comodo, e
#  sbagliato in due direzioni opposte:
#
#  - FALSI POSITIVI. L'indirizzo 10.0.0.5 corrisponde dentro 110.0.0.55, dentro
#    un numero di build, dentro un offset esadecimale. Il dominio evil.com
#    corrisponde dentro notevil.com.au. In un report con centomila righe questo
#    non e' rumore fastidioso: e' rumore che nasconde i match veri.
#  - FALSI NEGATIVI. Gli IoC arrivano quasi sempre defanged — hxxp://,
#    1[.]2[.]3[.]4, evil[.]com — perche' e' cosi' che si scrivono in una mail o
#    in un bollettino. Cercati alla lettera non corrispondono a niente, e
#    l'analista conclude che il disco e' pulito.
#
#  Qui ogni indicatore viene TIPIZZATO al caricamento e confrontato con i
#  confini giusti per il suo tipo, dopo essere stato normalizzato. Si accettano
#  liste piatte, bundle STIX 2.x e export MISP.
#
#  IOC_LIST resta l'elenco dei valori normalizzati: i moduli che fanno il
#  proprio match in Python continuano a funzionare, e guadagnano comunque la
#  normalizzazione del defanging.
# ================================================================

declare -a IOC_TYPES=()     # "tipo|valore" per ogni indicatore caricato
IOC_REGEX=""                # ERE unica, minuscola, usata da check_ioc

# defang_value <stringa> — rende inerte un indicatore per la visualizzazione.
# Serve nei report: un URL cliccabile o copiabile per sbaglio in un browser e'
# un rischio che non ha ragione di esistere in un documento di analisi.
defang_value() {
    local V="$1"
    V="${V//http:/hxxp:}"; V="${V//https:/hxxps:}"
    V="${V//./[.]}"
    V="${V//@/[at]}"
    printf '%s' "$V"
}

# Estensioni che rendono un valore un nome di file e non un dominio. Senza
# questo elenco "payload.dll" verrebbe classificato dominio: sintatticamente lo
# e'. La riga di riepilogo per tipo stampata al caricamento serve proprio a far
# accorgere di una classificazione sbagliata; per i casi davvero ambigui
# (payload.com) esiste il prefisso esplicito "file:payload.com".
_IOC_FILE_EXT="|exe|dll|sys|bat|cmd|ps1|psm1|vbs|vbe|js|jse|wsf|scr|hta|jar|lnk|tmp|dat|bin|log|txt|ini|cfg|conf|zip|rar|7z|gz|tar|iso|img|doc|docx|xls|xlsx|ppt|pptx|pdf|rtf|one|php|asp|aspx|jsp|py|sh|pl|rb|dmp|pf|evtx|reg|crt|pem|key|"

# Riconosce il tipo di un indicatore gia' normalizzato.
_ioc_type_of() {
    local V="${1,,}"
    case "$V" in
        /*/) echo regex; return ;;
        http://*|https://*|ftp://*) echo url; return ;;
    esac
    if [[ "$V" =~ ^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}(/[0-9]{1,2})?$ ]]; then
        echo ipv4; return
    fi
    if [[ "$V" =~ ^[0-9a-f]{64}$ ]]; then echo sha256; return; fi
    if [[ "$V" =~ ^[0-9a-f]{40}$ ]]; then echo sha1; return; fi
    if [[ "$V" =~ ^[0-9a-f]{32}$ ]]; then echo md5; return; fi
    if [[ "$V" =~ ^[^@[:space:]]+@[^@[:space:]]+\.[a-z]{2,}$ ]]; then echo email; return; fi
    if [[ "$V" == *:*:* ]]; then echo ipv6; return; fi
    if [[ "$V" =~ ^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$ ]]; then
        [[ "$_IOC_FILE_EXT" == *"|${V##*.}|"* ]] && { echo literal; return; }
        echo domain; return
    fi
    # Nomi di file, chiavi di registro, nomi di mutex: nessun confine sensato
    # da imporre, si resta sulla sottostringa e lo si dichiara.
    echo literal
}

# Escapa i metacaratteri ERE di un valore letterale.
_ioc_quote() { printf '%s' "$1" | sed 's/[][\.^$*+?(){}|\\/]/\\&/g'; }

# ERE con i confini adatti al tipo. Il confronto avviene su testo minuscolo.
_ioc_regex_for() {
    local TYPE="$1" V="${2,,}" Q
    Q=$(_ioc_quote "$V")
    case "$TYPE" in
        regex)  printf '%s' "${V:1:${#V}-2}" ;;
        ipv4)   # niente cifre o punti attaccati: 10.0.0.5 non deve valere
                # dentro 110.0.0.55.
                printf '(^|[^0-9.])%s($|[^0-9.])' "$Q" ;;
        md5|sha1|sha256)
                printf '(^|[^0-9a-f])%s($|[^0-9a-f])' "$Q" ;;
        domain) # I sottodomini contano (mail.evil.com per evil.com), il
                # suffisso no: evil.com.altro.net non e' evil.com.
                printf '(^|[^a-z0-9._-])(([a-z0-9-]+\.)*)%s($|[^a-z0-9.-])' "$Q" ;;
        email)  printf '(^|[^a-z0-9._%%+-])%s($|[^a-z0-9.-])' "$Q" ;;
        *)      printf '%s' "$Q" ;;
    esac
}

# Normalizza un indicatore: toglie il defanging e gli spazi.
_ioc_normalize() {
    local V="$1"
    V="${V#"${V%%[![:space:]]*}"}"; V="${V%"${V##*[![:space:]]}"}"
    V="${V//hxxp/http}"; V="${V//hXXp/http}"; V="${V//HXXP/HTTP}"
    V="${V//\[.\]/.}"; V="${V//(.)/.}"; V="${V//\{.\}/.}"
    V="${V//\[:\]/:}"; V="${V//\[\/\]/\/}"
    V="${V//\[at\]/@}"; V="${V//\[@\]/@}"
    V="${V//\[dot\]/.}"
    printf '%s' "$V"
}

# Estrae type<TAB>value da un bundle STIX 2.x o da un export MISP.
_ioc_from_json() {
    "$PY3" - "$1" << 'PYEOF' 2>/dev/null
import sys, json, re

try:
    doc = json.load(open(sys.argv[1], encoding='utf-8', errors='replace'))
except Exception:
    sys.exit(1)

out = []

# STIX 2.x: il pattern e' un linguaggio a se'. Non lo si interpreta: se ne
# estraggono i confronti di uguaglianza, che coprono il caso reale degli
# indicatori distribuiti nei feed.
EQ = re.compile(r"([a-z0-9:._'\"-]+)\s*=\s*'([^']+)'", re.I)


def stix(obj):
    pat = obj.get('pattern') or ''
    for path, val in EQ.findall(pat):
        p = path.lower()
        if 'hashes' in p:
            out.append(('hash', val))
        elif 'ipv4' in p:
            out.append(('ipv4', val))
        elif 'ipv6' in p:
            out.append(('ipv6', val))
        elif 'domain-name' in p:
            out.append(('domain', val))
        elif 'url' in p:
            out.append(('url', val))
        elif 'email' in p:
            out.append(('email', val))
        elif 'file:name' in p or 'directory:path' in p:
            out.append(('literal', val))
        else:
            out.append(('literal', val))


def misp_attrs(container):
    for a in container.get('Attribute', []) or []:
        v = a.get('value')
        if v:
            out.append((a.get('type', 'literal'), v))
    for o in container.get('Object', []) or []:
        misp_attrs(o)


if isinstance(doc, dict) and doc.get('type') == 'bundle':
    for obj in doc.get('objects', []):
        if obj.get('type') == 'indicator':
            stix(obj)
elif isinstance(doc, dict) and ('Event' in doc or 'response' in doc):
    events = []
    if 'Event' in doc:
        events = [doc['Event']]
    else:
        for item in doc.get('response', []):
            if isinstance(item, dict) and 'Event' in item:
                events.append(item['Event'])
    for ev in events:
        misp_attrs(ev)
else:
    # JSON valido ma schema non riconosciuto. Codice diverso da "non e' JSON":
    # il chiamante deve fermarsi, non trattare il file come lista piatta —
    # finirebbe per caricare le graffe come indicatore letterale.
    sys.exit(2)

if not out:
    sys.exit(2)

for t, v in out:
    print('%s\t%s' % (t, v))
PYEOF
}

# Carica un file IoC: lista piatta (una entry per riga, # ignorate),
# bundle STIX 2.x o export MISP. Il formato viene riconosciuto dal contenuto.
load_ioc_file() {
    local IOCFILE="$1"
    if [[ ! -f "$IOCFILE" ]]; then
        warn "$(t ioc_not_found) $IOCFILE"
        return 1
    fi

    local SRC; SRC=$(mktemp); register_tmp "$SRC"
    local FORMAT="lista"
    if [[ "$(head -c 1 "$IOCFILE" 2>/dev/null)" == "{" ]]; then
        _ioc_from_json "$IOCFILE" > "$SRC"
        local RC=$?
        if [[ "$RC" -eq 2 ]]; then
            err "$(L "JSON riconosciuto ma non e' un bundle STIX 2.x ne' un export MISP:" \
                     "JSON recognised but it is neither a STIX 2.x bundle nor a MISP export:") $IOCFILE"
            info "$(L "Il file non viene caricato: interpretarlo come lista piatta trasformerebbe le righe JSON in indicatori letterali." \
                     "The file is not loaded: reading it as a flat list would turn JSON lines into literal indicators.")"
            return 1
        fi
        [[ "$RC" -eq 0 && -s "$SRC" ]] && FORMAT="STIX/MISP"
    fi
    if [[ "$FORMAT" == "lista" ]]; then
        # Lista piatta: il tipo si deduce dal valore.
        local LINE
        while IFS= read -r LINE || [[ -n "$LINE" ]]; do
            [[ -z "${LINE//[[:space:]]/}" || "$LINE" == \#* ]] && continue
            # Prefisso esplicito "tipo:valore" per i casi che l'euristica non
            # puo' risolvere (payload.com e' un nome di file o un dominio?).
            if [[ "$LINE" =~ ^(ipv4|ipv6|domain|url|email|md5|sha1|sha256|file|literal|regex):(.+)$ ]]; then
                printf '%s\t%s\n' "${BASH_REMATCH[1]}" "${BASH_REMATCH[2]}" >> "$SRC"
            else
                printf 'auto\t%s\n' "$LINE" >> "$SRC"
            fi
        done < "$IOCFILE"
    fi

    local NDUP=0
    local -A _seen=()
    local RAW_TYPE RAW_VAL VAL TYPE RE
    while IFS=$'\t' read -r RAW_TYPE RAW_VAL; do
        [[ -n "$RAW_VAL" ]] || continue
        VAL=$(_ioc_normalize "$RAW_VAL")
        [[ -n "$VAL" ]] || continue
        if [[ -n "${_seen[${VAL,,}]:-}" ]]; then NDUP=$((NDUP + 1)); continue; fi
        _seen[${VAL,,}]=1
        case "$RAW_TYPE" in
            auto|"")            TYPE=$(_ioc_type_of "$VAL") ;;
            file|filename)      TYPE=literal ;;
            ip-src|ip-dst|ip)   TYPE=$(_ioc_type_of "$VAL") ;;
            hash)               TYPE=$(_ioc_type_of "$VAL") ;;
            ipv4|ipv6|domain|url|email|md5|sha1|sha256|literal|regex) TYPE="$RAW_TYPE" ;;
            # MISP usa decine di tipi (hostname, uri, sha512, mutex...): quelli
            # che non conosciamo si fanno riconoscere dal valore, che e' piu'
            # affidabile di una mappatura incompleta.
            *)                  TYPE=$(_ioc_type_of "$VAL") ;;
        esac
        IOC_LIST+=("$VAL")
        IOC_TYPES+=("${TYPE}|${VAL}")
        RE=$(_ioc_regex_for "$TYPE" "$VAL")
        [[ -n "$RE" ]] && IOC_REGEX="${IOC_REGEX:+${IOC_REGEX}|}(${RE})"
    done < "$SRC"

    if [[ ${#IOC_LIST[@]} -eq 0 ]]; then
        warn "$(L "Nessun indicatore utilizzabile in" "No usable indicator in") $IOCFILE"
        return 1
    fi

    # Il conteggio per tipo dice subito se il file e' stato interpretato come
    # ci si aspettava: 40 "literal" al posto di 40 hash sono un errore di
    # formato, e senza questa riga si scoprirebbe solo dall'assenza di match.
    local BREAKDOWN; BREAKDOWN=$(printf '%s\n' "${IOC_TYPES[@]}" | cut -d'|' -f1 \
        | sort | uniq -c | awk '{printf "%s=%s ", $2, $1}')
    ok "$(t ioc_loaded) ${#IOC_LIST[@]} ($FORMAT) — ${BREAKDOWN}"
    [[ "$NDUP" -gt 0 ]] && info "$(L "Duplicati ignorati:" "Duplicates ignored:") $NDUP"
    log_msg "[IOC] $FORMAT: ${#IOC_LIST[@]} indicatori da $IOCFILE — $BREAKDOWN"
    return 0
}

# Vero se il testo contiene almeno un indicatore caricato.
#
# Puro bash, senza fork: viene chiamata riga per riga su report da decine di
# migliaia di righe, e un processo per riga renderebbe i moduli inutilizzabili.
check_ioc() {
    [[ -n "$IOC_REGEX" ]] || return 1
    local TEXT="${1,,}"
    [[ "$TEXT" =~ $IOC_REGEX ]]
}

# Il tipo e il valore dell'indicatore che corrisponde, per i report che vogliono
# dire QUALE indicatore ha fatto match e non solo che ce n'e' stato uno.
matched_ioc() {
    local TEXT="${1,,}" entry TYPE VAL RE
    for entry in "${IOC_TYPES[@]}"; do
        TYPE="${entry%%|*}"; VAL="${entry#*|}"
        RE=$(_ioc_regex_for "$TYPE" "$VAL")
        [[ -n "$RE" && "$TEXT" =~ $RE ]] && { printf '%s|%s' "$TYPE" "$VAL"; return 0; }
    done
    return 1
}

# ================================================================
#  TEMPO — timezone del volume e finestra di analisi
#
#  Due problemi distinti, spesso confusi.
#
#  1) IN CHE FUSO SONO LE DATE CHE LEGGO. Gli artefatti non concordano: gli
#     hive di registro e i log Windows portano quasi sempre UTC, syslog e le
#     shell history portano l'ora locale della macchina, i database SQLite dei
#     browser dipendono dal browser. Convertire tutto a UTC richiederebbe di
#     sapere, artefatto per artefatto, quale delle due si sta leggendo: una
#     conversione applicata alla cieca produrrebbe date sbagliate di qualche
#     ora, che e' molto peggio di date dichiarate ambigue. Qui quindi NON si
#     converte nulla: si rileva il fuso del volume e lo si dichiara, cosi'
#     l'analista sa cosa significa un'ora locale trovata in un report.
#
#  2) QUANTE DATE DEVO GUARDARE. Su un disco da un terabyte un report puo'
#     contenere centinaia di migliaia di righe che coprono anni, mentre
#     l'incidente sta in tre giorni. --since/--until tagliano tutto cio' che
#     cade fuori dalla finestra.
#
#  Il filtro confronta le date COSI' COME COMPAIONO nell'artefatto, senza
#  riportarle a un fuso comune: e' l'unico confronto che non introduce errori
#  inventati. Su una finestra di giorni la differenza e' irrilevante; se la
#  finestra e' di ore, va allargata di un margine pari all'offset del volume.
#  Il report lo dichiara esplicitamente.
# ================================================================

TIME_SINCE=""          # limite inferiore normalizzato (YYYY-MM-DDTHH:MM:SS)
TIME_UNTIL=""          # limite superiore normalizzato
VOLUME_TZ=""           # fuso del volume analizzato, se rilevato
VOLUME_TZ_SOURCE=""    # da dove e' stato letto (serve a poterlo contestare)

# Directory di stato: come per custody e hive, il percorso e' deterministico
# perche' il conteggio delle righe scartate viene incrementato da subshell.
_time_dir() {
    local D="${TMPDIR:-/tmp}/fiuto_time_$$"
    [[ -d "$D" ]] || mkdir -p "$D" 2>/dev/null || return 1
    echo "$D"
}

# --- normalizzazione dei limiti ---------------------------------------------

# parse_time_bound <stringa> [end]
# Accetta: YYYY-MM-DD, YYYY-MM-DD HH:MM[:SS], la stessa con 'T' e/o 'Z' finale,
# e le forme relative -7d / -36h / -90m rispetto all'istante di avvio.
# Con "end" una data senza orario diventa fine giornata invece che inizio,
# cosi' --until 2026-03-01 include il 1 marzo per intero: e' quello che
# chiunque si aspetta scrivendolo.
parse_time_bound() {
    local RAW="$1" KIND="${2:-start}"
    [[ -n "$RAW" ]] || return 1

    # Forme relative: -7d, -36h, -90m
    if [[ "$RAW" =~ ^-([0-9]+)([dhm])$ ]]; then
        local N="${BASH_REMATCH[1]}" U="${BASH_REMATCH[2]}" SPEC BSD
        # Le unita' di date(1) BSD non coincidono con le nostre: 'm' li' sono
        # MESI, i minuti sono 'M'. Usare la stessa lettera darebbe una finestra
        # sbagliata di ordini di grandezza, in silenzio.
        case "$U" in
            d) SPEC="$N days ago";    BSD="-${N}d" ;;
            h) SPEC="$N hours ago";   BSD="-${N}H" ;;
            m) SPEC="$N minutes ago"; BSD="-${N}M" ;;
        esac
        date -d "$SPEC" "+%Y-%m-%dT%H:%M:%S" 2>/dev/null && return 0
        # date(1) BSD (macOS): non ha -d.
        date -v"$BSD" "+%Y-%m-%dT%H:%M:%S" 2>/dev/null && return 0
        return 1
    fi

    local S="${RAW%Z}"
    S="${S//T/ }"
    if [[ "$S" =~ ^([0-9]{4}-[0-9]{2}-[0-9]{2})$ ]]; then
        [[ "$KIND" == "end" ]] && { echo "${S}T23:59:59"; return 0; }
        echo "${S}T00:00:00"; return 0
    fi
    if [[ "$S" =~ ^([0-9]{4}-[0-9]{2}-[0-9]{2})\ ([0-9]{2}:[0-9]{2})$ ]]; then
        echo "${BASH_REMATCH[1]}T${BASH_REMATCH[2]}:$([[ "$KIND" == "end" ]] && echo 59 || echo 00)"
        return 0
    fi
    if [[ "$S" =~ ^([0-9]{4}-[0-9]{2}-[0-9]{2})\ ([0-9]{2}:[0-9]{2}:[0-9]{2})$ ]]; then
        echo "${BASH_REMATCH[1]}T${BASH_REMATCH[2]}"; return 0
    fi
    return 1
}

# Vero se e' stata impostata almeno una delle due estremita'.
time_window_active() { [[ -n "$TIME_SINCE" || -n "$TIME_UNTIL" ]]; }

# Descrizione leggibile della finestra, per console e report.
time_window_label() {
    time_window_active || return 0
    if [[ -n "$TIME_SINCE" && -n "$TIME_UNTIL" ]]; then
        echo "${TIME_SINCE/T/ } → ${TIME_UNTIL/T/ }"
    elif [[ -n "$TIME_SINCE" ]]; then
        echo "$(L "dal" "from") ${TIME_SINCE/T/ }"
    else
        echo "$(L "fino al" "until") ${TIME_UNTIL/T/ }"
    fi
}

# --- rilevamento del fuso del volume ----------------------------------------

# Legge il fuso dal volume analizzato. Non converte niente: serve a dichiarare
# cosa significano le ore locali che compaiono nei report.
detect_volume_timezone() {
    VOLUME_TZ=""; VOLUME_TZ_SOURCE=""
    [[ -n "$WIN_ROOT" ]] || return 0

    case "$OS_TYPE" in
        linux)
            local TZF; TZF=$(ci_find_file "$WIN_ROOT/etc" "timezone" 2>/dev/null)
            if [[ -n "$TZF" && -s "$TZF" ]]; then
                VOLUME_TZ=$(head -1 "$TZF" | tr -d '[:space:]')
                VOLUME_TZ_SOURCE="/etc/timezone"
            fi
            # Su systemd /etc/timezone puo' mancare: il fuso e' allora nel
            # target del symlink /etc/localtime.
            if [[ -z "$VOLUME_TZ" && -L "$WIN_ROOT/etc/localtime" ]]; then
                local T; T=$(readlink "$WIN_ROOT/etc/localtime" 2>/dev/null)
                VOLUME_TZ="${T##*/zoneinfo/}"
                VOLUME_TZ_SOURCE="/etc/localtime"
            fi
            ;;
        macos)
            if [[ -L "$WIN_ROOT/etc/localtime" || -L "$WIN_ROOT/private/etc/localtime" ]]; then
                local T
                T=$(readlink "$WIN_ROOT/etc/localtime" 2>/dev/null \
                    || readlink "$WIN_ROOT/private/etc/localtime" 2>/dev/null)
                VOLUME_TZ="${T##*/zoneinfo/}"
                VOLUME_TZ_SOURCE="/etc/localtime"
            fi
            ;;
        windows)
            local HIVE; HIVE=$(get_hive SYSTEM 2>/dev/null)
            [[ -n "$HIVE" && -f "$HIVE" ]] || return 0
            local OUT
            OUT=$("$PY3" - "$HIVE" << 'PYEOF' 2>/dev/null
import sys
try:
    from regipy.registry import RegistryHive
except Exception:
    sys.exit(0)

try:
    hive = RegistryHive(sys.argv[1])
except Exception:
    sys.exit(0)

# Il ControlSet corrente non e' sempre il 001: su una macchina che ha avuto un
# avvio fallito la differenza c'e' davvero.
sel = 1
try:
    for v in hive.get_key('\\Select').get_values():
        if v['name'].lower() == 'current':
            sel = int(v['value'])
except Exception:
    pass

name = bias = std_bias = None
for cs in (f'ControlSet{sel:03d}', 'ControlSet001', 'CurrentControlSet'):
    try:
        vals = hive.get_key(f'\\{cs}\\Control\\TimeZoneInformation').get_values()
    except Exception:
        continue
    for v in vals:
        n = v['name'].lower()
        if n == 'timezonekeyname' and v['value']:
            name = str(v['value']).rstrip('\x00').strip()
        elif n == 'bias':
            bias = v['value']
        elif n == 'standardbias':
            std_bias = v['value']
    if name or bias is not None:
        break

if not name and bias is None:
    sys.exit(0)

# Bias e' in minuti da sottrarre all'ora locale per ottenere UTC: il segno e'
# quindi invertito rispetto all'offset che si scrive di solito (UTC+1 -> -60).
off = ''
if isinstance(bias, int):
    total = -(bias + (std_bias if isinstance(std_bias, int) else 0))
    sign = '+' if total >= 0 else '-'
    total = abs(total)
    off = f'UTC{sign}{total // 60:02d}:{total % 60:02d}'
print('\t'.join([name or '', off]))
PYEOF
            )
            local TZNAME TZOFF
            IFS=$'\t' read -r TZNAME TZOFF <<< "$OUT"
            if [[ -n "${TZNAME:-}" || -n "${TZOFF:-}" ]]; then
                VOLUME_TZ="${TZNAME:-?}${TZOFF:+ (${TZOFF})}"
                VOLUME_TZ_SOURCE="SYSTEM\\Control\\TimeZoneInformation"
            fi
            ;;
    esac
    [[ -n "$VOLUME_TZ" ]] && log_msg "[TZ] volume: $VOLUME_TZ (${VOLUME_TZ_SOURCE})"
    return 0
}

# --- filtro ------------------------------------------------------------------

# Somma le righe scartate dal filtro, per poterlo dichiarare a fine sessione.
time_filtered_add() {
    local N="${1:-0}"
    [[ "$N" -gt 0 ]] 2>/dev/null || return 0
    local D; D=$(_time_dir) || return 0
    echo "$N" >> "${D}/dropped" 2>/dev/null || true
}

time_filtered_total() {
    local D="${TMPDIR:-/tmp}/fiuto_time_$$"
    [[ -f "${D}/dropped" ]] || { echo 0; return 0; }
    awk '{s+=$1} END{print s+0}' "${D}/dropped" 2>/dev/null || echo 0
}

# Nota HTML da inserire nei report quando la finestra e' attiva. Un report
# filtrato che non dichiara di esserlo e' una trappola: chi lo legge conclude
# che prima di quella data non e' successo niente.
time_window_html() {
    time_window_active || return 0
    printf "<div class='card' style='margin-bottom:1rem'><div style='padding:.9rem 1.5rem;font-size:.8rem;line-height:1.7'><b>%s</b> %s<br>%s%s</div></div>" \
        "$(L "Report filtrato per data:" "Report filtered by date:")" \
        "$(html_esc "$(time_window_label)")" \
        "$(L "Le righe con una data fuori dalla finestra non compaiono. Le righe prive di data sono state mantenute: non erano valutabili." \
             "Rows carrying a date outside the window are not shown. Rows without a date were kept: they could not be evaluated.")" \
        "$([[ -n "$VOLUME_TZ" ]] && printf " %s" "$(L "Il confronto usa le date come compaiono nell'artefatto, senza riportarle a un fuso comune; il volume risulta configurato su" "The comparison uses dates as they appear in the artefact, without normalising them to a common zone; the volume is configured for") $(html_esc "$VOLUME_TZ").")"
}

# ================================================================
#  MODALITA' --redact
#
#  I report contengono materiale che serve all'analisi e non deve uscire dallo
#  studio: hash NTLM da SAM e NTDS, PSK Wi-Fi, token OAuth, chiavi private,
#  segreti LSA. Finche' l'unica opzione e' "mandare il report intero o non
#  mandarlo", quel materiale finisce in una mail a un cliente o in un allegato
#  di un ticket. --redact esiste per rendere praticabile la terza via.
#
#  DUE SCELTE DI FONDO.
#
#  1. Il report originale NON viene toccato. Accanto ad esso nasce un
#     report.redacted.html. Oscurare sul posto distruggerebbe evidenza per una
#     necessita' di comunicazione, che e' un baratto che non spetta al tool
#     fare. Il manifesto di custodia elenca entrambi, con ruoli distinti.
#
#  2. Si oscura per CONTESTO, non per forma. Un hash SHA-256 e un hash NTLM
#     sono entrambi stringhe esadecimali: il primo e' l'impronta di integrita'
#     di un reperto e cancellarlo distruggerebbe la catena di custodia, il
#     secondo e' una credenziale. A distinguerli e' l'etichetta accanto, non il
#     numero di caratteri. Per questo le regole guardano il campo, non il
#     valore, e i conteggi per regola sono stampati: una regola che non scatta
#     mai o che scatta ovunque si vede subito.
#
#  Ogni copia oscurata dichiara in testa cosa e' stata: un file che sembra un
#  report ma non lo e' del tutto, senza dirlo, e' peggio di nessun file.
# ================================================================

REDACT=false          # --redact: genera una copia oscurata di ogni report
REDACT_DEFANG=false   # --defang: nella copia oscurata rende inerti URL e IP

# Numero di copie oscurate prodotte nella sessione (per il riepilogo finale).
_redact_dir() {
    local D="${TMPDIR:-/tmp}/fiuto_redact_$$"
    [[ -d "$D" ]] || mkdir -p "$D" 2>/dev/null || return 1
    echo "$D"
}

redacted_count() {
    local D="${TMPDIR:-/tmp}/fiuto_redact_$$"
    [[ -f "${D}/done" ]] && wc -l < "${D}/done" || echo 0
}

# redact_report <report.html>
# Scrive <report>.redacted.html accanto all'originale. L'originale resta
# intatto. Non fallisce mai in modo rumoroso: se non c'e' niente da oscurare
# la copia viene prodotta ugualmente, cosi' chi condivide non deve chiedersi
# perche' per un modulo il file manchi.
redact_report() {
    local HTML="$1"
    [[ -n "$HTML" && -f "$HTML" ]] || return 0
    case "$HTML" in *.redacted.html) return 0 ;; esac
    local OUT="${HTML%.html}.redacted.html"

    local N
    N=$(FIUTO_DEFANG="$REDACT_DEFANG" FIUTO_LANG="${LANG:-en}" \
        "$PY3" - "$HTML" "$OUT" << 'PYEOF' 2>/dev/null
import sys, os, re, html

src, dst = sys.argv[1], sys.argv[2]
DEFANG = os.environ.get('FIUTO_DEFANG', 'false') == 'true'
IT = os.environ.get('FIUTO_LANG', 'en') == 'it'


def L(i, e):
    return i if IT else e


try:
    text = open(src, encoding='utf-8', errors='replace').read()
except Exception:
    sys.exit(1)

MARK = '[REDACTED]'
counts = {}


def sub(name, pattern, repl, flags=re.I):
    global text
    text, n = re.subn(pattern, repl, text, flags=flags)
    if n:
        counts[name] = counts.get(name, 0) + n


# 1) Chiavi private in formato PEM: il blocco intero, non solo l'intestazione.
sub('chiave privata PEM',
    r'-----BEGIN [A-Z ]*PRIVATE KEY-----.*?-----END [A-Z ]*PRIVATE KEY-----',
    '-----BEGIN PRIVATE KEY----- ' + MARK + ' -----END PRIVATE KEY-----',
    re.I | re.DOTALL)

# 2) Coppie etichetta/valore. L'etichetta resta: sapere CHE c'era una password
#    e' un dato dell'analisi, il suo valore no.
LABELS = (r'password|passwd|pwd|psk|pre[- ]?shared[- ]?key|keymaterial|'
          r'key\s*material|secret|client_secret|access_token|refresh_token|'
          r'bearer|token|api[_-]?key|apikey|aws_secret_access_key|'
          r'connection\s*string|nt\s*hash|ntlm\s*hash|lm\s*hash|hash\s*nt')
sub('campo etichettato',
    r'(\b(?:' + LABELS + r')\b\s*(?:[:=]|&gt;|&#61;)\s*)(?:&quot;|["\']?)([^\s<"\'&]{3,})',
    lambda m: m.group(1) + MARK)

# 3) Riversamenti SAM/NTDS: utente:RID:LM:NT:::. Utente e RID restano — sono
#    l'informazione forense — gli hash no.
sub('hash SAM/NTDS',
    r'([^\s:<>]{1,64}):(\d+):([0-9a-f]{32}):([0-9a-f]{32}):::',
    lambda m: '%s:%s:%s:%s:::' % (m.group(1), m.group(2), MARK, MARK))

# 4) Materiale WLAN, sia come tag sia come entita' HTML.
sub('WLAN keyMaterial',
    r'(&lt;keyMaterial&gt;|<keyMaterial>)([^<&]+)',
    lambda m: m.group(1) + MARK)

# 5) Credenziali con forma inequivocabile: il contesto non serve, la stringa
#    stessa non e' altro che una credenziale.
sub('AWS access key', r'\bAKIA[0-9A-Z]{16}\b', MARK)
sub('JWT', r'\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b', MARK)
sub('Slack/GitHub token', r'\b(?:xox[baprs]-[A-Za-z0-9-]{10,}|gh[pousr]_[A-Za-z0-9]{20,})\b', MARK)

# 6) Defanging facoltativo di URL e IP nella copia condivisibile.
if DEFANG:
    sub('URL defanged', r'\bhttps?://', lambda m: m.group(0).replace('http', 'hxxp'), 0)
    sub('IPv4 defanged',
        r'(?<![\w.])((?:\d{1,3}\.){3}\d{1,3})(?![\w.])',
        lambda m: m.group(1).replace('.', '[.]'), 0)

rows = ''.join('<tr><td class="mono">%s</td><td class="mono">%d</td></tr>'
               % (html.escape(k), v) for k, v in sorted(counts.items()))
total = sum(counts.values())

banner = (
    '<div class="card" style="margin:1rem 0;border-color:#f0883e">'
    '<div style="padding:1rem 1.5rem;font-size:.82rem;line-height:1.7">'
    '<b>' + L('COPIA OSCURATA — non e\' il report originale',
              'REDACTED COPY — this is not the original report') + '</b><br>'
    + L('Da questa copia sono stati rimossi i valori delle credenziali (hash NTLM, PSK, '
        'token, chiavi private). Il report originale, integro, sta nello stesso percorso '
        'senza il suffisso <code>.redacted</code>: e\' quello da conservare come reperto e '
        'da citare in una perizia.',
        'Credential values (NTLM hashes, PSKs, tokens, private keys) have been removed from '
        'this copy. The original, intact report sits in the same path without the '
        '<code>.redacted</code> suffix: that is the one to retain as evidence and to cite '
        'in a report.')
    + '<br><br>'
    + L('Gli hash SHA-256 dei reperti <b>non</b> sono stati toccati: sono dati di '
        'integrita\', non segreti, e rimuoverli spezzerebbe la catena di custodia.',
        'Evidence SHA-256 hashes were <b>not</b> touched: they are integrity data, not '
        'secrets, and removing them would break the chain of custody.'))
if rows:
    banner += ('<br><br><table><tr><th>' + L('Regola', 'Rule') + '</th><th>'
               + L('Occorrenze', 'Occurrences') + '</th></tr>' + rows + '</table>')
else:
    banner += ('<br><br>' + L('Nessuna regola ha trovato materiale da oscurare in questo report.',
                              'No rule found material to redact in this report.'))
banner += '</div></div>'

# Il cartiglio va subito dopo l'apertura di <main>: e' la prima cosa che si
# legge, prima di qualunque dato.
if '<main>' in text:
    text = text.replace('<main>', '<main>\n' + banner, 1)
else:
    text = banner + text
text = text.replace('<title>', '<title>[REDACTED] ', 1)

with open(dst, 'w', encoding='utf-8') as fh:
    fh.write(text)
print(total)
PYEOF
    )

    [[ -f "$OUT" ]] || return 0
    local D; D=$(_redact_dir) && printf '%s\t%s\n' "${N:-0}" "$OUT" >> "${D}/done" 2>/dev/null
    evidence_note "$OUT" "copia oscurata" 2>/dev/null || true
    log_msg "[REDACT] $OUT — ${N:-0} occorrenze oscurate"
    return 0
}

# Riepilogo di fine sessione. Va detto quante copie esistono e dove: una copia
# oscurata che nessuno sa di avere non serve a niente, e una che si confonde
# con l'originale e' un rischio.
redact_summary() {
    [[ "$REDACT" == "true" ]] || return 0
    local D="${TMPDIR:-/tmp}/fiuto_redact_$$"
    [[ -f "${D}/done" ]] || return 0
    local NC NT
    NC=$(wc -l < "${D}/done")
    NT=$(awk -F'\t' '{s+=$1} END{print s+0}' "${D}/done")
    echo ""
    ok "$(L "Copie oscurate prodotte:" "Redacted copies produced:") ${BOLD}${NC}${RESET}  ·  $(L "valori rimossi:" "values removed:") ${BOLD}${NT}"
    info "$(L "Sono i file *.redacted.html. Gli originali restano integri accanto a loro." \
             "They are the *.redacted.html files. The originals remain intact beside them.")"
}

# ================================================================
#  IMMAGINI FORENSI E VOLUMI CIFRATI
#
#  Fino alla 2.2 FIUTO pretendeva un volume gia' montato. Il montaggio a mano
#  di una E01 e' una catena di quattro comandi in cui e' facile sbagliare
#  proprio il pezzo che conta: dimenticare il -r su losetup, o il ro su mount,
#  significa aver montato il reperto in scrittura senza accorgersene.
#
#  --image chiude quella catena. Ogni anello e' in SOLA LETTURA per costruzione
#  e non per convenzione: ewfmount lo e' per natura, losetup riceve -r,
#  cryptsetup --readonly, dislocker -r, mount -o ro. Non esiste un percorso di
#  codice che monti in scrittura.
#
#  COSA NON FA, dichiarato invece che scoperto a meta' analisi:
#  - non indovina chiavi. BitLocker e LUKS si aprono solo con --unlock;
#  - non tocca FileVault 2: richiede il keybag di macOS e strumenti che su
#    Linux non esistono in forma affidabile. Il volume viene riconosciuto e la
#    cosa viene detta, non aggirata;
#  - VMDK, VHDX e QCOW2 sono riconosciuti e rifiutati con l'indicazione del
#    comando per convertirli: fingere di supportarli produrrebbe letture
#    sbagliate su formati sparsi o a snapshot.
#
#  Il montaggio richiede privilegi (losetup, mount, cryptsetup). L'elenco delle
#  partizioni no: `--list-partitions` funziona da utente normale, cosi' si puo'
#  capire cosa c'e' in un'immagine prima di decidere se serve sudo.
# ================================================================

IMAGE_PATH=""              # --image: immagine da montare
IMAGE_PARTITION=""         # --partition N: quale partizione (numero di mmls)
IMAGE_UNLOCK=""            # --unlock: password/recovery key o file che la contiene
IMAGE_LIST_ONLY=false      # --list-partitions: elenca ed esci

# Pila di smontaggio. Su file e non in un array: il cleanup gira dalla trap
# EXIT, e deve funzionare anche se qualcosa e' stato aperto da una subshell.
_image_state() {
    local D="${TMPDIR:-/tmp}/fiuto_image_$$"
    [[ -d "$D" ]] || mkdir -p "$D" 2>/dev/null || return 1
    echo "$D"
}

# _image_push <tipo> <bersaglio> — tipo: fuse | crypt | mount | loop
_image_push() {
    local D; D=$(_image_state) || return 0
    printf '%s\t%s\n' "$1" "$2" >> "${D}/stack"
}

# Smonta tutto in ordine inverso. Idempotente: puo' essere chiamata due volte.
image_cleanup() {
    local D="${TMPDIR:-/tmp}/fiuto_image_$$"
    local S="${D}/stack"
    [[ -f "$S" ]] || return 0
    local KIND TARGET
    # tac e non un ciclo al contrario: l'ordine di chiusura e' l'inverso esatto
    # di quello di apertura, e sbagliarlo lascia loop device orfani.
    while IFS=$'\t' read -r KIND TARGET; do
        # Ogni ramo termina con true: una funzione di pulizia non deve MAI
        # propagare un errore. Gira dalla trap EXIT, spesso quando qualcosa e'
        # gia' andato storto, e un suo fallimento nasconderebbe la causa vera.
        case "$KIND" in
            mount) umount "$TARGET" 2>/dev/null || umount -l "$TARGET" 2>/dev/null || true ;;
            crypt) cryptsetup close "$TARGET" 2>/dev/null || true ;;
            loop)  losetup -d "$TARGET" 2>/dev/null || true ;;
            fuse)  fusermount -u "$TARGET" 2>/dev/null || umount "$TARGET" 2>/dev/null || true ;;
        esac
    done < <(tac "$S" 2>/dev/null)
    rm -f "$S"
    return 0
}

# Tipo dell'immagine dal contenuto, non dall'estensione: un .dd rinominato .img
# resta un raw, e una E01 chiamata .001 resta una E01.
image_type() {
    local F="$1"
    [[ -f "$F" ]] || { echo "assente"; return 1; }
    local MAGIC; MAGIC=$(head -c 8 "$F" 2>/dev/null | od -An -tx1 | tr -d ' \n')
    case "$MAGIC" in
        45564609*) echo "ewf" ; return ;;   # "EVF\x09" — EnCase E01
        45564632*) echo "ewf2"; return ;;   # "EVF2"    — Ex01
    esac
    case "$MAGIC" in
        4b444d56*) echo "vmdk" ; return ;;  # "KDMV"
        76686478*) echo "vhdx" ; return ;;  # "vhdx"
        514649fb*) echo "qcow2"; return ;;  # "QFI\xfb"
    esac
    echo "raw"
}

# Cifratura del contenuto a un dato offset. Riconoscere e dichiarare vale piu'
# che tentare: un volume BitLocker montato come raw non da' errore, da' dati
# illeggibili che sembrano corruzione.
image_crypto_at() {
    local DEV="$1" OFF="${2:-0}"
    local HDR; HDR=$(dd if="$DEV" bs=1 skip="$OFF" count=16 2>/dev/null | od -An -c | tr -s ' ')
    case "$HDR" in
        *'- F V E - F S -'*) echo bitlocker; return ;;
        *'M S W I N 4 . 1'*) : ;;
    esac
    local L; L=$(dd if="$DEV" bs=1 skip="$OFF" count=6 2>/dev/null | od -An -tx1 | tr -d ' \n')
    [[ "$L" == 4c554b53babe* ]] && { echo luks; return; }
    # APFS cifrato: il container si riconosce, la chiave sta nel keybag.
    # Il confronto passa dall'esadecimale: i byte grezzi in una command
    # substitution fanno emettere a bash un avviso sui null, e su una
    # partizione vuota (tutti zeri) l'avviso comparirebbe a ogni riga.
    local A; A=$(dd if="$DEV" bs=1 skip=$((OFF + 32)) count=4 2>/dev/null | od -An -tx1 | tr -d ' \n')
    [[ "$A" == "4e585342" ]] && { echo apfs; return; }
    echo none
}

# Elenco delle partizioni via mmls. Ritorna righe "slot<TAB>offset_byte<TAB>settori<TAB>descrizione".
image_partitions() {
    local DEV="$1"
    command -v mmls > /dev/null 2>&1 || return 1
    # L'output di mmls passa da un file e non da una pipe: lo stdin del
    # processo Python e' gia' occupato dall'heredoc con il programma.
    local TMP; TMP=$(mktemp) || return 1
    mmls -M "$DEV" > "$TMP" 2>/dev/null
    "$PY3" - "$TMP" << 'PYEOF'
import sys, re

sector = 512
rows = []
for line in open(sys.argv[1], encoding='utf-8', errors='replace'):
    m = re.search(r'Units are in (\d+)-byte sectors', line)
    if m:
        sector = int(m.group(1))
        continue
    # "  002:  000:000   0000002048   0000206847   0000204800   NTFS / exFAT"
    m = re.match(r'\s*(\d+):\s+\S+\s+(\d+)\s+(\d+)\s+(\d+)\s+(.*?)\s*$', line)
    if not m:
        continue
    slot, start, _end, length, desc = m.groups()
    # Le voci meta (tabella delle partizioni, spazio non allocato) non si
    # montano: elencarle come montabili farebbe scegliere quella sbagliata.
    if re.search(r'unallocated|meta|extended|primary table', desc, re.I):
        continue
    rows.append('%s\t%d\t%s\t%s' % (slot, int(start) * sector, length, desc))
print('\n'.join(rows))
PYEOF
    rm -f "$TMP"
}

# Stampa a console le partizioni trovate. Funziona senza privilegi.
image_show_partitions() {
    local DEV="$1"
    local ROWS; ROWS=$(image_partitions "$DEV")
    if [[ -z "$ROWS" ]]; then
        warn "$(L "Nessuna tabella delle partizioni: l'immagine sembra contenere un solo filesystem." \
                 "No partition table: the image seems to hold a single filesystem.")"
        return 1
    fi
    echo ""
    printf "  ${BOLD}%-5s %-14s %-12s %-8s %s${RESET}\n" \
        "$(L "SLOT" "SLOT")" "$(L "OFFSET" "OFFSET")" "$(L "SETTORI" "SECTORS")" \
        "$(L "CIFRAT." "CRYPTO")" "$(L "DESCRIZIONE" "DESCRIPTION")"
    local SLOT OFF LEN DESC
    while IFS=$'\t' read -r SLOT OFF LEN DESC; do
        local C; C=$(image_crypto_at "$DEV" "$OFF")
        local COL="$RESET"; [[ "$C" != "none" ]] && COL="$YELLOW"
        printf "  ${CYAN}%-5s${RESET} %-14s %-12s ${COL}%-8s${RESET} %s\n" \
            "$SLOT" "$OFF" "$LEN" "$([[ "$C" == "none" ]] && echo "-" || echo "$C")" "$DESC"
    done <<< "$ROWS"
    echo ""
    return 0
}

# Legge la chiave da --unlock: valore diretto o percorso di un file.
_image_key() {
    [[ -n "$IMAGE_UNLOCK" ]] || return 1
    if [[ -f "$IMAGE_UNLOCK" ]]; then
        head -1 "$IMAGE_UNLOCK"
    else
        printf '%s' "$IMAGE_UNLOCK"
    fi
}

# image_open — apre IMAGE_PATH e stampa il punto di mount su stdout.
# Ogni messaggio va su stderr: lo stdout e' il percorso, e sporcarlo
# significherebbe restituire una root inesistente.
image_open() {
    local IMG="$IMAGE_PATH"
    [[ -f "$IMG" ]] || { err "$(L "Immagine non trovata:" "Image not found:") $IMG" >&2; return 1; }

    local TYPE; TYPE=$(image_type "$IMG")
    info "$(L "Immagine:" "Image:") ${BOLD}$(basename "$IMG")${RESET} — $(L "formato" "format"): ${BOLD}${TYPE}" >&2

    case "$TYPE" in
        vmdk|vhdx|qcow2)
            err "$(L "Formato $TYPE non supportato." "Format $TYPE not supported.")" >&2
            info "$(L "Sono formati sparsi o a snapshot: leggerli come raw darebbe dati sbagliati senza errori. Converti prima:" \
                     "These are sparse or snapshot formats: reading them as raw would give wrong data with no error. Convert first:")" >&2
            info "  qemu-img convert -f $TYPE -O raw \"$(basename "$IMG")\" immagine.raw" >&2
            return 1 ;;
    esac

    # --- E01/Ex01: ewfmount espone un raw. E' FUSE, non serve root. ---------
    local RAW="$IMG"
    if [[ "$TYPE" == ewf || "$TYPE" == ewf2 ]]; then
        command -v ewfmount > /dev/null 2>&1 || {
            err "$(L "ewfmount non disponibile (pacchetto ewf-tools)." "ewfmount unavailable (ewf-tools package).")" >&2
            return 1; }
        local D; D=$(_image_state) || return 1
        local EWFDIR="${D}/ewf"; mkdir -p "$EWFDIR"
        if ! ewfmount "$IMG" "$EWFDIR" 2>/dev/null; then
            err "$(L "ewfmount non riuscito. Se l'immagine e' segmentata, indica il primo segmento (.E01)." \
                     "ewfmount failed. If the image is segmented, point at the first segment (.E01).")" >&2
            return 1
        fi
        _image_push fuse "$EWFDIR"
        RAW="${EWFDIR}/ewf1"
        [[ -e "$RAW" ]] || { err "$(L "ewfmount non ha esposto ewf1." "ewfmount did not expose ewf1.")" >&2; return 1; }
        ok "$(L "E01 aperta in sola lettura." "E01 opened read-only.")" >&2
    fi

    # --- scelta della partizione -------------------------------------------
    local ROWS; ROWS=$(image_partitions "$RAW")
    local OFFSET=0 PDESC="$(L "immagine intera" "whole image")"
    if [[ -n "$ROWS" ]]; then
        if $IMAGE_LIST_ONLY; then
            image_show_partitions "$RAW" >&2
            return 2
        fi
        local NPART; NPART=$(wc -l <<< "$ROWS")
        if [[ -n "$IMAGE_PARTITION" ]]; then
            local LINE; LINE=$(awk -F'\t' -v s="$IMAGE_PARTITION" '$1==s || $1+0==s+0' <<< "$ROWS" | head -1)
            [[ -n "$LINE" ]] || {
                err "$(L "Partizione inesistente:" "No such partition:") $IMAGE_PARTITION" >&2
                image_show_partitions "$RAW" >&2
                return 1; }
            OFFSET=$(cut -f2 <<< "$LINE"); PDESC=$(cut -f4 <<< "$LINE")
        elif [[ "$NPART" -eq 1 ]]; then
            OFFSET=$(cut -f2 <<< "$ROWS"); PDESC=$(cut -f4 <<< "$ROWS")
        else
            # Sceglierne una a caso su un disco multi-partizione significa
            # analizzare la partizione sbagliata e non accorgersene.
            err "$(L "L'immagine contiene $NPART partizioni: indica quale con --partition." \
                     "The image holds $NPART partitions: pick one with --partition.")" >&2
            image_show_partitions "$RAW" >&2
            return 1
        fi
    elif $IMAGE_LIST_ONLY; then
        image_show_partitions "$RAW" >&2
        return 2
    fi
    info "$(L "Partizione:" "Partition:") ${BOLD}${PDESC}${RESET} (offset ${OFFSET})" >&2

    # --- cifratura: si rileva PRIMA di chiedere i privilegi ----------------
    # "Questa partizione e' BitLocker e non mi hai dato la chiave" si sa da
    # utente normale, ed e' un'informazione piu' utile di "serve root": evita
    # di rilanciare con sudo per scoprire solo allora che manca la chiave.
    local CRYPTO; CRYPTO=$(image_crypto_at "$RAW" "$OFFSET")
    if [[ "$CRYPTO" == "apfs" ]]; then
        err "$(L "Container APFS cifrato (FileVault 2)." "Encrypted APFS container (FileVault 2).")" >&2
        info "$(L "Non supportato: la chiave sta nel keybag di macOS e su Linux non esistono strumenti affidabili per estrarla. Sblocca il volume su un Mac e ri-acquisiscilo, oppure analizzalo montato da macOS." \
                 "Not supported: the key lives in the macOS keybag and Linux has no reliable tooling for it. Unlock the volume on a Mac and re-acquire it, or analyse it mounted from macOS.")" >&2
        return 1
    fi
    if [[ "$CRYPTO" != "none" ]] && ! _image_key > /dev/null; then
        err "$(L "Partizione ${CRYPTO}: serve la chiave con --unlock <chiave|file>." \
                 "Partition is ${CRYPTO}: a key is required — --unlock <key|file>.")" >&2
        info "$(L "FIUTO non tenta di indovinarla." "FIUTO does not attempt to guess it.")" >&2
        return 1
    fi

    # --- da qui servono i privilegi ----------------------------------------
    if [[ "$(id -u)" -ne 0 ]]; then
        err "$(L "Il montaggio richiede privilegi di root (losetup, mount, cryptsetup)." \
                 "Mounting requires root privileges (losetup, mount, cryptsetup).")" >&2
        info "$(L "Con --list-partitions l'elenco funziona anche senza." \
                 "With --list-partitions the listing works without them.")" >&2
        return 1
    fi

    local D; D=$(_image_state) || return 1
    local MNT="${D}/mnt"; mkdir -p "$MNT"
    local SRC=""

    case "$CRYPTO" in
        none)
            SRC=$(losetup --find --show --read-only --offset "$OFFSET" "$RAW" 2>/dev/null) || {
                err "$(L "losetup non riuscito." "losetup failed.")" >&2; return 1; }
            _image_push loop "$SRC"
            ;;
        luks)
            warn "$(L "Partizione LUKS." "LUKS partition.")" >&2
            local KEY; KEY=$(_image_key) || {
                err "$(L "Serve la passphrase: --unlock <passphrase|file>." \
                         "Passphrase required: --unlock <passphrase|file>.")" >&2; return 1; }
            local LOOP; LOOP=$(losetup --find --show --read-only --offset "$OFFSET" "$RAW" 2>/dev/null) || {
                err "$(L "losetup non riuscito." "losetup failed.")" >&2; return 1; }
            _image_push loop "$LOOP"
            local NAME="fiuto_luks_$$"
            if ! printf '%s' "$KEY" | cryptsetup open --readonly --key-file=- "$LOOP" "$NAME" 2>/dev/null; then
                err "$(L "Apertura LUKS fallita: passphrase errata o header danneggiato." \
                         "LUKS unlock failed: wrong passphrase or damaged header.")" >&2
                return 1
            fi
            _image_push crypt "$NAME"
            SRC="/dev/mapper/${NAME}"
            ok "$(L "Volume LUKS aperto in sola lettura." "LUKS volume opened read-only.")" >&2
            ;;
        bitlocker)
            warn "$(L "Partizione BitLocker." "BitLocker partition.")" >&2
            command -v dislocker > /dev/null 2>&1 || {
                err "$(L "dislocker non disponibile." "dislocker unavailable.")" >&2; return 1; }
            local KEY; KEY=$(_image_key) || {
                err "$(L "Serve la chiave: --unlock <password|recovery key|file>." \
                         "Key required: --unlock <password|recovery key|file>.")" >&2; return 1; }
            local DIS="${D}/dislocker"; mkdir -p "$DIS"
            # Una recovery key BitLocker sono 48 cifre in 8 gruppi: la si passa
            # con -p, una password utente con -u. Distinguerle a occhio evita
            # un fallimento che sembrerebbe "chiave sbagliata".
            local FLAG="-u"
            [[ "$KEY" =~ ^[0-9]{6}(-[0-9]{6}){7}$ ]] && FLAG="-p"
            if ! dislocker -r -V "$RAW" -O "$OFFSET" "${FLAG}${KEY}" -- "$DIS" 2>/dev/null; then
                err "$(L "Apertura BitLocker fallita: chiave errata o volume non supportato." \
                         "BitLocker unlock failed: wrong key or unsupported volume.")" >&2
                return 1
            fi
            _image_push fuse "$DIS"
            SRC="${DIS}/dislocker-file"
            ok "$(L "Volume BitLocker aperto in sola lettura." "BitLocker volume opened read-only.")" >&2
            ;;
    esac

    # --- mount, sempre e solo in sola lettura ------------------------------
    # noexec/nodev/nosuid: il contenuto e' materiale ostile per definizione, e
    # non c'e' ragione perche' il kernel lo consideri eseguibile.
    local OPTS="ro,noexec,nodev,nosuid"
    if mount -o "$OPTS" "$SRC" "$MNT" 2>/dev/null; then
        :
    elif mount -t ntfs-3g -o "ro,noexec,nodev,nosuid,streams_interface=windows" "$SRC" "$MNT" 2>/dev/null; then
        :
    else
        err "$(L "Mount fallito. Filesystem non riconosciuto o servono moduli kernel assenti." \
                 "Mount failed. Unrecognised filesystem, or missing kernel modules.")" >&2
        return 1
    fi
    _image_push mount "$MNT"

    # Il montaggio e' un fatto della catena di custodia quanto il file.
    evidence_note "$IMG" "immagine forense montata" 2>/dev/null || true
    log_msg "[IMAGE] $IMG ($TYPE) offset=$OFFSET crypto=$CRYPTO -> $MNT"
    ok "$(L "Immagine montata in sola lettura:" "Image mounted read-only:") ${BOLD}${MNT}" >&2
    echo "$MNT"
}

# ================================================================
#  CATENA DI CUSTODIA
#
#  Un report forense vale quanto la tracciabilita' di cio' su cui si basa. Fino
#  alla 2.1 FIUTO calcolava qualche hash sparso in pochi moduli: non c'era modo
#  di rispondere a "quali file sono stati letti, in che stato erano, e i report
#  allegati sono quelli prodotti allora?".
#
#  Il manifesto risponde a tutte e tre. E' scritto a fine sessione in
#  evidence_manifest.json e contiene: identita' del tool e comando eseguito,
#  operatore e host di analisi, finestra temporale, volume analizzato, elenco
#  dei file di evidenza con dimensione, data e SHA256, ed elenco dei report
#  generati con il loro SHA256.
#
#  Aggancio: invece di modificare 90 moduli, si annotano i file nelle funzioni
#  che TUTTI attraversano per accedere a un'evidenza — ci_find_file,
#  query_sqlite, read_plist, recover_hive. La copertura e' quindi automatica
#  anche per i moduli futuri.
# ================================================================

CUSTODY=true              # --no-custody per disattivare
CUSTODY_HASH=true         # --no-hash: annota i file senza calcolarne l'hash
CUSTODY_HASH_LIMIT_MB=1024   # oltre questa soglia niente hash (--hash-limit)

# Directory di sessione. Come per il replay degli hive, il percorso e'
# deterministico e non tenuto in una variabile: le annotazioni arrivano da
# subshell (le funzioni sono usate in command substitution).
_custody_dir() {
    local D="${TMPDIR:-/tmp}/fiuto_custody_$$"
    [[ -d "$D" ]] || { mkdir -p "$D" 2>/dev/null || return 1; chmod 700 "$D" 2>/dev/null || true; }
    echo "$D"
}

# evidence_note <percorso> [ruolo]
# Annota un file come evidenza consultata. Non calcola nulla: l'hash arriva
# alla fine, una volta sola per file, per non pagarlo a ogni accesso.
evidence_note() {
    [[ "$CUSTODY" == "true" ]] || return 0
    local P="$1" ROLE="${2:-letto}"
    [[ -n "$P" && -f "$P" ]] || return 0
    # I file di lavoro di FIUTO non sono evidenza. L'esclusione deve pero'
    # essere precisa: scartare tutto cio' che sta sotto /tmp escluderebbe le
    # evidenze quando il volume e' montato li', cosa del tutto legittima.
    case "$P" in
        */fiuto_custody_*/*|*/fiuto_hives_*/*) return 0 ;;
    esac
    local D; D=$(_custody_dir) || return 0
    printf '%s\t%s\n' "$ROLE" "$P" >> "${D}/evidence.tsv" 2>/dev/null || true
}

# Quante evidenze sono state annotate finora.
custody_count() {
    local D="${TMPDIR:-/tmp}/fiuto_custody_$$"
    [[ -f "${D}/evidence.tsv" ]] && cut -f2 "${D}/evidence.tsv" | sort -u | wc -l || echo 0
}

# Scrive evidence_manifest.json in REPORT_BASE_DIR.
# Va chiamata a fine sessione, PRIMA della pulizia dei file temporanei.
write_evidence_manifest() {
    [[ "$CUSTODY" == "true" ]] || return 0
    [[ -n "$REPORT_BASE_DIR" && -d "$REPORT_BASE_DIR" ]] || return 0
    local D="${TMPDIR:-/tmp}/fiuto_custody_$$"
    local EV="${D}/evidence.tsv"
    [[ -s "$EV" ]] || return 0

    local OUT="${REPORT_BASE_DIR}/evidence_manifest.json"
    # Lo script legge il contesto da os.environ: va esportato esplicitamente.
    export FIUTO_VERSION PY3_VERSION REPORT_BASE_DIR WIN_ROOT OS_TYPE HOST_NAME
    export CUSTODY_HASH CUSTODY_HASH_LIMIT_MB CUSTODY_START_UTC
    export CUSTODY_CMDLINE CUSTODY_OPERATOR CUSTODY_HOST
    export TIME_SINCE TIME_UNTIL VOLUME_TZ VOLUME_TZ_SOURCE
    export IMAGE_PATH IMAGE_PARTITION
    FIUTO_TIME_DROPPED=$(time_filtered_total); export FIUTO_TIME_DROPPED
    local REPORTS; REPORTS=$(mktemp)
    printf '%s\n' "${GENERATED_REPORTS[@]:-}" > "$REPORTS"
    local REPLAY; REPLAY=$(mktemp)
    hive_replay_report > "$REPLAY" 2>/dev/null || true

    "$PY3" - "$EV" "$REPORTS" "$REPLAY" "$OUT" << 'PYEOF' 2>/dev/null
import sys, os, json, hashlib, datetime

ev_path, rep_path, replay_path, out_path = sys.argv[1:5]

env = os.environ
do_hash = env.get('CUSTODY_HASH', 'true') == 'true'
try:
    limit = int(env.get('CUSTODY_HASH_LIMIT_MB', '1024')) * 1024 * 1024
except ValueError:
    limit = 1024 * 1024 * 1024


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def describe(path, roles):
    rec = {"path": path, "roles": sorted(roles)}
    try:
        st = os.stat(path)
        rec["size_bytes"] = st.st_size
        rec["mtime_utc"] = datetime.datetime.utcfromtimestamp(st.st_mtime).strftime('%Y-%m-%dT%H:%M:%SZ')
    except Exception as exc:
        rec["error"] = str(exc)
        return rec
    if not do_hash:
        rec["sha256"] = None
        rec["sha256_omitted"] = "hashing disabilitato (--no-hash)"
    elif rec["size_bytes"] > limit:
        # Su un pagefile o un $MFT da diversi GB l'hash costa quanto tutto il
        # resto dell'analisi: si dichiara perche' manca invece di ometterlo.
        rec["sha256"] = None
        rec["sha256_omitted"] = f"file oltre la soglia di {limit // (1024*1024)} MB"
    else:
        try:
            rec["sha256"] = sha256(path)
        except Exception as exc:
            rec["sha256"] = None
            rec["sha256_omitted"] = f"lettura fallita: {exc}"
    return rec


# --- evidenze -------------------------------------------------------------
items = {}
with open(ev_path, encoding='utf-8', errors='replace') as fh:
    for line in fh:
        parts = line.rstrip('\n').split('\t', 1)
        if len(parts) != 2:
            continue
        role, path = parts
        items.setdefault(path, set()).add(role)

evidence = [describe(p, r) for p, r in sorted(items.items())]

# --- report generati ------------------------------------------------------
reports = []
seen = set()
with open(rep_path, encoding='utf-8', errors='replace') as fh:
    for line in fh:
        p = line.strip()
        if not p or p in seen or not os.path.isfile(p):
            continue
        seen.add(p)
        rec = {"path": p}
        try:
            rec["size_bytes"] = os.path.getsize(p)
            rec["sha256"] = sha256(p)
        except Exception as exc:
            rec["error"] = str(exc)
        reports.append(rec)

# --- replay dei transaction log ------------------------------------------
replay = []
try:
    with open(replay_path, encoding='utf-8', errors='replace') as fh:
        for line in fh:
            f = line.rstrip('\n').split('\t')
            if len(f) >= 3:
                replay.append({"status": f[0], "hive": f[1], "detail": f[2]})
except Exception:
    pass

manifest = {
    "manifest_version": 1,
    "tool": {
        "name": "FIUTO",
        "version": env.get('FIUTO_VERSION', ''),
        "python": env.get('PY3_VERSION', ''),
    },
    "session": {
        "started_utc": env.get('CUSTODY_START_UTC', ''),
        "completed_utc": datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
        "command": env.get('CUSTODY_CMDLINE', ''),
        "operator": env.get('CUSTODY_OPERATOR', ''),
        "analysis_host": env.get('CUSTODY_HOST', ''),
        "report_directory": env.get('REPORT_BASE_DIR', ''),
    },
    "subject": {
        # Se l'analisi e' partita da un'immagine, il reperto e' l'immagine e non
        # il punto di mount: quest'ultimo e' una directory temporanea che non
        # esistera' piu' quando qualcuno rileggera' il manifesto.
        "source_image": env.get('IMAGE_PATH', '') or None,
        "source_image_partition": env.get('IMAGE_PARTITION', '') or None,
        "volume_root": env.get('WIN_ROOT', ''),
        "detected_os": env.get('OS_TYPE', ''),
        "hostname_from_artefacts": env.get('HOST_NAME', ''),
        "timezone": env.get('VOLUME_TZ', '') or None,
        "timezone_source": env.get('VOLUME_TZ_SOURCE', '') or None,
    },
    # Un report filtrato che non dichiara il filtro fa concludere a chi legge
    # che fuori dalla finestra non e' successo nulla: il manifesto lo registra
    # anche quando la finestra non e' attiva, cosi' l'assenza e' un fatto.
    "analysis_window": {
        "since": env.get('TIME_SINCE', '') or None,
        "until": env.get('TIME_UNTIL', '') or None,
        "rows_excluded": int(env.get('FIUTO_TIME_DROPPED', '0') or 0),
        "comparison": ("Le date sono confrontate come compaiono nell'artefatto, "
                       "senza conversione a un fuso comune: gli artefatti di uno "
                       "stesso volume mescolano UTC e ora locale e una conversione "
                       "applicata alla cieca sposterebbe gli eventi di ore."),
    },
    "integrity_policy": {
        "algorithm": "SHA-256",
        "hashing_enabled": do_hash,
        "size_limit_bytes": limit,
        "note": ("Gli hash sono calcolati sui file COSI' COME LETTI dal volume montato. "
                 "FIUTO non modifica il volume: monta sempre in sola lettura e verifica "
                 "questi valori contro l'immagine di acquisizione."),
    },
    "registry_transaction_log_replay": replay,
    "evidence_files": evidence,
    "generated_reports": reports,
    "counts": {
        "evidence_files": len(evidence),
        "evidence_hashed": sum(1 for e in evidence if e.get("sha256")),
        "generated_reports": len(reports),
    },
}

with open(out_path, 'w', encoding='utf-8') as fh:
    json.dump(manifest, fh, indent=2, ensure_ascii=False)
    fh.write('\n')
PYEOF

    rm -f "$REPORTS" "$REPLAY"
    if [[ -s "$OUT" ]]; then
        local NEV NRE
        NEV=$("$PY3" -c "import json,sys; print(json.load(open(sys.argv[1]))['counts']['evidence_files'])" "$OUT" 2>/dev/null || echo "?")
        NRE=$("$PY3" -c "import json,sys; print(json.load(open(sys.argv[1]))['counts']['generated_reports'])" "$OUT" 2>/dev/null || echo "?")
        echo ""
        ok "$(L "Manifesto catena di custodia:" "Chain-of-custody manifest:") ${BOLD}${OUT}"
        info "$(L "Evidenze tracciate:" "Evidence files tracked:") ${BOLD}${NEV}${RESET}  ·  $(L "report:" "reports:") ${BOLD}${NRE}"
    fi
}
ci_find_file() {
    local BASE="$1"
    local REL="$2"
    if [[ -f "$BASE/$REL" ]]; then
        evidence_note "$BASE/$REL" "letto"
        echo "$BASE/$REL"
        return
    fi
    local DIR; DIR=$(dirname "$REL")
    local FILE; FILE=$(basename "$REL")
    local DIR_PATH; DIR_PATH=$(ci_find_dir "$BASE" "$DIR")
    [[ -z "$DIR_PATH" ]] && echo "" && return
    local FOUND; FOUND=$(find "$DIR_PATH" -maxdepth 1 -iname "$FILE" -type f 2>/dev/null | head -1)
    [[ -n "$FOUND" ]] && evidence_note "$FOUND" "letto"
    echo "$FOUND"
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

# NB: la definizione di ci_find_file sta più in alto (vicino a ci_find_dir che usa).
# Qui esisteva un secondo ci_find_file che sovrascriveva il primo con semantiche
# diverse (solo nome file, -maxdepth 1): rompeva silenziosamente le chiamate che
# passano un percorso relativo, es. ci_find_file "$HOME" ".local/share/recently-used.xbel".

# ----------------------------------------------------------------
# Rilevamento OS del volume montato.
# Restituisce: windows | linux | macos | unknown
# NB: macOS va testato PRIMA di Linux, perché un volume macOS contiene anche /etc.
# ----------------------------------------------------------------
detect_os_type() {
    local ROOT="$1"
    [[ -z "$ROOT" || ! -d "$ROOT" ]] && { echo "unknown"; return; }

    # --- macOS ---
    if [[ -f "$ROOT/System/Library/CoreServices/SystemVersion.plist" ]] \
       || [[ -d "$ROOT/private/var/db/dslocal/nodes/Default" ]] \
       || [[ -d "$ROOT/var/db/dslocal/nodes/Default" ]] \
       || [[ -n "$(ci_find_dir "$ROOT" "System/Library/CoreServices")" ]]; then
        echo "macos"; return
    fi

    # --- Windows ---
    if [[ -d "$ROOT/Windows/System32" || -d "$ROOT/Users" ]] \
       || [[ -n "$(ci_find_dir "$ROOT" "Windows/System32")" ]] \
       || [[ -n "$(ci_find_dir "$ROOT" "Users")" ]]; then
        echo "windows"; return
    fi

    # --- Linux ---
    if [[ -f "$ROOT/etc/os-release" || -f "$ROOT/etc/passwd" ]] \
       || { [[ -d "$ROOT/etc" && -d "$ROOT/var" && -d "$ROOT/bin" ]]; }; then
        echo "linux"; return
    fi

    echo "unknown"
}

# Cerca una root riconoscibile un livello sotto quella indicata.
#
# Gli export e i backup di un disco arrivano quasi sempre dentro una cartella
# di servizio: "ntfs", "C", il nome del disco, la data dell'acquisizione. Chi
# analizza indica la cartella che vede, e FIUTO rispondeva "nessun volume
# valido" senza dire ne' perche' ne' dove guardare — pur avendo la risposta a
# una directory di distanza.
#
# Si scende di UN livello soltanto, e solo se quel livello contiene davvero una
# struttura riconoscibile: scendere a tentoni troverebbe prima o poi qualcosa
# che somiglia a una root e la analizzerebbe al posto di quella giusta.
find_nested_root() {
    local ROOT="$1"
    [[ -d "$ROOT" ]] || return 1
    local D OS
    while IFS= read -r D; do
        [[ -d "$D" ]] || continue
        OS=$(detect_os_type "$D")
        if [[ "$OS" != "unknown" ]]; then
            printf '%s\t%s\n' "$D" "$OS"
            return 0
        fi
    done < <(find "$ROOT" -maxdepth 1 -mindepth 1 -type d 2>/dev/null | head -20)
    return 1
}

# Etichetta leggibile dell'OS rilevato (per badge/menu)
os_label() {
    case "${1:-$OS_TYPE}" in
        windows) echo "Windows" ;;
        linux)   echo "Linux" ;;
        macos)   echo "macOS" ;;
        *)       echo "$(L "sconosciuto" "unknown")" ;;
    esac
}

# Verifica che WIN_ROOT sia impostata e contenga una struttura Windows
check_win_root() {
    if [[ -z "$WIN_ROOT" ]]; then
        err "$(t root_not_set)"
        return 1
    fi
    if [[ ! -d "$WIN_ROOT/Users" && ! -d "$WIN_ROOT/Windows" ]]; then
        # cerca case-insensitive
        local FOUND
        FOUND=$(find "$WIN_ROOT" -maxdepth 2 -type d \( -iname "Users" -o -iname "Windows" \) 2>/dev/null | head -1)
        if [[ -z "$FOUND" ]]; then
            err "$(t root_invalid) $WIN_ROOT"
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

# ----------------------------------------------------------------
# Validazione root generica in base a OS_TYPE (usata dai moduli Linux/macOS).
# I moduli Windows continuano a usare check_win_root.
# ----------------------------------------------------------------
check_target_root() {
    if [[ -z "$WIN_ROOT" ]]; then
        err "$(t root_not_set)"
        return 1
    fi
    case "$OS_TYPE" in
        windows) check_win_root ;;
        linux)
            if [[ ! -d "$WIN_ROOT/etc" && -z "$(ci_find_dir "$WIN_ROOT" "etc")" \
                  && ! -d "$WIN_ROOT/home" && -z "$(ci_find_dir "$WIN_ROOT" "home")" ]]; then
                err "$(L "Struttura Linux non trovata in:" "Linux structure not found in:") $WIN_ROOT"
                return 1
            fi
            return 0 ;;
        macos)
            if [[ ! -d "$WIN_ROOT/Users" && -z "$(ci_find_dir "$WIN_ROOT" "Users")" \
                  && ! -d "$WIN_ROOT/System" && -z "$(ci_find_dir "$WIN_ROOT" "System")" ]]; then
                err "$(L "Struttura macOS non trovata in:" "macOS structure not found in:") $WIN_ROOT"
                return 1
            fi
            return 0 ;;
        *)
            err "$(L "Tipo OS del volume non riconosciuto." "Volume OS type not recognised.")"
            return 1 ;;
    esac
}

# ----------------------------------------------------------------
# Elenca le home degli utenti reali su un volume Linux.
# Usa /etc/passwd (UID >= 1000) quando disponibile, con fallback a /home/* + /root.
# ----------------------------------------------------------------
get_linux_user_homes() {
    local ETC PASSWD
    ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    [[ -n "$ETC" ]] && PASSWD=$(ci_find_file "$ETC" "passwd")
    local -A SEEN=()
    if [[ -n "$PASSWD" && -f "$PASSWD" ]]; then
        while IFS=':' read -r _user _pw _uid _gid _gecos _home _shell; do
            [[ -z "$_home" ]] && continue
            # root (uid 0) + utenti normali (uid >= 1000), esclusi gli pseudo-account
            if [[ "$_uid" == "0" || ( "$_uid" =~ ^[0-9]+$ && "$_uid" -ge 1000 && "$_uid" -lt 65534 ) ]]; then
                local ABS="$WIN_ROOT/${_home#/}"
                [[ -d "$ABS" && -z "${SEEN[$ABS]+x}" ]] && { SEEN[$ABS]=1; echo "$ABS"; }
            fi
        done < "$PASSWD"
    fi
    # Fallback / integrazione: /home/* e /root
    local HOMEDIR ROOTDIR
    HOMEDIR=$(ci_find_dir "$WIN_ROOT" "home")
    if [[ -n "$HOMEDIR" ]]; then
        for D in "$HOMEDIR"/*/; do
            [[ -d "$D" ]] || continue
            local A="${D%/}"
            [[ -z "${SEEN[$A]+x}" ]] && { SEEN[$A]=1; echo "$A"; }
        done
    fi
    ROOTDIR=$(ci_find_dir "$WIN_ROOT" "root")
    [[ -n "$ROOTDIR" && -d "$ROOTDIR" && -z "${SEEN[$ROOTDIR]+x}" ]] && echo "$ROOTDIR"
}

# Elenca le home degli utenti su un volume macOS (/Users/*, esclusi account speciali)
get_macos_user_homes() {
    local USERS_DIR
    USERS_DIR=$(ci_find_dir "$WIN_ROOT" "Users")
    [[ -z "$USERS_DIR" ]] && return
    for D in "$USERS_DIR"/*/; do
        [[ -d "$D" ]] || continue
        local U; U=$(basename "$D")
        case "${U,,}" in
            "shared"|"guest"|".localized") continue ;;
        esac
        echo "${D%/}"
    done
}

# Home utenti coerenti con l'OS del volume corrente
get_target_user_homes() {
    case "$OS_TYPE" in
        windows) get_user_homes ;;
        linux)   get_linux_user_homes ;;
        macos)   get_macos_user_homes ;;
    esac
}

# ----------------------------------------------------------------
# Legge un plist (binario o XML) e ne stampa una rappresentazione testuale.
# ----------------------------------------------------------------
read_plist() {
    local F="$1"
    [[ -f "$F" ]] || return 1
    evidence_note "$F" "plist"
    "$PY3" - "$F" << 'PYEOF'
import sys, plistlib
def walk(o, indent=0):
    pad = "  " * indent
    if isinstance(o, dict):
        for k, v in o.items():
            if isinstance(v, (dict, list)):
                print(f"{pad}{k}:")
                walk(v, indent + 1)
            else:
                print(f"{pad}{k}: {v!r}")
    elif isinstance(o, list):
        for i, v in enumerate(o):
            if isinstance(v, (dict, list)):
                print(f"{pad}[{i}]")
                walk(v, indent + 1)
            else:
                print(f"{pad}[{i}] {v!r}")
    else:
        print(f"{pad}{o!r}")
try:
    with open(sys.argv[1], 'rb') as fh:
        data = plistlib.load(fh)
    walk(data)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF
}

# ----------------------------------------------------------------
# Esegue una query SQL su un DB SQLite, operando su una COPIA temporanea
# (evita lock/WAL su volumi read-only). Output: righe tab-separated.
# ----------------------------------------------------------------
query_sqlite() {
    local DB="$1" SQL="$2"
    [[ -f "$DB" ]] || return 1
    evidence_note "$DB" "database SQLite"
    local TMP; TMP=$(mktemp)
    cp -f "$DB" "$TMP" 2>/dev/null || { rm -f "$TMP"; return 1; }
    # copia anche -wal/-shm se presenti, per leggere transazioni non ancora consolidate
    [[ -f "${DB}-wal" ]] && cp -f "${DB}-wal" "${TMP}-wal" 2>/dev/null
    [[ -f "${DB}-shm" ]] && cp -f "${DB}-shm" "${TMP}-shm" 2>/dev/null
    "$PY3" - "$TMP" "$SQL" << 'PYEOF'
import sys, sqlite3
try:
    con = sqlite3.connect(sys.argv[1])
    cur = con.execute(sys.argv[2])
    for row in cur.fetchall():
        print("\t".join("" if c is None else str(c) for c in row))
    con.close()
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF
    rm -f "$TMP" "${TMP}-wal" "${TMP}-shm"
}

# ================================================================
#  RECUPERO INFORMAZIONI MACCHINA
# ================================================================

gather_host_info() {
    [[ -n "$WIN_ROOT" ]] || return 1

    # Per i volumi non-Windows usa una raccolta info dedicata e termina qui.
    if [[ "$OS_TYPE" == "linux" ]]; then
        gather_host_info_linux
        return 0
    elif [[ "$OS_TYPE" == "macos" ]]; then
        gather_host_info_macos
        return 0
    fi

    local SYSTEM_HIVE; SYSTEM_HIVE=$(get_hive "SYSTEM")
    local SOFTWARE_HIVE; SOFTWARE_HIVE=$(get_hive "SOFTWARE")
    
    if [[ -z "$SYSTEM_HIVE" && -z "$SOFTWARE_HIVE" ]]; then
        warn "$(t hive_not_found)"
    else
        info "$(t retrieving_info)"
        
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
        local UNABLE_MSG="$([ "$LANG" = "it" ] && echo "Impossibile rilevare il nome macchina automaticamente." || echo "Unable to detect machine name automatically.")"
        warn "$UNABLE_MSG"
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} $(t hostname_prompt) [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi

    print_host_info_table
}

# Raccolta info host per volumi Linux (hostname + distro da /etc)
gather_host_info_linux() {
    info "$(t retrieving_info)"
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    if [[ -n "$ETC" ]]; then
        local HN; HN=$(ci_find_file "$ETC" "hostname")
        [[ -n "$HN" && -f "$HN" ]] && HOST_NAME=$(head -1 "$HN" 2>/dev/null | tr -d '[:space:]')
        local OSR; OSR=$(ci_find_file "$ETC" "os-release")
        if [[ -n "$OSR" && -f "$OSR" ]]; then
            OS_VER=$(grep -E '^PRETTY_NAME=' "$OSR" 2>/dev/null | head -1 | cut -d= -f2- | tr -d '"')
        fi
    fi
    [[ -z "$OS_VER" ]] && OS_VER="Linux"
    DOMAIN_NAME=""; IP_ADDR=""
    if [[ -z "$HOST_NAME" ]]; then
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} $(t hostname_prompt) [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi
    print_host_info_table
}

# Raccolta info host per volumi macOS (da SystemVersion.plist)
gather_host_info_macos() {
    info "$(t retrieving_info)"
    local SV
    SV=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "System/Library/CoreServices")" "SystemVersion.plist")
    if [[ -n "$SV" && -f "$SV" ]]; then
        local PLIST_TXT; PLIST_TXT=$(read_plist "$SV" 2>/dev/null)
        local PROD VER BUILD
        PROD=$(echo "$PLIST_TXT"  | grep -i 'ProductName:'        | head -1 | sed "s/.*: '\\?//; s/'\\?$//")
        VER=$(echo "$PLIST_TXT"   | grep -i 'ProductVersion:'     | head -1 | sed "s/.*: '\\?//; s/'\\?$//")
        BUILD=$(echo "$PLIST_TXT" | grep -i 'ProductBuildVersion:'| head -1 | sed "s/.*: '\\?//; s/'\\?$//")
        OS_VER=$(echo "${PROD} ${VER} (${BUILD})" | sed 's/  */ /g; s/ ()//')
    fi
    [[ -z "$OS_VER" ]] && OS_VER="macOS"
    # hostname: prova preferences.plist di SystemConfiguration
    local PREF
    PREF=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "Library/Preferences/SystemConfiguration")" "preferences.plist")
    if [[ -n "$PREF" && -f "$PREF" ]]; then
        HOST_NAME=$(read_plist "$PREF" 2>/dev/null | grep -iE 'HostName:|LocalHostName:|ComputerName:' | head -1 | sed "s/.*: '\\?//; s/'\\?$//")
    fi
    DOMAIN_NAME=""; IP_ADDR=""
    if [[ -z "$HOST_NAME" ]]; then
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} $(t hostname_prompt) [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi
    print_host_info_table
}

print_host_info_table() {
    local TITLE="$([ "$LANG" = "it" ] && echo "INFORMAZIONI MACCHINA TARGET" || echo "TARGET MACHINE INFORMATION")"
    echo -e "  ${CYAN}${BOLD}┌────────────────────────────────────────────────────────────────────────────┐${RESET}"
    local _T=${#TITLE} _LEFT=$(( (76 - ${#TITLE}) / 2 )) _RIGHT
    _RIGHT=$(( 76 - ${#TITLE} - _LEFT ))
    printf "  ${CYAN}${BOLD}│%*s%*s│${RESET}\n" $(( _LEFT + _T )) "$TITLE" $_RIGHT ""
    echo -e "  ${CYAN}${BOLD}├────────────────────────────────────────────────────────────────────────────┤${RESET}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : ${WHITE}${BOLD}%-52s${RESET} ${CYAN}${BOLD}│${RESET}\n" "$(t hostname)" "${HOST_NAME:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "$(t os)" "${OS_VER:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "$(t ip)" "${IP_ADDR:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "$(t domain)" "${DOMAIN_NAME:-N/A}"
    echo -e "  ${CYAN}${BOLD}└────────────────────────────────────────────────────────────────────────────┘${RESET}"
    echo ""
}

# Controlla se regipy è disponibile tramite il Python rilevato all'avvio ($PY3)
check_regipy() {
    "$PY3" -c "import regipy" 2>/dev/null
}

# ----------------------------------------------------------------
#  Replay dei transaction log del registro (.LOG1 / .LOG2)
#
#  Windows non scrive immediatamente nell'hive primario: accoda le modifiche
#  nei transaction log e le consolida solo a uno smontaggio pulito. Un hive
#  acquisito da una macchina spenta a caldo, da un'immagine o da uno snapshot
#  e' quindi quasi sempre "dirty": le scritture piu' recenti — tipicamente
#  proprio quelle dell'attaccante — esistono SOLO nei .LOG1/.LOG2.
#
#  Parsare l'hive grezzo le perde senza emettere alcun errore: il modulo
#  produce un report che sembra completo ma non lo e'. Qui i log vengono
#  riapplicati su una COPIA in area temporanea; il volume di evidenza resta
#  intatto e in sola lettura.
#
#  Nota: regipy scriverebbe di default l'hive ricostruito accanto
#  all'originale (<hive>.restored), cioe' SUL volume di evidenza. Il percorso
#  di destinazione viene quindi sempre passato esplicitamente.
# ----------------------------------------------------------------

# Area temporanea per gli hive ricostruiti.
#
# recover_hive viene invocata quasi sempre dentro una command substitution,
# quindi gira in una subshell: variabili globali e array modificati li' non
# risalgono al processo padre. Il percorso deve percio' essere DETERMINISTICO
# e non memorizzato in una variabile, altrimenti ogni chiamata creerebbe una
# nuova directory — rifacendo il replay ogni volta e, peggio, lasciando in
# giro copie di hive (che contengono credenziali) mai ripulite.
#
# $$ resta il PID della shell principale anche nelle subshell, quindi
# identifica la sessione in modo stabile.
_hive_replay_tmpdir() {
    local D="${TMPDIR:-/tmp}/fiuto_hives_$$"
    if [[ ! -d "$D" ]]; then
        mkdir -p "$D" 2>/dev/null || return 1
        chmod 700 "$D" 2>/dev/null || true   # gli hive ricostruiti contengono credenziali
    fi
    echo "$D"
}

# Percorso deterministico dell'hive ricostruito a partire dall'originale.
# Fa anche da cache: se il file esiste gia', il replay e' gia' stato fatto.
_hive_recovered_path() {
    local ORIG="$1" TMPD
    TMPD=$(_hive_replay_tmpdir) || return 1
    printf '%s/%s_%s.recovered' "$TMPD" \
        "$(printf '%s' "$ORIG" | sha256sum | cut -c1-16)" "$(basename "$ORIG")"
}

# Annota l'esito del replay su file: gli array in subshell non sopravvivono.
_hive_replay_note() {
    local TMPD; TMPD=$(_hive_replay_tmpdir) || return 0
    printf '%s\t%s\t%s\n' "$1" "$2" "$3" >> "${TMPD}/replay.log" 2>/dev/null || true
}

# Righe di esito del replay raccolte finora (una per hive, formato TSV).
hive_replay_report() {
    local TMPD="${TMPDIR:-/tmp}/fiuto_hives_$$"
    [[ -f "${TMPD}/replay.log" ]] && sort -u "${TMPD}/replay.log"
}

# recover_hive <path_hive>
# Stampa il percorso dell'hive da parsare: la copia ricostruita se erano
# presenti transaction log applicabili, altrimenti l'originale.
# Non fallisce mai: in caso di problema ripiega sull'originale e lo annota.
recover_hive() {
    local ORIG="$1"
    [[ -n "$ORIG" && -f "$ORIG" ]] || { echo ""; return; }
    evidence_note "$ORIG" "hive di registro"
    [[ "$HIVE_REPLAY" == "true" ]] || { echo "$ORIG"; return; }

    local OUT; OUT=$(_hive_recovered_path "$ORIG") || { echo "$ORIG"; return; }

    # Cache su disco: sopravvive alle subshell delle command substitution.
    [[ -s "$OUT" ]] && { echo "$OUT"; return; }
    # Marcatore "gia' valutato, si usa l'originale" (hive pulito o replay fallito):
    # evita di ritentare a ogni modulo.
    [[ -f "${OUT}.skip" ]] && { echo "$ORIG"; return; }

    # Con --jobs piu' moduli possono chiedere lo stesso hive nello stesso
    # istante. Senza lock il secondo leggerebbe la copia ricostruita mentre il
    # primo la sta ancora scrivendo: un hive troncato non da' errore, da'
    # risultati parziali — che e' peggio. mkdir e' atomico anche su NFS.
    local LOCK="${OUT}.lock" _held=false _tries=0
    while true; do
        if mkdir "$LOCK" 2>/dev/null; then _held=true; break; fi
        _tries=$((_tries + 1))
        # Un lock orfano (processo ucciso) non deve bloccare la sessione: dopo
        # due minuti si procede comunque, nel caso peggiore rifacendo il lavoro.
        [[ $_tries -gt 240 ]] && break
        sleep 0.5
        # Nel frattempo puo' aver finito qualcun altro.
        [[ -s "$OUT" ]] && { echo "$OUT"; return; }
        [[ -f "${OUT}.skip" ]] && { echo "$ORIG"; return; }
    done
    # Rilascia il lock solo chi lo detiene davvero: dopo un timeout il lock e'
    # di un altro processo, e rimuoverlo aprirebbe la corsa che il lock evita.
    _unlock() { [[ "$_held" == "true" ]] && rmdir "$LOCK" 2>/dev/null; return 0; }

    local DIR BASE LOG1 LOG2
    DIR=$(dirname "$ORIG")
    BASE=$(basename "$ORIG")
    LOG1=$(find "$DIR" -maxdepth 1 -iname "${BASE}.LOG1" -type f 2>/dev/null | head -1)
    LOG2=$(find "$DIR" -maxdepth 1 -iname "${BASE}.LOG2" -type f 2>/dev/null | head -1)

    # Nessun log, o log vuoti: l'hive e' gia' consolidato.
    if [[ ( -z "$LOG1" || ! -s "$LOG1" ) && ( -z "$LOG2" || ! -s "$LOG2" ) ]]; then
        : > "${OUT}.skip"
        _hive_replay_note "clean" "$BASE" "$(L "nessun transaction log da applicare" "no transaction log to apply")"
        _unlock; echo "$ORIG"
        return
    fi

    if ! check_regipy; then
        # L'avviso va emesso una sola volta per sessione: il marcatore sta su
        # disco perche' una variabile non sopravvivrebbe alla subshell.
        local WARNED; WARNED="$(_hive_replay_tmpdir)/.regipy_warned"
        if [[ ! -f "$WARNED" ]]; then
            : > "$WARNED"
            warn "$(L "regipy assente: i transaction log del registro non verranno applicati (dati recenti potenzialmente mancanti)." \
                     "regipy missing: registry transaction logs will not be applied (recent data potentially missing).")" >&2
        fi
        : > "${OUT}.skip"
        _hive_replay_note "skipped" "$BASE" "$(L "regipy non disponibile" "regipy unavailable")"
        _unlock; echo "$ORIG"
        return
    fi

    local RES
    RES=$("$PY3" - "$ORIG" "$OUT" "${LOG1:-}" "${LOG2:-}" << 'PYEOF' 2>/dev/null
import sys, os

hive_path, out_path = sys.argv[1], sys.argv[2]
log1 = sys.argv[3] if len(sys.argv) > 3 and sys.argv[3] else None
log2 = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] else None

# Se esiste solo il secondario, va passato come primario: regipy richiede
# comunque un primary_log_path valido.
if not log1 and log2:
    log1, log2 = log2, None

try:
    from regipy.recovery import apply_transaction_logs
    restored, dirty = apply_transaction_logs(
        hive_path, log1, secondary_log_path=log2, restored_hive_path=out_path
    )
    if restored and os.path.isfile(restored) and os.path.getsize(restored) > 0:
        print(f"OK\t{dirty}")
    else:
        print("FAIL\tnessun hive ricostruito prodotto")
except Exception as exc:
    print(f"FAIL\t{exc}")
PYEOF
    )

    local STATUS DETAIL
    STATUS="${RES%%	*}"
    DETAIL="${RES#*	}"

    if [[ "$STATUS" == "OK" && -s "$OUT" ]]; then
        _hive_replay_note "recovered" "$BASE" "${DETAIL} $(L "pagine dirty riapplicate" "dirty pages replayed")"
        info "$(L "Transaction log applicati a" "Transaction logs applied to") ${BOLD}${BASE}${RESET} — ${DETAIL} $(L "pagine dirty" "dirty pages")" >&2
        log_msg "[HIVE] replay OK: $ORIG -> $OUT (${DETAIL} dirty pages)"
        _unlock; echo "$OUT"
        return
    fi

    # Fallback non distruttivo: si continua sull'hive originale.
    rm -f "$OUT" 2>/dev/null
    : > "${OUT}.skip"
    _hive_replay_note "failed" "$BASE" "$DETAIL"
    warn "$(L "Replay dei transaction log fallito per" "Transaction log replay failed for") ${BASE}: ${DETAIL}" >&2
    log_msg "[HIVE] replay FAILED: $ORIG — $DETAIL"
    _unlock; echo "$ORIG"
}

# Torna il percorso di un hive di sistema, con i transaction log gia' applicati.
get_hive() {
    local NAME="$1"   # SOFTWARE, SYSTEM, SECURITY, SAM
    local FOUND=""
    local PATH1="$WIN_ROOT/Windows/System32/config/${NAME}"
    if [[ -f "$PATH1" ]]; then
        FOUND="$PATH1"
    else
        local PATH2
        PATH2=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
        [[ -n "$PATH2" ]] && FOUND=$(ci_find_file "$PATH2" "$NAME")
    fi
    [[ -n "$FOUND" ]] && recover_hive "$FOUND"
}

# Torna l'hive di un utente (NTUSER.DAT, UsrClass.dat) con i log applicati.
# $1 = directory da cui cercare · $2 = nome hive
get_user_hive() {
    local FOUND; FOUND=$(ci_find_file "$1" "$2")
    [[ -n "$FOUND" ]] && recover_hive "$FOUND"
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
#
# E' anche il punto di aggancio dell'export JSONL: ogni modulo passa di qui
# dopo aver scritto il proprio HTML, quindi l'export copre automaticamente
# tutti i moduli — inclusi quelli Windows, che generano l'HTML per conto
# proprio senza passare da finish_report.
# Effetti collaterali di un report appena scritto: export JSONL e copia
# oscurata. Stanno in una funzione a se' perche' l'esecuzione batch
# RIDEFINISCE register_report nel processo figlio — gli array bash non
# risalgono — e quella ridefinizione deve poterli richiamare.
#
# Non e' un dettaglio di stile: finche' gli effetti sono stati scritti dentro
# register_report, `--all --jsonl` non produceva alcun JSONL. Nessun errore,
# nessun avviso, semplicemente i file non c'erano.
_report_side_effects() {
    [[ "$EXPORT_JSONL" == "true" ]] && export_report_jsonl "$1"
    # La copia oscurata nasce qui e non dentro finish_report: i moduli Windows
    # scrivono l'HTML per conto proprio e passano solo di qua.
    [[ "$REDACT" == "true" ]] && redact_report "$1"
    return 0
}

register_report() {
    [[ -n "${1:-}" && -f "$1" ]] || return 0
    GENERATED_REPORTS+=("$1")
    _report_side_effects "$1"
    return 0
}

# HTML boilerplate header comune a tutti i report
html_header() {
    local TITLE="$1"
    local ACCENT="${2:-#58a6ff}"
    cat << HTMLEOF
<!DOCTYPE html>
<html lang="$(L "it" "en")">
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
  /* ── Barra di ricerca / Search bar ── */
  #fz-toolbar { position:fixed; top:.7rem; right:.9rem; z-index:50; display:flex; align-items:center;
    gap:.4rem; font-family:var(--mono); }
  #fz-csv { display:flex; align-items:center; gap:.35rem; background:var(--bg3);
    border:1px solid var(--border); border-radius:6px; color:var(--text); cursor:pointer;
    font-family:var(--mono); font-size:.74rem; padding:.4rem .6rem; box-shadow:0 4px 18px rgba(0,0,0,.45); }
  #fz-csv:hover { border-color:var(--accent3); color:var(--accent3); }
  #fz-csv svg { width:.95rem; height:.95rem; }
  #fz-search { display:flex; align-items:center; gap:.4rem; background:var(--bg3);
    border:1px solid var(--border); border-radius:6px; padding:.35rem .5rem;
    box-shadow:0 4px 18px rgba(0,0,0,.45); }
  #fz-search input { background:var(--bg); border:1px solid var(--border); color:var(--text);
    font-family:var(--mono); font-size:.78rem; padding:.3rem .55rem; border-radius:4px; width:15rem;
    outline:none; }
  #fz-search input:focus { border-color:var(--accent); }
  #fz-search button { background:var(--bg4); border:1px solid var(--border); color:var(--text);
    font-family:var(--mono); font-size:.8rem; line-height:1; padding:.28rem .5rem; border-radius:4px;
    cursor:pointer; }
  #fz-search button:hover { border-color:var(--accent); color:var(--accent); }
  #fz-close:hover { border-color:var(--accent2); color:var(--accent2); }
  #fz-count { font-size:.7rem; color:var(--text-dim); min-width:4.2rem; text-align:center; }
  #fz-search.fz-hidden { display:none; }
  #fz-toggle { display:none; align-items:center; justify-content:center; width:2.1rem; height:2.1rem;
    background:var(--bg3); border:1px solid var(--border); border-radius:6px; color:var(--text-mid);
    cursor:pointer; box-shadow:0 4px 18px rgba(0,0,0,.45); }
  #fz-toggle.fz-show { display:flex; }
  #fz-toggle:hover { border-color:var(--accent); color:var(--accent); }
  #fz-toggle svg { width:1.05rem; height:1.05rem; }
  mark.fz-hit { background:rgba(240,136,62,.32); color:inherit; border-radius:2px; padding:0 .05em; }
  mark.fz-hit.fz-active { background:var(--accent); color:#08121f; }
</style>
</head>
<body>
<div id="fz-toolbar">
  <button id="fz-csv" title="$(L "Scarica tutti i dati in CSV" "Download all data as CSV")"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"></path><path d="M7 11l5 5 5-5"></path><path d="M5 21h14"></path></svg>CSV</button>
  <div id="fz-search">
    <input id="fz-q" type="search" placeholder="$(L "Cerca nel report…" "Search report…")" autocomplete="off" spellcheck="false">
    <span id="fz-count"></span>
    <button id="fz-prev" title="$(L "Precedente" "Previous")">▲</button>
    <button id="fz-next" title="$(L "Successivo" "Next")">▼</button>
    <button id="fz-close" title="$(L "Chiudi ricerca" "Close search")">✕</button>
  </div>
  <button id="fz-toggle" title="$(L "Cerca nel report" "Search report")"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"></circle><line x1="21" y1="21" x2="16.5" y2="16.5"></line></svg></button>
</div>
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
    <div>$(L "Scansione:" "Scan:") <span class="val">${SCAN}</span></div>
    <div>$(L "Sorgente:" "Source:") <span class="val">${SRC}</span></div>$([ -n "$HASH" ] && echo "
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
  <div>fiuto.sh — $(L "generato il" "generated on") ${SCAN}</div>
  <div>$(L "Sorgente:" "Source:") <span>${SRC}</span></div>
</footer>
HTMLEOF
    # Script di ricerca/evidenziazione (heredoc quotato: nessuna espansione shell)
    cat << 'JSEOF'
<script>
(function(){
  var q=document.getElementById('fz-q'), cnt=document.getElementById('fz-count'),
      prev=document.getElementById('fz-prev'), next=document.getElementById('fz-next'),
      bar=document.getElementById('fz-search'), closeBtn=document.getElementById('fz-close'),
      toggle=document.getElementById('fz-toggle'), csvBtn=document.getElementById('fz-csv');
  // ── Export CSV: serializza TUTTE le tabelle e i blocchi <pre> del report ──
  function csvCell(v){ v=(v==null?'':String(v)); return /[",\r\n]/.test(v) ? '"'+v.replace(/"/g,'""')+'"' : v; }
  function exportCSV(){
    var root=document.querySelector('main')||document.body;
    var nodes=root.querySelectorAll('table, pre');
    var lines=[], lastLabel='';
    nodes.forEach(function(node){
      var card=node.closest ? node.closest('.card') : null, label='';
      if(card){
        var u=card.querySelector('.uname'), p=card.querySelector('.upath');
        label=[u&&u.textContent.trim(), p&&p.textContent.trim()].filter(Boolean).join(' — ');
      }
      if(label && label!==lastLabel){ lines.push('# '+csvCell(label)); lastLabel=label; }
      if(node.tagName==='TABLE'){
        node.querySelectorAll('tr').forEach(function(tr){
          var row=[];
          tr.querySelectorAll('th,td').forEach(function(c){ row.push(csvCell(c.textContent.replace(/\s+/g,' ').trim())); });
          if(row.length) lines.push(row.join(','));
        });
      } else {
        node.textContent.split('\n').forEach(function(l){
          l=l.replace(/^\s*\d+\s/,'').replace(/\s+$/,'');   // toglie il numero di riga
          if(l.trim()!=='') lines.push(csvCell(l));
        });
      }
      lines.push('');
    });
    if(!lines.length){ return; }
    var blob=new Blob(['\ufeff'+lines.join('\r\n')], {type:'text/csv;charset=utf-8'});
    var name=(document.title||'fiuto_report').replace(/\s*—.*$/,'').replace(/[^\w.-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,80)||'fiuto_report';
    var a=document.createElement('a');
    a.href=URL.createObjectURL(blob); a.download=name+'.csv';
    document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 200);
  }
  if(csvBtn) csvBtn.addEventListener('click', exportCSV);
  if(!q) return;
  var hits=[], idx=-1, timer=null;
  function openBar(){ bar.classList.remove('fz-hidden'); if(toggle) toggle.classList.remove('fz-show'); q.focus(); q.select(); }
  function closeBar(){ q.value=''; clearMarks(); cnt.textContent=''; bar.classList.add('fz-hidden'); if(toggle) toggle.classList.add('fz-show'); }
  function esc(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
  function clearMarks(){
    document.querySelectorAll('mark.fz-hit').forEach(function(m){
      var p=m.parentNode; if(!p) return;
      p.replaceChild(document.createTextNode(m.textContent), m); p.normalize();
    });
    hits=[]; idx=-1;
  }
  function activate(){
    hits.forEach(function(m){ m.classList.remove('fz-active'); });
    if(idx>=0 && idx<hits.length){
      var m=hits[idx]; m.classList.add('fz-active');
      m.scrollIntoView({block:'center', behavior:'smooth'});
      cnt.textContent=(idx+1)+'/'+hits.length;
    }
  }
  function go(d){ if(!hits.length) return; idx=(idx+d+hits.length)%hits.length; activate(); }
  function search(term){
    clearMarks();
    if(!term || term.length<2){ cnt.textContent=''; return; }
    var reTest=new RegExp(esc(term),'i'), reG=new RegExp(esc(term),'gi');
    var walker=document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
      acceptNode:function(node){
        if(!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        var p=node.parentNode;
        while(p){
          if(p.id==='fz-search') return NodeFilter.FILTER_REJECT;
          var t=p.nodeName;
          if(t==='SCRIPT'||t==='STYLE'||t==='MARK') return NodeFilter.FILTER_REJECT;
          p=p.parentNode;
        }
        return reTest.test(node.nodeValue)?NodeFilter.FILTER_ACCEPT:NodeFilter.FILTER_REJECT;
      }
    });
    var nodes=[], n;
    while(n=walker.nextNode()) nodes.push(n);
    nodes.forEach(function(node){
      var text=node.nodeValue, frag=document.createDocumentFragment(), last=0, m;
      reG.lastIndex=0;
      while((m=reG.exec(text))){
        if(m.index>last) frag.appendChild(document.createTextNode(text.slice(last,m.index)));
        var mk=document.createElement('mark'); mk.className='fz-hit'; mk.textContent=m[0];
        frag.appendChild(mk); hits.push(mk);
        last=m.index+m[0].length;
        if(m[0].length===0) reG.lastIndex++;
      }
      if(last<text.length) frag.appendChild(document.createTextNode(text.slice(last)));
      if(node.parentNode) node.parentNode.replaceChild(frag, node);
    });
    if(hits.length){ idx=0; activate(); } else { cnt.textContent='0'; }
  }
  q.addEventListener('input', function(){ clearTimeout(timer); timer=setTimeout(function(){ search(q.value); }, 200); });
  q.addEventListener('keydown', function(e){
    if(e.key==='Enter'){ e.preventDefault(); go(e.shiftKey?-1:1); }
    else if(e.key==='Escape'){ e.preventDefault(); closeBar(); }
  });
  prev.addEventListener('click', function(){ go(-1); });
  next.addEventListener('click', function(){ go(1); });
  if(closeBtn) closeBtn.addEventListener('click', closeBar);
  if(toggle) toggle.addEventListener('click', openBar);
  document.addEventListener('keydown', function(e){
    if((e.ctrlKey||e.metaKey) && (e.key==='f'||e.key==='F')){ e.preventDefault(); openBar(); }
  });
})();
</script>
</body></html>
JSEOF
}

# ================================================================
#  MENU PRINCIPALE
# ================================================================
# ================================================================
#  HELPER CONDIVISI PER I MODULI LINUX / macOS
# ================================================================

# Blocco <style> per i <pre> con numeri di riga ed evidenziazione (riuso dal modulo PS)
pre_style_block() {
    cat << 'EOF'
<style>
  .hist-pre{font-family:var(--mono);font-size:.75rem;line-height:1.7;padding:.8rem 1rem;
    overflow-x:auto;max-height:520px;overflow-y:auto;}
  .hist-pre::-webkit-scrollbar{width:5px;height:5px}
  .hist-pre::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
  .line{display:block;color:var(--text);white-space:pre;padding:.05rem .5rem;border-radius:2px}
  .line:hover{background:rgba(88,166,255,.05)}
  .line.sensitive{color:var(--accent2);background:rgba(255,123,114,.07);
    border-left:2px solid rgba(255,123,114,.5);padding-left:calc(.5rem - 2px)}
  .lnum{color:var(--text-dim);user-select:none;margin-right:1rem;font-size:.7rem}
  .grp{margin-bottom:1.5rem}
</style>
EOF
}

# Genera il contenuto di un <pre> (numeri di riga + escape HTML + evidenziazione IoC).
# $1 = file, $2 = keyword separate da '|' (case-insensitive) per marcare le righe sensibili.
render_pre_block() {
    local FILE="$1" KW="$2" MODE="${3:-}"
    local _PTMP; _PTMP=$(mktemp)
    FIUTO_SINCE="${TIME_SINCE:-}" FIUTO_UNTIL="${TIME_UNTIL:-}" \
    "$PY3" - "$FILE" "$KW" "$MODE" "$_PTMP" << 'PYEOF'
import sys, os, html, re, datetime
path, kw = sys.argv[1], sys.argv[2].lower()
mode = sys.argv[3] if len(sys.argv) > 3 else ''
drop_file = sys.argv[4] if len(sys.argv) > 4 else ''
keys = [k for k in kw.split('|') if k]

# Finestra --since/--until. I blocchi <pre> sono log e history: senza filtro
# qui l'HTML mostrerebbe righe che l'export JSONL, che gia' filtra, esclude —
# due viste dello stesso modulo che si contraddicono.
since = os.environ.get('FIUTO_SINCE', '')
until = os.environ.get('FIUTO_UNTIL', '')
_TS_ISO = re.compile(r'\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?')
_MONTHS = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
           'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
_TS_SYS = re.compile(r'\b(' + '|'.join(_MONTHS) + r')\s+(\d{1,2})\s+(\d{2}:\d{2}:\d{2})')
_YEAR = str(datetime.date.today().year)

def _line_in_window(line):
    """None se la riga non porta date: non valutabile, quindi si tiene."""
    found = []
    for m in _TS_ISO.finditer(line):
        t = m.group(0).replace(' ', 'T')
        found.append(t + 'T00:00:00'[len(t) - 10:] if len(t) < 19 else t)
    if not found:
        # syslog non scrive l'anno: si assume quello corrente, come fa
        # l'export JSONL. Approssimazione dichiarata, non nascosta.
        m = _TS_SYS.search(line)
        if m:
            found.append(f"{_YEAR}-{_MONTHS[m.group(1)]}-{int(m.group(2)):02d}T{m.group(3)}")
    if not found:
        return None
    return any((not since or t >= since) and (not until or t <= until) for t in found)

# Decodifica i timestamp UNIX nelle history di shell in formato leggibile.
# zsh extended_history:  ": <epoch>:<elapsed>;<comando>"
# bash con HISTTIMEFORMAT: una riga "#<epoch>" prima del comando
_ZSH = re.compile(r'^: (\d{9,12}):(\d+);(.*)$', re.S)
_BASH = re.compile(r'^#(\d{9,12})$')

# I REPL basati su GNU readline (python3, node, psql) scrivono le voci
# multi-riga con gli spazi e i backslash codificati in ottale: una riga di
# codice indentata diventa "\040\040\040\040value = ..." e il report la
# mostrerebbe cosi', illeggibile. Qui si decodifica.
#
# La sostituzione e' UNA sola passata con callback e non una catena di
# replace: decodificando prima \134 (backslash) e poi gli altri si
# reinterpreterebbero come escape i backslash appena prodotti.
_OCTAL = re.compile(r'\\([0-7]{3})')


def decode_readline(line):
    return _OCTAL.sub(lambda m: chr(int(m.group(1), 8)), line)
def fmt(ep):
    try:
        return datetime.datetime.utcfromtimestamp(int(ep)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ep
def decode_histts(line):
    m = _ZSH.match(line)
    if m:
        return f"[{fmt(m.group(1))}]  {m.group(3)}"
    m = _BASH.match(line)
    if m:
        return f"[{fmt(m.group(1))}]"
    return line

try:
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('utf-8', 'replace').replace('\r\n', '\n').replace('\r', '\n')
    out = []
    dropped = 0
    for i, line in enumerate(text.split('\n'), 1):
        if mode == 'histts':
            line = decode_histts(line)
        elif mode == 'histrl':
            line = decode_readline(line)
        if (since or until) and _line_in_window(line) is False:
            # Il numero di riga resta quello del file: i salti nella
            # numerazione rendono visibile che qualcosa e' stato tolto.
            dropped += 1
            continue
        esc = html.escape(line)
        css = 'line sensitive' if any(k in line.lower() for k in keys) else 'line'
        out.append(f'<span class="{css}"><span class="lnum">{i:5d}</span> {esc}</span>')
    print('\n'.join(out))
    if drop_file and dropped:
        with open(drop_file, 'w') as fh:
            fh.write(str(dropped))
except Exception as e:
    print(f'<span class="line bad">{html.escape(str(e))}</span>')
PYEOF
    if [[ -s "$_PTMP" ]]; then
        local _D; _D=$(cat "$_PTMP")
        time_filtered_add "$_D"
        printf "\n<span class='line' style='color:var(--text-mid)'>      %s %s</span>" \
            "$_D" "$(L "righe nascoste dal filtro --since/--until" "rows hidden by the --since/--until filter")"
    fi
    rm -f "$_PTMP"
}

# Stampa a console le righe di un file con evidenziazione IoC (rosso sulle corrispondenze).
# $1 = file, $2 = regex grep (-iE), $3 = max righe (default 200)
print_file_lines() {
    local FILE="$1" KW="$2" MAX="${3:-200}"
    [[ -f "$FILE" ]] || return
    local LN=0
    while IFS= read -r LINE || [[ -n "$LINE" ]]; do
        LN=$((LN + 1))
        if [[ $LN -gt $MAX ]]; then
            echo -e "      ${DIM}... ($(L "troncato a" "truncated at") $MAX $(L "righe" "lines"))${RESET}"
            break
        fi
        if [[ -n "$KW" ]] && printf '%s' "$LINE" | grep -qiE "$KW"; then
            printf "      ${RED}%5d  %s${RESET}\n" "$LN" "$LINE"
        else
            printf "      ${DIM}%5d${RESET}  %s\n" "$LN" "$LINE"
        fi
    done < "$FILE"
}

# Card HTML per un singolo file di testo (header con metadati + <pre> evidenziato).
# $1 = file, $2 = keyword IoC, $3 = icona (default ≣)
file_card_html() {
    local F="$1" KW="$2" ICON="${3:-≣}" MODE="${4:-}"
    evidence_note "$F" "riprodotto nel report"
    local SZ MT BODY
    SZ=$(stat -c %s "$F" 2>/dev/null || echo "?")
    MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1 || echo "?")
    BODY=$(render_pre_block "$F" "$KW" "$MODE")
    printf "<div class='card' style='margin-bottom:.8rem'><div class='card-header'><div class='uicon' style='font-size:.7rem'>%s</div><div><div class='uname' style='font-size:.85rem'>%s</div><div class='upath'>%s</div></div><div style='margin-left:auto;text-align:right;font-family:var(--mono);font-size:.65rem;color:var(--text-dim)'><div class='mid'>%s</div><div>%s B</div></div></div><div class='hist-content'><pre class='hist-pre'>%s</pre></div></div>" \
        "$ICON" "$(html_esc "$(basename "$F")")" "$(html_esc "$F")" "$MT" "$SZ" "$BODY"
}

# Card HTML generica con corpo arbitrario (tabella/pre già formattati).
# $1 = titolo, $2 = sottopath, $3 = badge, $4 = corpo HTML, $5 = icona
generic_card_html() {
    printf "<div class='card'><div class='card-header'><div class='uicon'>%s</div><div class='user-info'><div class='uname'>%s</div><div class='upath'>%s</div></div><div class='badge'>%s</div></div><div style='padding:1rem 1.5rem'>%s</div></div>" \
        "${5:-▣}" "$(html_esc "$1")" "$(html_esc "$2")" "$3" "$4"
}

# Scrive il report HTML finale e lo registra.
# $1 slug · $2 titolo · $3 icona · $4 sottotitolo · $5 stats_html · $6 body_html
finish_report() {
    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "$1")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    {
        html_header "$2"
        html_page_header "$3" "$2" "$4" "$SCAN" "$WIN_ROOT"
        [[ -n "$5" ]] && printf "<div class='statsbar'>%s</div>\n" "$5"
        echo "<main>"
        pre_style_block
        time_window_html
        printf '%s\n' "$6"
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# Helper per una stat della statsbar
stat_box() { printf "<div class='stat %s'><div class='label'>%s</div><div class='value'>%s</div></div>" "${3:-}" "$1" "$2"; }

# ================================================================
#  MODULI macOS
# ================================================================

# Renderizza una tabella HTML da righe tab-separated. $1=righe, $2.. = intestazioni
#
# E' anche il punto in cui si applica la finestra --since/--until: e' la
# funzione che quasi tutti i moduli usano per emettere dati datati, quindi il
# filtro copre l'intero toolkit senza toccare i moduli uno per uno.
_rows_to_table() {
    local ROWS="$1"; shift
    local _RTMP; _RTMP=$(mktemp); printf '%s\n' "$ROWS" > "$_RTMP"
    printf '%s\n' "$@" > "${_RTMP}.h"
    FIUTO_SINCE="${TIME_SINCE:-}" FIUTO_UNTIL="${TIME_UNTIL:-}" \
    "$PY3" - "$_RTMP" "${_RTMP}.h" "${_RTMP}.d" << 'PYEOF'
import sys, os, re, html
heads=[h.rstrip('\n') for h in open(sys.argv[2])]

since = os.environ.get('FIUTO_SINCE', '')
until = os.environ.get('FIUTO_UNTIL', '')
# Le date compaiono nelle forme piu' varie ma quasi sempre iniziano con
# YYYY-MM-DD: ci si limita a quelle, perche' un pattern piu' permissivo
# scarterebbe righe sulla base di numeri che date non sono.
TS = re.compile(r'\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?')

def in_window(cells):
    """None se la riga non porta date: non e' valutabile, quindi si tiene."""
    found = False
    for c in cells:
        for m in TS.finditer(c):
            found = True
            # Confronto lessicografico: su ISO 8601 equivale a quello
            # cronologico, e non richiede di parsare formati parziali.
            t = m.group(0).replace(' ', 'T')
            # Completa le forme parziali ("2026-03-01", "2026-03-01T10:30")
            # con l'inizio del periodo, cosi' tutte hanno la stessa lunghezza.
            t = t + 'T00:00:00'[len(t) - 10:] if len(t) < 19 else t
            if since and t < since:
                continue
            if until and t > until:
                continue
            return True
    return False if found else None

print("<table><tr>"+''.join(f'<th>{html.escape(h)}</th>' for h in heads)+"</tr>")
dropped = 0
for line in open(sys.argv[1], errors='replace'):
    if not line.strip(): continue
    cells=line.rstrip('\n').split('\t')
    if (since or until) and in_window(cells) is False:
        dropped += 1
        continue
    tds=''.join(f"<td class='mono'>{html.escape(c)}</td>" for c in cells)
    print(f"<tr>{tds}</tr>")
print("</table>")
with open(sys.argv[3], 'w') as fh:
    fh.write(str(dropped))
PYEOF
    if [[ -s "${_RTMP}.d" ]]; then
        local _DROP; _DROP=$(cat "${_RTMP}.d")
        if [[ "$_DROP" -gt 0 ]]; then
            time_filtered_add "$_DROP"
            printf "<div style='font-size:.72rem;color:var(--text-mid);padding:.4rem 0'>%s %s</div>" \
                "$_DROP" "$(L "righe nascoste dal filtro --since/--until" "rows hidden by the --since/--until filter")"
        fi
    fi
    rm -f "$_RTMP" "${_RTMP}.h" "${_RTMP}.d"
}

# ----------------------------------------------------------------
#  Export JSONL (schema Timesketch / plaso)
#
#  I report HTML sono ottimi per l'analista e inutilizzabili per una
#  pipeline: non si correlano con altre sorgenti e non si caricano in un
#  SIEM. Qui gli stessi eventi vengono riemessi in JSON Lines con i campi
#  attesi da Timesketch (datetime, timestamp_desc, message), cosi' il
#  risultato di FIUTO entra direttamente in una super-timeline.
#
#  L'estrazione lavora sull'HTML gia' prodotto invece che sui dati grezzi
#  dei singoli moduli: e' l'unico punto in cui il formato e' omogeneo per
#  tutti e 91 i moduli, e non richiede di toccarli uno per uno. Il codice
#  che la esegue sta in pylib_timeline, condiviso con l'executive summary.
# ----------------------------------------------------------------
export_report_jsonl() {
    local HTML="$1"
    local DIR; DIR=$(dirname "$HTML")
    local SLUG; SLUG=$(basename "$DIR" | sed -E 's/_[0-9]{8}_[0-9]{6}$//')
    local OUT="${DIR}/report.jsonl"

    local DROPF; DROPF=$(mktemp); register_tmp "$DROPF"
    FIUTO_SINCE="${TIME_SINCE:-}" FIUTO_UNTIL="${TIME_UNTIL:-}" FIUTO_TZ="${VOLUME_TZ:-}" \
    FIUTO_DROPFILE="$DROPF" \
    run_py_with_lib pylib_timeline "$HTML" "$SLUG" "${WIN_ROOT:-}" "${HOST_NAME:-}" "${OS_TYPE:-}" > "$OUT" << 'PYEOF' 2>/dev/null
import sys, os, json

html_path, slug = sys.argv[1], sys.argv[2]
volume  = sys.argv[3] if len(sys.argv) > 3 else ''
host    = sys.argv[4] if len(sys.argv) > 4 else ''
os_type = sys.argv[5] if len(sys.argv) > 5 else ''

# La stessa finestra applicata ai report HTML: se l'export non la rispettasse,
# la timeline caricata in Timesketch conterrebbe eventi che il report esclude.
since = os.environ.get('FIUTO_SINCE', '')
until = os.environ.get('FIUTO_UNTIL', '')
volume_tz = os.environ.get('FIUTO_TZ', '')

dropped = 0

for ev in extract_events(html_path):
    dt = ev['datetime']
    if (since and dt < since) or (until and dt > until):
        dropped += 1
        continue
    rec = {
        # Campi richiesti da Timesketch
        "datetime": dt,
        "timestamp_desc": f"FIUTO {slug}",
        "message": ev['message'],
        # Contesto aggiuntivo
        "data_type": f"fiuto:{slug}",
        "module": slug,
        "source_volume": volume,
        "hostname": host,
        "os": os_type,
    }
    if volume_tz:
        # Il campo datetime NON e' riportato a UTC: dichiarare il fuso del
        # volume e' l'unico modo perche' chi carica la timeline sappia in che
        # riferimento sono gli eventi presi dagli artefatti in ora locale.
        rec["volume_timezone"] = volume_tz
    if ev['year_inferred']:
        # L'anno non era nel dato di origine: va dichiarato, non nascosto.
        rec["year_inferred"] = True
    print(json.dumps(rec, ensure_ascii=False))

drop_file = os.environ.get('FIUTO_DROPFILE')
if drop_file and dropped:
    try:
        with open(drop_file, 'w') as fh:
            fh.write(str(dropped))
    except Exception:
        pass
PYEOF

    if [[ -s "$DROPF" ]]; then
        log_msg "[JSONL] $(cat "$DROPF") eventi esclusi dalla finestra --since/--until"
    fi
    rm -f "$DROPF"

    local N=0
    [[ -s "$OUT" ]] && N=$(wc -l < "$OUT")
    if [[ "$N" -eq 0 ]]; then
        rm -f "$OUT"
        return 0
    fi

    # Timeline unica di sessione: e' il file da caricare in Timesketch.
    # Con --jobs l'append arriva da piu' processi e le righe superano PIPE_BUF:
    # senza lock si intreccerebbero, producendo JSON non parsabile proprio nel
    # file destinato a essere caricato in un altro strumento.
    local COMBINED="${REPORT_BASE_DIR}/fiuto_timeline.jsonl"
    if command -v flock > /dev/null 2>&1; then
        flock "$COMBINED.lock" -c "cat '$OUT' >> '$COMBINED'" 2>/dev/null || \
            cat "$OUT" >> "$COMBINED" 2>/dev/null || true
    else
        cat "$OUT" >> "$COMBINED" 2>/dev/null || true
    fi
    ok "$(L "Export JSONL:" "JSONL export:") ${BOLD}${N}$(L " eventi" " events")${RESET} → $(basename "$OUT")"
    log_msg "[JSONL] $OUT — $N eventi"
}

# ================================================================
#  AUTODETECT ROOT WINDOWS
# ================================================================

debug_mounts() {
    echo ""
    section_header "$(L "DEBUG — Mount attivi su questo sistema" "DEBUG — Active Mounts on This System")" "$YELLOW"
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
    local ARROW_LABEL="$([ "$LANG" = "it" ] && echo "candidati" || echo "candidates")"
    local GRAY_LABEL="$([ "$LANG" = "it" ] && echo "esclusi" || echo "excluded")"
    echo -e "  ${DIM}(${CYAN}→${DIM} = ${ARROW_LABEL}; grigio = ${GRAY_LABEL})${RESET}"
    echo ""
    echo -ne "  ${YELLOW}$(t press_key)${RESET}"
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
        # Marcatori Windows
        if [[ -d "$MNT/Users" || -d "$MNT/Windows" ||
              -d "$MNT/users" || -d "$MNT/windows" ]]; then
            CANDIDATES+=("$MNT"); continue
        fi
        # Marcatori macOS
        if [[ -d "$MNT/System/Library/CoreServices" || -d "$MNT/private/var/db/dslocal" ]]; then
            CANDIDATES+=("$MNT"); continue
        fi
        # Marcatori Linux
        if [[ -f "$MNT/etc/os-release" || -f "$MNT/etc/passwd" ]]; then
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

# Conta gli utenti reali su un volume, in base al suo OS (per il menu di selezione)
_count_volume_users() {
    local MNT="$1" VOS="$2"
    case "$VOS" in
        windows|macos)
            local UD
            UD=$(find "$MNT" -maxdepth 1 -type d -iname "Users" 2>/dev/null | head -1)
            [[ -z "$UD" ]] && { echo 0; return; }
            find "$UD" -maxdepth 1 -mindepth 1 -type d 2>/dev/null \
                | grep -ciEv '/(Public|Default|Default User|All Users|Shared|Guest|\.localized)$' || echo 0 ;;
        linux)
            local HD C=0
            HD=$(find "$MNT" -maxdepth 1 -type d -iname "home" 2>/dev/null | head -1)
            [[ -n "$HD" ]] && C=$(find "$HD" -maxdepth 1 -mindepth 1 -type d 2>/dev/null | wc -l)
            [[ -d "$MNT/root" ]] && C=$((C + 1))
            echo "$C" ;;
        *) echo 0 ;;
    esac
}

autodetect_win_root() {
    local SILENT="${1:-}"
    local -a FOUND=()
    mapfile -t FOUND < <(_find_windows_mounts)
    if [[ ${#FOUND[@]} -eq 0 ]]; then
        [[ "$SILENT" != "silent" ]] && warn "$(L "Nessun volume analizzabile rilevato tra i filesystem montati." "No analysable volume detected among mounted filesystems.")"
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
    echo -e "  ${CYAN}${BOLD}$(L "Volumi rilevati:" "Detected volumes:")${RESET}"
    echo ""
    local IDX=1
    for MNT in "${FOUND[@]}"; do
        local LABEL; LABEL=$(basename "$MNT")
        local VOS; VOS=$(detect_os_type "$MNT")
        local BADGE_COLOR
        case "$VOS" in
            windows) BADGE_COLOR="$BLUE" ;;
            linux)   BADGE_COLOR="$YELLOW" ;;
            macos)   BADGE_COLOR="$WHITE" ;;
            *)       BADGE_COLOR="$DIM" ;;
        esac
        local USER_COUNT
        USER_COUNT=$(_count_volume_users "$MNT" "$VOS")
        echo -e "  ${GREEN}[${IDX}]${RESET}  ${BOLD}${MNT}${RESET} ${MAGENTA}(${LABEL})${RESET}  ${BADGE_COLOR}[$(os_label "$VOS")]${RESET}  ${CYAN}${USER_COUNT} $(L "utenti" "users")${RESET}"
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
        echo -e "  ${YELLOW}[${IDX}]${RESET}  ${BOLD}${WOLD}${RESET}  ${YELLOW}★ Windows.old${RESET} ${DIM}($(L "da" "from") ${BASE})${RESET}  ${CYAN}${USER_COUNT} $(L "utenti" "users")${RESET}"
        IDX=$((IDX + 1))
    done
    echo ""

    local CHOICE
    if [[ ${#ALL_PATHS[@]} -eq 1 ]]; then
        echo -ne "  ${YELLOW}[?]${RESET} $(L "Usare" "Use") ${BOLD}${ALL_PATHS[0]}${RESET} $(L "come root da analizzare? [S/n]:" "as analysis root? [Y/n]:") "
        read -r CHOICE || true
        [[ "${CHOICE,,}" == "n" ]] && return 1
        _apply_win_root "${ALL_PATHS[0]}"; return 0
    fi
    echo -ne "  ${YELLOW}[?]${RESET} $(L "Seleziona numero, inserisci path manuale, o [N] per saltare:" "Select number, enter manual path, or [N] to skip:") "
    read -r CHOICE || true
    case "${CHOICE,,}" in
        n|"") return 1 ;;
        [0-9]*)
            local SEL=$((CHOICE - 1))
            if [[ $SEL -ge 0 && $SEL -lt ${#ALL_PATHS[@]} ]]; then
                _apply_win_root "${ALL_PATHS[$SEL]}"; return 0
            else
                err "$(L "Selezione non valida" "Invalid selection")"; return 1
            fi ;;
        *)
            local MP; MP=$(realpath -m "$CHOICE" 2>/dev/null || echo "$CHOICE")
            [[ ! -d "$MP" ]] && err "$(L "Directory non trovata:" "Directory not found:") $MP" && return 1
            _apply_win_root "$MP"; return 0 ;;
    esac
}

# Imposta WIN_ROOT e innesca la raccolta informazioni
_apply_win_root() {
    local ROOT="$1"
    WIN_ROOT="$ROOT"
    OS_TYPE=$(detect_os_type "$ROOT")

    # Un export di disco sta quasi sempre dentro una cartella di servizio
    # ("ntfs", "C", il nome del disco). Prima di dichiarare la root inutile,
    # si guarda un livello sotto: la risposta e' spesso li', e tacerla
    # lascerebbe l'analista davanti a un menu vuoto senza sapere perche'.
    if [[ "$OS_TYPE" == "unknown" ]]; then
        local _NESTED _NPATH _NOS
        if _NESTED=$(find_nested_root "$ROOT"); then
            IFS=$'\t' read -r _NPATH _NOS <<< "$_NESTED"
            warn "$(L "In questa directory non c'e' una struttura di sistema riconoscibile." \
                     "No recognisable system structure in this directory.")"
            info "$(L "Ne ho trovata una un livello sotto:" "I found one one level down:") ${BOLD}${_NPATH}${RESET}  ${CYAN}[${_NOS}]${RESET}"
            if ask_yn "$(L "Uso quella?" "Use that one?")"; then
                ROOT="$_NPATH"
                WIN_ROOT="$ROOT"
                OS_TYPE="$_NOS"
            fi
        else
            warn "$(L "Nessuna struttura Windows, Linux o macOS riconoscibile in:" \
                     "No recognisable Windows, Linux or macOS structure in:") $ROOT"
            info "$(L "Attesi al primo livello: Windows/System32 o Users (Windows), etc/passwd (Linux), System/Library/CoreServices (macOS)." \
                     "Expected at the top level: Windows/System32 or Users (Windows), etc/passwd (Linux), System/Library/CoreServices (macOS).")"
        fi
    fi

    ok "$(L "Root impostata:" "Root set:") ${BOLD}$WIN_ROOT${RESET}  ${CYAN}[$(os_label)]${RESET}"

    # Recupera info macchina (hostname, OS, IP, dominio)
    gather_host_info
    # Fuso del volume: non converte nulla, serve a dichiarare cosa significano
    # le ore locali che compaiono nei report. Va rifatto a ogni cambio di root.
    detect_volume_timezone
    if [[ -n "$VOLUME_TZ" ]]; then
        info "$(L "Fuso orario del volume:" "Volume timezone:") ${BOLD}${VOLUME_TZ}${RESET} ${DIM}(${VOLUME_TZ_SOURCE})${RESET}"
    fi
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
    echo -e "  ${CYAN}[*]${RESET} $(L "Ricerca volumi montati (Windows/Linux/macOS)..." "Searching mounted volumes (Windows/Linux/macOS)...")"
    if autodetect_win_root; then
        return 0
    fi
    # Fallback: input manuale
    echo ""
    echo -ne "  ${YELLOW}[?]${RESET} $(L "Inserisci il path della root da analizzare (es. /mnt/disk):" "Enter analysis root path (e.g. /mnt/disk):") "
    read -r INPUT_ROOT
    [[ -z "$INPUT_ROOT" ]] && return 1
    INPUT_ROOT=$(realpath -m "$INPUT_ROOT" 2>/dev/null || echo "$INPUT_ROOT")
    if [[ ! -d "$INPUT_ROOT" ]]; then
        err "$(L "Directory non trovata:" "Directory not found:") $INPUT_ROOT"
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

    # --report-dir salta l'intera interazione. Serve a due casi in cui il
    # prompt e' un ostacolo e non un aiuto: l'uso scriptato, e l'analisi di un
    # sistema vivo, dove il default (la directory di invocazione) finirebbe
    # DENTRO il volume analizzato.
    if [[ -n "${REPORT_DIR_FIXED:-}" ]]; then
        REPORT_BASE_DIR="$REPORT_DIR_FIXED"
        LOG_FILE="${REPORT_BASE_DIR}/fiuto_session_$(date +%Y%m%d_%H%M%S).log"
        log_msg "=== Log sessione inizializzato (--report-dir) ==="
        ok "$(L "Report dir (--report-dir):" "Report directory (--report-dir):") ${BOLD}$REPORT_BASE_DIR"
        return 0
    fi

    local SUGGESTED_DEFAULT="${INVOCATION_DIR}/${HOST_NAME:-CASE}_fiuto_${TS}"
    echo ""
    echo -e "  ${CYAN}${BOLD}╔══════════════════════════════════════════════════════╗${RESET}"
    echo -e "  ${CYAN}${BOLD}║  $(L "Configurazione cartella di output dei report" "Report output directory setup               ")        ║ ${RESET}"
    echo -e "  ${CYAN}${BOLD}╚══════════════════════════════════════════════════════╝${RESET}"
    echo ""
    local _PARENT_OK=false
    if [[ -d "$SUGGESTED_DEFAULT" ]]; then
        [[ -w "$SUGGESTED_DEFAULT" ]] && _PARENT_OK=true
    elif [[ -w "$(dirname "$SUGGESTED_DEFAULT")" ]]; then
        _PARENT_OK=true
    fi
    if $_PARENT_OK; then
        echo -e "  ${GREEN}[✓]${RESET} $(L "Directory suggerita:" "Suggested directory:") ${BOLD}${SUGGESTED_DEFAULT}${RESET}  ${GREEN}[$(L "scrivibile" "writable")]${RESET}"
    else
        echo -e "  ${RED}[!]${RESET} $(L "Directory suggerita:" "Suggested directory:") ${BOLD}${SUGGESTED_DEFAULT}${RESET}  ${RED}[$(L "non scrivibile o parent protetto" "not writable or protected parent")]${RESET}"
    fi
    echo ""
    echo -e "  ${DIM}$(L "I report di ogni modulo verranno salvati in sottocartelle con data/ora." "Each module report will be saved in subfolders with date/time.")${RESET}"
    echo -e "  ${DIM}$(L "Puoi inserire un percorso diverso oppure premere INVIO per usare quello suggerito." "You can enter a different path or press ENTER to use the suggested one.")${RESET}"
    echo ""
    echo -ne "  ${YELLOW}[?]${RESET} $(L "Cartella report" "Report directory") [${BOLD}${SUGGESTED_DEFAULT}${RESET}]: "
    local _INPUT
    read -r _INPUT
    local _CHOSEN
    if [[ -z "$_INPUT" ]]; then
        _CHOSEN="$SUGGESTED_DEFAULT"
    else
        _CHOSEN=$(realpath -m "$_INPUT" 2>/dev/null || echo "$_INPUT")
    fi
    local _RW_OK=false _RW_MSG=""
    if [[ -d "$_CHOSEN" ]]; then
        if [[ -w "$_CHOSEN" ]]; then
            _RW_OK=true
            _RW_MSG="${GREEN}[$(L "scrivibile" "writable")]${RESET}"
        else
            _RW_MSG="${RED}[$(L "SOLA LETTURA — i report NON potranno essere salvati!" "READ ONLY — reports CANNOT be saved!")]${RESET}"
        fi
    else
        local _P; _P=$(dirname "$_CHOSEN")
        if [[ -w "$_P" ]]; then
            _RW_OK=true
            _RW_MSG="${GREEN}[$(L "verrà creata — parent scrivibile" "will be created — parent writable")]${RESET}"
        else
            _RW_MSG="${RED}[$(L "parent '${_P}' NON scrivibile — i report NON potranno essere salvati!" "parent '${_P}' NOT writable — reports CANNOT be saved!")]${RESET}"
        fi
    fi
    echo ""
    echo -e "  ${CYAN}[→]${RESET} $(L "Cartella scelta:" "Selected directory:") ${BOLD}${_CHOSEN}${RESET}  ${_RW_MSG}"
    echo ""
    if ! $_RW_OK; then
        warn "$(L "Attenzione: la directory selezionata non è scrivibile." "Warning: the selected directory is not writable.")"
        echo -ne "  ${YELLOW}[?]${RESET} $(L "Vuoi comunque usarla? [s/N]:" "Use it anyway? [y/N]:") "
        local _CONF; read -r _CONF
        local _YES_KEY="$(L "s" "y")"
        [[ "${_CONF,,}" != "$_YES_KEY" ]] && { warn "$(L "Setup annullato. Ripeti con [P] dal menu." "Setup cancelled. Repeat with [P] from menu.")"; return 1; }
    fi
    REPORT_BASE_DIR="$_CHOSEN"
    LOG_FILE="${REPORT_BASE_DIR}/fiuto_session_$(date +%Y%m%d_%H%M%S).log"
    log_msg "=== Log sessione inizializzato ==="
    ok "$(L "Log sessione:" "Session log:") ${BOLD}$LOG_FILE"
    ok "$(L "Report dir impostata:" "Report directory set:") ${BOLD}$REPORT_BASE_DIR"
    sleep 1
}

# Esegue un modulo in modalità batch e registra il risultato in SUMMARY_TABLE.
# Durante l'esecuzione mostra il suggerimento ESC; premendo ESC il modulo viene
# interrotto e si passa al successivo.
run_batch_module() {
    local mod_num="$1"
    local mod_func="$2"
    local mod_name="$3"
    local total_mods="${4:-38}"
    local _ESC_HINT; _ESC_HINT="$(L "[ESC: salta modulo]" "[ESC: skip module]")"

    echo -ne "  ${CYAN}[*]${RESET} [${mod_num}/${total_mods}] $(L "Esecuzione modulo" "Running module") $mod_num ($mod_name)...  ${DIM}${_ESC_HINT}${RESET}\r"
    log_msg "[BATCH] Modulo $mod_num: $mod_name"

    # I report generati dal subshell vengono scritti su un file temp
    # perché gli array bash non si propagano al processo padre.
    local _REP_TMP; _REP_TMP=$(mktemp)

    (
        # Override locale: scrive il path nel file temp invece dell'array.
        # Gli effetti collaterali vanno richiamati a mano, o si perdono.
        register_report() {
            [[ -n "${1:-}" && -f "$1" ]] || return 0
            echo "$1" >> "$_REP_TMP"
            _report_side_effects "$1"
        }
        $mod_func >/dev/null 2>&1
    ) &
    local MOD_PID=$!

    # Monitoraggio tasto ESC (solo se /dev/tty è disponibile)
    local SKIPPED=0
    local _OLD_STTY=""
    if [[ -c /dev/tty ]]; then
        _OLD_STTY=$(stty -g </dev/tty 2>/dev/null) || true
        stty -echo -icanon min 0 time 0 </dev/tty 2>/dev/null || true
        while kill -0 "$MOD_PID" 2>/dev/null; do
            local KEY=""
            IFS= read -r -s -t 0.2 -N 1 KEY </dev/tty 2>/dev/null || true
            if [[ "$KEY" == $'\033' ]]; then
                kill "$MOD_PID" 2>/dev/null
                SKIPPED=1
                break
            fi
        done
        [[ -n "$_OLD_STTY" ]] && stty "$_OLD_STTY" </dev/tty 2>/dev/null || true
    fi

    wait "$MOD_PID" 2>/dev/null

    if [[ $SKIPPED -eq 1 ]]; then
        printf '\r\033[K'
        echo -e "  ${YELLOW}[⏭]${RESET} [${mod_num}/${total_mods}] $mod_name — $(L "annullato (ESC)" "cancelled (ESC)")"
        SUMMARY_TABLE+=("$mod_num|$mod_name|SKIP|$(L "annullato" "cancelled")")
        rm -f "$_REP_TMP"
        return
    fi

    # Importa i report generati nel subshell
    if [[ -s "$_REP_TMP" ]]; then
        while IFS= read -r _rep; do
            [[ -n "$_rep" && -f "$_rep" ]] && GENERATED_REPORTS+=("$_rep")
        done < "$_REP_TMP"
        local rep_path="${GENERATED_REPORTS[-1]}"
        printf '\r\033[K'
        echo -e "  ${GREEN}[✓]${RESET} [${mod_num}/${total_mods}] $mod_name — report: ${DIM}${rep_path}${RESET}"
        SUMMARY_TABLE+=("$mod_num|$mod_name|SI|$rep_path")
    else
        printf '\r\033[K'
        echo -e "  ${DIM}[i] [${mod_num}/${total_mods}] $mod_name — $(L "nessun risultato" "no results")${RESET}"
        SUMMARY_TABLE+=("$mod_num|$mod_name|NO|-")
    fi
    rm -f "$_REP_TMP"
}

# ================================================================
#  POOL DI JOB PER --all --jobs N
#
#  I moduli sono indipendenti: leggono file diversi e scrivono in cartelle
#  diverse. L'unico vincolo d'ordine e' la Master Timeline, marcata `defer`,
#  che aggrega il lavoro degli altri — e quindi va eseguita per ultima, da
#  sola, quando il pool si e' svuotato.
#
#  La parallelizzazione e' OPT-IN e non il default. Su un volume montato da
#  disco meccanico o via rete N processi che leggono insieme vanno piu' piano
#  di uno solo, e il guadagno dipende dal collo di bottiglia reale, che qui e'
#  quasi sempre l'I/O e non la CPU. Chi analizza sa qual e' il suo: sceglie.
#
#  In parallelo si perde l'interruzione con ESC: intercettarla richiede il
#  controllo esclusivo del terminale, che con N moduli concorrenti non c'e'.
#  Viene dichiarato all'avvio invece di lasciare che il tasto smetta di
#  funzionare senza spiegazione.
# ================================================================
run_batch_pool() {
    local _total="$1"; shift
    local -a _queue=("$@")
    local _njobs="${JOBS:-1}"

    echo -e "  ${CYAN}[*]${RESET} $(L "Esecuzione parallela:" "Parallel execution:") ${BOLD}${_njobs}${RESET} $(L "moduli alla volta" "modules at a time")"
    echo -e "  ${DIM}$(L "In parallelo l'interruzione con ESC non e' disponibile." \
                        "ESC interruption is not available in parallel mode.")${RESET}"
    echo ""

    # Un file per modulo: gli array bash non risalgono dai processi figli.
    local _POOLDIR; _POOLDIR=$(mktemp -d); register_tmp "$_POOLDIR"

    # I moduli `defer` restano fuori dal pool: aggregano gli altri e devono
    # vederli finiti. Sono gia' in coda, ma "in coda" non basta col parallelo.
    local -a _par=() _seq=()
    local _item _i _f _label _flags
    for _item in "${_queue[@]}"; do
        IFS='|' read -r _i _f _label _flags <<< "$_item"
        if [[ "${_flags:-}" == *defer* ]]; then _seq+=("$_item"); else _par+=("$_item"); fi
    done

    local _running=0
    for _item in "${_par[@]}"; do
        IFS='|' read -r _i _f _label _flags <<< "$_item"
        (
            register_report() {
                [[ -n "${1:-}" && -f "$1" ]] || return 0
                echo "$1" >> "${_POOLDIR}/${_i}.rep"
                _report_side_effects "$1"
            }
            "$_f" > /dev/null 2>&1
        ) &
        _running=$((_running + 1))
        if [[ "$_running" -ge "$_njobs" ]]; then
            wait -n 2>/dev/null || wait
            _running=$((_running - 1))
        fi
    done
    wait

    # Esiti raccolti nell'ordine dei moduli, non in quello di completamento:
    # un riepilogo che cambia ordine a ogni esecuzione non e' confrontabile.
    #
    # Va fatto PRIMA di lanciare i deferred: la Master Timeline aggrega leggendo
    # GENERATED_REPORTS, e se lo trovasse ancora vuoto produrrebbe una timeline
    # vuota senza segnalare nulla. E' il difetto che questo ordine evita.
    for _item in "${_par[@]}"; do
        IFS='|' read -r _i _f _label _flags <<< "$_item"
        local _rf="${_POOLDIR}/${_i}.rep"
        if [[ -s "$_rf" ]]; then
            local _rep
            while IFS= read -r _rep; do
                [[ -n "$_rep" && -f "$_rep" ]] && GENERATED_REPORTS+=("$_rep")
            done < "$_rf"
            local _last; _last=$(tail -1 "$_rf")
            echo -e "  ${GREEN}[✓]${RESET} [${_i}/${_total}] $_label — report: ${DIM}${_last}${RESET}"
            SUMMARY_TABLE+=("$_i|$_label|SI|$_last")
        else
            echo -e "  ${DIM}[i] [${_i}/${_total}] $_label — $(L "nessun risultato" "no results")${RESET}"
            SUMMARY_TABLE+=("$_i|$_label|NO|-")
        fi
    done

    # I deferred girano ora, in sequenza e nel percorso normale: il pool e'
    # vuoto e GENERATED_REPORTS contiene gli altri moduli, che e' esattamente
    # la condizione che si aspettano.
    for _item in "${_seq[@]}"; do
        IFS='|' read -r _i _f _label _flags <<< "$_item"
        run_batch_module "$_i" "$_f" "$_label" "$_total"
    done
}

# ================================================================
#  DASHBOARD "FULL" — indice navigabile con tab + iframe centrale.
#  Generata al termine di "esegui TUTTI i moduli" (Windows/Linux/macOS).
#  Costruita interamente da SUMMARY_TABLE (righe "num|nome|SI/NONE/SKIP|path").
# ================================================================
generate_full_dashboard() {
    [[ -z "$REPORT_BASE_DIR" ]] && return 0
    [[ ${#SUMMARY_TABLE[@]} -eq 0 ]] && return 0
    local OSL; OSL=$(os_label)
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local DASH="${REPORT_BASE_DIR}/index.html"
    local TABS="" COUNT_OK=0 COUNT_TOTAL=0

    for row in "${SUMMARY_TABLE[@]}"; do
        IFS='|' read -r mnum mname msy mpath <<< "$row"
        [[ -z "$mnum" ]] && continue
        COUNT_TOTAL=$((COUNT_TOTAL + 1))
        local NUM2; NUM2=$(printf '%02d' "$mnum" 2>/dev/null || echo "$mnum")
        local NAME_ESC; NAME_ESC=$(html_esc "$mname")
        if [[ "$msy" == "SI" ]]; then
            COUNT_OK=$((COUNT_OK + 1))
            local rel="${mpath#$REPORT_BASE_DIR/}"
            TABS+="<button class='tab' data-src='$(html_esc "$rel")'><span class='tn'>${NUM2}</span><span class='tl'>${NAME_ESC}</span><span class='dot ok'></span></button>"
        else
            local CLS="none" LBL
            [[ "$msy" == "SKIP" ]] && { CLS="skip"; LBL="skip"; } || LBL="—"
            TABS+="<button class='tab disabled' disabled title='$([ "$msy" = "SKIP" ] && echo "$(L "saltato" "skipped")" || echo "$(L "nessuna evidenza" "no findings")")'><span class='tn'>${NUM2}</span><span class='tl'>${NAME_ESC}</span><span class='dot ${CLS}'></span></button>"
        fi
    done

    local HOST_DISP="${HOST_NAME:-N/A}"
    # Icona "naso di cane che fiuta" (SVG inline, bianco su sfondo blu)
    local NOSE_SVG='<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M19 11c2-2 5-2 7 0" fill="none" stroke="#dbeafe" stroke-width="2.4" stroke-linecap="round" opacity=".85"/><path d="M38 11c2-2 5-2 7 0" fill="none" stroke="#dbeafe" stroke-width="2.4" stroke-linecap="round" opacity=".85"/><path fill="#fff" d="M32 50C17 39 10 32 10 25c0-6 6-9 12-7 4 1 7 4 10 7 3-3 6-6 10-7 6-2 12 1 12 7 0 7-7 14-22 25Z"/><ellipse cx="23" cy="28" rx="3" ry="4.3" fill="#1f6feb"/><ellipse cx="41" cy="28" rx="3" ry="4.3" fill="#1f6feb"/><path d="M32 33v10" stroke="#1f6feb" stroke-width="2.6" stroke-linecap="round"/></svg>'
    {
        cat << HTMLEOF
<!DOCTYPE html>
<html lang="$(L "it" "en")">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FIUTO — Full Report (${OSL})</title>
<style>
  :root{ --bg:#080b0f; --bg2:#0d1117; --bg3:#131920; --bg4:#1a2332; --border:#1e2d3d;
    --accent:#58a6ff; --accent2:#ff7b72; --accent3:#3fb950; --accent4:#f0883e;
    --text:#c9d1d9; --text-dim:#3d5166; --text-mid:#6e8898;
    --mono:'Fira Code',ui-monospace,monospace; --sans:'DM Sans',system-ui,sans-serif; }
  *{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%}
  body{background:var(--bg);color:var(--text);font-family:var(--sans);display:flex;flex-direction:column;height:100vh;overflow:hidden}
  header{background:var(--bg2);border-bottom:1px solid var(--border);padding:1rem 1.6rem;display:flex;align-items:center;gap:1.4rem;flex-shrink:0}
  .hicon{width:2.6rem;height:2.6rem;background:linear-gradient(135deg,var(--accent),#1f6feb);display:flex;align-items:center;justify-content:center;flex-shrink:0;clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%)}
  .hicon svg{width:62%;height:62%}
  .htxt h1{font-size:1.1rem;color:#fff;font-weight:700}
  .htxt .sub{font-size:.66rem;color:var(--text-dim);font-family:var(--mono);margin-top:.2rem}
  .hstats{margin-left:auto;display:flex;gap:1.6rem;font-family:var(--mono)}
  .hstats .s .v{font-size:1.3rem;font-weight:800}
  .hstats .s .l{font-size:.55rem;text-transform:uppercase;letter-spacing:.12em;color:var(--text-dim)}
  .s.ok .v{color:var(--accent3)} .s.tot .v{color:var(--accent)}
  #tabs{display:flex;flex-wrap:wrap;gap:.15rem .2rem;background:var(--bg2);border-bottom:1px solid var(--border);padding:.4rem .8rem;flex-shrink:0;max-height:45vh;overflow-y:auto;align-content:flex-start}
  #tabs::-webkit-scrollbar{width:6px}
  #tabs::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
  .tab{display:flex;align-items:center;gap:.4rem;white-space:nowrap;background:var(--bg3);border:1px solid var(--border);border-radius:5px;color:var(--text-mid);font-family:var(--mono);font-size:.72rem;padding:.4rem .6rem;cursor:pointer;transition:.15s}
  .tab:hover:not(.disabled){color:var(--text);border-color:var(--accent);background:rgba(88,166,255,.08)}
  .tab .tn{color:var(--text-dim);font-size:.64rem}
  .tab.active{color:#fff;border-color:var(--accent4);background:rgba(240,136,62,.12)}
  .tab.active .tn{color:var(--accent4)}
  .tab.disabled{opacity:.4;cursor:not-allowed}
  .dot{width:.45rem;height:.45rem;border-radius:50%;flex-shrink:0}
  .dot.ok{background:var(--accent3)} .dot.none{background:var(--border)} .dot.skip{background:var(--accent4)}
  main{flex:1;position:relative;background:var(--bg)}
  iframe{width:100%;height:100%;border:0;background:var(--bg);display:none}
  iframe.show{display:block}
  #placeholder{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1rem;text-align:center;padding:2rem}
  #placeholder .big{font-family:var(--mono);font-size:1.1rem;color:var(--text-mid)}
  #placeholder .small{font-size:.8rem;color:var(--text-dim);max-width:34rem;line-height:1.6}
  #placeholder .logo{width:4rem;height:4rem;background:linear-gradient(135deg,var(--accent),#1f6feb);display:flex;align-items:center;justify-content:center;clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%)}
  #placeholder .logo svg{width:62%;height:62%}
</style>
</head>
<body>
<header>
  <div class="hicon">${NOSE_SVG}</div>
  <div class="htxt">
    <h1>FIUTO — Full Report</h1>
    <div class="sub">${OSL} · ${HOST_DISP} · ${SCAN}</div>
  </div>
  <div class="hstats">
    <div class="s ok"><div class="v">${COUNT_OK}</div><div class="l">$(L "con evidenze" "with findings")</div></div>
    <div class="s tot"><div class="v">${COUNT_TOTAL}</div><div class="l">$(L "moduli" "modules")</div></div>
  </div>
</header>
<nav id="tabs">${TABS}</nav>
<main>
  <iframe id="viewer" title="report"></iframe>
  <div id="placeholder">
    <div class="logo">${NOSE_SVG}</div>
    <div class="big">$(L "Seleziona un modulo dalle tab in alto" "Select a module from the tabs above")</div>
    <div class="small">$(L "Il report verrà caricato qui al centro. Puoi passare da un modulo all'altro senza aprire file separati." "The report will load here in the center. Switch between modules without opening separate files.")</div>
  </div>
</main>
<script>
(function(){
  var tabs=document.querySelectorAll('.tab:not(.disabled)'),
      viewer=document.getElementById('viewer'),
      ph=document.getElementById('placeholder');
  tabs.forEach(function(t){
    t.addEventListener('click',function(){
      document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});
      t.classList.add('active');
      var src=t.getAttribute('data-src');
      if(src){ viewer.src=src; viewer.classList.add('show'); ph.style.display='none'; }
    });
  });
})();
</script>
</body></html>
HTMLEOF
    } > "$DASH"

    register_report "$DASH"
    echo ""
    ok "$(L "Dashboard FULL generata:" "FULL dashboard generated:") ${BOLD}$DASH"
    open_report_prompt "$DASH"
}


# Elenco dei moduli per --help, reso DAI REGISTRI.
#
# Era scritto a mano, ed era rimasto fermo a 39 voci con l'intestazione
# "moduli disponibili (1-50)" mentre i registri ne contavano 54, 24 e 22. E'
# la stessa deriva che nella Fase 1 aveva portato a togliere il menu scritto a
# mano: un elenco che non si aggiorna da solo non si aggiorna.
print_module_list() {
    local _reg _name _label _entry _f _nm _color _desc
    for _reg in MODULES_WIN:Windows MODULES_LINUX:Linux MODULES_MACOS:macOS; do
        _name="${_reg%%:*}"; _label="${_reg##*:}"
        local -n _R="$_name"
        echo ""
        printf "  ${BOLD}%s${RESET} ${DIM}(1-%d)${RESET}\n" "$_label" "${#_R[@]}"
        local _i=1 _line=""
        for _entry in "${_R[@]}"; do
            IFS='|' read -r _f _nm _color _desc <<< "$_entry"
            _line+=$(printf "%3d %-26s" "$_i" "$(reg_text "$_nm")")
            if (( _i % 3 == 0 )); then echo -e "   ${_line}"; _line=""; fi
            _i=$((_i + 1))
        done
        [[ -n "$_line" ]] && echo -e "   ${_line}"
        unset -n _R
    done
}

# Estrae dalla forma bilingue "italiano§english" la variante per la lingua
# corrente. Senza separatore il testo vale per entrambe.
reg_text() {
    local _s="$1"
    [[ "$_s" != *§* ]] && { printf '%s' "$_s"; return; }
    if [[ "${LANG:-en}" == "it" ]]; then
        printf '%s' "${_s%%§*}"
    else
        printf '%s' "${_s##*§}"
    fi
}

active_registry_name() {
    case "$OS_TYPE" in
        windows) echo "MODULES_WIN" ;;
        linux)   echo "MODULES_LINUX" ;;
        macos)   echo "MODULES_MACOS" ;;
        *)       echo "" ;;
    esac
}

# Renderizza il menu a partire da un registro (equivalente dinamico di print_menu)
render_menu_from_registry() {
    local -n _REG="$1"
    local _NOT_SET _WRITABLE _READONLY _NOT_CREATED _PARENT_RO _DIAG _RUN_ALL _QUIT _CHOICE_LABEL _REPORTS_LABEL
    _NOT_SET="$(L "non impostata" "not set")"
    _WRITABLE="$(L "scrivibile" "writable")"
    _READONLY="$(L "SOLA LETTURA" "READ ONLY")"
    _NOT_CREATED="$(L "OK (non ancora creata)" "OK (not yet created)")"
    _PARENT_RO="$(L "PARENT NON SCRIVIBILE" "PARENT NOT WRITABLE")"
    _DIAG="$(L "Diagnostica volumi montati" "Diagnose mounted volumes")"
    _RUN_ALL="$(L "Esegui TUTTI i moduli" "Run ALL modules")"
    _QUIT="$(L "Esci" "Quit")"
    _CHOICE_LABEL="$(L "Scelta" "Choice")"
    _REPORTS_LABEL="$(L "Report generati" "Generated reports")"
    local _OSL; _OSL=$(os_label)
    local _TITLE; _TITLE="$(L "SELEZIONA UN MODULO" "SELECT A MODULE")"

    # La cornice si calcola sul testo, non a mano. Con le larghezze fisse
    # "SELEZIONA UN MODULO" (19 caratteri) sforava il campo da 18 e spingeva
    # fuori il bordo destro: il riquadro non si chiudeva. Un valore scritto a
    # mano va rifatto a ogni traduzione e a ogni etichetta di OS nuova.
    local _INNER="   F I U T O  —  ${_OSL}  —  ${_TITLE}   "
    local _W=$(( ${#_INNER} > 50 ? ${#_INNER} : 50 ))
    local _BAR; _BAR=$(printf '═%.0s' $(seq 1 "$_W"))
    # Il riempimento si scrive a mano invece di usare %-*s: ${#stringa} conta i
    # CARATTERI, mentre la larghezza di printf conta i BYTE. Con tre em dash da
    # tre byte l'una il conto salta di sei, e la cornice si richiude storta —
    # in modo diverso a seconda della lingua e dell'etichetta di OS.
    local _PAD=$(( _W - ${#_INNER} ))
    echo -e "  ${CYAN}${BOLD}╔${_BAR}╗${RESET}"
    printf  "  ${CYAN}${BOLD}║%s%*s║${RESET}\n" "$_INNER" "$_PAD" ""
    echo -e "  ${CYAN}${BOLD}╚${_BAR}╝${RESET}"
    echo ""
    if [[ -n "$REPORT_BASE_DIR" ]]; then
        local _RW_LABEL _RW_COLOR
        if [[ -d "$REPORT_BASE_DIR" ]]; then
            if [[ -w "$REPORT_BASE_DIR" ]]; then _RW_LABEL="$_WRITABLE"; _RW_COLOR="$GREEN"
            else _RW_LABEL="$_READONLY"; _RW_COLOR="$RED"; fi
        else
            local _RD_PARENT; _RD_PARENT=$(dirname "$REPORT_BASE_DIR")
            if [[ -w "$_RD_PARENT" ]]; then _RW_LABEL="$_NOT_CREATED"; _RW_COLOR="$GREEN"
            else _RW_LABEL="$_PARENT_RO"; _RW_COLOR="$RED"; fi
        fi
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${DIM}${REPORT_BASE_DIR}${RESET}  ${_RW_COLOR}[${_RW_LABEL}]${RESET}"
    else
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${RED}$(L "non impostata — premi [P] per configurare" "not set — press [P] to configure")${RESET}"
    fi
    echo -e "  ${WHITE}[R]${RESET}  ${BOLD}$(L "Imposta root da analizzare" "Set analysis root")${RESET}  ${DIM}${WIN_ROOT:-($_NOT_SET)} [${_OSL}]${RESET}"
    echo -e "  ${YELLOW}[D]${RESET}  ${BOLD}$(L "Debug mount attivi" "Debug active mounts")${RESET}  ${DIM}${_DIAG}${RESET}"
    echo ""
    local _i=1 _entry _f _name _color _desc
    for _entry in "${_REG[@]}"; do
        IFS='|' read -r _f _name _color _desc <<< "$_entry"
        local _C="${!_color:-$RESET}"
        printf "  ${_C}[%2d]${RESET} %-28s ${DIM}%s${RESET}\n" \
            "$_i" "$(reg_text "$_name")" "$(reg_text "$_desc")"
        _i=$((_i + 1))
    done
    echo ""
    echo -e "  ${WHITE}${BOLD}[0]${RESET}  ${BOLD}${_RUN_ALL}${RESET}"
    if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
        echo -e "  ${WHITE}${BOLD}[S]${RESET}  ${BOLD}$(L "Executive summary dei report di questa sessione" "Executive summary of this session's reports")${RESET}"
    fi
    echo ""
    if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
        echo -e "  ${DIM}── ${_REPORTS_LABEL} (${#GENERATED_REPORTS[@]}) ──────────────────────────${RESET}"
        for _R in "${GENERATED_REPORTS[@]}"; do
            echo -e "  ${CYAN}↳${RESET} ${DIM}${_R}${RESET}"
        done
        echo ""
    fi
    echo -e "  ${RED}[Q]  ${_QUIT}${RESET}"
    echo ""
    echo -ne "  ${YELLOW}${_CHOICE_LABEL}:${RESET} "
}

# Esegue il modulo n-esimo (1-based) di un registro
dispatch_from_registry() {
    local _RNAME="$1" _N="$2"
    local -n _REG="$_RNAME"
    if ! [[ "$_N" =~ ^[0-9]+$ ]] || (( _N < 1 || _N > ${#_REG[@]} )); then
        err "$(L "Modulo sconosciuto:" "Unknown module:") $_N"
        return 1
    fi
    local _entry="${_REG[$((_N - 1))]}"
    local _f="${_entry%%|*}"
    "$_f"
}

# Esegue TUTTI i moduli di un registro in modalità batch (equivalente di run_all_modules)
run_all_from_registry() {
    local _RNAME="$1"
    local -n _REG="$_RNAME"
    clear
    print_banner
    info "$(t batch_running)"
    echo ""
    if [[ -z "$REPORT_BASE_DIR" ]]; then
        REPORT_BASE_DIR="${INVOCATION_DIR}/fiuto_reports_$(date +%Y%m%d_%H%M%S)"
        LOG_FILE="${REPORT_BASE_DIR}/fiuto_session_$(date +%Y%m%d_%H%M%S).log"
    fi
    info "$(t batch_report_dir) ${BOLD}$REPORT_BASE_DIR${RESET}"
    log_msg "$(t batch_started)$WIN_ROOT ==="
    sleep 1
    BATCH_MODE=true
    SUMMARY_TABLE=()
    echo ""
    local _total=${#_REG[@]} _entry _f _name _color _desc _guard _flags

    # I moduli marcati "defer" (la Master Timeline) aggregano il lavoro degli
    # altri, quindi devono girare per ultimi anche se stanno a meta' elenco.
    # Cosi' si possono aggiungere moduli in coda senza rinumerare la Master
    # Timeline: la numerazione e' un contratto con chi usa --module N.
    local -a _order=() _deferred=()
    local _n=1
    for _entry in "${_REG[@]}"; do
        IFS='|' read -r _f _name _color _desc _guard _flags <<< "$_entry"
        if [[ "${_flags:-}" == *defer* ]]; then
            _deferred+=("${_n}|${_entry}")
        else
            _order+=("${_n}|${_entry}")
        fi
        _n=$((_n + 1))
    done
    _order+=("${_deferred[@]}")

    # Le guardie si valutano prima, in sequenza: sono veloci e cosi' la coda
    # da eseguire e' nota, che serve al pool per non lanciare lavoro inutile.
    local -a _todo=()
    local _item _i
    for _item in "${_order[@]}"; do
        _i="${_item%%|*}"
        _entry="${_item#*|}"
        IFS='|' read -r _f _name _color _desc _guard _flags <<< "$_entry"
        local _label; _label=$(reg_text "$_name")
        if [[ -n "${_guard:-}" ]] && declare -F "$_guard" > /dev/null; then
            local _reason
            if ! _reason=$("$_guard"); then
                echo -e "  ${DIM}[i] [$_i/$_total] ${_label} — $(L "saltato" "skipped") (${_reason})${RESET}"
                SUMMARY_TABLE+=("$_i|$_label|SKIP|$_reason")
                continue
            fi
        fi
        _todo+=("${_i}|${_f}|${_label}|${_flags:-}")
    done

    if [[ "${JOBS:-1}" -gt 1 ]]; then
        run_batch_pool "$_total" "${_todo[@]}"
    else
        for _item in "${_todo[@]}"; do
            IFS='|' read -r _i _f _label _flags <<< "$_item"
            run_batch_module "$_i" "$_f" "$_label" "$_total"
        done
    fi
    BATCH_MODE=false
    echo ""
    section_header "$(L "Riepilogo Scansione Globale" "Global Scan Summary")" "$GREEN"
    local _hdr_mod _hdr_name _hdr_evid _hdr_file _lbl_found _lbl_none _lbl_skip
    _hdr_mod="$(L "MOD" "MOD")"; _hdr_name="$(L "NOME MODULO" "MODULE NAME")"
    _hdr_evid="$(L "EVIDENZE" "FINDINGS")"; _hdr_file="$(L "FILE GENERATI" "GENERATED FILES")"
    _lbl_found="$(L "TROVATE" "FOUND")"; _lbl_none="$(L "NESSUNA" "NONE")"; _lbl_skip="$(L "SALTATO" "SKIPPED")"
    printf "  ${BOLD}%-4s %-32s %-12s %s${RESET}\n" "$_hdr_mod" "$_hdr_name" "$_hdr_evid" "$_hdr_file"
    echo "  ─────────────────────────────────────────────────────────────────────────────────────────"
    for row in "${SUMMARY_TABLE[@]}"; do
        IFS='|' read -r mnum mname msy mpath <<< "$row"
        if [[ "$msy" == "SI" ]]; then
            local rel_path="${mpath#$REPORT_BASE_DIR/}"
            printf "  ${CYAN}%02d${RESET}   ${BOLD}%-32s${RESET} ${GREEN}%-12s${RESET} ${DIM}%s${RESET}\n" "$mnum" "$mname" "$_lbl_found" "$rel_path"
        elif [[ "$msy" == "SKIP" ]]; then
            printf "  ${CYAN}%02d${RESET}   %-32s ${YELLOW}%-12s${RESET} ${DIM}%s${RESET}\n" "$mnum" "$mname" "$_lbl_skip" "$mpath"
        else
            printf "  ${CYAN}%02d${RESET}   %-32s ${DIM}%-12s${RESET} ${DIM}-${RESET}\n" "$mnum" "$mname" "$_lbl_none"
        fi
    done
    echo ""
    ok "$(L "Report salvati integralmente in:" "All reports saved in:") ${BOLD}$REPORT_BASE_DIR"
    # Il riepilogo va generato PRIMA della dashboard: la dashboard elenca i
    # report esistenti, e il riepilogo e' il primo che va aperto.
    generate_executive_summary
    generate_full_dashboard
    redact_summary
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


# ================================================================
#  EXECUTIVE SUMMARY
#
#  Novanta report HTML sono un dump, non un'analisi. Chi apre la cartella al
#  termine di --all non sa da dove cominciare, e la differenza fra un tool DFIR
#  e un collector sta esattamente li'.
#
#  Il riepilogo risponde a due domande: *da dove comincio* e *cosa e' successo
#  insieme a cos'altro*. Non risponde a "la macchina e' compromessa": quella e'
#  una conclusione dell'analista, e il report lo dice a chiare lettere.
#
#  TRE SOSTRATI, in ordine di affidabilita' decrescente:
#
#  1. Le righe che i moduli hanno gia' marcato come notevoli (classe
#     'sensitive', applicata da render_pre_block sulle parole chiave che ogni
#     modulo passa). E' un giudizio espresso da chi conosce l'artefatto.
#  2. Gli IoC caricati con --ioc, se presenti.
#  3. Una tabella di regole esplicite, deliberatamente piccola. Le regole
#     girano SOLO sul dato (celle di tabella e righe dei log), mai sulla prosa
#     dei cartigli esplicativi: un motore che cercasse "malware" nell'intera
#     pagina scatterebbe sulle spiegazioni scritte da FIUTO stesso.
#
#  Piu' la correlazione temporale fra moduli diversi, che e' l'unica cosa che
#  nessun singolo report puo' mostrare.
# ================================================================

# Regole di detection. TSV per non litigare con le pipe dentro le regex.
#
#   slug_regex  severita'  MITRE  pattern  titolo§title  perche'§why
#
# severita': CRITICA | ALTA | MEDIA | BASSA
# Il pattern e' case-insensitive e viene applicato a ogni cella/riga di dato.
# Aggiungendo regole: preferire pattern che descrivono la FORMA del dato
# (un base64 lungo dopo -enc) a parole che potrebbero comparire in un nome di
# file legittimo. Una regola rumorosa in un executive summary costa piu' di
# una regola mancante: fa perdere fiducia in tutte le altre.
finding_rules_tsv() {
    cat << 'RULESEOF'
ps_history|ps_scriptblock|ps_transcripts	ALTA	T1059.001	(-enc(odedcommand)?\s+[A-Za-z0-9+/=]{24,})|frombase64string|downloadstring|iex\s*\(|invoke-expression	PowerShell offuscato o download-and-execute§Obfuscated or download-and-execute PowerShell	Comando codificato in base64 o che scarica ed esegue codice dalla rete: e' la forma piu' comune di esecuzione iniziale su Windows.§Base64-encoded command, or one that downloads and runs code from the network: the most common initial-execution shape on Windows.
ifeo	ALTA	T1546.012	debugger	Debugger IFEO impostato§IFEO debugger set	Un valore Debugger sotto Image File Execution Options fa partire un eseguibile arbitrario al posto del programma indicato. Ha usi legittimi rari e un uso malevolo frequente.§A Debugger value under Image File Execution Options launches an arbitrary executable in place of the named program. Legitimate uses are rare, malicious ones frequent.
wmi	ALTA	T1546.003	eventfilter|eventconsumer|__filtertoconsumerbinding	Sottoscrizione WMI permanente§Permanent WMI subscription	Le sottoscrizioni WMI sopravvivono al riavvio e non compaiono in nessuna delle cartelle di avvio che si guardano di solito.§WMI subscriptions survive reboot and appear in none of the startup locations usually inspected.
run_keys|services|scheduled_tasks	ALTA	T1547.001	((users\\[^\\]+\\appdata|windows\\temp|programdata|users\\public)\\[^\\]*\.(exe|dll|ps1|bat|vbs|js|scr))	Persistenza che punta a una cartella scrivibile dall'utente§Persistence pointing at a user-writable folder	Un meccanismo di avvio che esegue un binario in AppData, Temp, ProgramData o Public: il software installato normalmente non sta li'.§A startup mechanism running a binary in AppData, Temp, ProgramData or Public: normally installed software does not live there.
defender_quarantine	ALTA	T1204	.	Defender ha messo elementi in quarantena§Defender quarantined items	Su questa macchina l'antivirus ha rilevato e isolato qualcosa. E' un evento gia' accaduto, non un'ipotesi: da qui si ricava data, percorso e nome della minaccia.§On this machine the antivirus detected and isolated something. This already happened, it is not a hypothesis: date, path and threat name follow from here.
bits	MEDIA	T1197	https?://	Job BITS con destinazione remota§BITS job with a remote endpoint	BITS scarica in background con il processo di sistema: e' un canale di download che non compare nei log del browser.§BITS downloads in the background under a system process: a download channel that never shows up in browser logs.
xplat_esp_bootkit	CRITICA	T1542.003	anomal|unsigned|non firmat|sconosciut|unknown	Anomalia nella partizione EFI§Anomaly in the EFI System Partition	Il codice nella ESP gira prima del sistema operativo e di qualunque EDR, e sopravvive alla reinstallazione. Un'anomalia qui va verificata subito.§Code in the ESP runs before the OS and any EDR, and survives reinstallation. An anomaly here must be verified first.
lsa_secrets|sam|ntds|credential_manager|wlan|linux_cloud_credentials	BASSA	T1003	.	Materiale soggetto a restrizioni nel report§Restricted material in the report	Il report contiene hash di credenziali, chiavi o segreti. Non e' un indizio di compromissione: e' un vincolo su come il report va conservato e condiviso.§The report contains credential hashes, keys or secrets. This is not an indicator of compromise: it is a constraint on how the report may be stored and shared.
linux_pam	ALTA	T1556.003	pam_exec|pam_python|nullok|pam_permit	Configurazione PAM che permette un bypass§PAM configuration allowing a bypass	pam_exec e pam_python eseguono codice a ogni autenticazione; nullok e pam_permit accettano password vuote o qualsiasi credenziale.§pam_exec and pam_python run code on every authentication; nullok and pam_permit accept empty passwords or any credential.
linux_kernel_modules	ALTA	T1014	out-of-tree|unsigned|non firmato|taint	Modulo kernel non firmato o fuori albero§Unsigned or out-of-tree kernel module	Un LKM caricato fuori dai pacchetti della distribuzione puo' nascondere processi, file e connessioni a tutto cio' che gira in user space.§An LKM loaded outside distribution packages can hide processes, files and connections from everything running in user space.
linux_webserver	ALTA	T1505.003	(cmd|exec|shell|passthru|system)=|union\s+select|\.\./\.\./|<\?php	Richiesta compatibile con una webshell§Request consistent with a web shell	Parametri che passano comandi, path traversal o SQL nella query string. Vanno confrontati con la risposta HTTP: un 200 su una di queste e' molto peggio di un 404.§Parameters passing commands, path traversal or SQL in the query string. Compare with the HTTP status: a 200 on one of these is far worse than a 404.
linux_persistence|linux_suid_caps	ALTA	T1053.003	(/tmp/|/dev/shm/|/var/tmp/)	Persistenza o SUID in una directory temporanea§Persistence or SUID in a temporary directory	/tmp, /dev/shm e /var/tmp sono scrivibili da tutti: niente di legittimo installa li' un meccanismo di avvio o un binario SUID.§/tmp, /dev/shm and /var/tmp are world-writable: nothing legitimate installs a startup mechanism or a SUID binary there.
linux_ssh	MEDIA	T1098.004	ssh-(rsa|ed25519|dss)|ecdsa-sha2	Chiave in authorized_keys§Key in authorized_keys	Una chiave pubblica autorizzata e' un accesso permanente che non passa dalla password e non scade. Vanno verificate una per una con il proprietario dell'account.§An authorized public key is permanent access that bypasses the password and never expires. Each must be verified with the account owner.
macos_persistence|macos_loginitems	ALTA	T1543.001	(/tmp/|/users/shared/|/private/var/tmp/)	LaunchAgent o LoginItem in una directory scrivibile§LaunchAgent or LoginItem in a writable directory	Un elemento di avvio che punta a una cartella condivisa o temporanea invece che a /Applications o ai bundle di sistema.§A startup item pointing at a shared or temporary folder rather than /Applications or system bundles.
macos_tcc	MEDIA	T1123	kTCCServiceScreenCapture|kTCCServiceListenEvent|kTCCServiceAccessibility|kTCCServiceMicrophone	Permesso TCC ad alto impatto concesso§High-impact TCC permission granted	Accessibility, cattura schermo, tastiera e microfono: con questi permessi un'applicazione vede e registra tutto quello che fa l'utente.§Accessibility, screen capture, keystrokes and microphone: with these an application sees and records everything the user does.
xplat_sqlite_recovery	MEDIA	T1555	password|passwd|token|api[_-]?key|secret	Credenziale in un record SQLite cancellato§Credential in a deleted SQLite record	Una stringa che sembra una credenziale recuperata da spazio non allocato: era stata cancellata dall'applicazione ma e' rimasta nel file.§A credential-looking string recovered from unallocated space: the application deleted it but it stayed in the file.
sigma	ALTA	T1059	.	Match di una regola Sigma§Sigma rule match	Una detection della comunita' ha riconosciuto un evento negli EVTX. Il livello originale della regola e' nella tabella del modulo; qui la severita' e' uniforme perche' FIUTO non puo' giudicare la qualita' di una regola di terze parti.§A community detection matched an event in the EVTX. The rule's own level is in the module table; the severity here is uniform because FIUTO cannot judge the quality of a third-party rule.
yara	ALTA	T1204	.	Match di una regola YARA§YARA rule match	Una firma esterna ha riconosciuto un file sul volume. Il peso dipende da chi ha scritto la regola: la severita' qui e' uniforme perche' FIUTO non puo' giudicare la qualita' di una regola di terze parti.§An external signature recognised a file on the volume. Its weight depends on who wrote the rule: the severity here is uniform because FIUTO cannot judge the quality of a third-party rule.
usb|setupapi	BASSA	T1052.001	.	Supporti rimovibili collegati§Removable media connected	Da solo non significa nulla. Conta per la correlazione: un supporto collegato nella stessa finestra in cui compaiono LNK e attivita' USN e' il profilo di un'esfiltrazione.§On its own it means nothing. It matters for correlation: media connected in the same window as LNK and USN activity is the shape of an exfiltration.
RULESEOF
}

# Scenari di correlazione: insiemi di moduli che, se attivi nella stessa
# finestra, suggeriscono un'ipotesi. Sono IPOTESI DA VERIFICARE e il report le
# presenta come tali — la coincidenza temporale non e' una prova di nesso.
#
#   moduli_regex_separati_da_virgola	severita'	ipotesi§hypothesis
correlation_scenarios_tsv() {
    cat << 'SCENEOF'
usb|setupapi,lnk|jumplists|opensave|shellbags,usn|mft	ALTA	Supporto rimovibile collegato mentre venivano aperti file e il file system registrava modifiche: profilo compatibile con una copia di dati verso l'esterno.§Removable media connected while files were being opened and the file system recorded changes: consistent with copying data out.
ps_history|ps_scriptblock|ps_transcripts,browser|webcache|bits|cloud_sync	ALTA	Esecuzione PowerShell in prossimita' di attivita' di rete: possibile download ed esecuzione di codice.§PowerShell execution close to network activity: possible download and execution of code.
defender_quarantine,prefetch|amcache|bam|userassist	ALTA	Rilevamento antivirus in prossimita' di un'esecuzione registrata: il binario eseguito e quello messo in quarantena vanno confrontati.§Antivirus detection close to a recorded execution: the binary that ran and the one quarantined must be compared.
evtx|linux_logins|macos_logs,run_keys|services|scheduled_tasks|linux_persistence|macos_persistence	ALTA	Accesso al sistema seguito a breve da una modifica dei meccanismi di avvio: profilo di un impianto di persistenza dopo l'ingresso.§System access shortly followed by a change to startup mechanisms: the shape of persistence being installed after entry.
linux_ssh|linux_logins,linux_pam|linux_accounts	ALTA	Accesso remoto in prossimita' di modifiche ad autenticazione o utenze.§Remote access close to changes in authentication or accounts.
SCENEOF
}

# Finestra (minuti) entro cui due eventi di moduli diversi sono considerati
# correlati. 30 minuti e' un compromesso: piu' stretto perde le catene manuali
# di un operatore, piu' largo correla mezza giornata di lavoro normale.
CORRELATION_WINDOW_MIN=30

# Genera executive_summary.html e findings.json in REPORT_BASE_DIR.
generate_executive_summary() {
    [[ -n "$REPORT_BASE_DIR" && -d "$REPORT_BASE_DIR" ]] || return 0

    # I report da leggere sono quelli dei moduli: la dashboard e il riepilogo
    # stesso non sono fonti, sono aggregatori.
    local LIST; LIST=$(mktemp); register_tmp "$LIST"
    local R
    for R in "${GENERATED_REPORTS[@]:-}"; do
        [[ -f "$R" ]] || continue
        case "$(basename "$R")" in
            index.html|executive_summary.html) continue ;;
        esac
        printf '%s\n' "$R" >> "$LIST"
    done
    if [[ ! -s "$LIST" ]]; then
        warn "$(L "Nessun report da riepilogare." "No report to summarise.")"
        return 0
    fi

    local RULES; RULES=$(mktemp); register_tmp "$RULES"
    finding_rules_tsv > "$RULES"
    local SCEN; SCEN=$(mktemp); register_tmp "$SCEN"
    correlation_scenarios_tsv > "$SCEN"
    # Si passa la regex compilata dal motore IoC, non l'elenco dei valori: i
    # confini per tipo devono essere gli stessi di check_ioc, e due
    # implementazioni degli stessi confini divergono al primo aggiustamento.
    local IOCF; IOCF=$(mktemp); register_tmp "$IOCF"
    printf '%s' "${IOC_REGEX:-}" > "$IOCF"

    local JSON="${REPORT_BASE_DIR}/findings.json"
    local BODY; BODY=$(mktemp); register_tmp "$BODY"

    info "$(L "Correlazione dei report in corso..." "Correlating reports...")"
    FIUTO_LANG="${LANG:-en}" FIUTO_WINDOW="$CORRELATION_WINDOW_MIN" \
    run_py_with_lib pylib_timeline "$LIST" "$RULES" "$SCEN" "$IOCF" "$JSON" "$BODY" << 'PYEOF' 2>/dev/null
import sys, os, re, json, html, datetime

list_path, rules_path, scen_path, ioc_path, json_path, body_path = sys.argv[1:7]
IT = os.environ.get('FIUTO_LANG', 'en') == 'it'
WINDOW = int(os.environ.get('FIUTO_WINDOW', '30')) * 60


def L(it, en):
    return it if IT else en


def bilingual(s):
    return L(*s.split('§', 1)) if '§' in s else s


SEV_ORDER = {'CRITICA': 0, 'ALTA': 1, 'MEDIA': 2, 'BASSA': 3}
SEV_WEIGHT = {'CRITICA': 40, 'ALTA': 15, 'MEDIA': 5, 'BASSA': 1}
SEV_LABEL = {'CRITICA': L('CRITICA', 'CRITICAL'), 'ALTA': L('ALTA', 'HIGH'),
             'MEDIA': L('MEDIA', 'MEDIUM'), 'BASSA': L('BASSA', 'LOW')}
SEV_CSS = {'CRITICA': 'crit', 'ALTA': 'high', 'MEDIA': 'med', 'BASSA': 'low'}

MAX_EXAMPLES = 6          # esempi mostrati per finding
MAX_EXAMPLE_LEN = 300

reports = [p.strip() for p in open(list_path, encoding='utf-8') if p.strip()]
try:
    _ioc_src = open(ioc_path, encoding='utf-8', errors='replace').read().strip()
    ioc_rx = re.compile(_ioc_src, re.I) if _ioc_src else None
except re.error:
    ioc_rx = None

rules = []
for line in open(rules_path, encoding='utf-8'):
    f = line.rstrip('\n').split('\t')
    if len(f) != 6:
        continue
    slug_rx, sev, mitre, pat, title, why = f
    try:
        rules.append((re.compile(f'^({slug_rx})$'), sev, mitre,
                      re.compile(pat, re.I), title, why))
    except re.error:
        continue

scenarios = []
for line in open(scen_path, encoding='utf-8'):
    f = line.rstrip('\n').split('\t')
    if len(f) != 3:
        continue
    groups = [re.compile(f'^({g})$') for g in f[0].split(',')]
    scenarios.append((groups, f[1], f[2]))


def slug_of(path):
    return re.sub(r'_\d{8}_\d{6}$', '', os.path.basename(os.path.dirname(path)))


def rel(path):
    base = os.path.dirname(json_path)
    return os.path.relpath(path, base)


findings = []
events = []          # (epoch_ordinabile, slug, message)
per_module = {}

for path in reports:
    slug = slug_of(path)
    data = read_report_data(path)
    sens = read_sensitive_lines(path)
    evs = extract_events(path)
    for ev in evs:
        events.append((ev['datetime'], slug, ev['message']))
    per_module[slug] = {'report': rel(path), 'rows': len(data), 'events': len(evs)}

    # 1) righe marcate dal modulo stesso
    if sens:
        findings.append({
            'severity': 'MEDIA',
            'module': slug,
            'report': rel(path),
            'mitre': '',
            'title': L('Righe segnalate dal modulo', 'Rows flagged by the module'),
            'why': L('Il modulo ha marcato queste righe con le proprie parole chiave: '
                     'sono il punto in cui l\'autore dell\'artefatto guarderebbe per primo.',
                     'The module flagged these rows with its own keywords: they are where '
                     'someone who knows the artefact would look first.'),
            'count': len(sens),
            'source': 'modulo',
            'examples': [s[:MAX_EXAMPLE_LEN] for s in sens[:MAX_EXAMPLES]],
        })

    # 2) IoC caricati con --ioc
    if ioc_rx is not None:
        hits = [d for d in data if ioc_rx.search(d)]
        if hits:
            findings.append({
                'severity': 'ALTA',
                'module': slug,
                'report': rel(path),
                'mitre': '',
                'title': L('Corrispondenza con gli IoC forniti', 'Match against the supplied IoCs'),
                'why': L('Il dato contiene uno degli indicatori passati con --ioc.',
                         'The data contains one of the indicators passed with --ioc.'),
                'count': len(hits),
                'source': 'ioc',
                'examples': [h[:MAX_EXAMPLE_LEN] for h in hits[:MAX_EXAMPLES]],
            })

    # 3) regole esplicite, solo sul dato
    for slug_rx, sev, mitre, pat, title, why in rules:
        if not slug_rx.match(slug):
            continue
        hits = [d for d in data if pat.search(d)]
        if not hits:
            continue
        findings.append({
            'severity': sev,
            'module': slug,
            'report': rel(path),
            'mitre': mitre,
            'title': bilingual(title),
            'why': bilingual(why),
            'count': len(hits),
            'source': 'regola',
            'examples': [h[:MAX_EXAMPLE_LEN] for h in hits[:MAX_EXAMPLES]],
        })

findings.sort(key=lambda f: (SEV_ORDER.get(f['severity'], 9), -f['count'], f['module']))

# --- correlazione temporale ------------------------------------------------
# Gli eventi sono ordinati per stringa ISO, che per questo formato coincide con
# l'ordine cronologico. La finestra scorre e si chiude quando il salto supera
# CORRELATION_WINDOW_MIN.
_EPOCH = datetime.datetime(1970, 1, 1)


def to_secs(iso):
    """Secondi assoluti. Un'aritmetica approssimata sui campi della data
    (giorni=31 al mese) sballerebbe i salti attorno ai cambi di mese e
    spezzerebbe cluster legittimi senza che nulla lo segnali."""
    try:
        return int((datetime.datetime.fromisoformat(iso[:19]) - _EPOCH).total_seconds())
    except Exception:
        return None


events = [(to_secs(dt), dt, slug, msg) for dt, slug, msg in events]
events = sorted((e for e in events if e[0] is not None), key=lambda e: e[0])

clusters = []
cur = []
for e in events:
    if cur and e[0] - cur[-1][0] > WINDOW:
        clusters.append(cur)
        cur = []
    cur.append(e)
if cur:
    clusters.append(cur)

correlations = []
for cl in clusters:
    mods = sorted({e[2] for e in cl})
    if len(mods) < 2:
        continue
    for groups, sev, hypo in scenarios:
        matched = []
        for g in groups:
            m = [x for x in mods if g.match(x)]
            if not m:
                break
            matched.append(m[0])
        else:
            correlations.append({
                'severity': sev,
                'from': cl[0][1],
                'to': cl[-1][1],
                'modules': matched,
                'all_modules': mods,
                'events': len(cl),
                'hypothesis': bilingual(hypo),
            })
            break

correlations.sort(key=lambda c: (SEV_ORDER.get(c['severity'], 9), c['from']))

# --- punteggio -------------------------------------------------------------
# Somma di pesi, tetto 100. E' una CODA DI LAVORO, non una misura di
# compromissione: due macchine con lo stesso punteggio possono avere storie
# opposte. Il report riporta la formula perche' un numero senza definizione e'
# peggio di nessun numero.
score = min(100, sum(SEV_WEIGHT.get(f['severity'], 0) for f in findings)
            + sum(SEV_WEIGHT.get(c['severity'], 0) for c in correlations))

counts = {s: sum(1 for f in findings if f['severity'] == s) for s in SEV_ORDER}

with open(json_path, 'w', encoding='utf-8') as fh:
    json.dump({
        'schema_version': 1,
        'score': score,
        'score_formula': 'CRITICA=40 ALTA=15 MEDIA=5 BASSA=1, somma limitata a 100',
        'score_meaning': ('Ordina la coda di lavoro. NON misura la compromissione: '
                          'e\' la somma delle severita\' dei riscontri, non un giudizio.'),
        'counts': counts,
        'modules': per_module,
        'findings': findings,
        'correlations': correlations,
        'correlation_window_minutes': WINDOW // 60,
    }, fh, indent=2, ensure_ascii=False)


# --- corpo HTML ------------------------------------------------------------
def esc(s):
    return html.escape(str(s))


out = []
out.append("<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>")
out.append('<b>' + esc(L('Come leggere questa pagina', 'How to read this page')) + '</b><br>')
out.append(esc(L(
    "Questa pagina ordina il lavoro, non lo conclude. Ogni voce dice da quale modulo viene, "
    "su quale dato ha fatto match e perche' quel dato conta: il giudizio su cosa sia successo "
    "resta all'analista, che deve aprire il report del modulo e guardare il contesto. "
    "L'assenza di riscontri non e' un attestato di pulizia — significa che le regole applicate, "
    "che sono poche e deliberatamente conservative, non hanno trovato nulla.",
    "This page orders the work, it does not conclude it. Every entry states which module it comes "
    "from, which data it matched and why that data matters: the judgement on what happened stays "
    "with the analyst, who must open the module report and look at the context. "
    "No findings is not a clean bill of health — it means the applied rules, which are few and "
    "deliberately conservative, found nothing.")))
out.append('</div></div>')

if correlations:
    out.append("<h2 style='font-family:var(--mono);font-size:.9rem;margin:1.4rem 0 .7rem;color:var(--accent4)'>"
               + esc(L('Correlazioni temporali fra moduli', 'Cross-module time correlations')) + '</h2>')
    out.append("<div class='card' style='margin-bottom:1rem'><div style='padding:.8rem 1.5rem;font-size:.78rem;color:var(--text-mid);line-height:1.6'>"
               + esc(L(
                   "Moduli diversi con attivita' nella stessa finestra di %d minuti. La coincidenza "
                   "temporale non e' un nesso di causa: sono ipotesi da verificare aprendo i report "
                   "citati." % (WINDOW // 60),
                   "Different modules active within the same %d-minute window. Temporal coincidence "
                   "is not causation: these are hypotheses to verify by opening the reports named."
                   % (WINDOW // 60)))
               + '</div></div>')
    out.append("<div class='cards'>")
    for c in correlations[:30]:
        out.append("<div class='card'><div class='card-header'>")
        out.append("<div class='badge sev-%s'>%s</div>" % (SEV_CSS[c['severity']], esc(SEV_LABEL[c['severity']])))
        out.append("<div class='user-info' style='margin-left:.8rem'><div class='uname'>%s → %s</div>"
                   % (esc(c['from'].replace('T', ' ')), esc(c['to'].replace('T', ' '))))
        out.append("<div class='upath'>%s · %d %s</div></div></div>"
                   % (esc(' + '.join(c['modules'])), c['events'], esc(L('eventi', 'events'))))
        out.append("<div style='padding:1rem 1.5rem;font-size:.82rem;line-height:1.7'>%s" % esc(c['hypothesis']))
        others = [m for m in c['all_modules'] if m not in c['modules']]
        if others:
            out.append("<div style='margin-top:.6rem;font-size:.72rem;color:var(--text-dim)'>%s %s</div>"
                       % (esc(L('Altri moduli attivi nella finestra:', 'Other modules active in the window:')),
                          esc(', '.join(others[:12]))))
        out.append('</div></div>')
    out.append('</div>')

out.append("<h2 style='font-family:var(--mono);font-size:.9rem;margin:1.6rem 0 .7rem;color:var(--accent)'>"
           + esc(L('Riscontri per severita\'', 'Findings by severity')) + '</h2>')

if not findings:
    out.append("<div class='card'><div style='padding:1.4rem 1.5rem;font-size:.85rem'>%s</div></div>"
               % esc(L('Nessuna regola ha prodotto riscontri. Vedi l\'avvertenza in testa: '
                       'non equivale a un esito negativo dell\'analisi.',
                       'No rule produced a finding. See the notice above: this is not equivalent '
                       'to a negative analysis result.')))
else:
    out.append("<div class='cards'>")
    for f in findings:
        out.append("<div class='card'><div class='card-header'>")
        out.append("<div class='badge sev-%s'>%s</div>" % (SEV_CSS[f['severity']], esc(SEV_LABEL[f['severity']])))
        out.append("<div class='user-info' style='margin-left:.8rem'><div class='uname'>%s</div>" % esc(f['title']))
        sub = f['module']
        if f['mitre']:
            sub += ' · MITRE ' + f['mitre']
        out.append("<div class='upath'>%s</div></div>" % esc(sub))
        out.append("<div style='margin-left:auto;font-family:var(--mono);font-size:1.1rem;font-weight:800'>%d</div>"
                   % f['count'])
        out.append('</div>')
        out.append("<div style='padding:.9rem 1.5rem;font-size:.8rem;line-height:1.7'>%s" % esc(f['why']))
        if f['examples']:
            out.append("<pre style='margin-top:.7rem;padding:.7rem;background:var(--bg);border:1px solid var(--border);"
                       "font-size:.7rem;overflow-x:auto;white-space:pre-wrap;word-break:break-all'>%s</pre>"
                       % esc('\n'.join(f['examples'])))
            if f['count'] > len(f['examples']):
                out.append("<div style='font-size:.7rem;color:var(--text-dim)'>%s</div>"
                           % esc(L('… e altri %d nel report del modulo' % (f['count'] - len(f['examples'])),
                                   '… and %d more in the module report' % (f['count'] - len(f['examples'])))))
        out.append("<div style='margin-top:.7rem;font-size:.72rem'><a href='%s' style='color:var(--accent)'>%s</a></div>"
                   % (esc(f['report']), esc(L('Apri il report del modulo', 'Open the module report'))))
        out.append('</div></div>')
    out.append('</div>')

with open(body_path, 'w', encoding='utf-8') as fh:
    fh.write('\n'.join(out))
PYEOF

    if [[ ! -s "$JSON" ]]; then
        warn "$(L "Riepilogo non generato." "Summary not generated.")"
        return 0
    fi

    local SCORE NCRIT NHIGH NMED NLOW NCORR
    read -r SCORE NCRIT NHIGH NMED NLOW NCORR < <("$PY3" -c '
import json, sys
d = json.load(open(sys.argv[1]))
c = d["counts"]
print(d["score"], c["CRITICA"], c["ALTA"], c["MEDIA"], c["BASSA"], len(d["correlations"]))' "$JSON" 2>/dev/null)

    local OUT="${REPORT_BASE_DIR}/executive_summary.html"
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local STATS
    STATS="$(stat_box "$(L "Priorità" "Priority")" "$SCORE" "$([[ "${SCORE:-0}" -ge 40 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Critici" "Critical")" "${NCRIT:-0}")"
    STATS+="$(stat_box "$(L "Alti" "High")" "${NHIGH:-0}")"
    STATS+="$(stat_box "$(L "Medi" "Medium")" "${NMED:-0}" "info")"
    STATS+="$(stat_box "$(L "Bassi" "Low")" "${NLOW:-0}" "info")"
    STATS+="$(stat_box "$(L "Correlazioni" "Correlations")" "${NCORR:-0}" "$([[ "${NCORR:-0}" -gt 0 ]] && echo warn || echo info)")"

    {
        html_header "Executive Summary"
        html_page_header "SUM" "Executive Summary" \
            "$(L "Riepilogo di sessione — priorità e correlazioni" "Session summary — priorities and correlations")" \
            "$SCAN" "$WIN_ROOT"
        printf "<div class='statsbar'>%s</div>\n" "$STATS"
        echo "<main>"
        pre_style_block
        _summary_style_block
        time_window_html
        cat "$BODY"
        printf "<div class='card' style='margin-top:1.2rem'><div style='padding:1rem 1.5rem;font-size:.75rem;color:var(--text-mid);line-height:1.7'><b>%s</b> %s</div></div>\n" \
            "$(L "Come è calcolata la priorità:" "How the priority is computed:")" \
            "$(L "somma dei pesi dei riscontri (critico 40, alto 15, medio 5, basso 1), limitata a 100. Ordina la coda di lavoro; non misura la compromissione. Due macchine con lo stesso valore possono avere storie opposte." \
                 "sum of finding weights (critical 40, high 15, medium 5, low 1), capped at 100. It orders the work queue; it does not measure compromise. Two machines with the same value can have opposite stories.")"
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$OUT"

    register_report "$OUT"
    echo ""
    ok "$(L "Executive summary:" "Executive summary:") ${BOLD}${OUT}"
    info "$(L "Riscontri:" "Findings:") ${BOLD}$(( ${NCRIT:-0} + ${NHIGH:-0} + ${NMED:-0} + ${NLOW:-0} ))${RESET}  ·  $(L "critici/alti:" "critical/high:") ${BOLD}$(( ${NCRIT:-0} + ${NHIGH:-0} ))${RESET}  ·  $(L "correlazioni:" "correlations:") ${BOLD}${NCORR:-0}"
    log_msg "[SUMMARY] score=$SCORE crit=${NCRIT:-0} high=${NHIGH:-0} corr=${NCORR:-0}"
    open_report_prompt "$OUT"
}

# Stili delle sole severita': il resto lo porta html_header.
_summary_style_block() {
    cat << 'CSSEOF'
<style>
  .badge.sev-crit { background:#ff7b72; color:#0d1117; }
  .badge.sev-high { background:#f0883e; color:#0d1117; }
  .badge.sev-med  { background:#58a6ff; color:#0d1117; }
  .badge.sev-low  { background:#1e2d3d; color:#c9d1d9; }
</style>
CSSEOF
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
            dim_msg "$USERNAME — $(L "PSReadLine non trovata" "PSReadLine not found")"
            continue
        fi
        mapfile -t HIST_FILES < <(find "$PSRL_DIR" -maxdepth 1 -iname "*_history.txt" -type f 2>/dev/null)
        if [[ ${#HIST_FILES[@]} -eq 0 ]]; then
            warn "$USERNAME — $(L "PSReadLine trovata ma nessun history" "PSReadLine found but no history")"
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
    info "$(L "Utenti con history:" "Users with history:") ${BOLD}$TOTAL_USERS${RESET}  |  File totali: ${BOLD}$TOTAL_FILES"

    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun file history trovato." "No history file found.")"; return 0; }
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
                  <div class='uname' style='font-size:.85rem'>$(html_esc "$FNAME")</div>
                  <div class='upath'>$(html_esc "$FPATH")</div>
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
              <div class='uname'>$(html_esc "$USER")</div>
              <div class='upath'>$(html_esc "$PSRL_PATH")</div>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 7 — Notepad TabState
# ================================================================
module_notepad_tabstate() {
    section_header "$(L "Notepad TabState — Tab Rimasti Aperti" "Notepad TabState — Open Tabs")" "$MAGENTA"
    check_win_root || return 1

    local NOTEPAD_PKG="Microsoft.WindowsNotepad_8wekyb3d8bbwe"
    local TABSTATE_REL="AppData/Local/Packages"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_PATHS=() RES_COUNTS=()
    declare -a RES_FILES=() RES_SIZES=() RES_MTIMES=() RES_CTIMES=()

    # Parser .bin inline (stessa logica di notepad_tabstate.sh)
    local PARSER_PY=""
    PARSER_PY=$(mktemp /tmp/npad_parse_XXXXXX.py)
    register_tmp "$PARSER_PY"
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
        [[ -z "$PACKAGES_DIR" ]] && { dim_msg "$USERNAME — $(L "AppData\\Local\\Packages non trovata" "AppData\\Local\\Packages not found")"; continue; }
        local NOTEPAD_DIR
        NOTEPAD_DIR=$(find "$PACKAGES_DIR" -maxdepth 1 -iname "${NOTEPAD_PKG}*" -type d 2>/dev/null | head -1)
        [[ -z "$NOTEPAD_DIR" ]] && { dim_msg "$USERNAME — $(L "Notepad UWP non installato" "Notepad UWP not installed")"; continue; }
        local TABSTATE_DIR
        TABSTATE_DIR=$(ci_find_dir "$NOTEPAD_DIR" "LocalState/TabState")
        [[ -z "$TABSTATE_DIR" || ! -d "$TABSTATE_DIR" ]] && { warn "$USERNAME — $(L "TabState non trovata" "TabState not found")"; continue; }
        mapfile -t BIN_FILES < <(find "$TABSTATE_DIR" -maxdepth 1 -iname "*.bin" -type f -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        if [[ ${#BIN_FILES[@]} -eq 0 ]]; then
            # Fallback se printf %T@ non è supportato (BSD/macOS)
            mapfile -t BIN_FILES < <(find "$TABSTATE_DIR" -maxdepth 1 -iname "*.bin" -type f -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
        fi
        local COUNT=${#BIN_FILES[@]}
        [[ $COUNT -eq 0 ]] && { warn "$USERNAME — $(L "TabState vuota" "TabState empty")"; continue; }
        ok "$USERNAME — $COUNT $(L "file .bin trovati" ".bin files found")"
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
    info "$(L "Utenti:" "Users:") ${BOLD}$TOTAL_USERS${RESET}  |  File .bin: ${BOLD}$TOTAL_FILES"
    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun tab Notepad trovato." "No Notepad tab found.")"; return 0; }
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
                [[ -n "$PPATH" ]] && SPATH_HTML="<div class='dim mono' style='font-size:.65rem;margin:.3rem 0'>📄 $(html_esc "$PPATH")</div>"
                if [[ -n "$PTEXT" ]]; then
                    CONTENT_HTML="${TBADGE}${SPATH_HTML}<pre style='font-family:var(--mono);font-size:.75rem;white-space:pre-wrap;max-height:250px;overflow-y:auto;margin-top:.4rem;color:var(--text)'>$(html_esc "$PTEXT")</pre>"
                else
                    CONTENT_HTML="${TBADGE}<span class='dim' style='margin-left:.5rem;font-size:.7rem'>nessun testo leggibile</span>"
                fi
            fi
            ROWS+="<tr>
              <td class='mono ok' style='white-space:nowrap;font-size:.7rem'>${FCTIME}</td>
              <td class='mono' style='font-size:.75rem'>$(html_esc "$FNAME")</td>
              <td class='mono mid' style='white-space:nowrap'>${FMTIME}</td>
              <td class='mono mid'>${FSIZE} B</td>
            </tr>
            <tr><td colspan='4' style='padding:.5rem 1rem 1rem'>${CONTENT_HTML}</td></tr>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>NT</div>
            <div><div class='uname'>$(html_esc "$USER")</div><div class='upath'>$(html_esc "$PATH_FULL")</div></div>
            <div class='badge'>${COUNT} tab</div>
          </div>
          <table><thead><tr><th>$(L "Creato" "Created")</th><th>GUID / File</th><th>$(L "Modificato" "Modified")</th><th>Dim.</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 2 — IFEO (Image File Execution Options)
# ================================================================
module_ifeo() {
    section_header "IFEO — Image File Execution Options" "$RED"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE
    HIVE=$(get_hive "SOFTWARE")
    if [[ -z "$HIVE" || ! -f "$HIVE" ]]; then
        err "$(L "Hive SOFTWARE non trovato in" "SOFTWARE hive not found in") $WIN_ROOT"
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
    info "$(L "Voci con Debugger/GFlags:" "Entries with Debugger/GFlags:") ${BOLD}$FOUND_COUNT${RESET}  |  Sospette: ${RED}${BOLD}$SUSPICIOUS_COUNT${RESET}"

    [[ $FOUND_COUNT -eq 0 ]] && { ok "$(L "Nessuna voce IFEO con Debugger/GFlags trovata." "No IFEO entries with Debugger/GFlags found.")"; return 0; }
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
        local DBG_HTML="-"
        [[ -n "$DBG" ]] && DBG_HTML="<span class='bad'>$(html_esc "$DBG")</span>"
        ROWS+="<tr ${ROW_CLASS}>
          <td class='mono'>$(html_esc "$EXE")$(${IS_SENS} && echo " <span class='badge warn'>⚠ sensibile</span>" || true)</td>
          <td class='mono'>${DBG_HTML}</td>
          <td class='mono mid'>$([[ -n "$GFLAG" ]] && html_esc "$GFLAG" || echo "-")</td>
          <td class='mono dim'>$([[ -n "$VDLL" ]] && html_esc "$VDLL" || echo "-")</td>
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
          <thead><tr><th>$(L "Eseguibile" "Executable")</th><th>Debugger</th><th>GlobalFlag</th><th>VerifierDLLs</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 3 — BAM (Background Activity Moderator)
# ================================================================
module_bam() {
    section_header "BAM — Background Activity Moderator" "$BLUE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE
    HIVE=$(get_hive "SYSTEM")
    if [[ -z "$HIVE" || ! -f "$HIVE" ]]; then
        err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found")"
        return 1
    fi

    info "$(L "Parsing hive SYSTEM:" "Parsing SYSTEM hive:") $HIVE"

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
        warn "$(L "Nessuna voce BAM trovata." "No BAM entries found.")"
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
    info "$(L "Totale eseguibili tracciati:" "Total tracked executables:") ${BOLD}$TOTAL${RESET}  |  SID distinti: ${BOLD}$SID_COUNT"
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
                    <div><div class='uname'>SID: $(html_esc "$PREV_SID")</div></div>
                    <div class='badge'>${SID_EXEC_COUNT} eseguibili</div>
                  </div>
                  <table><thead><tr><th>Timestamp (UTC)</th><th>$(L "Eseguibile" "Executable")</th></tr></thead>
                  <tbody>${SID_ROWS}</tbody></table></div>"
            fi
            PREV_SID="$SID"; SID_ROWS=""; SID_EXEC_COUNT=0
        fi
        SID_EXEC_COUNT=$((SID_EXEC_COUNT + 1))
        local EXE_ESC; EXE_ESC=$(html_esc "$EXEPATH")
        local EXE_HTML="<span class='mono'>${EXE_ESC}</span>"
        # Evidenzia percorsi sospetti
        if echo "$EXEPATH" | grep -qi "temp\|appdata\|public\|programdata\|downloads"; then
            EXE_HTML="<span class='mono warn'>${EXE_ESC}</span>"
        fi
        SID_ROWS+="<tr><td class='mono ok' style='white-space:nowrap'>$(html_esc "$TIMESTAMP")</td><td>${EXE_HTML}</td></tr>"
    done
    # Ultima card
    if [[ -n "$PREV_SID" ]]; then
        CARDS_HTML+="<div class='card'>
          <div class='card-header'><div class='uicon'>▣</div>
            <div><div class='uname'>SID: $(html_esc "$PREV_SID")</div></div>
            <div class='badge'>${SID_EXEC_COUNT} eseguibili</div>
          </div>
          <table><thead><tr><th>Timestamp (UTC)</th><th>$(L "Eseguibile" "Executable")</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 5 — Run Keys & Persistenza nel Registro
# ================================================================
module_run_keys() {
    section_header "$(L "Run Keys & Persistenza Registro" "Run Keys & Registry Persistence")" "$ORANGE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SW
    HIVE_SW=$(get_hive "SOFTWARE")
    [[ -z "$HIVE_SW" ]] && { err "$(L "Hive SOFTWARE non trovato" "SOFTWARE hive not found")"; return 1; }

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

    info "$(L "Scansione hive SOFTWARE (HKLM)..." "Scanning SOFTWARE hive (HKLM)...")"
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
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
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
    info "$(L "Voci totali trovate:" "Total entries found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { ok "$(L "Nessuna voce Run trovata." "No Run entries found.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "run_keys")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r HKEY VNAME VVAL <<< "$E"
        local CSS=""
        echo "$VVAL" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|powershell\|cmd.exe\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin" && CSS="class='bad'"
        ROWS+="<tr><td class='mono dim'>$(html_esc "$HKEY")</td><td class='mono'>$(html_esc "$VNAME")</td>
               <td class='mono' ${CSS}>$(html_esc "$VVAL")</td></tr>"
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
          <thead><tr><th>$(L "Chiave" "Key")</th><th>$(L "Nome valore" "Value name")</th><th>$(L "Dati (comando)" "Data (command)")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 6 — Prefetch
# ================================================================
module_prefetch() {
    section_header "$(L "Prefetch — Eseguibili Tracciati" "Prefetch — Execution History")" "$GREEN"
    check_win_root || return 1

    local PREF_DIR
    PREF_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/Prefetch")
    if [[ -z "$PREF_DIR" || ! -d "$PREF_DIR" ]]; then
        warn "$(L "Directory Prefetch non trovata (potrebbe essere disabilitato o sistema su SSD)" "Prefetch directory not found (may be disabled or SSD system)")"
        return 0
    fi

    mapfile -t PF_FILES < <(find "$PREF_DIR" -maxdepth 1 -iname "*.pf" -type f 2>/dev/null | sort)
    local TOTAL=${#PF_FILES[@]}
    if [[ $TOTAL -eq 0 ]]; then
        warn "$(L "Nessun file .pf trovato" "No .pf files found")"
        return 0
    fi

    info "$(L "Trovati" "Found") $TOTAL file .pf in: $PREF_DIR"

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
    info "$(L "File Prefetch:" "Prefetch Files:") ${BOLD}$TOTAL"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "prefetch")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${PF_ENTRIES[@]}"; do
        IFS='|' read -r EXENAME FNAME FSIZE FMTIME <<< "$E"
        local CSS=""
        echo "$EXENAME" | grep -qi "powershell\|cmd\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin\|mimikatz\|procdump\|psexec\|wce\|fgdump" && CSS="class='warn'"
        ROWS+="<tr><td class='mono' ${CSS}>$(html_esc "$EXENAME")</td>
               <td class='mono dim'>$(html_esc "$FNAME")</td>
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
          <thead><tr><th>$(L "Eseguibile" "Executable")</th><th>File .pf</th><th>$(L "Ultima esecuzione (mtime)" "Last run (mtime)")</th><th>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        warn "$(L "Directory Tasks non trovata" "Tasks directory not found")"
        return 0
    fi

    mapfile -t TASK_FILES < <(find "$TASKS_DIR" -type f ! -iname "*.job" -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
    if [[ ${#TASK_FILES[@]} -eq 0 ]]; then
        mapfile -t TASK_FILES < <(find "$TASKS_DIR" -type f ! -iname "*.job" -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
    fi
    local TOTAL=${#TASK_FILES[@]}
    info "$(L "Trovati" "Found") $TOTAL task in: $TASKS_DIR"

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
    info "$(L "Task totali:" "Total tasks:") ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"

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
            <th>$(L "Dettagli" "Details")</th>
          </tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 9 — USB / Dispositivi Rimovibili
# ================================================================
module_usb() {
    section_header "$(L "USB — Dispositivi Rimovibili" "USB — Removable Devices")" "$BLUE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SYS
    HIVE_SYS=$(get_hive "SYSTEM")
    [[ -z "$HIVE_SYS" ]] && { err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found")"; return 1; }

    info "$(L "Parsing dispositivi USB da hive SYSTEM..." "Parsing USB devices from SYSTEM hive...")"

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
    info "$(L "Dispositivi USB storici:" "Historical USB devices:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun dispositivo USB trovato nell'hive." "No USB device found in hive.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "usb")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for USB_E in "${USB_ROWS[@]}"; do
        IFS='§' read -r DTYPE SERIAL FNAME <<< "$USB_E"
        ROWS+="<tr>
          <td class='mono'>$(html_esc "$DTYPE")</td>
          <td class='mono dim'>$(html_esc "$SERIAL")</td>
          <td class='mono'>$([[ -n "$FNAME" ]] && html_esc "$FNAME" || echo "-")</td>
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
          <thead><tr><th>$(L "Tipo" "Type")</th><th>$(L "Seriale" "Serial")</th><th>$(L "Nome" "Name")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        [[ -z "$RECENT_DIR" ]] && { dim_msg "$USERNAME — $(L "Recent non trovata" "Recent not found")"; continue; }

        mapfile -t LNK_FILES < <(find "$RECENT_DIR" -maxdepth 1 -iname "*.lnk" -type f -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2- | head -50)
        if [[ ${#LNK_FILES[@]} -eq 0 ]]; then
            mapfile -t LNK_FILES < <(find "$RECENT_DIR" -maxdepth 1 -iname "*.lnk" -type f -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null | head -50)
        fi
        local COUNT=${#LNK_FILES[@]}
        [[ $COUNT -eq 0 ]] && { dim_msg "$USERNAME — $(L "nessun .lnk trovato" "no .lnk files found")"; continue; }

        ok "$USERNAME — $COUNT $(L "file .lnk recenti" "recent .lnk files")"
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
    info "$(L "Utenti:" "Users:") ${BOLD}$TOTAL_USERS${RESET}  |  .lnk trovati: ${BOLD}$TOTAL_LNK"
    [[ $TOTAL_LNK -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "lnk")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | sort -t'|' -k3,3r)

    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r UNAME FNAME FMTIME TARGET <<< "$E"
        ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
               <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
               <td class='mono'>$(html_esc "$FNAME")</td>
               <td class='mono mid' style='word-break:break-all;font-size:.72rem'>$([[ -n "$TARGET" ]] && html_esc "$TARGET" || echo "-")</td></tr>"
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
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:16%;white-space:nowrap'>$(L "Data accesso" "Access date")</th><th style='width:24%'>File .lnk</th><th>Target path</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 4 — Cache RDP (Terminal Server Client)
# ================================================================
module_rdp_cache() {
    section_header "$(L "Cache RDP — Terminal Server Client" "RDP Cache — Terminal Server Client")" "$CYAN"
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
            dim_msg "$USERNAME — $(L "Cache RDP non trovata" "RDP cache not found")"
            continue
        fi
        mapfile -t CACHE_FILES < <(find "$CACHE_DIR" -maxdepth 1 -type f \( -iname "*.bmc" -o -iname "*.bin" \) -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        if [[ ${#CACHE_FILES[@]} -eq 0 ]]; then
            mapfile -t CACHE_FILES < <(find "$CACHE_DIR" -maxdepth 1 -type f \( -iname "*.bmc" -o -iname "*.bin" \) -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
        fi
        local COUNT=${#CACHE_FILES[@]}
        if [[ $COUNT -eq 0 ]]; then
            warn "$USERNAME — $(L "Directory cache trovata ma vuota" "Cache directory found but empty")"
            continue
        fi
        ok "$USERNAME — $COUNT $(L "file cache trovati in:" "cache files found in:") $CACHE_DIR"
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
    info "$(L "Ricerca server RDP recenti nel registro (NTUSER.DAT)..." "Searching recent RDP servers in registry (NTUSER.DAT)...")"
    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local NTUSER
        NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
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
                echo -e "  ${GREEN}${BOLD}$USERNAME — $(L "Server RDP trovati:" "RDP servers found:")${RESET}"
                while IFS=$'\t' read -r HOST UNAME; do
                    printf "      ${CYAN}%-40s${RESET}  ${DIM}utente: %s${RESET}\n" "$HOST" "${UNAME:--}"
                done <<< "$RDP_SERVERS"
            fi
        fi
    done < <(get_user_homes)

    separator
    info "Utenti con cache RDP: ${BOLD}$TOTAL_USERS${RESET}  |  $(L "File totali:" "Total files:") ${BOLD}$TOTAL_FILES"

    if [[ $TOTAL_FILES -gt 0 ]]; then
        echo ""
        info "${BOLD}$(L "Per analizzare le tile bitmap della cache usa bmc-tools:" "To analyze cache bitmap tiles use bmc-tools:")${RESET}"
        echo -e "    ${DIM}git clone https://github.com/ANSSI-FR/bmc-tools${RESET}"
        echo -e "    ${DIM}${PY3} bmc-tools.py -s <dir_cache> -d ./output/ -b${RESET}"
    fi

    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun file cache RDP trovato." "No RDP cache files found.")"; return 0; }
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
              <td class='mono'>$(html_esc "$FNAME")</td>
              <td class='mono mid' style='white-space:nowrap'>${FSIZE} B</td>
              <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
            </tr>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>RD</div>
            <div><div class='uname'>$(html_esc "$USER")</div><div class='upath'>$(html_esc "$DIR")</div></div>
            <div class='badge'>${#FENTRIES[@]} file</div>
          </div>
          <table><thead><tr><th>File</th><th>$(L "Dimensione" "Size")</th><th>$(L "Ultima modifica" "Last modified")</th></tr></thead>
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
            ${PY3} bmc-tools.py -s &lt;dir_cache&gt; -d ./output/ -b
          </div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 11 — Services (Servizi Windows)
# ================================================================
module_services() {
    section_header "$(L "Services — Servizi Windows" "Services — Windows Services")" "$RED"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SYS
    HIVE_SYS=$(get_hive "SYSTEM")
    [[ -z "$HIVE_SYS" || ! -f "$HIVE_SYS" ]] && { err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found")"; return 1; }

    info "$(L "Parsing servizi da hive SYSTEM:" "Parsing services from SYSTEM hive:") $HIVE_SYS"
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
    info "Servizi totali: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun servizio trovato." "No services found.")"; return 0; }

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
            <th style='width:13%'>$(L "Tipo" "Type")</th>
            <th>$(L "Dettagli" "Details")</th>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 0 — Esegui tutti i moduli
# ================================================================
# ================================================================
#  MODULO 12 — Event Log (Security / System / PowerShell / RDP)
# ================================================================
module_evtx() {
    section_header "$(L "Event Log — Analisi .evtx" "Event Log — .evtx Analysis")" "$RED"
    check_win_root || return 1

    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        warn "$(L "python-evtx non trovato." "python-evtx not found.")"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install python-evtx"
        return 1
    fi

    local EVTX_DIR
    EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    if [[ -z "$EVTX_DIR" || ! -d "$EVTX_DIR" ]]; then
        err "$(L "Directory Event Log non trovata" "Event Log directory not found")"; return 1
    fi
    info "$(L "Directory log:" "Log directory:") $EVTX_DIR"
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
    register_tmp "$PARSER_EVTX"; register_tmp "$TMP_ALL"

    local TOTAL_FOUND=0 SUSP_COUNT=0

    for EVTX_NAME in "${!EVTX_MAP[@]}"; do
        local EVTX_FILE
        EVTX_FILE=$(ci_find_file "$EVTX_DIR" "$EVTX_NAME")
        [[ -z "$EVTX_FILE" ]] && { dim_msg "$EVTX_NAME — $(L "non trovato" "not found")"; continue; }

        local EIDS_CSV="${EVTX_MAP[$EVTX_NAME]// /,}"
        local LABEL="${EVTX_NAME%.evtx}"; LABEL="${LABEL//%4/\/}"
        info "Parsing: $LABEL"

        local TMP_OUT; TMP_OUT=$(mktemp /tmp/dfir_evtxo_XXXXXX.json)
        "$PY3" "$PARSER_EVTX" "$EVTX_FILE" "$EIDS_CSV" "$TMP_OUT" 2>/dev/null \
            || echo "[]" > "$TMP_OUT"

        local COUNT
        COUNT=$("$PY3" -c "import json,sys; d=json.load(open(sys.argv[1])); print(len(d))" \
            "$TMP_OUT" 2>/dev/null || echo 0)
        [[ "$COUNT" -eq 0 ]] && { dim_msg "  $(L "Nessun evento corrispondente" "No matching events")"; rm -f "$TMP_OUT"; continue; }

        TOTAL_FOUND=$((TOTAL_FOUND + COUNT))
        ok "  $COUNT $(L "eventi trovati" "events found")"

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
    info "Totale eventi: ${BOLD}$TOTAL_FOUND${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL_FOUND -eq 0 ]] && { warn "$(L "Nessun evento trovato." "No events found.")"; return 0; }
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
        <div class='stitle'>$(L "Eventi per timestamp — EID · Timestamp · Sorgente · Dettagli" "Events by timestamp — EID · Timestamp · Source · Details")</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:11%'>EID</th>
            <th style='width:13%'>Timestamp</th>
            <th style='width:18%'>Log</th>
            <th>$(L "Dettagli" "Details")</th>
          </tr></thead>
          <tbody>${TABLE_ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 13 — Amcache + Shimcache (timeline esecuzione binari)
# ================================================================
module_amcache() {
    section_header "$(L "Amcache + Shimcache — Timeline Binari" "Amcache + Shimcache — Binary Timeline")" "$YELLOW"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
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
        info "$(L "Amcache trovato:" "Amcache found:") $AMC_PATH"
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
        ok "$(L "Amcache: $AMC_COUNT entry" "Amcache: $AMC_COUNT entries")"
    else
        warn "$(L "Amcache.hve non trovato" "Amcache.hve not found")"
    fi

    echo ""

    # ── Shimcache ─────────────────────────────────────────────────────
    local SYS_HIVE
    SYS_HIVE=$(get_hive "SYSTEM")
    declare -a SHIM_ENTRIES=()
    local SHIM_COUNT=0

    if [[ -n "$SYS_HIVE" ]]; then
        info "$(L "Shimcache da hive SYSTEM:" "Shimcache from SYSTEM hive:") $SYS_HIVE"
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
        ok "$(L "Shimcache: $SHIM_COUNT entry" "Shimcache: $SHIM_COUNT entries")"
    else
        warn "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found") — Shimcache saltato"
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
          <thead><tr><th>Path</th><th>LinkDate</th><th>$(L "SHA1 (troncato)" "SHA1 (truncated)")</th><th>Publisher</th></tr></thead>
          <tbody>${AMC_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>Shimcache — Eseguibili visti dal sistema</div>
        <div class='card'><table>
          <thead><tr><th>Path</th><th>Timestamp</th></tr></thead>
          <tbody>${SHIM_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 14 — Recycle Bin ($Recycle.Bin)
# ================================================================
module_recycle_bin() {
    section_header "$(L "Recycle Bin — File Eliminati" "Recycle Bin — Deleted Files")" "$GREEN"
    check_win_root || return 1


    # Cerca $Recycle.Bin sul volume (può stare nella root)
    local RB_DIR
    RB_DIR=$(find "$WIN_ROOT" -maxdepth 1 -iname "\$Recycle.Bin" -type d 2>/dev/null | head -1)
    if [[ -z "$RB_DIR" ]]; then
        warn "$(L "\$Recycle.Bin non trovato nella root del volume" "\$Recycle.Bin not found in volume root")"
        return 0
    fi

    info "Directory: $RB_DIR"
    echo ""

    declare -a ALL_ITEMS=()
    local TOTAL=0 SUSP_COUNT=0

    # Parser $I files — formato binario con path originale e timestamp
    local IPARSE=""
    IPARSE=$(mktemp /tmp/dfir_rb_XXXXXX.py)
    register_tmp "$IPARSE"
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
        [[ $SID_COUNT -gt 0 ]] && ok "  $(L "SID $SID: $SID_COUNT file" "SID $SID: $SID_COUNT files")"
    done

    separator
    info "File nel cestino: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Cestino vuoto o nessun \$I file trovato." "Recycle bin empty or no \$I file found.")"; return 0; }
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
          <td class='mono dim' style='font-size:.68rem;white-space:nowrap'>$(html_esc "$SID")</td>
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
        <div class='stitle'>$(L "File nel Cestino — SID · Path originale · Data eliminazione · Dimensione" "Recycle Bin Files — SID · Original path · Deletion date · Size")</div>
        <div class='card'><table>
          <thead><tr><th style='width:16%'>SID</th><th>$(L "Path originale" "Original path")</th><th style='width:14%'>$(L "Eliminato" "Deleted")</th><th style='width:8%'>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 15 — WMI Subscriptions (persistenza invisibile)
# ================================================================
module_wmi() {
    section_header "$(L "WMI Subscriptions — Persistenza Invisibile" "WMI Subscriptions — Fileless Persistence")" "$RED"
    check_win_root || return 1


    # Repository WMI
    local WMI_DIR
    WMI_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/wbem/Repository")
    if [[ -z "$WMI_DIR" || ! -d "$WMI_DIR" ]]; then
        warn "$(L "Repository WMI non trovato" "WMI repository not found")"
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
        warn "$(L "OBJECTS.DATA non trovato — analisi tramite strings sul repository" "OBJECTS.DATA not found — analysis via strings on repository")"
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
        info "$(L "Fallback: strings sul repository..." "Fallback: strings on repository...")"
        WMI_DATA=$(find "$WMI_DIR" -type f -print0 2>/dev/null | \
            xargs -0 strings 2>/dev/null | \
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
    info "Pattern WMI trovati: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna subscription WMI sospetta trovata." "No suspicious WMI subscription found.")"; return 0; }
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
          <td class='mono' style='white-space:nowrap;vertical-align:top;padding-top:.45rem;color:var(--accent4)'>$(html_esc "$WMI_TYPE")</td>
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
        <div class='stitle'>$(L "WMI Event Subscriptions — Tipo · Contenuto" "WMI Event Subscriptions — Type · Content")</div>
        <div class='card'><table>
          <thead><tr><th style='width:20%'>$(L "Tipo" "Type")</th><th>$(L "Valore / Query / Script" "Value / Query / Script")</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        warn "$(L "SRUDB.dat non trovato" "SRUDB.dat not found")"
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
        info "$(L "Estrazione tramite strings..." "Extracting via strings...")"
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
            <div class='card'><table><thead><tr><th>$(L "Stringa" "String")</th></tr></thead><tbody>${ROWS}</tbody></table></div>
            <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px'>
              <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Per analisi completa con traffico di rete e CPU per applicazione:</div>
              <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>pip install pyesedb<br>oppure usa: srum-dump (Mark Baggett) / ese2csv</div>
            </div></main>"
            html_footer "$SCAN" "$WIN_ROOT"
        } > "$REPORT_HTML"
        register_report "$REPORT_HTML"
        ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        err "$(L "sqlite3 non disponibile" "sqlite3 not available") nel Python rilevato"
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

    register_tmp "$HIST_PARSER"; register_tmp "$TMP_ROWS"
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
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
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
    info "Visite totali: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history browser trovata." "No browser history found.")"; return 0; }
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
        <div class='stitle'>$(L "Cronologia navigazione — Timestamp · Utente/Browser · URL · Titolo · Visite" "Browsing history — Timestamp · User/Browser · URL · Title · Visits")</div>
        <div class='card'><table id='brtable'>
          <thead><tr>
            <th style='width:12%'>Timestamp</th>
            <th style='width:12%'>$(L "Utente/Browser" "User/Browser")</th>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    local TOTAL_UA=0 TOTAL_RUN=0 TOTAL_TP=0 TOTAL_WW=0
    declare -a UA_ENTRIES=()   # UserAssist
    declare -a RUN_ENTRIES=()  # RunMRU
    declare -a TP_ENTRIES=()   # TypedPaths
    declare -a WW_ENTRIES=()   # WordWheelQuery

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && { dim_msg "$USERNAME — NTUSER.DAT $(L "non trovato" "not found")"; continue; }

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
    [[ $GRAND -eq 0 ]] && { warn "$(L "Nessuna voce trovata." "No entries found.")"; return 0; }
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
        UA_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
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
        RUN_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='word-break:break-all'><span ${SUSP_CSS}>$(html_esc "$CMD")</span></td></tr>"
    done

    # TypedPaths rows
    local TP_ROWS=""
    for E in "${TP_ENTRIES[@]}"; do
        IFS='§' read -r U K V <<< "$E"
        TP_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono mid' style='white-space:nowrap'>$(html_esc "$K")</td>
          <td class='mono' style='word-break:break-all'>$(html_esc "$V")</td></tr>"
    done

    # WordWheelQuery rows
    local WW_ROWS=""
    for E in "${WW_ENTRIES[@]}"; do
        IFS='§' read -r U TERM <<< "$E"
        WW_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
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
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Eseguibile" "Executable")</th><th style='width:6%'>$(L "Avvii" "Runs")</th><th style='width:16%'>$(L "Ultimo uso" "Last used")</th></tr></thead>
          <tbody>${UA_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>RunMRU — Comandi digitati in Esegui (Win+R)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Comando" "Command")</th></tr></thead>
          <tbody>${RUN_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>TypedPaths — Percorsi digitati in Explorer</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:8%'>Slot</th><th>$(L "Percorso" "Path")</th></tr></thead>
          <tbody>${TP_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>WordWheelQuery — Ricerche in Start Menu</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Termine cercato" "Search term")</th></tr></thead>
          <tbody>${WW_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 19 — ShellBags (navigazione cartelle, anche cancellate)
# ================================================================
module_shellbags() {
    section_header "$(L "ShellBags — Navigazione Cartelle" "ShellBags — Folder Navigation")" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    declare -a ALL_BAGS=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        # UsrClass.dat contiene le ShellBags principali (Win7+)
        local USRCLASS; USRCLASS=$(get_user_hive \
            "$(ci_find_dir "$USER_DIR" "AppData/Local/Microsoft/Windows")" \
            "UsrClass.dat")
        # Fallback: anche in NTUSER.DAT ci sono bag per desktop/drive locali
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")

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
            dim_msg "$USERNAME — $(L "nessuna ShellBag trovata" "no ShellBag found")"
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
    info "ShellBags totali: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna ShellBag trovata." "No ShellBags found.")"; return 0; }
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
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono ${SUSP_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
          <td class='mono dim' style='white-space:nowrap;font-size:.68rem'>$(html_esc "$L")</td>
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
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Percorso ricostruito" "Reconstructed path")</th><th style='width:12%'>Hive</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 20 — SAM: Hash password account locali
# ================================================================
module_sam() {
    section_header "$(L "SAM — Hash Account Locali" "SAM — Local Account Hashes")" "$RED"
    check_win_root || return 1

    local SAM_HIVE;  SAM_HIVE=$(get_hive "SAM")
    local SYS_HIVE;  SYS_HIVE=$(get_hive "SYSTEM")

    if [[ -z "$SAM_HIVE" || ! -f "$SAM_HIVE" ]]; then
        err "Hive SAM $(L "non trovato" "not found")"
        return 1
    fi
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found") (necessario per la SysKey)"
        return 1
    fi

    # Verifica impacket
    if ! "$PY3" -c "from impacket.examples.secretsdump import LocalOperations" 2>/dev/null; then
        warn "impacket $(L "non trovato" "not found")."
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install impacket"
        info "$(L "Estrazione parziale tramite regipy (senza hash decifrati)..." "Partial extraction via regipy (without decrypted hashes)...")"

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
        warn "$(L "Per gli hash installa impacket e rilancia il modulo." "For hashes install impacket and relaunch the module.")"
        return 0
    fi

    info "$(L "Estrazione hash con impacket secretsdump..." "Extracting hashes with impacket secretsdump...")"
    echo ""

    local TMP_DIR; TMP_DIR=$(mktemp -d /tmp/dfir_sam_XXXXXX)
    register_tmp "$TMP_DIR"
    cp "$SAM_HIVE" "$TMP_DIR/SAM" 2>/dev/null || { err "Copia hive fallita"; return 1; }
    cp "$SYS_HIVE" "$TMP_DIR/SYSTEM" 2>/dev/null || { err "Copia hive fallita"; return 1; }
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
    info "$(L "Account trovati:" "Accounts found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun account estratto." "No accounts extracted.")"; return 0; }
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
          <td class='mono'>$(html_esc "$UNAME")</td>
          <td class='mono mid' style='white-space:nowrap'>${RID}</td>
          <td class='mono dim' style='font-size:.65rem;word-break:break-all'>$(html_esc "$LM")</td>
          <td class='${NT_CSS}' style='font-size:.72rem;word-break:break-all'>$(html_esc "$NT")</td>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
            info "$(L "Estrazione \$MFT tramite ntfscat dal device" "Extracting \$MFT via ntfscat from device") ${DEV}..."
            TMP_MFT=$(mktemp /tmp/dfir_MFT_XXXXXX)
            if ntfscat -f "$DEV" '$MFT' > "$TMP_MFT" 2>/dev/null && [[ -s "$TMP_MFT" ]]; then
                MFT_FILE="$TMP_MFT"
                ok "$(L "\$MFT estratto con ntfscat" "\$MFT extracted with ntfscat") ($(du -h "$MFT_FILE" | cut -f1))"
                register_tmp "$TMP_MFT"
            else
                rm -f "$TMP_MFT"
                TMP_MFT=""
            fi
        fi

        if [[ -z "$MFT_FILE" ]]; then
            warn "$(L "\$MFT non accessibile — il volume è montato con ntfs3 (driver kernel)" "\$MFT not accessible — volume mounted with ntfs3 (kernel driver)")"
            echo ""
            info "$(L "Il driver ntfs3 non espone i file di sistema (\$MFT, \$LogFile, ecc.)." "The ntfs3 driver does not expose system files (\$MFT, \$LogFile, etc.).")"
            info "$(L "Per accedere all'\$MFT hai due opzioni:" "To access \$MFT you have two options:")"
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

    info "$(L "MFT trovato:" "MFT found:") $MFT_FILE  ($(du -h "$MFT_FILE" 2>/dev/null | cut -f1 || echo '?'))"

    # Verifica python-mft
    local HAS_PYMFT=false
    "$PY3" -c "import mft" 2>/dev/null && HAS_PYMFT=true

    if ! $HAS_PYMFT; then
        warn "python-mft non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install mft"
        info "$(L "Analisi rapida tramite strings (parziale)..." "Quick analysis via strings (partial)...")"
        echo ""
        # Estrai nomi file e timestamp con strings come fallback
        local STRINGS_OUT
        STRINGS_OUT=$(strings "$MFT_FILE" 2>/dev/null | \
            grep -iE '\.(exe|dll|bat|ps1|vbs|py|zip|rar|7z|iso)$' | \
            sort -u | head -200 || true)
        echo "$STRINGS_OUT" | while IFS= read -r L; do
            printf "  ${DIM}%s${RESET}\n" "$L"
        done | head -50
        warn "$(L "Per l'analisi completa installa python-mft." "For complete analysis install python-mft.")"
        return 0
    fi

    # Calcola il numero approssimativo di record (ogni record MFT è 1KB)
    local MFT_SIZE; MFT_SIZE=$(stat -c "%s" "$MFT_FILE" 2>/dev/null || echo 0)
    local MFT_RECORDS=$((MFT_SIZE / 1024))
    
    info "$(L "Parsing MFT con python-mft (~${MFT_RECORDS} record presenti)..." "Parsing MFT with python-mft (~${MFT_RECORDS} records present)...")"
    if [[ $MFT_SIZE -gt 500000000 ]]; then
        warn "$(L "⏳ L'MFT supera i 500MB! Il parsing completo può richiedere vari minuti." "⏳ MFT exceeds 500MB! Full parsing may take several minutes.")"
        info "   $(L "(Specialmente su volumi montati con BitLocker/FUSE). Attendi senza interrompere..." "(Especially on volumes mounted with BitLocker/FUSE). Wait without interrupting...")"
    fi

    local TMP_MFT_CSV; TMP_MFT_CSV=$(mktemp /tmp/dfir_mft_XXXXXX.csv)
    local TMP_MFT_ROWS; TMP_MFT_ROWS=$(mktemp /tmp/dfir_mft_rows_XXXXXX.html)
    register_tmp "$TMP_MFT_CSV"; register_tmp "$TMP_MFT_ROWS"

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

    info "File processati: ${BOLD}$TOTAL_LINES${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT${RESET}  |  Timestomping: ${RED}${BOLD}$STOMP_COUNT"

    # Stampa terminale: solo sospetti e timestomped
    echo ""
    awk -F'\t' '$8=="1" || $7=="TIMESTOMP" {print}' "$TMP_MFT_CSV" 2>/dev/null | head -30 | \
    while IFS=$'\t' read -r NAME EXT CREATED MODIFIED MFTMOD FNCREATED STOMPED SUSP; do
        local MARK=""
        [[ "$STOMPED" == "TIMESTOMP" ]] && MARK="${RED}[STOMP]${RESET} "
        printf "  %s${DIM}%-40s${RESET}  ${CYAN}%s${RESET}  %s\n" \
            "$MARK" "${NAME: -40}" "$CREATED" "$EXT"
    done

    [[ $TOTAL_LINES -eq 0 ]] && { warn "$(L "Nessun record MFT processato." "No MFT records processed.")"; return 0; }
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
            <th style='width:57%'>$(L "Nome file" "Filename")</th>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    ok "Dati completi salvati: ${BOLD}${REPORT_HTML%.html}.csv"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 22 — OpenSaveMRU / LastVisitedMRU (dialoghi file)
# ================================================================
module_opensave() {
    section_header "$(L "OpenSaveMRU / LastVisitedMRU — Dialoghi File" "OpenSaveMRU / LastVisitedMRU — File Dialogs")" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    declare -a OS_ENTRIES=()   # OpenSaveMRU
    declare -a LV_ENTRIES=()   # LastVisitedMRU
    local TOTAL_OS=0 TOTAL_LV=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
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
    [[ $GRAND -eq 0 ]] && { warn "$(L "Nessun dato trovato." "No data found.")"; return 0; }
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
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='white-space:nowrap;color:var(--accent4)'>.$(html_esc "$EXT")</td>
          <td class='mono ${SUSP_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
        </tr>"
    done

    local LV_ROWS=""
    for E in "${LV_ENTRIES[@]}"; do
        IFS='§' read -r U APP P <<< "$E"
        LV_ROWS+="<tr>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
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
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:8%'>$(L "Estensione" "Extension")</th><th>Path</th></tr></thead>
          <tbody>${OS_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>LastVisitedMRU — Ultima cartella visitata per applicazione</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:20%'>Applicazione</th><th>$(L "Ultima cartella" "Last folder")</th></tr></thead>
          <tbody>${LV_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 23 — USN Journal ($UsnJrnl:$J)
# ================================================================
module_usn() {
    section_header "$(L "USN Journal — Change Log NTFS" "USN Journal — NTFS Change Log")" "$CYAN"
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
            info "$(L "Estrazione \$UsnJrnl via ntfscat da" "Extracting \$UsnJrnl via ntfscat from") $DEV..."
            local TMP_USN; TMP_USN=$(mktemp /tmp/dfir_usn_XXXXXX)
            if ntfscat -f "$DEV" '$Extend/$UsnJrnl:$J' > "$TMP_USN" 2>/dev/null && [[ -s "$TMP_USN" ]]; then
                UJFILE="$TMP_USN"
                ok "$(L "\$UsnJrnl estratto" "\$UsnJrnl extracted") ($(du -h "$UJFILE" | cut -f1))"
                register_tmp "$TMP_USN"
            else
                rm -f "$TMP_USN"
                warn "$(L "\$UsnJrnl non accessibile. Monta con: mount -t ntfs-3g -o ro,show_sys_files" "\$UsnJrnl not accessible. Mount with: mount -t ntfs-3g -o ro,show_sys_files")"
                return 0
            fi
        else
            warn "$(L "\$UsnJrnl non accessibile — volume montato con ntfs3 (driver kernel)" "\$UsnJrnl not accessible — volume mounted with ntfs3 (kernel driver)")"
            [[ -n "$DEV" ]] && info "Prova: ntfscat -f \"$DEV\" '\$Extend/\$UsnJrnl:\$J' > /tmp/J"
            return 0
        fi
    fi

    info "USN Journal: $UJFILE  ($(du -h "$UJFILE" 2>/dev/null | cut -f1 || echo '?'))"

    local TMP_USN_CSV; TMP_USN_CSV=$(mktemp /tmp/dfir_usn_XXXXXX.csv)
    local TMP_USN_ROWS; TMP_USN_ROWS=$(mktemp /tmp/dfir_usn_rows_XXXXXX.html)
    register_tmp "$TMP_USN_CSV"; register_tmp "$TMP_USN_ROWS"

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

    info "$(L "Record USN:" "USN Records:") ${BOLD}$TOTAL_LINES${RESET}  |  Sospetti (ext+azione): ${RED}${BOLD}$SUSP_COUNT"
    echo ""
    awk -F'\t' '$5=="1"' "$TMP_USN_CSV" 2>/dev/null | head -25 | \
    while IFS=$'\t' read -r TS NAME EXT REASON SUSP; do
        printf "  ${RED}%-20s${RESET}  ${CYAN}%-30s${RESET}  %s\n" "$TS" "$NAME" "$REASON"
    done

    [[ $TOTAL_LINES -eq 0 ]] && { warn "$(L "Nessun record USN processato." "No USN records processed.")"; return 0; }
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
        <div class='stitle'>$(L "Change Journal NTFS — Timestamp · File · Ext · Azione" "NTFS Change Journal — Timestamp · File · Ext · Action")</div>
        <div class='card'><table>
          <thead><tr><th style='width:14%'>Timestamp</th><th>$(L "Nome file" "Filename")</th><th style='width:6%'>Ext</th><th>$(L "Motivo" "Reason")</th></tr></thead>
          <tbody>${TABLE_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun record</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        warn "$(L "ntds.dit non trovato — questo non è un Domain Controller (o la struttura NTDS/ è assente)" "ntds.dit not found — this is not a Domain Controller (or NTDS/ structure is absent)")"
        return 0
    fi
    local SYS_HIVE; SYS_HIVE=$(get_hive "SYSTEM")
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found") (necessario per il BootKey)"; return 1
    fi
    info "NTDS.dit: $NTDS_FILE"
    info "SYSTEM:   $SYS_HIVE"
    echo ""

    if ! "$PY3" -c "from impacket.examples.secretsdump import NTDSHashes" 2>/dev/null; then
        warn "$(L "impacket non trovato — installalo con:" "impacket not found — install it with:") ${PY3} -m pip install impacket"
        info "$(L "Fallback: estrazione account senza hash tramite strings..." "Fallback: account extraction without hashes via strings...")"
        local ACC_STRINGS
        ACC_STRINGS=$(strings "$NTDS_FILE" 2>/dev/null | grep -E '^[A-Za-z0-9_.-]{3,20}\$?$' | sort -u | head -100 || true)
        echo "$ACC_STRINGS" | while IFS= read -r L; do
            printf "  ${DIM}%s${RESET}\n" "$L"
        done
        warn "$(L "Installa impacket per l'estrazione degli hash NTLM." "Install impacket for NTLM hash extraction.")"
        return 0
    fi

    info "$(L "Estrazione hash con impacket NTDSHashes..." "Extracting hashes with impacket NTDSHashes...")"
    local TMP_DIR; TMP_DIR=$(mktemp -d /tmp/dfir_ntds_XXXXXX)
    register_tmp "$TMP_DIR"

    # Copia con timeout: su share SMB da DC live la copia può bloccarsi indefinitamente
    info "$(L "Copia ntds.dit in /tmp (timeout 180s — file di rete, attendere)..." "Copying ntds.dit to /tmp (timeout 180s — network file, please wait)...")"
    if ! portable_timeout 180 cp "$NTDS_FILE" "$TMP_DIR/ntds.dit" 2>/dev/null; then
        err "$(L "Timeout o errore nella copia di ntds.dit (share lenta o file bloccato da Windows)" "Timeout or error copying ntds.dit (slow share or file locked by Windows)")"
        info "$(L "Suggerimento: esegui il modulo su un'immagine montata offline oppure tramite VSS snapshot" "Hint: run the module on an offline mounted image or via VSS snapshot")"
        return 1
    fi
    info "$(L "Copia SYSTEM hive..." "Copying SYSTEM hive...")"
    if ! portable_timeout 60 cp "$SYS_HIVE" "$TMP_DIR/SYSTEM" 2>/dev/null; then
        err "$(L "Timeout o errore nella copia del SYSTEM hive" "Timeout or error copying SYSTEM hive")"
        return 1
    fi
    chmod 600 "$TMP_DIR/ntds.dit" "$TMP_DIR/SYSTEM" 2>/dev/null || true

    # Patch dirty state ESE: NTDSHashes usa ESENT_DB internamente, che rifiuta
    # database in stato DirtyShutdown (3) — condizione normale su DC live.
    # Offset 52 (0x34): campo dbstate nello standard ESE/JET Blue.
    info "$(L "Verifica/patch dirty state ESE (necessario per DC live)..." "Checking/patching ESE dirty state (required for live DC)...")"
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

    info "$(L "Dump hash NTLM (può richiedere diversi minuti per domini grandi)..." "Dumping NTLM hashes (may take several minutes for large domains)...")"
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
            err "$(L "ntds.dit è bloccato dal servizio Active Directory (DC live, Errno 35 / EAGAIN)." "ntds.dit is locked by Active Directory service (live DC, Errno 35 / EAGAIN).")"
            info "$(L "Soluzioni: 1) ntdsutil IFM sul DC  2) impacket-secretsdump remoto  3) immagine disco offline" "Solutions: 1) ntdsutil IFM on DC  2) remote impacket-secretsdump  3) offline disk image")"
        else
            err "$(L "Dump vuoto — ntds.dit potrebbe essere in stato dirty o corrotto." "Empty dump — ntds.dit may be dirty or corrupted.")"
            info "$(L "Prova il modulo 38 (PAD Offline) che usa analisi ESE con recovery avanzato." "Try module 38 (PAD Offline) which uses ESE analysis with advanced recovery.")"
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
    info "$(L "Account AD estratti:" "AD accounts extracted:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun account estratto." "No accounts extracted.")"; return 0; }
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
          <td class='mono'>$(html_esc "$UNAME")</td>
          <td class='mono mid' style='white-space:nowrap'>${RID}</td>
          <td class='${NT_CSS}' style='font-size:.72rem;word-break:break-all'>$(html_esc "$NT")</td>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 25 — Hibernation / Pagefile (artefatti memoria)
# ================================================================
module_hiberfil() {
    section_header "$(L "Hibernation / Pagefile — Artefatti Memoria" "Hibernation / Pagefile — Memory Artifacts")" "$BLUE"
    check_win_root || return 1


    local HFILE PFILE
    HFILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "hiberfil.sys" -type f 2>/dev/null | head -1)
    PFILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "pagefile.sys" -type f 2>/dev/null | head -1)

    [[ -z "$HFILE" && -z "$PFILE" ]] && {
        warn "$(L "hiberfil.sys e pagefile.sys non trovati nella root del volume" "hiberfil.sys and pagefile.sys not found in volume root")"
        info "$(L "Il volume potrebbe essere montato senza visibilità sui file di sistema" "The volume may be mounted without visibility on system files")"
        return 0
    }

    declare -a FILE_ENTRIES=()
    local TOTAL_STRINGS=0

    for MFILE in "$HFILE" "$PFILE"; do
        [[ -z "$MFILE" || ! -f "$MFILE" ]] && continue
        local MNAME; MNAME=$(basename "$MFILE")
        local MSIZE; MSIZE=$(du -h "$MFILE" 2>/dev/null | cut -f1 || echo "?")
        ok "$MNAME $(L "trovato:" "found:") $MFILE  ($MSIZE)"

        # Verifica magic bytes per hiberfil.sys
        local MAGIC=""
        if [[ "$MNAME" =~ hiberfil ]]; then
            MAGIC=$(dd if="$MFILE" bs=4 count=1 2>/dev/null | od -An -tx1 | tr -d ' \n' | cut -c1-8 || echo "")
            case "${MAGIC^^}" in
                "5041474D") info "  $(L "Formato: HIBR (hibernated)" "Format: HIBR (hibernated)")" ;;
                "52535452") info "  $(L "Formato: RSTR (restored)" "Format: RSTR (restored)")" ;;
                *) info "  $(L "Magic: 0x${MAGIC} (sconosciuto)" "Magic: 0x${MAGIC} (unknown)")" ;;
            esac
        fi

        info "  $(L "Estrazione stringhe significative..." "Extracting significant strings...")"
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
    info "$(L "Stringhe estratte totali:" "Total extracted strings:") ${BOLD}$TOTAL_STRINGS"
    info "${BOLD}$(L "Per analisi completa della memoria usa Volatility3:" "For complete memory analysis use Volatility3:")${RESET}"
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
            STR_ROWS+="<tr><td class='mono' style='white-space:nowrap;color:var(--accent4)'>$(html_esc "$LBL")</td>
              <td><span class='${CSS}' style='font-size:.72rem'>$(html_esc "${VAL:0:200}")</span></td></tr>"
        done <<< "$DECODED"
        CARDS_HTML+="<div class='card'>
          <div class='card-header'><div class='uicon'>MEM</div>
            <div><div class='uname'>$(html_esc "$MNAME")</div><div class='upath'>$(html_esc "$MFILE")</div></div>
            <div class='badge'>${MSIZE}</div>
          </div>
          <table><thead><tr><th style='width:14%'>$(L "Tipo" "Type")</th><th>$(L "Valore estratto" "Extracted value")</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
    info "$(L "File WER trovati:" "WER files found:") ${BOLD}$TOTAL_FILES"

    if [[ $TOTAL_FILES -eq 0 ]]; then
        ok "$(L "Nessun file WER trovato." "No WER files found.")"
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
    info "$(L "Processati" "Processed") ${BOLD}$TOTAL_FILES${RESET} report di errore."

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
              <div class='uname'>$(html_esc "$FDIR_NAME")</div>
              <div class='upath'>$(L "Sorgente:" "Source:") $(html_esc "$OWNER") &nbsp;·&nbsp; $(L "Data:" "Date:") ${FMTIME}</div>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
    info "$(L "Blob DPAPI trovati:" "DPAPI blobs found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun blob DPAPI trovato." "No DPAPI blobs found.")"; return 0; }
    warn "$(L "I blob sono cifrati con DPAPI — richiedono la master key utente per la decifratura" "Blobs are DPAPI-encrypted — require the user master key for decryption")"
    info "$(L "Tool utili: dpapick3, mimikatz (sekurlsa::dpapi), impacket dpapi" "Useful tools: dpapick3, mimikatz (sekurlsa::dpapi), impacket dpapi")"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "credential_manager")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente
    mapfile -t ALL_CREDS < <(printf '%s\n' "${ALL_CREDS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+7*[""])[6],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_CREDS[@]}"; do
        IFS='§' read -r UNAME FNAME CPATH REL MKGUID FSIZE FMTIME <<< "$E"
        ROWS+="<tr>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
          <td class='mono' style='font-size:.72rem;word-break:break-all'>$(html_esc "$FNAME")</td>
          <td class='mono mid' style='font-size:.68rem;word-break:break-all'>$(html_esc "$REL")</td>
          <td class='mono dim' style='font-size:.65rem;word-break:break-all'>$([[ -n "$MKGUID" ]] && html_esc "$MKGUID" || echo "-")</td>
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
        <div class='stitle'>$(L "Blob DPAPI — Utente · File · Path · MasterKey GUID · Dim · Timestamp" "DPAPI Blobs — User · File · Path · MasterKey GUID · Size · Timestamp")</div>
        <div class='card'><table>
          <thead><tr><th>$(L "Utente" "User")</th><th>File</th><th>Store</th><th>MK GUID</th><th>Dim.</th><th>$(L "Modificato" "Modified")</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 28 — WLAN & VPN Profiles
# ================================================================
module_wlan() {
    section_header "$(L "WLAN & VPN — Profili di Rete" "WLAN & VPN — Network Profiles")" "$CYAN"
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
        warn "$(L "WLAN Profiles non trovati (ProgramData/Microsoft/Wlansvc/Profiles)" "WLAN Profiles not found (ProgramData/Microsoft/Wlansvc/Profiles)")"
    fi

    # VPN / Network profiles dal registro SOFTWARE
    echo ""
    info "$(L "Profili di rete (NetworkList) dal registro..." "Network profiles (NetworkList) from registry...")"
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
    info "$(L "Profili di rete totali:" "Total network profiles:") ${BOLD}$TOTAL_WLAN"
    [[ $TOTAL_WLAN -eq 0 ]] && { warn "$(L "Nessun profilo trovato." "No profiles found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wlan")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_WLAN[@]}"; do
        IFS='§' read -r TYPE NAME F2 F3 F4 F5 <<< "$E"
        local CSS=""; [[ "$TYPE" == "WLAN" && -n "$F4" ]] && CSS="style='background:rgba(255,123,114,.07)'"
        local KEYHTML=""; [[ "$TYPE" == "WLAN" && -n "$F4" ]] && KEYHTML="<span class='bad' style='font-size:.68rem'>KEY: $(html_esc "${F4:0:40}")</span>"
        ROWS+="<tr ${CSS}>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$TYPE")</td>
          <td class='mono'>$(html_esc "$NAME") ${KEYHTML}</td>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$F2")</td>
          <td class='mono mid' style='white-space:nowrap'>$(html_esc "$F3")</td>
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
          <thead><tr><th style='width:8%'>$(L "Tipo" "Type")</th><th>Nome / SSID</th><th style='width:12%'>Auth/Cat</th><th style='width:15%'>Gateway/Enc</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 29 — AppX / UWP Packages
# ================================================================
module_appx() {
    section_header "$(L "AppX / UWP — Pacchetti Store Installati" "AppX / UWP — Installed Store Packages")" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"; return 1
    fi


    declare -a ALL_PKGS=()
    local TOTAL=0 SUSP_COUNT=0

    # SOFTWARE hive (sistema)
    local SW_HIVE; SW_HIVE=$(get_hive "SOFTWARE")
    if [[ -n "$SW_HIVE" ]]; then
        info "$(L "Parsing AppX packages da SOFTWARE hive..." "Parsing AppX packages from SOFTWARE hive...")"
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
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
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
    info "Pacchetti AppX: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun pacchetto AppX trovato." "No AppX packages found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "appx")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_PKGS[@]}"; do
        IFS='§' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER IS_SUSP <<< "$E"
        local RS=""; [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        local NC="mono"; [[ "$IS_SUSP" == "true" ]] && NC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap;font-size:.68rem'>$(html_esc "$SCOPE")</td>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 30 — Browser Downloads & Login Data
# ================================================================
module_browser_extra() {
    section_header "Browser Downloads & Login Data" "$CYAN"
    check_win_root || return 1

    if ! "$PY3" -c "import sqlite3" 2>/dev/null; then
        err "$(L "sqlite3 non disponibile" "sqlite3 not available")"; return 1
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
    register_tmp "$TMP_DL_ROWS"; register_tmp "$TMP_LG_ROWS"
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
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
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
        ORDER BY h.visit_date DESC LIMIT 100000
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
                        ok "$USERNAME / $BLABEL — $(L "download trovati" "downloads found")"
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
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("""
        SELECT tab_url, target_path, start_time, total_bytes
        FROM downloads ORDER BY start_time DESC LIMIT 100000
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
                    ok "$USERNAME / $BLABEL — $(L "download trovati" "downloads found")"
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
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("SELECT origin_url, username_value, length(password_value), date_password_changed FROM logins ORDER BY date_password_changed DESC LIMIT 100000")
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
                ok "$USERNAME / $BLABEL — $(L "credenziali salvate trovate" "saved credentials found")"
            fi
        done
    done < <(get_user_homes)

    separator
    info "$(L "Download:" "Downloads:") ${BOLD}$TOTAL_DL${RESET}  |  Credenziali salvate: ${BOLD}$TOTAL_LG"
    [[ $((TOTAL_DL + TOTAL_LG)) -eq 0 ]] && { warn "$(L "Nessun dato trovato." "No data found.")"; return 0; }
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
        <div class='stitle'>$(L "Download — Timestamp · Utente/Browser · URL · Destinazione" "Downloads — Timestamp · User/Browser · URL · Destination")</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Timestamp</th><th style='width:14%'>$(L "Utente/Browser" "User/Browser")</th><th>URL</th><th style='width:20%'>$(L "Destinazione" "Destination")</th></tr></thead>
          <tbody>${DL_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun download trovato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>$(L "Login Data — Credenziali salvate (password cifrate DPAPI)" "Login Data — Saved credentials (DPAPI-encrypted passwords)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Timestamp</th><th style='width:14%'>$(L "Utente/Browser" "User/Browser")</th><th>URL</th><th style='width:15%'>Username</th><th style='width:10%'>Password</th></tr></thead>
          <tbody>${LG_ROWS:-<tr><td colspan='5' class='dim' style='padding:1rem'>Nessuna credenziale trovata</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}


# ================================================================
#  MODULO 31 — Clipboard History (Win10+)
# ================================================================
module_clipboard() {
    section_header "$(L "Clipboard History — Cronologia Appunti" "Clipboard History")" "$YELLOW"
    check_win_root || return 1


    declare -a ALL_CLIPS=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local CB_DIR
        CB_DIR=$(ci_find_dir "$USER_DIR" "AppData/Local/Microsoft/Windows/Clipboard")
        [[ -z "$CB_DIR" || ! -d "$CB_DIR" ]] && continue
        info "$(L "Clipboard dir trovata:" "Clipboard dir found:") $CB_DIR"

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
    info "Voci clipboard: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna voce clipboard trovata (feature Win10 1809+, va abilitata)." "No clipboard entries found (Win10 1809+ feature, must be enabled).")"; return 0; }
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
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
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
        <div class='stitle'>$(L "Clipboard History — Utente · Timestamp · Contenuto (troncato a 300 char)" "Clipboard History — User · Timestamp · Content (truncated to 300 chars)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:14%'>Timestamp</th><th>$(L "Contenuto" "Content")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 32 — Office MRU & Recent Docs
# ================================================================
module_office_mru() {
    section_header "$(L "Office MRU — File Recenti Office" "Office MRU — Recent Office Files")" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"; return 1
    fi


    declare -a ALL_MRU=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
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
    info "File Office recenti: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun file Office recente trovato." "No recent Office files found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "office_mru")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_MRU[@]}"; do
        IFS='§' read -r U APP VER FPATH IS_SUSP <<< "$E"
        local RS="" NC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && NC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$APP")</td>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$VER")</td>
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
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:12%'>App</th><th style='width:7%'>$(L "Versione" "Version")</th><th>Path file</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 33 — Defender Quarantine
# ================================================================
module_defender_quarantine() {
    section_header "$(L "Defender Quarantine — File in Quarantena" "Defender Quarantine — Quarantined Files")" "$RED"
    check_win_root || return 1


    local QENTRIES_DIR
    QENTRIES_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Windows Defender/Quarantine/Entries")
    local QRESOURCE_DIR
    QRESOURCE_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Windows Defender/Quarantine/ResourceData")

    if [[ -z "$QENTRIES_DIR" || ! -d "$QENTRIES_DIR" ]]; then
        warn "$(L "Directory quarantena Defender non trovata" "Defender quarantine directory not found")"
        info "$(L "Percorso atteso: ProgramData/Microsoft/Windows Defender/Quarantine/" "Expected path: ProgramData/Microsoft/Windows Defender/Quarantine/")"
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
    info "$(L "Voci in quarantena trovate:" "Quarantine entries found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna voce trovata nella quarantena." "No entries found in quarantine.")"; return 0; }

    if [[ -n "$QRESOURCE_DIR" ]]; then
        info "${BOLD}$(L "Payload cifrati disponibili in:" "Encrypted payloads available in:")${RESET} $QRESOURCE_DIR"
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
        <div class='stitle'>$(L "File in quarantena — recuperabili offline da ResourceData/ (cifrati XOR)" "Quarantined files — recoverable offline from ResourceData/ (XOR-encrypted)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:14%'>$(L "Rilevato" "Detected")</th><th style='width:22%'>$(L "Minaccia" "Threat")</th><th>$(L "Path originale" "Original path")</th><th style='width:18%'>SHA256</th><th style='width:6%'>Dim.</th></tr></thead>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
        warn "$(L "Directory EVTX non trovata" "EVTX directory not found")"
        return 1
    fi

    local PS_EVTX
    PS_EVTX=$(find "$EVTX_DIR" -maxdepth 1 -iname "Microsoft-Windows-PowerShell%4Operational.evtx" -type f 2>/dev/null | head -1)
    if [[ -z "$PS_EVTX" || ! -f "$PS_EVTX" ]]; then
        warn "$(L "Log PowerShell Operational non trovato" "PowerShell Operational log not found")"
        info "$(L "Atteso: Windows/System32/winevt/Logs/Microsoft-Windows-PowerShell%4Operational.evtx" "Expected: Windows/System32/winevt/Logs/Microsoft-Windows-PowerShell%4Operational.evtx")"
        return 0
    fi

    info "Parsing: $PS_EVTX"

    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        warn "$(L "python-evtx non trovato." "python-evtx not found.") Installalo con: ${PY3} -m pip install python-evtx"
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
        [[ "$TYPE" == "ERROR" ]] && { err "$(L "Errore parser:" "Parser error:") $SCRIPT"; break; }
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
    info "Script block unici: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun EID 4104 trovato (Script Block Logging potrebbe non essere abilitato)." "No EID 4104 found (Script Block Logging may not be enabled).")"; return 0; }
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
        local PATH_ATTR; PATH_ATTR=$(html_attr "$FPATH")
        ROWS+="<tr ${RS}>
          <td class='mono ok' style='white-space:nowrap;font-size:.68rem'>${TS}</td>
          <td class='mono dim' style='white-space:nowrap;font-size:.65rem;max-width:220px;overflow:hidden;text-overflow:ellipsis' title='${PATH_ATTR}'>${PATH_ESC}</td>
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
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
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
    info "JumpList entries: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna JumpList trovata." "No JumpLists found.")"; return 0; }
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
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$JT")</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.68rem'>${FMTIME}</td>
          <td class='mono dim' style='font-size:.68rem'>$(html_esc "$FNAME")</td>
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
        <div class='stitle'>$(L "JumpLists — Utente · Tipo · Timestamp · File · Percorsi estratti (strings)" "JumpLists — User · Type · Timestamp · File · Extracted paths (strings)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:10%'>$(L "Utente" "User")</th><th style='width:7%'>$(L "Tipo" "Type")</th><th style='width:14%'>Modificato</th><th style='width:14%'>File JL</th><th>Percorsi estratti</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 36 — Network Artifacts (DNS Cache, Interfacce, NetworkList)
# ================================================================
module_network_artifacts() {
    section_header "$(L "Network Artifacts — DNS · Interfacce · NetworkList" "Network Artifacts — DNS · Interfaces · NetworkList")" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"; return 1
    fi

    local SYS_HIVE; SYS_HIVE=$(get_hive "SOFTWARE")
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "$(L "Hive SOFTWARE non trovato" "SOFTWARE hive not found")"; return 1
    fi

    info "$(L "Parsing NetworkList da hive SOFTWARE:" "Parsing NetworkList from SOFTWARE hive:") $SYS_HIVE"

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
        info "$(L "Parsing interfacce TCP/IP da SYSTEM hive:" "Parsing TCP/IP interfaces from SYSTEM hive:") $SYS_HIVE2"
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
    info "$(L "Profili rete:" "Network profiles:") ${BOLD}$TOTAL_P${RESET}  |  Firme: ${BOLD}$TOTAL_S${RESET}  |  Interfacce TCP/IP: ${BOLD}$TOTAL_I"
    [[ $((TOTAL_P + TOTAL_S + TOTAL_I)) -eq 0 ]] && { warn "$(L "Nessun dato di rete trovato." "No network data found.")"; return 0; }
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
        ROWS_P+="<tr><td class='mono'>$(html_esc "$NAME")</td><td class='mono dim'>$(html_esc "$DESC")</td>
          <td class='mono ${CAT_COLOR}'>$(html_esc "$CAT")</td>
          <td class='mono dim' style='font-size:.68rem'>${DFIRST}</td>
          <td class='mono ok'  style='font-size:.68rem'>${DLAST}</td></tr>"
    done
    for E in "${SIGNATURES[@]}"; do
        IFS='§' read -r NAME DNS_S MAC SSID <<< "$E"
        ROWS_S+="<tr><td class='mono'>$(html_esc "$NAME")</td><td class='mono dim'>$(html_esc "$DNS_S")</td>
          <td class='mono warn'>$(html_esc "$MAC")</td><td class='mono'>$(html_esc "$SSID")</td></tr>"
    done
    for E in "${IFACES[@]}"; do
        IFS='§' read -r CS IFACE IP GW DNS_I DHCP <<< "$E"
        ROWS_I+="<tr><td class='mono dim'>$(html_esc "$CS")</td><td class='mono dim' style='font-size:.65rem'>$(html_esc "$IFACE")</td>
          <td class='mono ok'>$(html_esc "$IP")</td><td class='mono'>$(html_esc "$GW")</td>
          <td class='mono dim'>$(html_esc "$DNS_I")</td><td class='mono'>$(html_esc "$DHCP")</td></tr>"
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
              <thead><tr><th>$(L "Nome" "Name")</th><th>$(L "Descrizione" "Description")</th><th>$(L "Categoria" "Category")</th><th>$(L "Prima connessione" "First connection")</th><th>$(L "Ultima connessione" "Last connection")</th></tr></thead>
              <tbody>${ROWS_P}</tbody></table></div>"
        fi
        if [[ -n "$ROWS_S" ]]; then
            echo "<div class='stitle' style='margin-top:2rem'>Firme di rete (Signatures\\Unmanaged)</div>
            <div class='card'><table>
              <thead><tr><th>$(L "Descrizione" "Description")</th><th>DNS Suffix</th><th>MAC Gateway</th><th>SSID</th></tr></thead>
              <tbody>${ROWS_S}</tbody></table></div>"
        fi
        if [[ -n "$ROWS_I" ]]; then
            echo "<div class='stitle' style='margin-top:2rem'>Interfacce TCP/IP (SYSTEM\\Tcpip\\Parameters)</div>
            <div class='card'><table>
              <thead><tr><th>ControlSet</th><th>$(L "GUID interfaccia" "Interface GUID")</th><th>IP</th><th>Gateway</th><th>DNS</th><th>DHCP</th></tr></thead>
              <tbody>${ROWS_I}</tbody></table></div>"
        fi
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 37 — Master Timeline (aggregazione cross-moduli)
# ================================================================
module_master_timeline() {
    section_header "$(L "Master Timeline — Aggregazione Cross-Moduli" "Master Timeline — Cross-Module Aggregation")" "$YELLOW"

    if [[ ${#GENERATED_REPORTS[@]} -eq 0 ]]; then
        warn "$(L "Nessun report generato in questa sessione." "No reports generated in this session.")"
        info "$(L "Esegui prima i moduli che ti interessano, poi richiama la Master Timeline." "Run the desired modules first, then call the Master Timeline.")"
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
            # Descrizione: celle che non contengono solo il timestamp né un numero puro
            desc_parts = [c[:80] for c in cells if c and not TS_PAT.fullmatch(c.strip()) and not re.fullmatch(r'\d+', c.strip())]
            desc = ' | '.join(desc_parts)[:200]
            desc = ' '.join(desc.split())   # normalizza spazi
            key = (ts, mod, desc[:40])
            if key in seen:
                continue
            seen.add(key)
            sys.stdout.buffer.write(f"{ts}\t{mod}\t{desc}\n".encode('utf-8', errors='replace'))
        # Report AI Chat (modulo ai_chat): eventi nei blocchi <div class='msg'>
        for mm in re.finditer(
                r"<span class='role'>(.*?)</span>.*?<span class='mts'>(.*?)</span>"
                r".*?<pre class='msg-t'>(.*?)</pre>", content, re.DOTALL):
            role = strip_tags(mm.group(1))
            ts_m = TS_PAT.search(html_mod.unescape(mm.group(2)))
            if not ts_m:
                continue
            ts = ts_m.group(0).replace('T', ' ')[:19]
            text = ' '.join(strip_tags(mm.group(3)).split())[:200]
            desc = (role + ': ' + text) if role and role != '—' else text
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

    # Ordina per timestamp decrescente, scrive su file temp per evitare
    # problemi con here-string su variabili di grandi dimensioni
    local _TL_TMP; _TL_TMP=$(mktemp)
    printf '%s\n' "$TL_RAW" | sort -t$'\t' -k1 -r | head -100000 > "$_TL_TMP"
    local SHOW_COUNT; SHOW_COUNT=$(wc -l < "$_TL_TMP")

    separator
    info "$(L "Eventi con timestamp estratti:" "Events with extracted timestamps:") ${BOLD}$TOTAL_EVENTS${RESET} ($(L "mostrati" "showing") $SHOW_COUNT)"
    [[ $TOTAL_EVENTS -eq 0 ]] && { rm -f "$_TL_TMP"; warn "$(L "Nessun timestamp trovato nei report. Verifica che i moduli abbiano generato dati." "No timestamps found in reports. Verify that modules have generated data.")"; return 0; }
    ask_yn "$(L "Generare report HTML Master Timeline?" "Generate Master Timeline HTML report?")" || { rm -f "$_TL_TMP"; return 0; }

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "master_timeline")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local NR="${#GENERATED_REPORTS[@]}"

    {
        html_header "Master Timeline"
        html_page_header "TL" "Master Timeline — <span>$(L "Aggregazione Cross-Moduli" "Cross-Module Aggregation")</span>" \
            "$(L "Aggregazione di" "Aggregation of") ${NR} $(L "report" "reports")" "$SCAN" "$WIN_ROOT"
        cat << STATSEOF
<div class='statsbar'>
  <div class='stat'><div class='label'>$(L "Report aggregati" "Aggregated reports")</div><div class='value'>${NR}</div></div>
  <div class='stat info'><div class='label'>$(L "Eventi con timestamp" "Events with timestamp")</div><div class='value'>${TOTAL_EVENTS}</div></div>
  <div class='stat ok'><div class='label'>$(L "Mostrati nel report" "Shown in report")</div><div class='value'>${SHOW_COUNT}</div></div>
</div>
<main>
<div class='stitle'>$(L "Timeline cronologica (decrescente) — ordinata per data" "Chronological timeline (descending) — sorted by date")</div>
<div class='card'>
<table>
  <thead><tr>
    <th style='width:16%'>Timestamp</th>
    <th style='width:14%'>$(L "Modulo" "Module")</th>
    <th>$(L "Descrizione" "Description")</th>
  </tr></thead>
  <tbody>
STATSEOF

        # Genera righe HTML direttamente da file temp — escaping inline senza subshell
        while IFS=$'\t' read -r _TS _MOD _DESC; do
            [[ -z "$_TS" ]] && continue
            # Escaping inline: nessun subshell, nessun fork
            local _TE="${_TS//&/&amp;}"; _TE="${_TE//</&lt;}"; _TE="${_TE//>/&gt;}"
            local _ME="${_MOD//&/&amp;}"; _ME="${_ME//</&lt;}"; _ME="${_ME//>/&gt;}"
            local _D="${_DESC:0:150}"
            local _DE="${_D//&/&amp;}"; _DE="${_DE//</&lt;}"; _DE="${_DE//>/&gt;}"
            local _BC="var(--accent)"
            case "${_MOD,,}" in
                *evtx*|*service*|*ifeo*|*wmi*|*sam*|*ntds*) _BC="var(--accent2)" ;;
                *ps_hist*|*script*|*prefetch*|*amcache*)     _BC="var(--accent4)" ;;
                *browser*|*lnk*|*shell*|*userassist*)        _BC="var(--accent3)" ;;
                *ai_chat*)                                    _BC="#a371f7" ;;
            esac
            printf '<tr><td class="mono ok" style="white-space:nowrap;font-size:.72rem">%s</td>' "$_TE"
            printf '<td><span style="background:rgba(88,166,255,.1);border:1px solid rgba(88,166,255,.2);'
            printf 'color:%s;font-family:var(--mono);font-size:.65rem;padding:.15rem .5rem;border-radius:2px">%s</span></td>' "$_BC" "$_ME"
            printf '<td class="mono" style="font-size:.72rem;color:var(--text)">%s</td></tr>\n' "$_DE"
        done < "$_TL_TMP"

        cat << 'TABLEEOF'
  </tbody>
</table>
</div>
</main>
TABLEEOF

        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    rm -f "$_TL_TMP"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
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
        warn "$(L "ntds.dit non trovato sotto" "ntds.dit not found under") $WIN_ROOT"
        info "$(L "Percorso atteso: Windows/NTDS/ntds.dit" "Expected path: Windows/NTDS/ntds.dit")"
        info "$(L "Questo modulo richiede un Domain Controller (non una workstation)" "This module requires a Domain Controller (not a workstation)")"
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
        warn "$(L "SYSTEM hive non trovato — attributi cifrati (hash) non disponibili; tutti gli altri OK" "SYSTEM hive not found — encrypted attributes (hashes) unavailable; all others OK")"
        SYSTEM_HIVE=""
    fi

    # ── Avviso dirty database ─────────────────────────────────────
    warn "$(L "Se ntds.dit proviene da un DC attivo, alcune transazioni potrebbero essere nei log (.jrs)." "If ntds.dit comes from a live DC, some transactions may be in logs (.jrs).")"
    info "Per massima integrità usa una VSS snapshot o esegui 'esentutl /r edb /l <logdir>' prima."

    # ── Verifica dipendenza impacket ──────────────────────────────
    if ! "$PY3" -c "from impacket.ese import ESENT_DB" 2>/dev/null; then
        warn "impacket non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install impacket"
        return 1
    fi

    info "$(L "Avvio analisi NTDS.dit (può richiedere qualche minuto su database grandi)..." "Starting NTDS.dit analysis (may take a few minutes on large databases)...")"
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
    info "$(L "Tentativo copia ntds.dit in /tmp (timeout 180s)..." "Attempting to copy ntds.dit to /tmp (timeout 180s)...")"
    if portable_timeout 180 cp "$NTDS_PATH" "$TMP_NTDS" 2>/dev/null; then
        chmod 600 "$TMP_NTDS" 2>/dev/null || true
        NTDS_TO_USE="$TMP_NTDS"
        info "$(L "Copia completata — uso copia locale." "Copy completed — using local copy.")"
    else
        warn "$(L "Copia non riuscita (file bloccato da Windows o rete lenta) — uso path originale." "Copy failed (file locked by Windows or slow network) — using original path.")"
        info "$(L "Il dirty state verrà corretto in memoria (patch in-memory, senza write access)." "Dirty state will be corrected in memory (in-memory patch, no write access).")"
    fi

    TMP_NTDS_SIZE=$(wc -c < "$TMP_NTDS" 2>/dev/null | tr -d ' ' || echo "0")
    ORIG_NTDS_SIZE=$(wc -c < "$NTDS_TO_USE" 2>/dev/null | tr -d ' ' || echo "0")
    info "$(L "Dimensione NTDS.dit prima del parsing:" "NTDS.dit size before parsing:") ${TMP_NTDS_SIZE} bytes (originale: ${ORIG_NTDS_SIZE})"
    
    # ── Script Python: analisi completa + generazione HTML ───────
    local PYERR_FILE; PYERR_FILE="${TMP_PAD_DIR}/pyerr.log"
    "$PY3" - "$NTDS_TO_USE" "${SYSTEM_HIVE}" "$REPORT_HTML" "$SCAN" "$WIN_ROOT" "$NTDS_HASH" "$(dirname "$NTDS_PATH")" "$TMP_PAD_DIR" << 'PYEOF' 2>"$PYERR_FILE"
import sys, os, struct, datetime, traceback, html as html_mod

# Bilingual helper: reads LANG env var set by fiuto.sh
def L(it, en): return it if os.environ.get('LANG', 'en') == 'it' else en

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
    sys.stderr.write("INFO: no ESE log found — only dirty state patch will be applied\n")

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
    sys.stderr.write("FATAL: no ESE library available. Install impacket or libesedb-python.\n")
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
    sys.stderr.write("FATAL: unable to open NTDS.dit: no backend available.\n")
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
        return f"<tr><td colspan='5' class='dim' style='padding:1rem;text-align:center'>{L('Nessun ACE pericoloso rilevato (o nTSecurityDescriptor non trovato)', 'No dangerous ACE found (or nTSecurityDescriptor not present)')}</td></tr>"
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

# ── Frammenti HTML pre-generati (evitano f-string annidati con lo stesso quoting) ──
_krbtgt_risk_html = (
    f'<span class="tag tag-bad">{L("GOLDEN TICKET RISK: rotazione &gt;180gg", "GOLDEN TICKET RISK: rotation &gt;180d")}</span>'
    if (krbtgt_days and krbtgt_days > 180) else ''
)
def _empty_row(cols, it, en):
    return f"<tr><td colspan='{cols}' class='dim' style='padding:1rem;text-align:center'>{L(it, en)}</td></tr>"

_no_priv_row = _empty_row(6, 'Nessun utente privilegiato trovato', 'No privileged users found')
_no_gpo_row  = _empty_row(3, 'Nessuna GPO trovata', 'No GPOs found')
_no_comp_row = _empty_row(3, 'Nessun computer trovato', 'No computers found')

# ── Blocco HTML per catalogo corrotto (pre-generato per evitare f-string annidati) ──
_catalog_corrupt_html = ""
if _catalog_corrupt:
    _objects_count = len(objects)
    _catalog_corrupt_html = f"""
<div class="stitle" style="color:var(--accent2)">&#9888; NTDS.dit &mdash; {L("Catalogo ESE corrotto: analisi privilegiata non disponibile", "Corrupt ESE catalogue: privileged analysis unavailable")}</div>
<div class="card" style="border-color:rgba(255,123,114,.5);background:rgba(255,123,114,.06)">
<div style="padding:1.4rem 1.6rem;font-family:var(--mono);font-size:.82rem;line-height:1.9">
<div style="color:var(--accent2);font-weight:700;font-size:.95rem;margin-bottom:.9rem">Stato: DirtyShutdown &mdash; Pagine ESE corrotte nel catalogo di sistema</div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem">
  <div style="background:rgba(0,0,0,.25);border-radius:6px;padding:.9rem 1rem">
    <div style="color:var(--accent3);font-weight:700;margin-bottom:.5rem">&#10003; Dati estratti con successo</div>
    <div style="color:var(--text-mid)">&#x2022; Password KRBTGT (pwdLastSet)</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Livello funzionale dominio", "Domain functional level")}</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Oggetti GPO e percorsi", "GPO objects and paths")}</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Elenco computer (parziale)", "Computer list (partial)")}</div>
    <div style="color:var(--text-mid)">&#x2022; {_objects_count} oggetti datatable recuperati</div>
  </div>
  <div style="background:rgba(0,0,0,.25);border-radius:6px;padding:.9rem 1rem">
    <div style="color:var(--accent2);font-weight:700;margin-bottom:.5rem">&#10007; Dati non disponibili</div>
    <div style="color:var(--text-mid)">&#x2022; sAMAccountName (colonna non mappata)</div>
    <div style="color:var(--text-mid)">&#x2022; samAccountType / userAccountControl</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Membership gruppi privilegiati", "Privileged group membership")}</div>
    <div style="color:var(--text-mid)">&#x2022; Link-table group membership</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Hash delle password", "Password hashes")}</div>
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
  <div style="color:var(--text-dim);margin:.2rem 0 .2rem 1rem">{L("Sul DC sorgente o una VM Windows con ntds.dit + log ESE nella stessa cartella:", "On the source DC or a Windows VM with ntds.dit + ESE log files in the same folder:")}</div>
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
  {L("nessun log pu&ograve; ripristinare dati che non esistono pi&ugrave; sul disco.", "no log files can restore data that no longer exists on disk.")}</span>
</div>
</div></div>"""

# ── HTML output ────────────────────────────────────────────────────
html_out = f"""<!DOCTYPE html>
<html lang="{L("it", "en")}"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>PAD Offline — DFIR Report</title>
<style>{CSS}</style></head><body>
<header>
  <div class="hicon">AD</div>
  <div class="htxt">
    <h1>PAD Offline — <span>{L("Panoramica Active Directory", "Active Directory Overview")}</span></h1>
    <div class="sub">NTDS.dit offline forensic analysis</div>
  </div>
  <div class="hmeta">
    <div>{L("Scansione:", "Scan:")} <span class="val">{h(scan_date)}</span></div>
    <div>{L("Sorgente:", "Source:")} <span class="val">{h(source_root)}</span></div>
    <div style="font-size:.6rem">SHA256: <span class="val" style="word-break:break-all">{h(ntds_hash)}</span></div>
  </div>
</header>
<div class="statsbar">
  <div class="stat"><div class="label">{L("Utenti Privilegiati", "Privileged Users")}</div><div class="value">{total_priv}</div></div>
  <div class="stat"><div class="label">{L("Gruppi Monitorati", "Monitored Groups")}</div><div class="value info" style="color:var(--accent)">{len(PRIV_GROUPS)}</div></div>
  <div class="stat"><div class="label">{L("KRBTGT (giorni)", "KRBTGT (days)")}</div><div class="value {krbtgt_cls}">{krbtgt_days if krbtgt_days is not None else 'N/A'}</div></div>
  <div class="stat"><div class="label">{L("Password Stale (&gt;24h)", "Stale Passwords (&gt;24h)")}</div><div class="value">{total_stale}</div></div>
  <div class="stat"><div class="label">{L("GPO Modificate (30gg)", "Modified GPOs (30d)")}</div><div class="value {"" if gpo_recent==0 else ""}">{gpo_recent}/{total_gpo}</div></div>
  <div class="stat ok"><div class="label">{L("Computer nel Dominio", "Domain Computers")}</div><div class="value">{total_comp}</div></div>
</div>
<main>

<div class="stitle">{L("Configurazione Dominio", "Domain Configuration")}</div>
<div class="card"><table>
<thead><tr><th>{L("Parametro", "Parameter")}</th><th>{L("Valore", "Value")}</th><th>Note</th></tr></thead>
<tbody>
<tr><td class="mono dim">Domain Functional Level</td>
    <td class="mono {fl_badge_cls}">{h(domain_fl)}</td><td></td></tr>
<tr><td class="mono dim">KRBTGT pwdLastSet</td>
    <td class="mono {krbtgt_cls}">{h(krbtgt_pwd)}</td>
    <td>{krbtgt_badge}{_krbtgt_risk_html}</td></tr>
<tr><td class="mono dim">AD Recycle Bin</td>
    <td class="mono {"ok" if recycle_bin=="Enabled" else "bad"}">{h(recycle_bin)}</td>
    <td>{"" if recycle_bin=="Enabled" else "<span class='tag tag-warn'>Oggetti eliminati non recuperabili</span>"}</td></tr>
<tr><td class="mono dim">{L("Utenti Privilegiati", "Privileged Users")} Disabilitati</td>
    <td class="mono {"warn" if total_disabled>0 else "ok"}">{total_disabled}</td><td></td></tr>
</tbody></table></div>

{_catalog_corrupt_html}
<div class="stitle">{L("Utenti Privilegiati", "Privileged Users")} — {total_priv} account ({len(PRIV_GROUPS)} gruppi monitorati)</div>
<div class="card"><table>
<thead><tr><th>sAMAccountName</th><th>{L("Gruppi (Direct/Nested)", "Groups (Direct/Nested)")}</th><th>pwdLastSet</th><th>lastLogon</th><th>Flags</th><th>SID</th></tr></thead>
<tbody>{rows_priv if rows_priv else _no_priv_row}</tbody>
</table></div>

<div class="stitle">{L("ACL Domain Root — ACE con diritti pericolosi", "ACL Domain Root — Dangerous ACE rights")}</div>
<div class="card"><table>
<thead><tr><th>Trustee</th><th>SID</th><th>{L("Diritti", "Rights")}</th><th>Access Mask</th><th>Flags</th></tr></thead>
<tbody>{ace_rows(root_aces)}</tbody>
</table></div>

<div class="stitle">{L("ACL OU Domain Controllers — ACE con diritti pericolosi", "ACL OU Domain Controllers — Dangerous ACE rights")}</div>
<div class="card"><table>
<thead><tr><th>Trustee</th><th>SID</th><th>{L("Diritti", "Rights")}</th><th>Access Mask</th><th>Flags</th></tr></thead>
<tbody>{ace_rows(dc_ou_aces)}</tbody>
</table></div>

<div class="stitle">Group Policy Objects — {total_gpo} {L("GPO totali", "total GPOs")}</div>
<div class="card"><table>
<thead><tr><th>Display Name</th><th>{L("Ultima Modifica", "Last Modified")}</th><th>Path SYSVOL</th></tr></thead>
<tbody>{rows_gpo if rows_gpo else _no_gpo_row}</tbody>
</table></div>

<div class="stitle">{L("Computer nel Dominio", "Domain Computers")} — {total_comp} oggetti</div>
<div class="card"><table>
<thead><tr><th>Computer Name</th><th>Operating System</th><th>Distinguished Name</th></tr></thead>
<tbody>{rows_comp if rows_comp else _no_comp_row}</tbody>
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
            warn "$(L "NTDS.dit con catalogo ESE parzialmente corrotto — dati utente non disponibili." "NTDS.dit with partially corrupted ESE catalog — user data unavailable.")"
            info "$(L "Le colonne sAMAccountName/samAccountType sono su pagine di catalogo non leggibili." "Columns sAMAccountName/samAccountType are on unreadable catalog pages.")"
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
            info "$(L "Il report HTML contiene la guida completa al recupero." "The HTML report contains the complete recovery guide.")"
        fi

        info "$(L "Report generato:" "Report generated:") ${BOLD}$REPORT_HTML${RESET}"
        open_report_prompt "$REPORT_HTML"
    else
        local _PYERR _PYINFO
        _PYERR=$(grep -v '^STATS:' "$PYERR_FILE" 2>/dev/null | grep -iv '^INFO:' | head -10 || true)
        _PYINFO=$(grep -i '^INFO:' "$PYERR_FILE" 2>/dev/null | head -10 || true)
        [[ -n "$_PYINFO" ]] && info "$(L "Debug ESE:" "ESE debug:") $_PYINFO"
        [[ -n "$_PYERR" ]] && warn "$(L "Errore Python:" "Python error:") $_PYERR"

        # Diagnosi specifica per errori noti
        if echo "${_PYERR}" | grep -qiE 'Errno 35|EAGAIN|temporarily unavailable|Resource temporarily'; then
            err "$(L "ntds.dit è bloccato dal servizio Active Directory (Errno 35 / EAGAIN)." "ntds.dit is locked by Active Directory service (Errno 35 / EAGAIN).")"
            echo ""
            echo -e "  ${YELLOW}$(L "Il client SMB rispetta il lock esclusivo tenuto da lsass/ntds.exe." "The SMB client respects the exclusive lock held by lsass/ntds.exe.")${RESET}"
            echo -e "  ${YELLOW}$(L "Il file non è apribile nemmeno in lettura finché il DC è attivo." "The file cannot be opened even for reading while the DC is active.")${RESET}"
            echo ""
            echo -e "  ${BOLD}$(L "Soluzioni per DC live:" "Solutions for live DC:")${RESET}"
            echo -e "  ${CYAN}$(L "1) VSS snapshot (consigliato):" "1) VSS snapshot (recommended):")${RESET}"
            echo -e "     $(L "Sul DC (PowerShell admin):" "On the DC (PowerShell admin):")"
            echo -e "     ${DIM}ntdsutil \"activate instance ntds\" \"ifm\" \"create full C:\\\\ifm_snapshot\" quit quit${RESET}"
            echo -e "     $(L "Poi monta/copia la cartella C:\\\\ifm_snapshot e rilancia su quel path." "Then mount/copy the C:\\\\ifm_snapshot folder and relaunch on that path.")"
            echo ""
            echo -e "  ${CYAN}$(L "2) impacket secretsdump remoto (solo hash NTLM, senza analisi ACL/GPO):" "2) remote impacket secretsdump (NTLM hashes only, no ACL/GPO analysis):")${RESET}"
            echo -e "     ${DIM}impacket-secretsdump -just-dc-ntlm DOMINIO/utente:password@IP_DC${RESET}"
            echo ""
            echo -e "  ${CYAN}$(L "3) Se hai già un'immagine disco (VHDX/E01) del DC:" "3) If you already have a disk image (VHDX/E01) of the DC:")${RESET}"
            echo -e "     $(L "Monta offline con" "Mount offline with") ${BOLD}./mount_image.sh${RESET} $(L "e rilancia il modulo su quel mount point." "and relaunch the module on that mount point.")"
        elif echo "${_PYERR}" | grep -qiE 'dirty|state|DirtyShutdown|Unknown state'; then
            err "$(L "Generazione report fallita — database ESE in stato dirty non recuperabile." "Report generation failed — ESE database in unrecoverable dirty state.")"
            info "$(L "Suggerimento: copia il file e prova 'esentutl /r edb /l .' in un ambiente Windows" "Hint: copy the file and try 'esentutl /r edb /l .' in a Windows environment")"
        else
            err "$(L "Generazione report fallita. Verifica che ntds.dit sia leggibile e non corrotto." "Report generation failed. Verify ntds.dit is readable and not corrupted.")"
            info "$(L "Suggerimento: copia il file e prova 'esentutl /r edb' in un ambiente Windows" "Hint: copy the file and try 'esentutl /r edb' in a Windows environment")"
        fi
    fi
    rm -rf "$TMP_PAD_DIR" 2>/dev/null || true
}

# ================================================================
#  MODULO 39 — AI Chat / Query History
# ================================================================
# Recupera la cronologia delle chat con assistenti AI lasciata sul
# disco dagli applicativi desktop/CLI piu' diffusi. Riferimento:
#   https://r3nzsec.github.io/irflow-timeline/dfir-tips/ai-query-history
#
# Tutti gli artefatti sono cercati sotto le home utente Windows
# (Users\<utente>\...). I formati JSON/JSONL vengono parsati e i
# messaggi estratti; i DB SQLite (state.vscdb, store.db, state.sqlite)
# vengono interrogati per le chiavi di chat; i bundle cifrati / LevelDB /
# protobuf Cascade sono solo inventariati (path, dimensione, SHA256).
module_ai_chat() {
    section_header "AI Chat / Query History" "$CYAN"
    check_win_root || return 1

    local MANIFEST; MANIFEST=$(mktemp); register_tmp "$MANIFEST"
    local IOCTMP;   IOCTMP=$(mktemp);   register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"

    TOTAL_FILES=0
    local TOTAL_USERS=0

    # Aggiunge al manifest le righe lette da stdin (una per path trovato).
    # Gira nella shell corrente (chiamata con < <(...)) cosi' i contatori
    # globali vengono aggiornati. Formato manifest: tool<TAB>kind<TAB>user<TAB>path
    _ai_collect() {
        local _tool="$1" _kind="$2" _user="$3" _f
        while IFS= read -r _f; do
            [[ -n "$_f" ]] || continue
            printf '%s\t%s\t%s\t%s\n' "$_tool" "$_kind" "$_user" "$_f" >> "$MANIFEST"
            TOTAL_FILES=$((TOTAL_FILES+1))
            _USER_FOUND=$((_USER_FOUND+1))
        done
    }

    info "$(L "Ricerca artefatti chat AI nelle home utente..." "Searching AI chat artifacts in user homes...")"

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local _USER_FOUND=0

        # ── Claude Code (CLI) ──────────────────────────────────────
        _ai_collect "Claude Code (CLI)" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.claude/history.jsonl' -type f 2>/dev/null)
        _ai_collect "Claude Code (CLI)" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 6 -ipath '*/.claude/projects/*' -iname '*.jsonl' -type f 2>/dev/null)
        # ── Claude Desktop ─────────────────────────────────────────
        _ai_collect "Claude Desktop" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 7 -ipath '*/AppData/Roaming/Claude/claude-code-sessions/*' -iname '*.json' -type f 2>/dev/null)
        _ai_collect "Claude Desktop" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 7 -ipath '*/AppData/Roaming/Claude/local-agent-mode-sessions/*' -iname '*.json' -type f 2>/dev/null)
        # ── ChatGPT Desktop (standalone) ───────────────────────────
        _ai_collect "ChatGPT Desktop" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/*' \( -iname '*.ldb' -o -iname '*.log' \) -type f 2>/dev/null)
        _ai_collect "ChatGPT Desktop" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 12 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/IndexedDB/*.indexeddb.blob/*' -type f 2>/dev/null)
        _ai_collect "ChatGPT Desktop" inventory "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/*' -iname 'conversations-v2-*' -type f 2>/dev/null)
        _ai_collect "ChatGPT Desktop" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/*' \( -iname '*.sqlite' -o -iname '*.db' \) -type f 2>/dev/null)
        # ── ChatGPT (Microsoft Store) ──────────────────────────────
        _ai_collect "ChatGPT (Store)" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/*' \( -iname '*.ldb' -o -iname '*.log' \) -type f 2>/dev/null)
        _ai_collect "ChatGPT (Store)" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 14 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/IndexedDB/*.indexeddb.blob/*' -type f 2>/dev/null)
        _ai_collect "ChatGPT (Store)" inventory "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/*' -iname 'conversations-v2-*' -type f 2>/dev/null)
        _ai_collect "ChatGPT (Store)" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/*' \( -iname '*.sqlite' -o -iname '*.db' \) -type f 2>/dev/null)
        # ── OpenAI Codex ───────────────────────────────────────────
        _ai_collect "OpenAI Codex" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.codex/history.jsonl' -type f 2>/dev/null)
        _ai_collect "OpenAI Codex" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.codex/session_index.jsonl' -type f 2>/dev/null)
        _ai_collect "OpenAI Codex" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 8 \( -ipath '*/.codex/sessions/*' -o -ipath '*/.codex/archived_sessions/*' \) -iname 'rollout-*.jsonl' -type f 2>/dev/null)
        _ai_collect "OpenAI Codex" sqlite_codex "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.codex/state.sqlite' -type f 2>/dev/null)
        # ── Gemini CLI ─────────────────────────────────────────────
        _ai_collect "Gemini CLI" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 6 -ipath '*/.gemini/tmp/*/chats/*' -iname 'session-*.json' -type f 2>/dev/null)
        _ai_collect "Gemini CLI" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 5 -ipath '*/.gemini/tmp/*' \( -iname 'logs.json' -o -iname 'checkpoint-*.json' \) -type f 2>/dev/null)
        # ── Cursor ─────────────────────────────────────────────────
        _ai_collect "Cursor" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 7 -ipath '*/.cursor/projects/*/agent-transcripts/*' -iname '*.jsonl' -type f 2>/dev/null)
        _ai_collect "Cursor" sqlite_store "$USERNAME" < <(find "$USER_DIR" -maxdepth 6 -ipath '*/.cursor/chats/*' -iname 'store.db' -type f 2>/dev/null)
        _ai_collect "Cursor" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/Cursor/User/*Storage/*' -iname 'state.vscdb' -type f 2>/dev/null)
        # ── GitHub Copilot (VS Code / VSCodium) ────────────────────
        local _APP
        for _APP in "Code" "Code - Insiders" "VSCodium"; do
            _ai_collect "GitHub Copilot ($_APP)" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath "*/AppData/Roaming/$_APP/User/workspaceStorage/*/chatSessions/*" -iname '*.json' -type f 2>/dev/null)
            _ai_collect "GitHub Copilot ($_APP)" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath "*/AppData/Roaming/$_APP/User/globalStorage/emptyWindowChatSessions/*" -iname '*.json' -type f 2>/dev/null)
            _ai_collect "GitHub Copilot ($_APP)" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath "*/AppData/Roaming/$_APP/User/*Storage/*" -iname 'state.vscdb' -type f 2>/dev/null)
        done
        # ── Windsurf ───────────────────────────────────────────────
        _ai_collect "Windsurf" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/Windsurf/User/*Storage/*' -iname 'state.vscdb' -type f 2>/dev/null)
        _ai_collect "Windsurf" inventory "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/windsurf.cascade/*' -iname '*.pb' -type f 2>/dev/null)
        # ── Continue.dev ───────────────────────────────────────────
        _ai_collect "Continue.dev" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 4 -ipath '*/.continue/sessions/*' -iname '*.json' -type f 2>/dev/null)

        if [[ $_USER_FOUND -gt 0 ]]; then
            ok "$USERNAME — ${BOLD}${_USER_FOUND}${RESET} $(L "artefatti AI" "AI artifacts")"
            TOTAL_USERS=$((TOTAL_USERS+1))
        else
            dim_msg "$USERNAME — $(L "nessun artefatto AI" "no AI artifacts")"
        fi
    done < <(get_user_homes)
    unset -f _ai_collect

    separator
    info "$(L "Utenti con chat AI:" "Users with AI chats:") ${BOLD}$TOTAL_USERS${RESET}  |  $(L "File totali:" "Total files:") ${BOLD}$TOTAL_FILES"

    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun artefatto di chat AI trovato." "No AI chat artifact found.")"; return 0; }
    ask_yn "$(L "Generare report HTML?" "Generate HTML report?")" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ai_chat")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Parser Python: legge il manifest + il file IoC e produce
    #   riga 1: STATS:<file>|<utenti>|<tool>|<sensibili>|<ioc>
    #   resto : HTML delle card (contenuto di <div class='cards'>)
    local PARSER_OUT
    PARSER_OUT=$(run_py_with_lib pylib_leveldb "$MANIFEST" "$IOCTMP" << 'PYEOF'
import sys, os, re, json, html, sqlite3, shutil, tempfile, hashlib, datetime
from collections import OrderedDict

ISO_TS = re.compile(r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}')

def norm_ts(ts):
    """Normalizza un timestamp eterogeneo (ISO o epoch s/ms) in
    'YYYY-MM-DD HH:MM:SS' (UTC). Ritorna la stringa grezza se non parsabile."""
    if not ts:
        return ''
    ts = str(ts).strip()
    m = ISO_TS.search(ts)
    if m:
        return m.group(0).replace('T', ' ')[:19]
    try:
        v = float(ts)
        if v > 1e12:      # millisecondi
            v /= 1000.0
        if 1e8 < v < 1e11:
            return datetime.datetime.fromtimestamp(
                v, datetime.timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        pass
    return ts

manifest_path, ioc_path = sys.argv[1], sys.argv[2]

IOCS = []
try:
    with open(ioc_path, encoding='utf-8', errors='replace') as f:
        for ln in f:
            ln = ln.strip().lower()
            if ln and not ln.startswith('#'):
                IOCS.append(ln)
except Exception:
    pass

SENS = ['api_key','apikey','api-key','secret','password','passwd','token',
        'bearer ','authorization','sk-','ghp_','gho_','xoxb','xoxp','aws_secret',
        'private_key','-----begin','client_secret','credential','passphrase',
        'connectionstring']

MAX_MSGS = 300
MAX_TXT  = 6000

TEXT_KEYS  = ('text','content','display','message','prompt','body','parts',
              'summary','thinking','reasoning','response','rawText','value')
ROLE_KEYS  = ('role','type','sender','author','speaker','kind')
TS_KEYS    = ('timestamp','time','created_at','createdAt','ts','date',
              'updated_at','updatedAt','startTime','responseTime')
ARRAY_KEYS = ('messages','requests','history','chats','entries','items',
              'turns','conversation','events','log','records','responses')

def gstr(d, keys):
    for k in keys:
        if k in d and isinstance(d[k], (str, int, float)):
            return str(d[k])
    return ''

def collect_text(v, depth=0, acc=None):
    if acc is None: acc = []
    if depth > 6 or len(acc) > 40: return acc
    if isinstance(v, str):
        s = v.strip()
        if s: acc.append(s)
    elif isinstance(v, dict):
        for k in TEXT_KEYS:
            if k in v:
                collect_text(v[k], depth+1, acc)
    elif isinstance(v, list):
        for it in v:
            collect_text(it, depth+1, acc)
    return acc

def mk_msg(m):
    if isinstance(m, str):
        s = m.strip()
        return ('', '', s) if s else None
    if not isinstance(m, dict):
        return None
    txt = '\n'.join(collect_text(m)).strip()
    if not txt:
        return None
    role = gstr(m, ROLE_KEYS)
    # Voci della prompt-history (es. Claude history.jsonl: {"display": "..."}) →
    # sono prompt dell'utente anche se prive di campo ruolo.
    if not role and ('display' in m or 'prompt' in m):
        role = 'user'
    return (gstr(m, TS_KEYS), role, txt)

def extract(obj, out, depth=0):
    if len(out) >= MAX_MSGS or depth > 5: return
    if isinstance(obj, dict):
        handled = False
        for ak in ARRAY_KEYS:
            v = obj.get(ak)
            if isinstance(v, list) and v and isinstance(v[0], dict):
                for m in v:
                    mm = mk_msg(m)
                    if mm: out.append(mm)
                    if len(out) >= MAX_MSGS: return
                handled = True
        if not handled:
            if any(k in obj for k in TEXT_KEYS):
                mm = mk_msg(obj)
                if mm: out.append(mm)
            else:
                for vv in obj.values():
                    extract(vv, out, depth+1)
    elif isinstance(obj, list):
        if obj and isinstance(obj[0], dict) and any(
                isinstance(it, dict) and any(k in it for k in TEXT_KEYS) for it in obj[:5]):
            for m in obj:
                mm = mk_msg(m)
                if mm: out.append(mm)
                if len(out) >= MAX_MSGS: return
        else:
            for it in obj:
                extract(it, out, depth+1)

# Tipi di voce NON conversazionali nei session log (Claude Code, ecc.) da saltare.
JSONL_SKIP_TYPES = {'queue-operation', 'file-history-snapshot', 'ai-title',
                    'last-prompt', 'attachment', 'summary', 'progress'}

def parse_jsonl(path):
    out = []
    with open(path, encoding='utf-8', errors='replace') as f:
        for line in f:
            line = line.strip()
            if not line: continue
            try:
                obj = json.loads(line)
            except Exception:
                out.append(('', '', line[:MAX_TXT])); continue
            if isinstance(obj, dict) and obj.get('type') in JSONL_SKIP_TYPES:
                continue
            if isinstance(obj, dict) and any(k in obj for k in TEXT_KEYS):
                mm = mk_msg(obj)
                if mm: out.append(mm)
            else:
                extract(obj, out)
            if len(out) >= MAX_MSGS: break
    return out

def parse_json(path):
    out = []
    with open(path, encoding='utf-8', errors='replace') as f:
        obj = json.load(f)
    extract(obj, out)
    return out

def sqlite_rows(path, like_keys, tables):
    _fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd); out = []
    try:
        shutil.copy2(path, tmp)
        con = sqlite3.connect(tmp)
        con.text_factory = lambda b: b.decode('utf-8', 'replace') if isinstance(b, bytes) else b
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        present = {r[0] for r in cur.fetchall()}
        for t in tables:
            if t not in present: continue
            cur.execute("PRAGMA table_info(%s)" % t)
            cols = [c[1] for c in cur.fetchall()]
            if 'key' not in cols or 'value' not in cols: continue
            where = ' OR '.join("lower(key) LIKE ?" for _ in like_keys)
            cur.execute("SELECT key,value FROM %s WHERE %s" % (t, where),
                        tuple('%%%s%%' % k for k in like_keys))
            for k, v in cur.fetchall():
                if v is None: continue
                vs = v if isinstance(v, str) else str(v)
                sub = []
                try:
                    extract(json.loads(vs), sub)
                except Exception:
                    pass
                if sub:
                    for ts, role, txt in sub[:60]:
                        out.append((ts, role or k, txt))
                else:
                    out.append(('', k, vs[:MAX_TXT]))
                if len(out) >= MAX_MSGS: break
        con.close()
    except Exception as e:
        out.append(('', 'error', 'SQLite: %s' % e))
    finally:
        try: os.remove(tmp)
        except Exception: pass
    return out

def sqlite_meta(path):
    _fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd); out = []
    try:
        shutil.copy2(path, tmp)
        con = sqlite3.connect(tmp); cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        info = []
        for (t,) in cur.fetchall():
            try:
                cur.execute("SELECT count(*) FROM %s" % t)
                info.append('%s: %d' % (t, cur.fetchone()[0]))
            except Exception:
                info.append(t)
        con.close()
        out.append(('', 'metadata', 'Tabelle/rows: ' + ', '.join(info)))
    except Exception as e:
        out.append(('', 'error', str(e)))
    finally:
        try: os.remove(tmp)
        except Exception: pass
    return out

def sha256_short(p):
    try:
        h = hashlib.sha256()
        with open(p, 'rb') as f:
            for ch in iter(lambda: f.read(65536), b''):
                h.update(ch)
        return h.hexdigest()[:16]
    except Exception:
        return '?'


VSCDB_KEYS = ['chat','copilot','aichat','interactive.session','cascade',
              'composer','aiservice','windsurf','continue','prompt']
STORE_KEYS = ['composer','message','chat','bubble']

def parse_file(kind, path):
    if kind == 'jsonl':         return parse_jsonl(path)
    if kind == 'json':          return parse_json(path)
    if kind == 'sqlite_vscdb':  return sqlite_rows(path, VSCDB_KEYS, ['ItemTable', 'cursorDiskKV'])
    if kind == 'sqlite_store':  return sqlite_rows(path, STORE_KEYS, ['cursorDiskKV', 'ItemTable'])
    if kind == 'sqlite_codex':  return sqlite_meta(path)
    if kind == 'leveldb':       return carve_leveldb(path)
    return []  # inventory

def role_class(role):
    r = (role or '').lower()
    if 'user' in r or 'human' in r or 'prompt' in r: return 'r-user'
    if 'assist' in r or 'model' in r or 'gemini' in r or 'ai' in r or 'response' in r: return 'r-ai'
    if 'tool' in r or 'function' in r or 'shell' in r: return 'r-tool'
    if 'error' in r: return 'r-err'
    return 'r-meta'

def ai_product(tool):
    """Nome leggibile dell'assistente AL dal tool del record."""
    t = (tool or '').lower()
    if 'chatgpt' in t or ('openai' in t and 'codex' not in t): return 'ChatGPT'
    if 'copilot' in t:  return 'Copilot'
    if 'claude' in t:   return 'Claude'
    if 'gemini' in t:   return 'Gemini'
    if 'codex' in t:    return 'Codex'
    if 'cursor' in t:   return 'Cursor'
    if 'windsurf' in t: return 'Windsurf'
    if 'continue' in t: return 'Continue'
    return re.sub(r'\s*\(.*\)$', '', tool).strip() or 'AI'

def render(msgs, ai='AI'):
    rows, sens_n, ioc_n = [], 0, 0
    for ts, role, txt in msgs:
        t = txt[:MAX_TXT]
        low = t.lower()
        is_sens = any(s in low for s in SENS)
        is_ioc  = bool(IOCS) and any(i in low for i in IOCS)
        if is_sens: sens_n += 1
        if is_ioc:  ioc_n += 1
        rc = role_class(role)
        cls = 'msg ' + rc
        if is_sens: cls += ' sensitive'
        if is_ioc:  cls += ' ioc'
        if rc == 'r-user':
            label = 'UTENTE'
        elif rc == 'r-ai':
            label = 'AI — ' + html.escape(ai)
        else:
            label = html.escape(role) if role else '—'
        nts   = norm_ts(ts)
        meta  = html.escape(nts) if nts else ''
        tag   = ''
        if is_ioc:  tag += "<span class='tag tag-ioc'>IOC</span>"
        if is_sens: tag += "<span class='tag tag-sec'>SENS</span>"
        rows.append(
            "<div class='%s'><div class='msg-h'><span class='role'>%s</span>%s"
            "<span class='mts'>%s</span></div><pre class='msg-t'>%s</pre></div>"
            % (cls, label, tag, meta, html.escape(t)))
    return rows, sens_n, ioc_n

data = OrderedDict()
with open(manifest_path, encoding='utf-8', errors='replace') as f:
    for line in f:
        line = line.rstrip('\n')
        if not line: continue
        parts = line.split('\t')
        if len(parts) < 4: continue
        tool, kind, user = parts[0], parts[1], parts[2]
        path = '\t'.join(parts[3:])
        data.setdefault(user, OrderedDict()).setdefault(tool, []).append((kind, path))

tot_files = tot_sens = tot_ioc = 0
tools_seen = set()
cards = []

for user, tools in data.items():
    tool_blocks = []
    user_files = 0
    for tool, files in tools.items():
        tools_seen.add(tool)
        file_blocks = []
        for kind, path in files:
            tot_files += 1; user_files += 1
            try:
                st = os.stat(path)
                size = st.st_size
                mt = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                size, mt = 0, '?'
            try:
                msgs = parse_file(kind, path)
            except Exception as e:
                msgs = [('', 'error', '%s' % e)]
            rows, sn, ic = render(msgs, ai_product(tool))
            tot_sens += sn; tot_ioc += ic
            sha = sha256_short(path)
            base = html.escape(os.path.basename(path))
            meta = ("<span>%s</span> &middot; <span>%d B</span> &middot; "
                    "<span>mtime %s</span> &middot; <span>sha256 %s</span>"
                    % (html.escape(kind), size, mt, sha))
            badges = "<span class='fb'>%d msg</span>" % len(rows)
            if sn: badges += "<span class='fb fb-sec'>%d sens</span>" % sn
            if ic: badges += "<span class='fb fb-ioc'>%d ioc</span>" % ic
            if rows:
                note = ''
                if kind == 'leveldb':
                    note = ("<div class='inv'>%s</div>" % html.escape(
                        "Contenuto ricavato per string-carving dal LevelDB "
                        "(best-effort, non strutturato — verificare manualmente)."))
                body = note + "<div class='thread'>%s</div>" % ''.join(rows)
            elif kind in ('inventory', 'pb_inventory'):
                body = ("<div class='inv'>%s</div>" %
                        html.escape("Solo inventario — contenuto cifrato/proprietario non decodificato "
                                    "(LevelDB / bundle conversations-v2 / protobuf Cascade)."))
            else:
                body = "<div class='inv'>%s</div>" % html.escape("Nessun messaggio estratto.")
            file_blocks.append(
                "<div class='fblk'><div class='fhdr'><div class='fname'>%s</div>"
                "<div class='fbadges'>%s</div></div><div class='fpath'>%s</div>"
                "<div class='fmeta'>%s</div>%s</div>"
                % (base, badges, html.escape(path), meta, body))
        tool_blocks.append(
            "<div class='toolsec'><div class='toolname'>%s <span class='dim'>(%d)</span></div>%s</div>"
            % (html.escape(tool), len(files), ''.join(file_blocks)))
    cards.append(
        "<div class='card'><div class='card-header'><div class='uicon'>AI</div>"
        "<div class='user-info'><div class='uname'>%s</div>"
        "<div class='upath'>%d file</div></div>"
        "<div class='badge'>%d tool</div></div>"
        "<div style='padding:1rem 1.5rem'>%s</div></div>"
        % (html.escape(user), user_files, len(tools), ''.join(tool_blocks)))

print("STATS:%d|%d|%d|%d|%d" % (tot_files, len(data), len(tools_seen), tot_sens, tot_ioc))
sys.stdout.write(''.join(cards))
PYEOF
)

    local STATS_LINE; STATS_LINE=$(printf '%s\n' "$PARSER_OUT" | head -1)
    local CARDS_HTML; CARDS_HTML=$(printf '%s\n' "$PARSER_OUT" | tail -n +2)
    local _S="${STATS_LINE#STATS:}"
    local ST_FILES ST_USERS ST_TOOLS ST_SENS ST_IOC
    IFS='|' read -r ST_FILES ST_USERS ST_TOOLS ST_SENS ST_IOC <<< "$_S"

    {
        html_header "AI Chat History" "#a371f7"
        html_page_header "AI" "AI Chat / <span>Query History</span> Forensics" \
            "Claude · ChatGPT · Copilot · Cursor · Gemini · Codex · Windsurf · Continue" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>"
        echo "  <div class='stat'><div class='label'>$(L "File" "Files")</div><div class='value'>${ST_FILES:-0}</div></div>"
        echo "  <div class='stat info'><div class='label'>$(L "Utenti" "Users")</div><div class='value'>${ST_USERS:-0}</div></div>"
        echo "  <div class='stat info'><div class='label'>Tool</div><div class='value'>${ST_TOOLS:-0}</div></div>"
        echo "  <div class='stat'><div class='label'>$(L "Sensibili" "Sensitive")</div><div class='value'>${ST_SENS:-0}</div></div>"
        echo "  <div class='stat ok'><div class='label'>IoC</div><div class='value'>${ST_IOC:-0}</div></div>"
        echo "</div>"
        echo "<main><div class='stitle'>$(L "Risultati per utente" "Results per user")</div>"
        cat << 'CSSEOF'
<style>
  .toolsec{margin-bottom:1rem;border:1px solid var(--border);border-radius:6px;overflow:hidden}
  .toolname{font-family:var(--mono);font-size:.8rem;font-weight:700;color:var(--accent);
    background:var(--bg3);padding:.5rem .9rem;border-bottom:1px solid var(--border)}
  .toolname .dim{color:var(--text-dim);font-weight:400}
  .fblk{padding:.7rem .9rem;border-bottom:1px solid rgba(30,45,61,.5)}
  .fblk:last-child{border-bottom:none}
  .fhdr{display:flex;align-items:center;gap:1rem}
  .fname{font-family:var(--mono);font-size:.78rem;color:#fff;font-weight:500}
  .fbadges{margin-left:auto;display:flex;gap:.35rem}
  .fb{font-family:var(--mono);font-size:.6rem;padding:.1rem .5rem;border-radius:2px;
    background:rgba(88,166,255,.1);border:1px solid rgba(88,166,255,.3);color:var(--accent)}
  .fb-sec{background:rgba(240,136,62,.1);border-color:rgba(240,136,62,.4);color:var(--accent4)}
  .fb-ioc{background:rgba(255,123,114,.12);border-color:rgba(255,123,114,.4);color:var(--accent2)}
  .fpath{font-family:var(--mono);font-size:.6rem;color:var(--text-dim);word-break:break-all;margin:.2rem 0}
  .fmeta{font-family:var(--mono);font-size:.6rem;color:var(--text-mid);margin-bottom:.5rem}
  .thread{display:flex;flex-direction:column;gap:.4rem;max-height:520px;overflow-y:auto;
    padding:.3rem;background:var(--bg);border-radius:4px}
  .thread::-webkit-scrollbar{width:5px}
  .thread::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
  .msg{border-left:3px solid var(--border);padding:.35rem .6rem;border-radius:0 3px 3px 0;background:var(--bg2)}
  .msg-h{display:flex;align-items:center;gap:.5rem;margin-bottom:.2rem}
  .role{font-family:var(--mono);font-size:.6rem;text-transform:uppercase;letter-spacing:.08em;color:var(--text-mid)}
  .mts{margin-left:auto;font-family:var(--mono);font-size:.58rem;color:var(--text-dim)}
  .msg-t{font-family:var(--mono);font-size:.72rem;line-height:1.5;white-space:pre-wrap;word-break:break-word;color:var(--text)}
  .r-user{border-left:3px solid var(--accent);background:rgba(88,166,255,.08);margin-left:2.5rem}
  .r-user .role{color:var(--accent)}
  .r-user .role::before{content:'\1F464  '}
  .r-ai{border-left:3px solid var(--accent3);background:rgba(63,185,80,.07);margin-right:2.5rem}
  .r-ai .role{color:var(--accent3)}
  .r-ai .role::before{content:'\1F916  '}
  .r-tool{border-left-color:var(--accent4)}
  .r-tool .role{color:var(--accent4)}
  .r-err{border-left-color:var(--accent2)}
  .msg.sensitive{background:rgba(240,136,62,.07);border-left-color:var(--accent4)}
  .msg.ioc{background:rgba(255,123,114,.1);border-left-color:var(--accent2)}
  .tag{font-family:var(--mono);font-size:.55rem;padding:0 .35rem;border-radius:2px;font-weight:700}
  .tag-sec{background:rgba(240,136,62,.2);color:var(--accent4)}
  .tag-ioc{background:rgba(255,123,114,.2);color:var(--accent2)}
  .inv{font-family:var(--mono);font-size:.68rem;color:var(--text-dim);font-style:italic;padding:.3rem 0}
</style>
CSSEOF
        echo "<div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 40 — SetupAPI Device Log
#
#  setupapi.dev.log registra l'installazione dei driver di ogni dispositivo
#  collegato alla macchina. E' l'unica fonte che data con precisione la PRIMA
#  connessione di un dispositivo USB: il registro (modulo 8, USBSTOR) conserva
#  l'ultima connessione e i metadati, ma non il momento in cui il dispositivo
#  e' comparso per la prima volta.
#
#  In un caso di esfiltrazione la differenza conta: "questa chiavetta e' stata
#  collegata per la prima volta il giorno X" e' un'affermazione che si sostiene
#  solo con questo log.
# ================================================================
module_setupapi() {
    section_header "SetupAPI — $(L "Prima installazione dispositivi" "Device First Install")" "$BLUE"
    check_win_root || return 1

    local -a LOGS=()
    local D F
    D=$(ci_find_dir "$WIN_ROOT" "Windows/INF")
    if [[ -n "$D" ]]; then
        while IFS= read -r F; do
            [[ -n "$F" ]] && LOGS+=("$F")
        done < <(find "$D" -maxdepth 1 -type f -iname 'setupapi.dev*.log' 2>/dev/null)
    fi
    # Percorso legacy (XP/2003)
    D=$(ci_find_dir "$WIN_ROOT" "Windows")
    if [[ -n "$D" ]]; then
        F=$(ci_find_file "$D" "setupapi.log")
        [[ -n "$F" ]] && LOGS+=("$F")
    fi

    if [[ ${#LOGS[@]} -eq 0 ]]; then
        warn "$(L "Nessun setupapi.dev.log trovato." "No setupapi.dev.log found.")"
        return 0
    fi
    info "$(L "Log trovati:" "Logs found:") ${BOLD}${#LOGS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${LOGS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os

out_path = sys.argv[1]
logs = sys.argv[2:]

# Struttura del log:
#   >>>  [Device Install (Hardware initiated) - SWD\WPDBUSENUM\{GUID}#...]
#   >>>  Section start 2024/01/15 10:23:45.123
DEV = re.compile(r'^>>>\s+\[(?P<what>[^\]]+)\]')
SEC = re.compile(r'^>>>\s+Section start\s+(?P<ts>\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})')

# Bus che identificano supporti rimovibili: sono quelli che interessano in
# un'indagine su esfiltrazione.
REMOVABLE = ('USBSTOR', 'USB\\', 'WPDBUSENUM', 'SWD\\WPDBUSENUM', 'SCSI\\DISK')

rows = []
for path in logs:
    src = os.path.basename(path)
    try:
        fh = open(path, encoding='utf-8', errors='replace')
    except Exception:
        continue
    pending = None
    with fh:
        for line in fh:
            line = line.rstrip('\n')
            m = DEV.match(line)
            if m:
                pending = m.group('what').strip()
                continue
            m = SEC.match(line)
            if m and pending:
                ts = m.group('ts').replace('/', '-')
                what = pending
                pending = None
                # "Device Install (Hardware initiated) - <istanza>"
                inst = what.split(' - ', 1)[1] if ' - ' in what else what
                kind = what.split(' - ', 1)[0] if ' - ' in what else ''
                up = inst.upper()
                removable = '1' if any(k in up for k in REMOVABLE) else '0'
                rows.append((ts, kind, inst, removable, src))

rows.sort(key=lambda r: r[0])
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ') for x in r) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna installazione di dispositivo interpretabile." "No parsable device installation.")"
        return 0
    fi
    local NREM; NREM=$(awk -F'\t' '$4=="1"' "$OUT" | wc -l)

    ok "$(L "Installazioni registrate:" "Recorded installations:") ${BOLD}$TOTAL"
    info "$(L "Di cui supporti rimovibili:" "Of which removable media:") ${BOLD}$NREM"
    echo ""
    awk -F'\t' '$4=="1"{printf "      %s  %s\n", $1, substr($3,1,90)}' "$OUT" | head -25 | while IFS= read -r LN; do
        echo -e "      ${CYAN}${LN}${RESET}"
    done

    # Marca gli IoC eventualmente caricati.
    if [[ ${#IOC_LIST[@]} -gt 0 ]]; then
        local NIOC=0 LINE
        while IFS= read -r LINE; do
            check_ioc "$LINE" && NIOC=$((NIOC + 1))
        done < "$OUT"
        [[ $NIOC -gt 0 ]] && warn "$(L "Righe con match IoC:" "Lines matching IoC:") ${BOLD}$NIOC"
    fi

    ask_yn "Generare report HTML?" || return 0

    # Prima i rimovibili: sono il motivo per cui si guarda questo log.
    local ROWS; ROWS=$( { awk -F'\t' '$4=="1"{print $1"\t"$2"\t"$3"\t"$5}' "$OUT";
                          awk -F'\t' '$4=="0"{print $1"\t"$2"\t"$3"\t"$5}' "$OUT"; } )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Prima installazione" "First install")" "$(L "Tipo" "Kind")" \
        "$(L "Istanza dispositivo" "Device instance")" "Log")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' conta" "Why it matters")</b><br>"
    NOTE+="$(L "Il registro USBSTOR (modulo 8) conserva l'ULTIMA connessione di un dispositivo. Questo log conserva la PRIMA: e' l'unica fonte che permette di affermare quando un supporto e' comparso per la prima volta sulla macchina. I due vanno letti insieme." \
        "The USBSTOR registry (module 8) keeps a device's LAST connection. This log keeps the FIRST: it is the only source supporting a claim about when a device first appeared on the machine. Read the two together.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Installazioni" "Installations")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Rimovibili" "Removable")" "$NREM" "warn")"
    STATS+="$(stat_box "Log" "${#LOGS[@]}" "info")"
    finish_report "setupapi" "SetupAPI Device Log" "DEV" "Windows/INF/setupapi.dev.log" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Installazioni dispositivi" "Device installations")" "${LOGS[0]}" "$TOTAL" "$TABLE" "⇄")</div>"
}

# ================================================================
#  MODULO 41 — PowerShell Transcript
#
#  Il modulo 1 legge PSReadLine, che conserva solo le righe DIGITATE nella
#  console. I transcript (Start-Transcript, o la policy "Turn on PowerShell
#  Transcription") registrano invece la sessione completa: comandi, output,
#  utente, host e processo.
#
#  Differenza pratica: PSReadLine non vede nulla di quanto eseguito da script,
#  da -EncodedCommand o da una sessione remota; i transcript sì, e includono
#  anche l'output — spesso l'unica traccia rimasta di cosa un comando abbia
#  effettivamente restituito.
# ================================================================
module_ps_transcripts() {
    section_header "PowerShell Transcript" "$MAGENTA"
    check_win_root || return 1

    local -a FILES=()
    local F

    # I transcript finiscono in Documenti per default, ma la policy consente
    # una OutputDirectory arbitraria: si cercano anche nelle posizioni comuni.
    local -a ROOTS=()
    while IFS= read -r F; do
        [[ -n "$F" ]] && ROOTS+=("$F")
    done < <(get_user_homes)
    local D
    for D in "Windows/Temp" "Temp" "ProgramData/Microsoft/Windows/PowerShell" "Transcripts" "PSTranscripts"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ROOTS+=("$R")
    done

    local R
    for R in "${ROOTS[@]}"; do
        while IFS= read -r F; do
            [[ -n "$F" ]] && FILES+=("$F")
        done < <(find "$R" -maxdepth 5 -type f -iname 'PowerShell_transcript*.txt' 2>/dev/null)
    done

    if [[ ${#FILES[@]} -eq 0 ]]; then
        warn "$(L "Nessun transcript PowerShell trovato." "No PowerShell transcript found.")"
        info "$(L "I transcript esistono solo se attivati (Start-Transcript o policy di trascrizione)." \
                 "Transcripts exist only if enabled (Start-Transcript or transcription policy).")"
        return 0
    fi
    info "$(L "Transcript trovati:" "Transcripts found:") ${BOLD}${#FILES[@]}"

    # Parole che segnalano attivita' offensiva dentro una sessione PowerShell.
    local KW="downloadstring|downloadfile|iex|invoke-expression|invoke-webrequest|frombase64string|encodedcommand|-enc |bypass|hidden|net user|net localgroup|mimikatz|invoke-mimikatz|add-mppreference|set-mppreference|disable-windowsoptionalfeature|schtasks|reg add|vssadmin|bcdedit|wevtutil|certutil|bitsadmin|password|secret|token"
    local -a IOCS=("${IOC_LIST[@]:-}")
    local I
    for I in "${IOCS[@]}"; do
        [[ -n "$I" ]] && KW+="|${I,,}"
    done

    local BODY="" TOTAL=0 SUSP=0
    local META; META=$(mktemp); register_tmp "$META"
    for F in "${FILES[@]}"; do
        [[ -s "$F" ]] || continue
        TOTAL=$((TOTAL + 1))
        local N; N=$(grep -icE "$KW" "$F" 2>/dev/null || echo 0)
        [[ "$N" -gt 0 ]] && SUSP=$((SUSP + 1))

        # L'intestazione del transcript contiene utente, host e PID.
        local U H P ST
        U=$(grep -m1 -iE '^Username:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        H=$(grep -m1 -iE '^Host Application:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        P=$(grep -m1 -iE '^Process ID:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        ST=$(grep -m1 -iE '^Start time:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        printf '%s\t%s\t%s\t%s\t%s\t%s\n' \
            "${ST:-?}" "${U:-?}" "${P:-?}" "$N" "${H:0:120}" "$F" >> "$META"

        if [[ "$N" -gt 0 ]]; then
            ok "$(basename "$F") — ${BOLD}${N}$(L " righe sospette" " suspicious lines")"
        else
            dim_msg "$(basename "$F")"
        fi
        BODY+=$(file_card_html "$F" "$KW" "PS")
    done

    separator
    info "$(L "Transcript analizzati:" "Transcripts analysed:") ${BOLD}$TOTAL"
    [[ $SUSP -gt 0 ]] && warn "$(L "Con contenuto sospetto:" "With suspicious content:") ${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && return 0

    ask_yn "Generare report HTML?" || return 0

    local SUMTABLE; SUMTABLE=$(_rows_to_table "$(sort "$META")" \
        "$(L "Inizio" "Start")" "$(L "Utente" "User")" "PID" \
        "$(L "Righe sospette" "Suspicious lines")" "Host Application" "$(L "Percorso" "Path")")

    local STATS
    STATS="$(stat_box "Transcript" "$TOTAL")"
    STATS+="$(stat_box "$(L "Sospetti" "Suspicious")" "$SUSP" "$([[ $SUSP -gt 0 ]] && echo warn || echo info)")"
    finish_report "ps_transcripts" "PowerShell Transcript" "PS" "PowerShell_transcript*.txt" "$STATS" \
        "<div class='cards'>$(generic_card_html "$(L "Sessioni" "Sessions")" "$(L "intestazioni dei transcript" "transcript headers")" "$TOTAL" "$SUMTABLE" "▤")</div><div class='cards'>${BODY}</div>"
}

# ================================================================
#  MODULO 42 — LSA Secrets & Cached Domain Credentials (DCC2)
#
#  Il modulo 20 estrae gli hash degli account LOCALI dall'hive SAM. L'hive
#  SECURITY contiene altro, e spesso di piu' valore:
#
#  - LSA Secrets: password IN CHIARO degli account di servizio, delle
#    scheduled task configurate con credenziali, del DefaultPassword di
#    autologon, delle VPN e dei trust di dominio.
#  - Cached Domain Credentials (DCC2/MSCACHE v2): hash degli ultimi utenti di
#    dominio autenticatisi sulla macchina — inclusi amministratori di dominio
#    che si sono collegati a una workstation compromessa.
#
#  Su una workstation membro di dominio questo modulo e' spesso il punto in cui
#  si capisce fin dove l'attaccante poteva arrivare.
# ================================================================
module_lsa_secrets() {
    section_header "LSA Secrets & Cached Domain Credentials" "$RED"
    check_win_root || return 1

    local SEC_HIVE SYS_HIVE
    SEC_HIVE=$(get_hive SECURITY)
    SYS_HIVE=$(get_hive SYSTEM)

    if [[ -z "$SEC_HIVE" ]]; then
        warn "$(L "Hive SECURITY non trovato." "SECURITY hive not found.")"
        return 0
    fi
    if [[ -z "$SYS_HIVE" ]]; then
        warn "$(L "Hive SYSTEM non trovato: serve per la bootkey." "SYSTEM hive not found: required for the boot key.")"
        return 0
    fi
    info "SECURITY: ${DIM}${SEC_HIVE}${RESET}"
    info "SYSTEM:   ${DIM}${SYS_HIVE}${RESET}"

    if ! "$PY3" -c "import impacket" 2>/dev/null; then
        warn "$(L "impacket non disponibile: LSA Secrets e DCC2 non estraibili." \
                 "impacket unavailable: LSA Secrets and DCC2 cannot be extracted.")"
        info "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install impacket"
        return 0
    fi

    info "$(L "Estrazione in corso (impacket)..." "Extracting (impacket)...")"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$SYS_HIVE" "$SEC_HIVE" > "$OUT" << 'PYEOF' 2>/dev/null
import sys, io, os, contextlib

system_hive, security_hive = sys.argv[1], sys.argv[2]

def L(it, en):
    return it if os.environ.get('LANG', 'en') == 'it' else en

try:
    from impacket.examples.secretsdump import LocalOperations, LSASecrets
except Exception as exc:
    print(f"ERROR\t{L('impacket non importabile', 'impacket not importable')}: {exc}")
    sys.exit(0)

try:
    ops = LocalOperations(system_hive)
    boot_key = ops.getBootKey()
except Exception as exc:
    print(f"ERROR\t{L('bootkey non ricavabile dall hive SYSTEM', 'boot key not derivable from SYSTEM hive')}: {exc}")
    sys.exit(0)

# impacket scrive i risultati su stdout: li si cattura per classificarli
# invece di lasciarli finire grezzi nel report.
buf = io.StringIO()
try:
    lsa = LSASecrets(security_hive, boot_key, None, isRemote=False, perSecretCallback=lambda *a: None)
    with contextlib.redirect_stdout(buf):
        try:
            lsa.dumpCachedHashes()
        except Exception:
            pass
        try:
            lsa.dumpSecrets()
        except Exception:
            pass
except Exception as exc:
    print(f"ERROR\t{L('estrazione fallita', 'extraction failed')}: {exc}")
    sys.exit(0)

captured = buf.getvalue().splitlines()

# I DCC2 hanno forma  utente/DOMINIO:$DCC2$10240#utente#hash
# Gli LSA secret sono blocchi "NOME\n(valore)".
for line in captured:
    line = line.rstrip()
    if not line.strip():
        continue
    low = line.lower()
    if '$dcc2$' in low:
        kind = 'DCC2'
    elif line.startswith('[*]') or line.startswith('[-]'):
        continue
    else:
        kind = 'LSA'
    print(f"{kind}\t{line}")
PYEOF

    if grep -q '^ERROR' "$OUT" 2>/dev/null; then
        warn "$(sed -n 's/^ERROR\t//p' "$OUT" | head -1)"
        return 0
    fi

    local NDCC NLSA TOTAL
    NDCC=$(awk -F'\t' '$1=="DCC2"' "$OUT" | wc -l)
    NLSA=$(awk -F'\t' '$1=="LSA"' "$OUT" | wc -l)
    TOTAL=$((NDCC + NLSA))
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun segreto estratto (hive vuoto o non decifrabile)." "No secret extracted (empty or undecryptable hive).")"
        return 0
    fi

    ok "$(L "Cached domain credentials (DCC2):" "Cached domain credentials (DCC2):") ${BOLD}$NDCC"
    ok "$(L "LSA secrets:" "LSA secrets:") ${BOLD}$NLSA"
    [[ "$NDCC" -gt 0 ]] && warn "$(L "I DCC2 sono craccabili offline: hashcat -m 2100" "DCC2 are crackable offline: hashcat -m 2100")"

    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Materiale altamente sensibile" "Highly sensitive material")</b><br>"
    BODY+="$(L "Questo report puo' contenere password in chiaro di account di servizio e hash di credenziali di dominio. Trattalo come materiale riservato del caso: non allegarlo a ticket, non condividerlo su canali non cifrati." \
        "This report may contain cleartext service-account passwords and domain credential hashes. Treat it as restricted case material: do not attach it to tickets or share it over unencrypted channels.")<br><br>"
    BODY+="$(L "Cracking offline:" "Offline cracking:") <code>hashcat -m 2100 dcc2.txt wordlist.txt</code>"
    BODY+="</div></div>"

    if [[ "$NDCC" -gt 0 ]]; then
        local R; R=$(awk -F'\t' '$1=="DCC2"{print $2}' "$OUT")
        BODY+="<div class='cards'>$(generic_card_html "Cached Domain Credentials (DCC2)" "$SEC_HIVE" "$NDCC" "$(_rows_to_table "$R" "$(L "Voce" "Entry")")" "⚿")</div>"
    fi
    if [[ "$NLSA" -gt 0 ]]; then
        local R2; R2=$(awk -F'\t' '$1=="LSA"{print $2}' "$OUT")
        BODY+="<div class='cards'>$(generic_card_html "LSA Secrets" "$SEC_HIVE" "$NLSA" "$(_rows_to_table "$R2" "$(L "Voce" "Entry")")" "⚿")</div>"
    fi

    local STATS
    STATS="$(stat_box "DCC2" "$NDCC" "$([[ "$NDCC" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "LSA Secrets" "$NLSA" "$([[ "$NLSA" -gt 0 ]] && echo warn || echo info)")"
    finish_report "lsa_secrets" "LSA Secrets & DCC2" "LSA" "Windows/System32/config/SECURITY" "$STATS" "$BODY"
}

# ================================================================
#  MODULO 43 — Volume Shadow Copies
#
#  Le shadow copy sono fotografie precedenti del volume. Valgono molto in DFIR
#  per due ragioni:
#
#  - contengono file che l'attaccante ha poi cancellato o modificato, e hive di
#    registro anteriori alla compromissione;
#  - permettono l'analisi DIFFERENZIALE: cosa e' cambiato fra due istanti.
#
#  Sono anche un bersaglio: la loro cancellazione (vssadmin delete shadows) e'
#  un passo standard del ransomware, quindi l'ASSENZA di shadow copy su una
#  macchina che dovrebbe averne e' essa stessa un indicatore.
#
#  Nota sui limiti: montare una shadow copy richiede libvshadow (vshadowmount),
#  privilegi e un mount point. Questo modulo inventaria e prepara i comandi;
#  non monta nulla da solo, perche' un mount silenzioso su una postazione
#  forense non e' un effetto collaterale accettabile.
# ================================================================
module_vss() {
    section_header "Volume Shadow Copies" "$CYAN"
    check_win_root || return 1

    local SVI; SVI=$(ci_find_dir "$WIN_ROOT" "System Volume Information")
    local -a STORES=()
    local F

    if [[ -n "$SVI" ]]; then
        # I contenitori delle shadow copy hanno il GUID del provider VSS nel nome.
        while IFS= read -r F; do
            [[ -n "$F" ]] && STORES+=("$F")
        done < <(find "$SVI" -maxdepth 1 -type f -iname '*3808876b-c176-4e48-b7ae-04046e6cc752*' 2>/dev/null)
    fi

    local NSTORE=${#STORES[@]}
    local ROWS="" TOTBYTES=0
    for F in "${STORES[@]}"; do
        local SZ MT
        SZ=$(stat -c %s "$F" 2>/dev/null || echo 0)
        MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1 || echo "?")
        TOTBYTES=$((TOTBYTES + SZ))
        ROWS+="${MT}	$(basename "$F")	${SZ}
"
    done

    # vshadowinfo legge il catalogo VSS, ma opera sul DEVICE o sull'immagine
    # raw, non su un filesystem gia' montato: da qui si puo' solo segnalarne la
    # disponibilita' e preparare il comando, non eseguirlo.
    local HAVE_VSHADOW=false
    command -v vshadowinfo > /dev/null 2>&1 && HAVE_VSHADOW=true

    separator
    if [[ "$NSTORE" -eq 0 ]]; then
        warn "$(L "Nessun contenitore di shadow copy trovato in System Volume Information." \
                 "No shadow copy container found in System Volume Information.")"
        echo ""
        warn "$(L "L'assenza non e' neutra: la cancellazione delle shadow copy (vssadmin delete shadows) e' un passo tipico del ransomware. Verifica gli Event ID 524/8224 nel modulo Event Log e la presenza di vssadmin/wmic nelle history." \
                 "Absence is not neutral: deleting shadow copies (vssadmin delete shadows) is a standard ransomware step. Check Event IDs 524/8224 in the Event Log module and vssadmin/wmic usage in command histories.")"
    else
        ok "$(L "Contenitori shadow copy:" "Shadow copy containers:") ${BOLD}$NSTORE"
        info "$(L "Spazio occupato:" "Space used:") ${BOLD}$(numfmt --to=iec "$TOTBYTES" 2>/dev/null || echo "$TOTBYTES B")"
        echo ""
        info "$(L "Le date reali degli snapshot stanno nel catalogo VSS, leggibile solo dal device o dall'immagine raw:" \
                 "Real snapshot dates live in the VSS catalogue, readable only from the device or raw image:")"
        if $HAVE_VSHADOW; then
            echo -e "      ${CYAN}vshadowinfo <device|immagine.raw>${RESET}"
        else
            info "$(L "vshadowinfo non installato (pacchetto libvshadow-utils)." \
                     "vshadowinfo not installed (libvshadow-utils package).")"
        fi
    fi

    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    if [[ "$NSTORE" -eq 0 ]]; then
        BODY+="<b>$(L "Nessuna shadow copy presente" "No shadow copy present")</b><br>"
        BODY+="$(L "L'assenza va interpretata, non archiviata: la cancellazione delle shadow copy e' un passo standard del ransomware prima della cifratura. Da correlare con gli Event ID 524 e 8224 e con l'uso di vssadmin, wmic shadowcopy o Win32_ShadowCopy nelle history dei comandi." \
            "Absence needs interpreting, not filing: deleting shadow copies is a standard ransomware step before encryption. Correlate with Event IDs 524 and 8224 and with use of vssadmin, wmic shadowcopy or Win32_ShadowCopy in command histories.")"
    else
        BODY+="<b>$(L "Come sfruttarle" "How to use them")</b><br>"
        BODY+="$(L "Ogni shadow copy e' una versione precedente del volume: contiene file poi cancellati e hive di registro anteriori alla compromissione. Il vero valore e' il confronto differenziale fra due istanti." \
            "Each shadow copy is an earlier version of the volume: it holds files later deleted and registry hives predating the compromise. The real value is the differential comparison between two points in time.")<br><br>"
        BODY+="<b>$(L "Procedura consigliata" "Suggested procedure")</b><br>"
        BODY+="<code>vshadowinfo /dev/sdX1</code> — $(L "elenca gli snapshot e le loro date" "list snapshots and their dates")<br>"
        BODY+="<code>vshadowmount /dev/sdX1 /mnt/vss</code> — $(L "espone gli snapshot come file vssN" "expose snapshots as vssN files")<br>"
        BODY+="<code>mount -o ro,loop /mnt/vss/vss1 /mnt/snap1</code><br>"
        BODY+="<code>./fiuto.sh /mnt/snap1 --all</code> — $(L "rilancia FIUTO sullo snapshot" "re-run FIUTO on the snapshot")<br><br>"
        BODY+="$(L "Confrontando i report di due snapshot si isola cosa e' cambiato nella finestra di compromissione." \
            "Comparing the reports of two snapshots isolates what changed during the compromise window.")"
    fi
    BODY+="</div></div>"

    if [[ "$NSTORE" -gt 0 ]]; then
        local TABLE; TABLE=$(_rows_to_table "$ROWS" \
            "$(L "Ultima modifica" "Last modified")" "$(L "Contenitore" "Container")" "$(L "Byte" "Bytes")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Contenitori shadow copy" "Shadow copy containers")" "${SVI:-System Volume Information}" "$NSTORE" "$TABLE" "◫")</div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Contenitori" "Containers")" "$NSTORE" "$([[ "$NSTORE" -eq 0 ]] && echo warn || echo info)")"
    [[ "$NSTORE" -gt 0 ]] && STATS+="$(stat_box "$(L "Spazio" "Space")" "$(numfmt --to=iec "$TOTBYTES" 2>/dev/null || echo "$TOTBYTES")" "info")"
    finish_report "vss" "Volume Shadow Copies" "VSS" "System Volume Information" "$STATS" "$BODY"
}

# ================================================================
#  MODULO 44 — Outlook PST / OST
#
#  FIUTO non aveva alcuna copertura della posta locale, che in un caso di
#  phishing, BEC o esfiltrazione e' spesso il punto di partenza: il messaggio
#  di ingresso, l'allegato eseguito, le regole di inoltro create
#  dall'attaccante e la corrispondenza esfiltrata stanno tutti qui.
#
#  Un OST conserva anche messaggi cancellati lato server ma non ancora
#  sincronizzati: a volte e' l'unica copia rimasta.
#
#  Richiede pypff (libpff). Senza, il modulo si limita all'inventario dei file
#  con dimensioni e date, che e' comunque utile per l'acquisizione.
# ================================================================
module_pst_ost() {
    section_header "Outlook PST / OST" "$YELLOW"
    check_win_root || return 1

    local -a STORES=()
    local F HOME_DIR
    while IFS= read -r HOME_DIR; do
        while IFS= read -r F; do
            [[ -n "$F" ]] && STORES+=("$F")
        done < <(find "$HOME_DIR" -maxdepth 6 -type f \( -iname '*.pst' -o -iname '*.ost' \) 2>/dev/null)
    done < <(get_user_homes)

    if [[ ${#STORES[@]} -eq 0 ]]; then
        warn "$(L "Nessun archivio PST/OST trovato." "No PST/OST store found.")"
        return 0
    fi
    ok "$(L "Archivi trovati:" "Stores found:") ${BOLD}${#STORES[@]}"

    # Inventario: utile anche senza pypff (dimensioni, date, hash per la catena
    # di custodia, stima dei tempi di elaborazione).
    local INV="" TOTBYTES=0
    for F in "${STORES[@]}"; do
        local SZ MT
        SZ=$(stat -c %s "$F" 2>/dev/null || echo 0)
        MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1 || echo "?")
        TOTBYTES=$((TOTBYTES + SZ))
        info "$(basename "$F") — ${BOLD}$(numfmt --to=iec "$SZ" 2>/dev/null || echo "$SZ")${RESET} ${DIM}${MT}${RESET}"
        INV+="${MT}	$(basename "$F")	${SZ}	${F}
"
    done

    local HAVE_PYPFF=false
    "$PY3" -c "import pypff" 2>/dev/null && HAVE_PYPFF=true

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local NMSG=0 NSUSP=0
    if $HAVE_PYPFF; then
        info "$(L "Estrazione messaggi con pypff (puo' richiedere tempo)..." "Extracting messages with pypff (may take a while)...")"
        local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
        printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
        "$PY3" - "$OUT" "$IOCTMP" "${STORES[@]}" << 'PYEOF' 2>/dev/null
import sys, os, datetime

out_path, ioc_path = sys.argv[1], sys.argv[2]
stores = sys.argv[3:]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

import pypff

# Indicatori tipici di phishing/BEC negli header e negli allegati.
SUSP_EXT = ('.exe', '.scr', '.js', '.vbs', '.jse', '.wsf', '.hta', '.lnk',
            '.iso', '.img', '.7z', '.ace', '.docm', '.xlsm', '.pptm', '.jar', '.ps1')

rows = []

def walk(folder, path, src):
    try:
        subs = folder.number_of_sub_folders
    except Exception:
        subs = 0
    for i in range(subs):
        try:
            sub = folder.get_sub_folder(i)
        except Exception:
            continue
        name = ''
        try:
            name = sub.name or ''
        except Exception:
            pass
        walk(sub, f"{path}/{name}", src)
    try:
        n = folder.number_of_sub_messages
    except Exception:
        return
    for i in range(n):
        try:
            msg = folder.get_sub_message(i)
        except Exception:
            continue
        def g(attr):
            try:
                v = getattr(msg, attr)
                return v if isinstance(v, str) else (v or '')
            except Exception:
                return ''
        subject = g('subject')
        sender = g('sender_name')
        try:
            when = msg.get_delivery_time()
            when = when.strftime('%Y-%m-%d %H:%M:%S') if when else ''
        except Exception:
            when = ''
        # Allegati
        atts = []
        try:
            for a in range(msg.number_of_attachments):
                try:
                    att = msg.get_attachment(a)
                    an = ''
                    try:
                        an = att.get_name() or ''
                    except Exception:
                        pass
                    atts.append(an)
                except Exception:
                    continue
        except Exception:
            pass
        att_s = '; '.join(x for x in atts if x)
        flags = []
        low_att = att_s.lower()
        if any(low_att.endswith(e) or (e + ';') in low_att or (e + ' ') in low_att for e in SUSP_EXT):
            flags.append('ALLEGATO_RISCHIOSO')
        blob = f"{subject} {sender} {att_s}".lower()
        if iocs and any(i in blob for i in iocs):
            flags.append('IOC')
        rows.append((when, path, sender[:120], subject[:200], att_s[:200], ';'.join(flags), src))

for sp in stores:
    src = os.path.basename(sp)
    try:
        pff = pypff.file()
        pff.open(sp)
        root = pff.get_root_folder()
    except Exception:
        continue
    try:
        walk(root, '', src)
    except Exception:
        pass
    try:
        pff.close()
    except Exception:
        pass

rows.sort(key=lambda r: r[0])
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
PYEOF
        [[ -s "$OUT" ]] && NMSG=$(wc -l < "$OUT")
        [[ "$NMSG" -gt 0 ]] && NSUSP=$(awk -F'\t' '$6!=""' "$OUT" | wc -l)
        ok "$(L "Messaggi estratti:" "Messages extracted:") ${BOLD}$NMSG"
        [[ "$NSUSP" -gt 0 ]] && warn "$(L "Messaggi con allegato rischioso o match IoC:" "Messages with risky attachment or IoC match:") ${BOLD}$NSUSP"
    else
        warn "$(L "pypff non disponibile: estratto solo l'inventario degli archivi." \
                 "pypff unavailable: only the store inventory was extracted.")"
        info "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install libpff-python"
    fi

    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    local INVTABLE; INVTABLE=$(_rows_to_table "$INV" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Archivio" "Store")" "$(L "Byte" "Bytes")" "$(L "Percorso" "Path")")
    BODY+="<div class='cards'>$(generic_card_html "$(L "Archivi di posta" "Mail stores")" "$(L "inventario" "inventory")" "${#STORES[@]}" "$INVTABLE" "✉")</div>"

    if [[ "$NMSG" -gt 0 ]]; then
        local ROWS; ROWS=$( { awk -F'\t' '$6!=""{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6}' "$OUT";
                              awk -F'\t' '$6==""{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6}' "$OUT"; } | head -20000 )
        local TABLE; TABLE=$(_rows_to_table "$ROWS" \
            "$(L "Data" "Date")" "$(L "Cartella" "Folder")" "$(L "Mittente" "Sender")" \
            "$(L "Oggetto" "Subject")" "$(L "Allegati" "Attachments")" "$(L "Segnalazioni" "Flags")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Messaggi" "Messages")" "$(L "sospetti in testa" "suspicious first")" "$NMSG" "$TABLE" "✉")</div>"
    else
        BODY+="<div class='card'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
        BODY+="$(L "Contenuto dei messaggi non estratto: manca pypff (libpff). Gli archivi sono comunque inventariati sopra e possono essere analizzati a parte con readpst o pffexport." \
            "Message content not extracted: pypff (libpff) is missing. The stores are inventoried above and can be analysed separately with readpst or pffexport.")<br><br>"
        BODY+="<code>${PY3} -m pip install libpff-python</code><br>"
        BODY+="<code>pffexport -q archivio.pst</code>"
        BODY+="</div></div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Archivi" "Stores")" "${#STORES[@]}")"
    STATS+="$(stat_box "$(L "Dimensione" "Size")" "$(numfmt --to=iec "$TOTBYTES" 2>/dev/null || echo "$TOTBYTES")" "info")"
    if [[ "$NMSG" -gt 0 ]]; then
        STATS+="$(stat_box "$(L "Messaggi" "Messages")" "$NMSG" "info")"
        STATS+="$(stat_box "$(L "Sospetti" "Suspicious")" "$NSUSP" "$([[ "$NSUSP" -gt 0 ]] && echo warn || echo info)")"
    fi
    finish_report "pst_ost" "Outlook PST / OST" "PST" "*.pst · *.ost" "$STATS" "$BODY"
}

# ================================================================
#  MODULO 45 — Cloud Sync (OneDrive / Dropbox / Google Drive / Box)
#
#  L'esfiltrazione moderna raramente passa da una chiavetta: passa da una
#  cartella sincronizzata. Un file trascinato in OneDrive esce dal perimetro
#  senza generare traffico riconoscibile come esfiltrazione e senza toccare
#  nessuno degli artefatti USB.
#
#  Il modulo ricostruisce quali provider erano configurati, con quali account,
#  e — dove il formato lo consente — l'elenco dei file sincronizzati.
#
#  Nota sui formati: il database di Google Drive e' SQLite e si legge per
#  intero. I log .odl di OneDrive sono binari con stringhe offuscate: qui se ne
#  estraggono solo i nomi di file leggibili, e il report lo dichiara.
# ================================================================
module_cloud_sync() {
    section_header "Cloud Sync — $(L "Provider e file sincronizzati" "Providers and synced files")" "$BLUE"
    check_win_root || return 1

    local BODY="" NPROV=0 NFILES=0 NACC=0
    local ACCROWS="" FILEROWS=""
    local HOME_DIR

    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local LOCAL ROAM
        LOCAL=$(ci_find_dir "$HOME_DIR" "AppData/Local")
        ROAM=$(ci_find_dir "$HOME_DIR" "AppData/Roaming")

        # ---------- OneDrive ----------
        local OD; OD=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "Microsoft/OneDrive")
        if [[ -n "$OD" ]]; then
            NPROV=$((NPROV + 1))
            # settings/<Personal|Business1>/*.ini contiene account e cartella locale
            local INI
            while IFS= read -r INI; do
                [[ -s "$INI" ]] || continue
                local EMAIL FOLDER
                EMAIL=$(grep -aoE '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' "$INI" 2>/dev/null | head -1)
                FOLDER=$(grep -aoE '[A-Z]:\\[^"|]{3,120}' "$INI" 2>/dev/null | head -1)
                [[ -z "$EMAIL$FOLDER" ]] && continue
                NACC=$((NACC + 1))
                ACCROWS+="OneDrive	${U}	${EMAIL:-?}	${FOLDER:-?}	$(basename "$INI")
"
            done < <(find "$OD" -maxdepth 3 -type f \( -iname '*.ini' -o -iname 'global.ini' \) 2>/dev/null)

            # Log .odl: estrazione euristica dei nomi di file
            local ODLDIR; ODLDIR=$(ci_find_dir "$OD" "logs")
            if [[ -n "$ODLDIR" ]]; then
                local NODL; NODL=$(find "$ODLDIR" -type f \( -iname '*.odl' -o -iname '*.odlgz' -o -iname '*.aodl' \) 2>/dev/null | wc -l)
                if [[ "$NODL" -gt 0 ]]; then
                    info "OneDrive ($U) — ${BOLD}${NODL}${RESET} $(L "file di log" "log files")"
                    local NAMES
                    NAMES=$(find "$ODLDIR" -type f \( -iname '*.odl' -o -iname '*.aodl' \) -print0 2>/dev/null \
                            | xargs -0 strings -n 6 2>/dev/null \
                            | grep -aoE '[A-Za-z0-9 _().+-]{3,80}\.(docx?|xlsx?|pptx?|pdf|zip|7z|rar|txt|csv|jpg|jpeg|png|eml|msg|pst|ost|key|pem|sql|bak|exe|dll|ps1|vbs|js)' \
                            | sort -u | head -3000)
                    local N; N=$(printf '%s\n' "$NAMES" | grep -c . || true)
                    if [[ "$N" -gt 0 ]]; then
                        NFILES=$((NFILES + N))
                        while IFS= read -r FN; do
                            [[ -n "$FN" ]] && FILEROWS+="OneDrive	${U}	${FN}
"
                        done <<< "$NAMES"
                    fi
                fi
            fi
        fi

        # ---------- Dropbox ----------
        local DBX; DBX=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "Dropbox")
        local DBXR; DBXR=$(ci_find_dir "${ROAM:-$HOME_DIR}" "Dropbox")
        if [[ -n "$DBX" || -n "$DBXR" ]]; then
            NPROV=$((NPROV + 1))
            local IJ; IJ=$(ci_find_file "${DBX:-$DBXR}" "info.json")
            if [[ -s "$IJ" ]]; then
                local PATHS
                PATHS=$("$PY3" -c "
import json,sys
try:
    d=json.load(open(sys.argv[1],encoding='utf-8',errors='replace'))
    for k,v in d.items():
        print(f\"{k}\t{v.get('path','')}\t{v.get('subscription_type','')}\")
except Exception:
    pass" "$IJ" 2>/dev/null)
                while IFS=$'\t' read -r KIND P SUB; do
                    [[ -z "$KIND" ]] && continue
                    NACC=$((NACC + 1))
                    ACCROWS+="Dropbox	${U}	${KIND} ${SUB}	${P}	info.json
"
                done <<< "$PATHS"
            fi
        fi

        # ---------- Google Drive ----------
        local GD; GD=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "Google/DriveFS")
        if [[ -n "$GD" ]]; then
            NPROV=$((NPROV + 1))
            local MDB
            while IFS= read -r MDB; do
                [[ -s "$MDB" ]] || continue
                # Schema DriveFS: items(stable_id, ..., local_title, modified_date, size)
                local ROWS
                ROWS=$(query_sqlite "$MDB" "SELECT local_title, datetime(modified_date,'unixepoch'), file_size FROM items WHERE local_title IS NOT NULL ORDER BY modified_date DESC LIMIT 20000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                local N; N=$(printf '%s\n' "$ROWS" | grep -c . || true)
                NFILES=$((NFILES + N))
                info "Google Drive ($U) — ${BOLD}${N}${RESET} $(L "elementi" "items")"
                while IFS=$'\t' read -r T D S; do
                    [[ -n "$T" ]] && FILEROWS+="Google Drive	${U}	${T}  (${D:-?}, ${S:-?} B)
"
                done <<< "$ROWS"
            done < <(find "$GD" -maxdepth 3 -type f -name 'metadata_sqlite_db' 2>/dev/null)
        fi

        # ---------- Box / iCloud: solo presenza ----------
        local P
        for P in "Box/Box" "Apple/CloudDocs" "iCloudDrive"; do
            local X; X=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "$P")
            [[ -n "$X" ]] && { NPROV=$((NPROV + 1)); ACCROWS+="${P%%/*}	${U}	-	${X}	$(L "solo presenza" "presence only")
"; }
        done
    done < <(get_user_homes)

    separator
    if [[ "$NPROV" -eq 0 ]]; then
        warn "$(L "Nessun provider di sincronizzazione cloud rilevato." "No cloud sync provider detected.")"
        return 0
    fi
    ok "$(L "Provider rilevati:" "Providers detected:") ${BOLD}$NPROV"
    info "$(L "Account/configurazioni:" "Accounts/configurations:") ${BOLD}$NACC"
    info "$(L "Nomi di file recuperati:" "File names recovered:") ${BOLD}$NFILES"

    ask_yn "Generare report HTML?" || return 0

    BODY="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Perche' conta" "Why it matters")</b><br>"
    BODY+="$(L "Un file trascinato in una cartella sincronizzata esce dal perimetro senza toccare nessun artefatto USB e senza generare traffico riconoscibile come esfiltrazione. Questi elenchi vanno confrontati con i dati che l'organizzazione considera riservati." \
        "A file dropped into a synced folder leaves the perimeter without touching any USB artefact and without generating traffic recognisable as exfiltration. Cross-check these lists against the data the organisation treats as confidential.")<br><br>"
    BODY+="<b>$(L "Attendibilita' delle fonti" "Source reliability")</b><br>"
    BODY+="$(L "Google Drive: database SQLite, elenco completo e datato. Dropbox: configurazione JSON, account e percorsi. OneDrive: i log .odl sono binari con stringhe offuscate, qui se ne estraggono solo i nomi leggibili — sono indizi, non un inventario completo, e non hanno data." \
        "Google Drive: SQLite database, complete and dated listing. Dropbox: JSON configuration, accounts and paths. OneDrive: .odl logs are binary with obfuscated strings; only readable names are extracted here — these are leads, not a complete inventory, and carry no timestamp.")"
    BODY+="</div></div>"

    if [[ -n "$ACCROWS" ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "Account e configurazioni" "Accounts and configurations")" "AppData" "$NACC" \
            "$(_rows_to_table "$ACCROWS" "Provider" "$(L "Utente Windows" "Windows user")" "Account" "$(L "Cartella / dettaglio" "Folder / detail")" "$(L "Origine" "Source")")" "☁")</div>"
    fi
    if [[ -n "$FILEROWS" ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "File sincronizzati" "Synced files")" "$(L "Google Drive completo · OneDrive euristico" "Google Drive complete · OneDrive heuristic")" "$NFILES" \
            "$(_rows_to_table "$FILEROWS" "Provider" "$(L "Utente Windows" "Windows user")" "$(L "File" "File")")" "⇪")</div>"
    fi

    local STATS
    STATS="$(stat_box "Provider" "$NPROV")"
    STATS+="$(stat_box "Account" "$NACC" "info")"
    STATS+="$(stat_box "$(L "File" "Files")" "$NFILES" "warn")"
    finish_report "cloud_sync" "Cloud Sync" "CLD" "OneDrive · Dropbox · Google Drive · Box" "$STATS" "$BODY"
}

# ================================================================
#  MODULO 46 — BITS Jobs (Background Intelligent Transfer Service)
#
#  BITS e' il servizio che Windows usa per scaricare aggiornamenti in
#  background. Essendo un componente firmato e legittimo, viene usato dagli
#  attaccanti per scaricare payload e per mantenere persistenza: un job BITS
#  con /SetNotifyCmdLine rieseguita un comando a ogni completamento, e il
#  traffico appare come una normale attivita' di sistema (T1197).
#
#  La coda dei job sta in qmgr.db (Windows 10+, formato ESE) o nei vecchi
#  qmgr0.dat/qmgr1.dat. Qui si estraggono URL, percorsi locali e nomi dei job
#  dalle stringhe UTF-16LE: e' l'approccio che funziona su tutti i formati e
#  senza dipendenze, a costo di non ricostruire la struttura dei record.
# ================================================================
module_bits() {
    section_header "BITS Jobs" "$ORANGE"
    check_win_root || return 1

    local -a DBS=()
    local D F
    D=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Network/Downloader")
    [[ -z "$D" ]] && D=$(ci_find_dir "$WIN_ROOT" "Documents and Settings/All Users/Application Data/Microsoft/Network/Downloader")
    if [[ -n "$D" ]]; then
        while IFS= read -r F; do
            [[ -n "$F" ]] && DBS+=("$F")
        done < <(find "$D" -maxdepth 1 -type f \( -iname 'qmgr.db' -o -iname 'qmgr[01].dat' \) 2>/dev/null)
    fi

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna coda BITS trovata (qmgr.db / qmgr0.dat)." "No BITS queue found (qmgr.db / qmgr0.dat).")"
        return 0
    fi
    info "$(L "Code BITS trovate:" "BITS queues found:") ${BOLD}${#DBS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os

out_path = sys.argv[1]
dbs = sys.argv[2:]

# Le stringhe in qmgr sono UTF-16LE. Si estraggono URL, percorsi locali e
# nomi di job, poi si classifica cosa merita attenzione.
URL  = re.compile(r'(?:https?|ftp)://[!-~]{4,400}')
WPATH = re.compile(r'[A-Za-z]:\\[^\x00<>|?*"]{3,250}')

# Domini di aggiornamento legittimi: servono a separare il rumore dal resto,
# non a dichiarare "sicuro" cio' che vi corrisponde.
MS_HOSTS = ('microsoft.com', 'windowsupdate.com', 'msftncsi.com', 'msedge.net',
            'windows.com', 'msn.com', 'live.com', 'office.net', 'office.com',
            'azureedge.net', 'akamaized.net', 'delivery.mp.microsoft.com')

# Estensioni che, scaricate via BITS, sono di per se' un segnale.
RISKY_EXT = ('.exe', '.dll', '.ps1', '.bat', '.cmd', '.scr', '.vbs', '.js',
             '.hta', '.jar', '.zip', '.7z', '.tmp', '.dat')

rows = []
seen = set()

def add(kind, value, src, flags):
    key = (kind, value)
    if key in seen:
        return
    seen.add(key)
    rows.append((kind, value, ';'.join(flags), src))

for db in dbs:
    src = os.path.basename(db)
    try:
        with open(db, 'rb') as fh:
            raw = fh.read(512 * 1024 * 1024)
    except Exception:
        continue
    # UTF-16LE e, per sicurezza, anche ASCII
    for enc in ('utf-16-le', 'latin-1'):
        try:
            text = raw.decode(enc, 'ignore')
        except Exception:
            continue
        for m in URL.finditer(text):
            u = m.group(0).rstrip('\x00').strip()
            low = u.lower()
            flags = []
            if not any(h in low for h in MS_HOSTS):
                flags.append('HOST_NON_MICROSOFT')
            if any(low.split('?')[0].endswith(e) for e in RISKY_EXT):
                flags.append('ESTENSIONE_A_RISCHIO')
            if low.startswith('http://'):
                flags.append('IN_CHIARO')
            add('URL', u[:400], src, flags)
        for m in WPATH.finditer(text):
            p = m.group(0).rstrip('\x00').strip()
            low = p.lower()
            if not any(low.endswith(e) for e in RISKY_EXT) and '\\temp\\' not in low and '\\users\\public\\' not in low:
                continue
            flags = []
            if '\\temp\\' in low or '\\users\\public\\' in low or '\\programdata\\' in low:
                flags.append('PERCORSO_SOSPETTO')
            if any(low.endswith(e) for e in RISKY_EXT):
                flags.append('ESTENSIONE_A_RISCHIO')
            add('PATH', p[:250], src, flags)

# Prima le voci segnalate.
rows.sort(key=lambda r: (r[2] == '', r[0], r[1]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ') for x in r) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun URL o percorso estratto dalla coda BITS." "No URL or path extracted from the BITS queue.")"
        return 0
    fi
    local NFLAG NURL
    NFLAG=$(awk -F'\t' '$3!=""' "$OUT" | wc -l)
    NURL=$(awk -F'\t' '$1=="URL"' "$OUT" | wc -l)

    ok "$(L "Voci estratte:" "Entries extracted:") ${BOLD}$TOTAL"
    info "URL: ${BOLD}${NURL}"
    if [[ "$NFLAG" -gt 0 ]]; then
        warn "$(L "Voci segnalate:" "Flagged entries:") ${BOLD}$NFLAG"
        awk -F'\t' '$3!=""{printf "      [%s] %s\n", $3, substr($2,1,100)}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${ORANGE}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4}' "$OUT" | head -20000)
    local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Tipo" "Type")" "$(L "Valore" "Value")" "$(L "Segnalazioni" "Flags")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Come leggere le segnalazioni" "How to read the flags")</b><br>"
    NOTE+="$(L "HOST_NON_MICROSOFT non significa 'malevolo': molti software legittimi usano BITS. Significa che quella voce non e' spiegabile come Windows Update e va verificata. Un job BITS verso un host non Microsoft che scarica un eseguibile in %TEMP% e' invece un indicatore forte." \
        "HOST_NON_MICROSOFT does not mean 'malicious': plenty of legitimate software uses BITS. It means the entry is not explainable as Windows Update and needs checking. A BITS job to a non-Microsoft host downloading an executable into %TEMP% is a strong indicator instead.")<br><br>"
    NOTE+="$(L "Da correlare con gli Event ID 3, 59, 60 del log Microsoft-Windows-Bits-Client/Operational (modulo Event Log)." \
        "Correlate with Event IDs 3, 59, 60 in Microsoft-Windows-Bits-Client/Operational (Event Log module).")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Voci" "Entries")" "$TOTAL")"
    STATS+="$(stat_box "URL" "$NURL" "info")"
    STATS+="$(stat_box "$(L "Segnalate" "Flagged")" "$NFLAG" "$([[ "$NFLAG" -gt 0 ]] && echo warn || echo info)")"
    finish_report "bits" "BITS Jobs" "BIT" "ProgramData/Microsoft/Network/Downloader" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Coda BITS" "BITS queue")" "${DBS[0]}" "$TOTAL" "$TABLE" "⇩")</div>"
}

# ================================================================
#  MODULO 47 — Thumbcache / IconCache
#
#  Windows conserva le miniature dei file visualizzati in Esplora risorse in
#  database centralizzati. La miniatura SOPRAVVIVE alla cancellazione del file
#  originale: e' spesso l'unica prova visiva residua di un documento o di
#  un'immagine poi eliminata, e non e' toccata dalla pulizia del cestino.
#
#  I formati dei record cambiano fra Vista, 7, 8 e 10/11. Invece di parsare la
#  struttura — fragile fra versioni — qui si esegue il CARVING delle immagini
#  incorporate cercandone le firme. E' robusto su tutte le versioni; in cambio
#  non si ottiene la corrispondenza miniatura-nome file, che nel formato non
#  e' comunque presente in chiaro (la chiave e' un hash del percorso).
#
#  Le immagini estratte vengono scritte accanto al report e mostrate in
#  galleria, cosi' l'analista le sfoglia invece di aprirle una per una.
# ================================================================
module_thumbcache() {
    section_header "Thumbcache / IconCache" "$GREEN"
    check_win_root || return 1

    local -a DBS=()
    local HOME_DIR F
    while IFS= read -r HOME_DIR; do
        local EXPL; EXPL=$(ci_find_dir "$HOME_DIR" "AppData/Local/Microsoft/Windows/Explorer")
        [[ -z "$EXPL" ]] && continue
        while IFS= read -r F; do
            [[ -n "$F" ]] && DBS+=("$F")
        done < <(find "$EXPL" -maxdepth 1 -type f \( -iname 'thumbcache_*.db' -o -iname 'iconcache_*.db' \) -size +1k 2>/dev/null)
    done < <(get_user_homes)

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessun database thumbcache/iconcache trovato." "No thumbcache/iconcache database found.")"
        return 0
    fi
    info "$(L "Database trovati:" "Databases found:") ${BOLD}${#DBS[@]}"

    # Le immagini estratte vanno accanto al report: si prepara la cartella qui,
    # perche' prepare_report_dir ne crea una nuova a ogni chiamata.
    [[ -n "$REPORT_BASE_DIR" && ! -d "$REPORT_BASE_DIR" ]] && mkdir -p "$REPORT_BASE_DIR"
    local RDIR="${REPORT_BASE_DIR}/thumbcache_$(date +%Y%m%d_%H%M%S)"
    mkdir -p "$RDIR/images" 2>/dev/null || {
        err "$(L "Impossibile creare la cartella del report." "Cannot create the report directory.")"
        return 1
    }

    info "$(L "Carving delle miniature in corso..." "Carving thumbnails...")"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "$RDIR/images" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, os, struct

out_path, img_dir = sys.argv[1], sys.argv[2]
dbs = sys.argv[3:]

MAX_IMAGES = 4000          # tetto: una thumbcache puo' contenerne decine di migliaia
MIN_SIZE   = 512           # sotto questa soglia sono icone di sistema, rumore

def carve(data):
    """Genera (offset, estensione, byte) per ogni immagine incorporata."""
    n = len(data)
    i = 0
    while i < n - 8:
        # PNG
        if data[i:i+8] == b'\x89PNG\r\n\x1a\n':
            end = data.find(b'IEND', i)
            if end > 0:
                end += 8
                yield i, 'png', data[i:end]
                i = end
                continue
        # JPEG
        if data[i:i+3] == b'\xff\xd8\xff':
            end = data.find(b'\xff\xd9', i + 3)
            if end > 0:
                end += 2
                yield i, 'jpg', data[i:end]
                i = end
                continue
        # BMP: 'BM' + dimensione dichiarata nell'header
        if data[i:i+2] == b'BM' and i + 6 <= n:
            try:
                size = struct.unpack_from('<I', data, i + 2)[0]
            except Exception:
                size = 0
            if 100 < size < 8 * 1024 * 1024 and i + size <= n:
                yield i, 'bmp', data[i:i+size]
                i += size
                continue
        i += 1

rows = []
count = 0
for db in dbs:
    base = os.path.basename(db)
    try:
        with open(db, 'rb') as fh:
            data = fh.read()
    except Exception:
        continue

    # L'intestazione dichiara versione e tipo di cache (la dimensione delle
    # miniature): utile per sapere quale cache si sta guardando.
    version = cache_type = ''
    if data[:4] == b'CMMM':
        try:
            version, cache_type = struct.unpack_from('<II', data, 4)
        except Exception:
            pass

    for off, ext, blob in carve(data):
        if len(blob) < MIN_SIZE:
            continue
        if count >= MAX_IMAGES:
            break
        name = f"{base.replace('.db','')}_{off:08x}.{ext}"
        try:
            with open(os.path.join(img_dir, name), 'wb') as out:
                out.write(blob)
        except Exception:
            continue
        count += 1
        rows.append((base, str(version), str(cache_type), ext, str(len(blob)), name))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(r) + '\n')
PYEOF

    local N=0
    [[ -s "$OUT" ]] && N=$(wc -l < "$OUT")
    if [[ "$N" -eq 0 ]]; then
        warn "$(L "Nessuna miniatura estratta." "No thumbnail extracted.")"
        rmdir "$RDIR/images" "$RDIR" 2>/dev/null
        return 0
    fi
    ok "$(L "Miniature estratte:" "Thumbnails extracted:") ${BOLD}$N"
    info "$(L "Salvate in:" "Saved to:") ${DIM}${RDIR}/images${RESET}"

    ask_yn "Generare report HTML?" || return 0

    # Galleria: le miniature sono l'informazione, non la tabella.
    local GAL="" ROW
    local SHOWN=0
    while IFS=$'\t' read -r SRC VER CT EXT SZ NAME; do
        [[ -z "$NAME" ]] && continue
        SHOWN=$((SHOWN + 1))
        [[ $SHOWN -gt 1500 ]] && break
        GAL+="<figure style='margin:0;text-align:center'><img src='images/$(html_attr "$NAME")' loading='lazy' style='max-width:120px;max-height:120px;border:1px solid var(--border);border-radius:4px;background:#0d1117'><figcaption style='font-family:var(--mono);font-size:.55rem;color:var(--text-dim);word-break:break-all'>$(html_esc "${SRC}")</figcaption></figure>"
    done < "$OUT"

    local SUM; SUM=$(awk -F'\t' '{c[$1"\t"$2"\t"$3]++} END{for(k in c) print k"\t"c[k]}' "$OUT" | sort)
    local SUMTABLE; SUMTABLE=$(_rows_to_table "$SUM" "Database" "$(L "Versione" "Version")" "$(L "Tipo cache" "Cache type")" "$(L "Miniature" "Thumbnails")")

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Cosa sono e cosa non sono" "What these are and are not")</b><br>"
    BODY+="$(L "Ogni miniatura e' la prova che un file e' stato visualizzato in Esplora risorse: sopravvive alla cancellazione dell'originale e allo svuotamento del cestino. NON e' pero' possibile risalire al nome del file: nel formato la chiave e' un hash del percorso, non il percorso stesso. La miniatura prova l'esistenza e il contenuto, non l'ubicazione." \
        "Each thumbnail proves a file was viewed in Explorer: it survives deletion of the original and emptying the Recycle Bin. It is NOT possible to recover the file name: the format keys entries by a hash of the path, not the path itself. A thumbnail proves existence and content, not location.")<br><br>"
    BODY+="$(L "Estrazione per carving delle firme immagine, non per parsing dei record: robusta su tutte le versioni di Windows." \
        "Extraction is by image-signature carving, not record parsing: robust across all Windows versions.")"
    BODY+="</div></div>"
    BODY+="<div class='cards'>$(generic_card_html "$(L "Database analizzati" "Databases analysed")" "AppData/Local/Microsoft/Windows/Explorer" "${#DBS[@]}" "$SUMTABLE" "▦")</div>"
    BODY+="<div class='stitle'>$(L "Galleria miniature" "Thumbnail gallery")</div>"
    BODY+="<div class='card'><div style='padding:1rem;display:flex;flex-wrap:wrap;gap:.6rem'>${GAL}</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Miniature" "Thumbnails")" "$N")"
    STATS+="$(stat_box "Database" "${#DBS[@]}" "info")"
    [[ "$N" -gt 1500 ]] && STATS+="$(stat_box "$(L "In galleria" "In gallery")" "1500" "info")"

    # Il report va scritto nella cartella gia' creata, accanto alle immagini.
    local REPORT_HTML="${RDIR}/report.html"
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    {
        html_header "Thumbcache"
        html_page_header "THU" "Thumbcache / IconCache" "thumbcache_*.db" "$SCAN" "$WIN_ROOT"
        printf "<div class='statsbar'>%s</div>\n" "$STATS"
        echo "<main>"
        pre_style_block
        printf '%s\n' "$BODY"
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# ================================================================
#  MODULO 48 — Chat Desktop (Slack / Teams / Discord)
#
#  Le app di messaggistica aziendale sono Electron e conservano i messaggi in
#  LevelDB, esattamente come ChatGPT Desktop (modulo 39). In un'indagine
#  contano per due ragioni:
#
#  - social engineering interno: il messaggio che ha convinto la vittima ad
#    aprire l'allegato o ad autorizzare un pagamento spesso arriva da qui,
#    non dalla posta;
#  - esfiltrazione: file e credenziali condivisi in chat privata non passano
#    dal gateway di posta e non lasciano traccia negli artefatti USB.
#
#  Il carving riusa la libreria condivisa pylib_leveldb: le stringhe
#  recuperate sono frammenti di conversazione, non un export strutturato —
#  mancano interlocutori e timestamp, e il report lo dichiara.
# ================================================================
module_chat_desktop() {
    section_header "Chat Desktop — Slack / Teams / Discord" "$MAGENTA"
    check_win_root || return 1

    local MANIFEST; MANIFEST=$(mktemp); register_tmp "$MANIFEST"
    local HOME_DIR NAPP=0
    local -a FOUND_APPS=()

    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local ROAM LOCAL
        ROAM=$(ci_find_dir "$HOME_DIR" "AppData/Roaming")
        LOCAL=$(ci_find_dir "$HOME_DIR" "AppData/Local")

        # app|percorso relativo alla radice indicata
        local SPEC APP REL BASE DIR N
        for SPEC in "Slack|Slack|roam" "Discord|discord|roam" "Teams|Microsoft/Teams|roam" \
                    "Teams (new)|Packages|local"; do
            IFS='|' read -r APP REL BASE <<< "$SPEC"
            local ROOT; [[ "$BASE" == "roam" ]] && ROOT="$ROAM" || ROOT="$LOCAL"
            [[ -z "$ROOT" ]] && continue
            local APPDIR; APPDIR=$(ci_find_dir "$ROOT" "$REL")
            [[ -z "$APPDIR" ]] && continue

            # Teams "new" sta dentro Packages/MSTeams_*: si restringe la ricerca.
            if [[ "$APP" == "Teams (new)" ]]; then
                APPDIR=$(find "$ROOT" -maxdepth 1 -type d -iname 'MSTeams_*' 2>/dev/null | head -1)
                [[ -z "$APPDIR" ]] && continue
            fi

            N=0
            while IFS= read -r F; do
                [[ -s "$F" ]] || continue
                printf '%s\t%s\t%s\n' "$APP" "$U" "$F" >> "$MANIFEST"
                N=$((N + 1))
            done < <(find "$APPDIR" -maxdepth 8 -type f \( -iname '*.ldb' -o -iname '*.log' \) \
                          \( -ipath '*Local Storage*' -o -ipath '*IndexedDB*' -o -ipath '*leveldb*' \) 2>/dev/null)
            if [[ "$N" -gt 0 ]]; then
                NAPP=$((NAPP + 1))
                FOUND_APPS+=("$APP ($U): $N")
                ok "$APP — $U: ${BOLD}${N}${RESET} $(L "file LevelDB" "LevelDB files")"
            fi
        done
    done < <(get_user_homes)

    if [[ ! -s "$MANIFEST" ]]; then
        warn "$(L "Nessuna app di chat desktop con dati LevelDB trovata." "No desktop chat app with LevelDB data found.")"
        return 0
    fi
    local NFILES; NFILES=$(wc -l < "$MANIFEST")
    info "$(L "File da analizzare:" "Files to analyse:") ${BOLD}$NFILES"
    info "$(L "Carving in corso (puo' richiedere tempo)..." "Carving (may take a while)...")"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    run_py_with_lib pylib_leveldb "$MANIFEST" "$IOCTMP" "$OUT" << 'PYEOF' 2>/dev/null
import sys, os

manifest, ioc_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

# Termini che, dentro una chat aziendale, meritano di essere portati in cima.
SENSITIVE = ('password', 'passwd', 'credenziali', 'credential', 'token', 'api key',
             'apikey', 'secret', 'iban', 'bonifico', 'wire transfer', 'urgente',
             'urgent', 'riservato', 'confidential', 'vpn', 'rdp', '2fa', 'otp',
             'codice di verifica', 'verification code', 'invoice', 'fattura')

rows = []
seen = set()
for line in open(manifest, encoding='utf-8', errors='replace'):
    parts = line.rstrip('\n').split('\t')
    if len(parts) != 3:
        continue
    app, user, path = parts
    try:
        items = carve_leveldb(path, cap=1200)
    except Exception:
        continue
    src = os.path.basename(path)
    for _, role, text in items:
        if role in ('error', 'info'):
            continue
        t = text.strip()
        if not t:
            continue
        key = (app, user, t.lower()[:90])
        if key in seen:
            continue
        seen.add(key)
        low = t.lower()
        flags = []
        if any(s in low for s in SENSITIVE):
            flags.append('SENSIBILE')
        if iocs and any(i in low for i in iocs):
            flags.append('IOC')
        rows.append((app, user, role, t[:1500], ';'.join(flags), src))

# Prima i messaggi segnalati.
rows.sort(key=lambda r: (r[4] == '', r[0], r[1]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun frammento di conversazione recuperato." "No conversation fragment recovered.")"
        return 0
    fi
    local NFLAG; NFLAG=$(awk -F'\t' '$5!=""' "$OUT" | wc -l)

    ok "$(L "Frammenti recuperati:" "Fragments recovered:") ${BOLD}$TOTAL"
    if [[ "$NFLAG" -gt 0 ]]; then
        warn "$(L "Frammenti segnalati (termini sensibili o IoC):" "Flagged fragments (sensitive terms or IoC):") ${BOLD}$NFLAG"
        awk -F'\t' '$5!=""{printf "      [%s] %s — %s\n", $5, $1, substr($4,1,80)}' "$OUT" | head -15 | while IFS= read -r LN; do
            echo -e "      ${MAGENTA}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6}' "$OUT" | head -20000)
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "App" "$(L "Utente" "User")" "$(L "Ruolo" "Role")" "$(L "Testo" "Text")" \
        "$(L "Segnalazioni" "Flags")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Cosa sono questi frammenti" "What these fragments are")</b><br>"
    NOTE+="$(L "Sono stringhe carvate dai database LevelDB delle app: NON sono un export strutturato della conversazione. Mancano interlocutore, canale e timestamp, l'ordine non e' quello cronologico e i messaggi possono risultare troncati. Servono a stabilire che una conversazione e' avvenuta e cosa conteneva, non a ricostruirne il filo." \
        "These are strings carved from the apps' LevelDB stores: they are NOT a structured conversation export. Counterpart, channel and timestamp are missing, the order is not chronological and messages may be truncated. They establish that a conversation happened and what it contained, not its thread.")<br><br>"
    NOTE+="$(L "Il ruolo indicato e' dedotto dalla presenza di markdown nel testo: e' un'euristica, non un dato del formato." \
        "The role shown is inferred from markdown present in the text: it is a heuristic, not a field of the format.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Frammenti" "Fragments")" "$TOTAL")"
    STATS+="$(stat_box "App" "$NAPP" "info")"
    STATS+="$(stat_box "$(L "Segnalati" "Flagged")" "$NFLAG" "$([[ "$NFLAG" -gt 0 ]] && echo warn || echo info)")"
    finish_report "chat_desktop" "Chat Desktop" "CHT" "Slack · Teams · Discord (LevelDB)" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Frammenti di conversazione" "Conversation fragments")" "$(L "segnalati in testa" "flagged first")" "$TOTAL" "$TABLE" "💬")</div>"
}

# ================================================================
#  MODULO 49 — WebCacheV01.dat (Internet Explorer / Edge Legacy)
#
#  Il modulo 17 legge la cronologia dai database SQLite di Chrome, Edge
#  Chromium e Firefox. Prima di Edge Chromium, pero', Windows registrava la
#  navigazione in WebCacheV01.dat, un database ESE che contiene:
#
#  - la cronologia di IE e Edge Legacy;
#  - i download;
#  - i cookie;
#  - e soprattutto la cronologia di TUTTO cio' che usa WinINET, quindi anche
#    la navigazione fatta da applicazioni, script e malware che si appoggiano
#    alle API di sistema invece che a un browser.
#
#  Quest'ultimo punto e' il motivo per cui il file resta rilevante anche su
#  macchine dove nessuno usa piu' Internet Explorer.
# ================================================================
module_webcache() {
    section_header "WebCacheV01 — IE / Edge Legacy / WinINET" "$CYAN"
    check_win_root || return 1

    local -a DBS=()
    local HOME_DIR F
    while IFS= read -r HOME_DIR; do
        local D; D=$(ci_find_dir "$HOME_DIR" "AppData/Local/Microsoft/Windows/WebCache")
        [[ -z "$D" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && DBS+=("$F")
        done < <(find "$D" -maxdepth 1 -type f -iname 'WebCacheV*.dat' 2>/dev/null)
    done < <(get_user_homes)

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessun WebCacheV01.dat trovato." "No WebCacheV01.dat found.")"
        return 0
    fi
    info "$(L "Database trovati:" "Databases found:") ${BOLD}${#DBS[@]}"

    local HAVE_ESE=false
    "$PY3" -c "import pyesedb" 2>/dev/null && HAVE_ESE=true
    $HAVE_ESE || info "$(L "libesedb (pyesedb) non disponibile: si usera' l'estrazione dalle stringhe." \
                          "libesedb (pyesedb) unavailable: string extraction will be used.")"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    # Il parser effettivamente riuscito viene riportato dallo script.
    local METHOD; METHOD=$("$PY3" - "$OUT" "$($HAVE_ESE && echo 1 || echo 0)" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, os, re, datetime

out_path = sys.argv[1]
have_ese = sys.argv[2] == '1'
dbs = sys.argv[3:]

def filetime(v):
    """FILETIME (100ns dal 1601) -> stringa UTC."""
    try:
        v = int(v)
        if v <= 0:
            return ''
        return (datetime.datetime(1601, 1, 1)
                + datetime.timedelta(microseconds=v / 10)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ''

rows = []
seen = set()

def add(kind, when, url, extra, src):
    key = (kind, url[:150], when)
    if key in seen:
        return
    seen.add(key)
    rows.append((when, kind, url[:500], extra[:200], src))

def parse_ese():
    import pyesedb
    for db in dbs:
        src = os.path.basename(db)
        try:
            f = pyesedb.file()
            f.open(db)
        except Exception:
            continue
        for i in range(f.get_number_of_tables()):
            try:
                t = f.get_table(i)
                name = t.get_name()
            except Exception:
                continue
            # Le tabelle utili sono Container_N (cronologia, cookie, download).
            if not name.startswith('Container_'):
                continue
            try:
                cols = [t.get_column(c).get_name() for c in range(t.get_number_of_columns())]
            except Exception:
                continue
            idx = {c: n for n, c in enumerate(cols)}
            for r in range(t.get_number_of_records()):
                try:
                    rec = t.get_record(r)
                except Exception:
                    continue
                def val(col):
                    n = idx.get(col)
                    if n is None:
                        return ''
                    try:
                        v = rec.get_value_data_as_string(n)
                        return v if v else ''
                    except Exception:
                        try:
                            v = rec.get_value_data_as_integer(n)
                            return str(v) if v is not None else ''
                        except Exception:
                            return ''
                url = val('Url')
                if not url:
                    continue
                # WinINET prefissa le voci con "Visited:", "Cookie:", "iecompat:"...
                kind = 'history'
                low = url.lower()
                if low.startswith('cookie:'):
                    kind = 'cookie'
                elif low.startswith('visited:'):
                    kind = 'visited'
                elif low.startswith('iedownload') or 'download' in name.lower():
                    kind = 'download'
                when = filetime(val('AccessedTime')) or filetime(val('ModifiedTime')) or filetime(val('CreationTime'))
                extra = ' '.join(x for x in (val('Filename'), val('AccessCount') and f"hits={val('AccessCount')}") if x)
                add(kind, when, url, extra, src)
        try:
            f.close()
        except Exception:
            pass


def parse_strings():
    """Gli URL restano in chiaro nel file anche quando l'ESE non e' apribile."""
    URL = re.compile(r'(?:https?|ftp)://[!-~]{4,400}')
    for db in dbs:
        src = os.path.basename(db)
        try:
            with open(db, 'rb') as fh:
                raw = fh.read(1024 * 1024 * 1024)
        except Exception:
            continue
        for enc in ('utf-16-le', 'latin-1'):
            try:
                text = raw.decode(enc, 'ignore')
            except Exception:
                continue
            for m in URL.finditer(text):
                add('string', '', m.group(0).rstrip('\x00').strip(), '', src)

# Un WebCacheV01 acquisito da macchina accesa e' quasi sempre dirty: pyesedb
# non lo apre. Ripiegare sulle stringhe invece di restituire un report vuoto
# e' la differenza fra "nessun dato" e "dati parziali ma reali".
method = 'strings'
if have_ese:
    parse_ese()
    if rows:
        method = 'ESE'
if not rows:
    parse_strings()
print(method if rows else 'nessuno')

rows.sort(key=lambda r: (r[0] == '', r[0]), reverse=False)
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')
PYEOF
    )

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna voce estratta da WebCacheV01." "No entry extracted from WebCacheV01.")"
        return 0
    fi
    ok "$(L "Voci estratte:" "Entries extracted:") ${BOLD}$TOTAL${RESET} ($(L "parser" "parser"): ${METHOD:-strings})"
    if [[ "$METHOD" != "ESE" ]] && $HAVE_ESE; then
        warn "$(L "Database non apribile con libesedb (probabilmente dirty): estrazione dalle stringhe, senza date." \
                 "Database not openable with libesedb (likely dirty): string extraction, without dates.")"
    fi
    if [[ ${#IOC_LIST[@]} -gt 0 ]]; then
        local NIOC=0 LINE
        while IFS= read -r LINE; do
            check_ioc "$LINE" && NIOC=$((NIOC + 1))
        done < "$OUT"
        [[ $NIOC -gt 0 ]] && warn "$(L "Voci con match IoC:" "Entries matching IoC:") ${BOLD}$NIOC"
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(head -30000 "$OUT")
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data (UTC)" "Date (UTC)")" "$(L "Tipo" "Type")" "URL" "$(L "Dettaglio" "Detail")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' guardarlo anche senza Internet Explorer" "Why look at it even without Internet Explorer")</b><br>"
    NOTE+="$(L "WebCacheV01 non registra solo la navigazione del browser: raccoglie tutto cio' che passa dalle API WinINET, quindi anche le richieste fatte da applicazioni, script e malware che si appoggiano alle librerie di sistema. Su una macchina dove nessuno usa piu' IE, una voce qui e' spesso proprio codice non-browser che ha contattato la rete." \
        "WebCacheV01 does not only record browser activity: it collects everything going through the WinINET APIs, including requests made by applications, scripts and malware relying on system libraries. On a machine where nobody uses IE any more, an entry here is often exactly the non-browser code that reached out to the network.")"
    if [[ "$METHOD" != "ESE" ]]; then
        NOTE+="<br><br><b>$(L "Estrazione parziale" "Partial extraction")</b><br>"
        if $HAVE_ESE; then
            NOTE+="$(L "libesedb era disponibile ma non ha potuto aprire il database (file dirty o danneggiato, tipico di un'acquisizione a caldo): si e' ripiegato sull'estrazione degli URL dalle stringhe, senza date, tipo di voce e conteggi." \
                "libesedb was available but could not open the database (dirty or damaged file, typical of a live acquisition): extraction fell back to URLs from strings, without dates, entry type and counts.")"
        else
            NOTE+="$(L "libesedb non era disponibile: sono stati estratti solo gli URL dalle stringhe, senza date, tipo di voce e conteggi di accesso." \
                "libesedb was unavailable: only URLs were extracted from strings, without dates, entry type and access counts.")"
        fi
    fi
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Voci" "Entries")" "$TOTAL")"
    STATS+="$(stat_box "Database" "${#DBS[@]}" "info")"
    STATS+="$(stat_box "$(L "Parser" "Parser")" "${METHOD:-strings}" "$([[ "$METHOD" == "ESE" ]] && echo info || echo warn)")"
    finish_report "webcache" "WebCacheV01" "WEB" "AppData/Local/Microsoft/Windows/WebCache" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Voci WinINET" "WinINET entries")" "${DBS[0]}" "$TOTAL" "$TABLE" "🌐")</div>"
}

# ================================================================
#  MODULO 50 — Windows Search Index (Windows.edb)
#
#  L'indice di ricerca di Windows conserva nome, percorso, autore e — per molti
#  formati — un estratto del CONTENUTO dei file indicizzati. L'indice non viene
#  ripulito quando un file viene cancellato: la voce sopravvive fino al
#  successivo passaggio dell'indicizzatore.
#
#  Di conseguenza Windows.edb e' spesso l'unica fonte che conserva testo di
#  documenti eliminati, e i percorsi di file che non esistono piu' sul volume.
#
#  Il file e' un database ESE che puo' superare il gigabyte. Con libesedb si
#  legge la tabella SystemIndex_Gthr; senza, si ripiega sull'estrazione dei
#  percorsi dalle stringhe, che resta utile per sapere COSA c'era.
# ================================================================
module_search_index() {
    section_header "Windows Search Index" "$YELLOW"
    check_win_root || return 1

    local -a DBS=()
    local D F
    for D in "ProgramData/Microsoft/Search/Data/Applications/Windows" \
             "Documents and Settings/All Users/Application Data/Microsoft/Search/Data/Applications/Windows"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$R" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && DBS+=("$F")
        done < <(find "$R" -maxdepth 1 -type f -iname 'Windows.edb' 2>/dev/null)
    done

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Windows.edb non trovato (indicizzazione disattivata o percorso non standard)." \
                 "Windows.edb not found (indexing disabled or non-standard path).")"
        return 0
    fi
    local SZ; SZ=$(stat -c %s "${DBS[0]}" 2>/dev/null || echo 0)
    info "Windows.edb — ${BOLD}$(numfmt --to=iec "$SZ" 2>/dev/null || echo "$SZ")"

    local HAVE_ESE=false
    "$PY3" -c "import pyesedb" 2>/dev/null && HAVE_ESE=true
    $HAVE_ESE || info "$(L "libesedb (pyesedb) non disponibile: si usera' l'estrazione dalle stringhe." \
                          "libesedb (pyesedb) unavailable: string extraction will be used.")"

    info "$(L "Analisi in corso (il database puo' essere molto grande)..." "Analysing (the database can be very large)...")"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local METHOD; METHOD=$("$PY3" - "$OUT" "$($HAVE_ESE && echo 1 || echo 0)" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, os, re

out_path = sys.argv[1]
have_ese = sys.argv[2] == '1'
dbs = sys.argv[3:]

MAX_ROWS = 60000
rows = []
seen = set()

def add(path, extra, src):
    key = path.lower()[:200]
    if key in seen:
        return
    seen.add(key)
    rows.append((path[:400], extra[:200], src))

def parse_ese():
    import pyesedb
    for db in dbs:
        src = os.path.basename(db)
        try:
            f = pyesedb.file()
            f.open(db)
        except Exception:
            continue
        for i in range(f.get_number_of_tables()):
            if len(rows) >= MAX_ROWS:
                break
            try:
                t = f.get_table(i)
                tname = t.get_name()
            except Exception:
                continue
            # SystemIndex_Gthr contiene i percorsi indicizzati e i tempi di
            # ultima indicizzazione; SystemIndex_PropertyStore le proprieta'.
            if 'Gthr' not in tname and 'PropertyStore' not in tname:
                continue
            try:
                cols = [t.get_column(c).get_name() for c in range(t.get_number_of_columns())]
                ncols = len(cols)
            except Exception:
                continue
            for r in range(min(t.get_number_of_records(), MAX_ROWS)):
                try:
                    rec = t.get_record(r)
                except Exception:
                    continue
                path = ''
                extra = ''
                for n in range(ncols):
                    try:
                        v = rec.get_value_data_as_string(n)
                    except Exception:
                        continue
                    if not v:
                        continue
                    if not path and ('://' in v or re.match(r'^[A-Za-z]:\\', v)):
                        path = v
                    elif len(v) > 8 and not extra:
                        extra = v
                if path:
                    add(path, f"{tname} {extra}".strip(), src)
                if len(rows) >= MAX_ROWS:
                    break
        try:
            f.close()
        except Exception:
            pass


def parse_strings():
    """I percorsi indicizzati restano leggibili anche se l'ESE non si apre."""
    PATH = re.compile(r'(?:file:///)?[A-Za-z]:\\[^\x00<>|?*"\r\n]{4,250}')
    for db in dbs:
        src = os.path.basename(db)
        try:
            with open(db, 'rb') as fh:
                # Tetto di lettura: Windows.edb puo' superare il gigabyte.
                raw = fh.read(2 * 1024 * 1024 * 1024)
        except Exception:
            continue
        for enc in ('utf-16-le', 'latin-1'):
            try:
                text = raw.decode(enc, 'ignore')
            except Exception:
                continue
            for m in PATH.finditer(text):
                add(m.group(0).strip(), '', src)
                if len(rows) >= MAX_ROWS:
                    break

# Windows.edb e' quasi sempre in uso al momento dell'acquisizione, quindi
# spesso dirty e non apribile: ripiegare sulle stringhe evita un report vuoto.
method = 'strings'
if have_ese:
    parse_ese()
    if rows:
        method = 'ESE'
if not rows:
    parse_strings()
print(method if rows else 'nessuno')

rows.sort()
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')
PYEOF
    )

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna voce estratta dall'indice." "No entry extracted from the index.")"
        return 0
    fi
    ok "$(L "Voci indicizzate estratte:" "Indexed entries extracted:") ${BOLD}$TOTAL${RESET} ($(L "parser" "parser"): ${METHOD:-strings})"
    if [[ "$METHOD" != "ESE" ]] && $HAVE_ESE; then
        warn "$(L "Database non apribile con libesedb (in uso al momento dell'acquisizione): estrazione dalle stringhe." \
                 "Database not openable with libesedb (in use when acquired): string extraction.")"
    fi

    # Il valore forense sta nelle voci che puntano a file non piu' presenti, ma
    # il confronto non e' automatizzabile in modo affidabile: i percorsi
    # indicizzati usano la lettera di unita' vista dal sistema (C:\...), che non
    # corrisponde al punto di mount in analisi. Meglio dirlo che dedurre a caso.
    info "$(L "Il confronto con i file ancora presenti richiede la mappatura delle lettere di unita': va fatto a mano." \
             "Comparing against files still present requires drive-letter mapping: do it manually.")"

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(head -30000 "$OUT")
    local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Percorso indicizzato" "Indexed path")" "$(L "Dettaglio" "Detail")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Come usarlo" "How to use it")</b><br>"
    NOTE+="$(L "L'indice non viene ripulito alla cancellazione di un file: una voce che punta a un percorso non piu' esistente sul volume e' la traccia di un file eliminato. Il confronto va fatto tenendo conto della lettera di unita': il volume montato in analisi non corrisponde necessariamente a C: come lo vedeva il sistema." \
        "The index is not purged when a file is deleted: an entry pointing to a path no longer present on the volume is the trace of a removed file. Compare with the drive letter in mind: the mounted volume does not necessarily correspond to C: as the system saw it.")"
    if [[ "$METHOD" != "ESE" ]]; then
        NOTE+="<br><br>$(L "Sono stati estratti i soli percorsi: gli estratti di contenuto e le date di indicizzazione richiedono il parsing ESE, non riuscito su questo database." \
            "Only paths were extracted: content excerpts and indexing dates require ESE parsing, which did not succeed on this database.")"
    fi
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Voci" "Entries")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Dimensione DB" "DB size")" "$(numfmt --to=iec "$SZ" 2>/dev/null || echo "$SZ")" "info")"
    STATS+="$(stat_box "$(L "Parser" "Parser")" "${METHOD:-strings}" "$([[ "$METHOD" == "ESE" ]] && echo info || echo warn)")"
    finish_report "search_index" "Windows Search Index" "IDX" "Windows.edb" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Percorsi indicizzati" "Indexed paths")" "${DBS[0]}" "$TOTAL" "$TABLE" "⌕")</div>"
}

# ================================================================
#  MODULO 51 — Sigma sugli Event Log (EVTX)
#
#  Sigma e' il formato in cui la comunita' pubblica le detection: SigmaHQ,
#  i CERT e i vendor distribuiscono migliaia di regole YAML. Applicarle agli
#  EVTX di un disco acquisito e' quello che fanno Chainsaw e Hayabusa, ed e' il
#  passo che trasforma una raccolta di log in un triage.
#
#  IL SOTTOINSIEME SUPPORTATO E' DICHIARATO, NON IMPLICITO. Sigma e' un
#  linguaggio ampio: modificatori base64, CIDR, condizioni con parentesi,
#  aggregazioni temporali. Implementarne una parte e far finta di supportarlo
#  tutto significherebbe che una regola non valutata risulta "non scattata" —
#  cioe' un falso negativo silenzioso, il difetto peggiore che una detection
#  possa avere. Qui le regole che il motore non sa valutare vengono CONTATE ED
#  ELENCATE nel report con il motivo.
#
#  Supportato:
#    - selezioni: mappa campo/valore, liste di valori (OR), liste di mappe (OR)
#    - modificatori: contains, startswith, endswith, re, all, cased
#    - condizioni: "sel", "a and b", "a or b", "a and not b", "not a",
#                  "1 of x*", "all of x*", "1 of them", "all of them"
#    - null come valore (campo assente o vuoto)
#
#  Non supportato (regola scartata e dichiarata):
#    - condizioni con parentesi o aggregazioni (| count, near, timeframe)
#    - modificatori base64/base64offset, utf16, wide, cidr, gt/lt
#    - logsource non mappabile a un canale EVTX presente sul volume
#
#  Il compilatore vive in src/lib/19-pylib-sigma.sh: e' la parte rischiosa —
#  un modificatore interpretato male produce un falso negativo invisibile — e
#  li' e' esercitabile dai test con eventi sintetici, senza un .evtx. Qui resta
#  la lettura degli EVTX e la presentazione.
# ================================================================

_guard_sigma() {
    if [[ -z "${SIGMA_RULES:-}" ]]; then
        L "nessuna regola (--sigma)" "no rules (--sigma)"
        return 1
    fi
    return 0
}

module_sigma() {
    section_header "Sigma — Event Log" "$RED"
    check_win_root || return 1

    if [[ -z "${SIGMA_RULES:-}" ]]; then
        warn "$(L "Nessuna regola indicata." "No rules given.")"
        info "$(L "Uso: --sigma /percorso/regole/  (file .yml o directory, anche annidata)" \
                 "Usage: --sigma /path/rules/  (a .yml file or a directory, nested is fine)")"
        return 0
    fi
    [[ -e "$SIGMA_RULES" ]] || { err "$(L "Percorso regole inesistente:" "Rules path does not exist:") $SIGMA_RULES"; return 1; }

    if ! "$PY3" -c "import yaml" 2>/dev/null; then
        err "$(L "PyYAML non disponibile: le regole Sigma sono file YAML." \
                 "PyYAML unavailable: Sigma rules are YAML files.")"
        info "$(L "Installa con:" "Install with:") ${PY3} -m pip install pyyaml"
        return 1
    fi
    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        err "$(L "python-evtx non disponibile: senza non si leggono gli EVTX." \
                 "python-evtx unavailable: without it EVTX cannot be read.")"
        info "$(L "Installa con:" "Install with:") ${PY3} -m pip install python-evtx"
        return 1
    fi

    local EVTX_DIR; EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    [[ -z "$EVTX_DIR" ]] && EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
    if [[ -z "$EVTX_DIR" ]]; then
        warn "$(L "Directory dei log eventi non trovata." "Event log directory not found.")"
        return 0
    fi
    info "$(L "Log eventi:" "Event logs:") ${BOLD}${EVTX_DIR}"
    info "$(L "Regole:" "Rules:") ${BOLD}${SIGMA_RULES}"
    info "$(L "Valutazione in corso (dipende dal numero di regole e dalla dimensione dei log)..." \
             "Evaluating (depends on rule count and log size)...")"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local STATS; STATS=$(mktemp); register_tmp "$STATS"

    run_py_with_lib pylib_sigma "$SIGMA_RULES" "$EVTX_DIR" "$OUT" "$STATS" "$SIGMA_MAX_RECORDS" << 'PYEOF'
import sys, os, re, json, glob

rules_path, evtx_dir, out_path, stats_path = sys.argv[1:5]
MAX_RECORDS = int(sys.argv[5])


import Evtx.Evtx as evtx

NS = 'http://schemas.microsoft.com/win/2004/08/events/event'

available = {os.path.basename(p).lower(): p
             for p in glob.glob(os.path.join(evtx_dir, '*.evtx'))}
rules, rejected = load_sigma_rules(rules_path, available)

stats = {
    'rules_loaded': len(rules) + len(rejected),
    'rules_active': len(rules),
    'rules_rejected': rejected[:400],
    'rules_rejected_total': len(rejected),
    'channels': [],
    'records_read': 0,
    'stopped_at_cap': False,
}

if not rules:
    json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
    sys.exit(0)

# --- valutazione -----------------------------------------------------------
by_file = {}
for r in rules:
    for f in r['files']:
        by_file.setdefault(f, []).append(r)

_TAG = re.compile(r'\{[^}]+\}')
rows = []
seen = set()

for path, rlist in sorted(by_file.items()):
    n = 0
    try:
        with evtx.Evtx(path) as log:
            for rec in log.records():
                if stats['records_read'] >= MAX_RECORDS:
                    stats['stopped_at_cap'] = True
                    break
                n += 1
                stats['records_read'] += 1
                try:
                    root = rec.lxml()
                except Exception:
                    continue
                sysel = root.find('{%s}System' % NS)
                if sysel is None:
                    continue
                ev = {}
                eid = ''
                ts = ''
                for child in sysel:
                    tag = _TAG.sub('', child.tag)
                    if tag == 'EventID':
                        eid = (child.text or '').strip()
                        ev['EventID'] = eid
                    elif tag == 'TimeCreated':
                        ts = child.attrib.get('SystemTime', '')[:19]
                    elif tag == 'Provider':
                        ev['Provider_Name'] = child.attrib.get('Name', '')
                    elif tag == 'Channel':
                        ev['Channel'] = (child.text or '').strip()
                    elif tag == 'Computer':
                        ev['Computer'] = (child.text or '').strip()
                de = root.find('.//{%s}EventData' % NS)
                if de is not None:
                    for it in de:
                        key = it.attrib.get('Name') or _TAG.sub('', it.tag)
                        ev[key] = (it.text or '').strip()
                for r in rlist:
                    try:
                        if not r['pred'](ev):
                            continue
                    except Exception:
                        continue
                    key = (r['title'], ts, eid)
                    if key in seen:
                        continue
                    seen.add(key)
                    # Il contesto e' cio' che rende il match verificabile: i
                    # campi piu' parlanti, non il record intero.
                    ctx = ' | '.join(
                        '%s=%s' % (k, v[:120])
                        for k, v in ev.items()
                        if k in ('Image', 'CommandLine', 'ParentImage', 'ParentCommandLine',
                                 'TargetUserName', 'SubjectUserName', 'ServiceName',
                                 'ScriptBlockText', 'TargetFilename', 'DestinationIp',
                                 'QueryName', 'TargetObject', 'Details') and v)[:600]
                    rows.append((ts.replace('T', ' '), r['level'], r['title'],
                                 os.path.basename(path), eid, ctx,
                                 ','.join(t for t in r['tags'] if t.startswith('attack.'))[:100]))
    except Exception as exc:
        stats['channels'].append({'file': os.path.basename(path), 'records': n,
                                  'error': str(exc)[:150]})
        continue
    stats['channels'].append({'file': os.path.basename(path), 'records': n,
                              'rules': len(rlist)})
    if stats['stopped_at_cap']:
        break

LEVEL_ORDER = {'critical': 0, 'high': 1, 'medium': 2, 'low': 3, 'informational': 4}
rows.sort(key=lambda r: (LEVEL_ORDER.get(r[1], 9), r[0]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
stats['matches'] = len(rows)
json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
PYEOF

    [[ -s "$STATS" ]] || { err "$(L "Valutazione non riuscita." "Evaluation failed.")"; return 1; }
    local NLOAD NACT NREJ NREC CAPPED
    read -r NLOAD NACT NREJ NREC CAPPED < <("$PY3" -c '
import json, sys
s = json.load(open(sys.argv[1]))
print(s["rules_loaded"], s["rules_active"], s["rules_rejected_total"],
      s["records_read"], str(s["stopped_at_cap"]).lower())' "$STATS" 2>/dev/null)

    if [[ "${NACT:-0}" -eq 0 ]]; then
        warn "$(L "Nessuna regola valutabile su questo volume." "No rule evaluable on this volume.")"
        info "$(L "Regole caricate:" "Rules loaded:") ${NLOAD:-0}  ·  $(L "scartate:" "rejected:") ${NREJ:-0}"
        "$PY3" -c '
import json, sys
for r in json.load(open(sys.argv[1]))["rules_rejected"][:10]:
    print("      %s — %s" % (r["rule"][:60], r["reason"]))' "$STATS" 2>/dev/null
        return 0
    fi

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    ok "$(L "Regole attive:" "Active rules:") ${BOLD}${NACT}${RESET}/${NLOAD}  ·  $(L "record letti:" "records read:") ${BOLD}${NREC}"
    [[ "${NREJ:-0}" -gt 0 ]] && warn "$(L "Regole non valutate (elencate nel report):" "Rules not evaluated (listed in the report):") ${BOLD}${NREJ}"
    [[ "$CAPPED" == "true" ]] && warn "$(L "Raggiunto il tetto di record: la valutazione e' PARZIALE." \
                                          "Record cap reached: the evaluation is PARTIAL.")"

    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessun match." "No match.")"
    else
        warn "$(L "Match:" "Matches:") ${BOLD}${TOTAL}"
        awk -F'\t' '{printf "      [%s] %s  %s\n", $2, $1, substr($3,1,70)}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local TABLE
    if [[ "$TOTAL" -gt 0 ]]; then
        TABLE=$(_rows_to_table "$(head -20000 "$OUT")" \
            "$(L "Data" "Date")" "$(L "Livello" "Level")" "$(L "Regola" "Rule")" \
            "$(L "Canale" "Channel")" "EventID" "$(L "Contesto" "Context")" "ATT&CK")
    else
        TABLE="<div style='padding:.6rem 0;font-size:.85rem'>$(L "Nessun match." "No match.")</div>"
    fi

    local COV; COV=$("$PY3" - "$STATS" "${LANG:-en}" << 'PYEOF' 2>/dev/null
import json, sys, html

s = json.load(open(sys.argv[1]))
it = sys.argv[2] == 'it'


def L(i, e):
    return i if it else e


o = ['<div class="card" style="margin-bottom:1rem">'
     '<div style="padding:1rem 1.5rem;font-size:.8rem;line-height:1.7">']
o.append('<b>' + L('Copertura effettiva', 'Actual coverage') + '</b><br>')
o.append(L("Sigma e' un linguaggio ampio e questo motore ne implementa un sottoinsieme. "
           "Una regola non valutata non e' una regola non scattata: le regole scartate "
           "sono contate ed elencate qui sotto con il motivo, perche' un falso negativo "
           "silenzioso e' il difetto peggiore che una detection possa avere.",
           'Sigma is a broad language and this engine implements a subset of it. '
           'A rule that was not evaluated is not a rule that did not fire: rejected rules '
           'are counted and listed below with the reason, because a silent false negative '
           'is the worst defect a detection can have.') + '<br><br>')
o.append('<b>%d</b> %s · <b>%d</b> %s · <b>%d</b> %s<br><br>' % (
    s['rules_loaded'], L('regole caricate', 'rules loaded'),
    s['rules_active'], L('valutate', 'evaluated'),
    s['rules_rejected_total'], L('scartate', 'rejected')))

o.append('<table><tr><th>' + L('Canale', 'Channel') + '</th><th>'
         + L('Record letti', 'Records read') + '</th><th>'
         + L('Regole applicate', 'Rules applied') + '</th></tr>')
for c in s['channels']:
    o.append('<tr><td class="mono">%s</td><td class="mono">%d</td><td class="mono">%s</td></tr>'
             % (html.escape(c['file']), c.get('records', 0),
                html.escape(str(c.get('rules', c.get('error', '-'))))))
o.append('</table>')
if s['stopped_at_cap']:
    o.append('<br><b>' + L("Raggiunto il tetto di record: la valutazione e' PARZIALE.",
                           'Record cap reached: the evaluation is PARTIAL.') + '</b>')

if s['rules_rejected']:
    reasons = {}
    for r in s['rules_rejected']:
        key = r['reason'].split(':')[0].split('(')[0].strip()
        reasons.setdefault(key, []).append(r['rule'])
    o.append('<br><br><b>' + L('Regole scartate, per motivo', 'Rejected rules, by reason') + '</b>')
    o.append('<table><tr><th>' + L('Motivo', 'Reason') + '</th><th>'
             + L('Regole', 'Rules') + '</th><th>' + L('Esempi', 'Examples') + '</th></tr>')
    for reason, names in sorted(reasons.items(), key=lambda kv: -len(kv[1])):
        o.append('<tr><td class="mono">%s</td><td class="mono">%d</td><td class="mono">%s</td></tr>'
                 % (html.escape(reason), len(names),
                    html.escape(', '.join(n[:50] for n in names[:3]))))
    o.append('</table>')
    if s['rules_rejected_total'] > len(s['rules_rejected']):
        o.append('<br>' + L('Elenco troncato: scartate in totale %d.'
                            % s['rules_rejected_total'],
                            'List truncated: %d rejected in total.'
                            % s['rules_rejected_total']))
o.append('</div></div>')
print(''.join(o))
PYEOF
    )

    local SB
    SB="$(stat_box "Match" "$TOTAL" "$([[ "$TOTAL" -gt 0 ]] && echo warn || echo ok)")"
    SB+="$(stat_box "$(L "Regole valutate" "Rules evaluated")" "${NACT:-0}" "info")"
    SB+="$(stat_box "$(L "Regole scartate" "Rules rejected")" "${NREJ:-0}" "$([[ "${NREJ:-0}" -gt 0 ]] && echo warn || echo info)")"
    SB+="$(stat_box "$(L "Record letti" "Records read")" "${NREC:-0}" "info")"
    finish_report "sigma" "Sigma" "SIG" "$(basename "$SIGMA_RULES")" "$SB" \
        "${COV}<div class='cards'>$(generic_card_html "$(L "Riscontri" "Matches")" "$SIGMA_RULES" "$TOTAL" "$TABLE" "⚐")</div>"
}

# ================================================================
#  MODULI LINUX
# ================================================================

# --- LINUX 1 — System Logs (/var/log testuali) ---
module_linux_syslog() {
    section_header "Linux — System Logs" "$GREEN"
    check_target_root || return 1
    local LOGDIR; LOGDIR=$(ci_find_dir "$WIN_ROOT" "var/log")
    [[ -z "$LOGDIR" ]] && { warn "$(L "Directory var/log non trovata." "var/log directory not found.")"; return 0; }

    local TARGETS=(syslog messages kern.log dmesg auth.log secure boot.log faillog dpkg.log yum.log)
    local KW="fail|error|denied|refused|invalid|segfault|root|sudo|su:|authentication failure|break-in|illegal"
    local BODY="" FOUND=0
    for NAME in "${TARGETS[@]}"; do
        local F; F=$(ci_find_file "$LOGDIR" "$NAME")
        [[ -z "$F" || ! -s "$F" ]] && continue
        FOUND=$((FOUND + 1))
        ok "$NAME — ${BOLD}$(stat -c %s "$F" 2>/dev/null) B"
        print_file_lines "$F" "$KW" 60
        echo ""
        BODY+=$(file_card_html "$F" "$KW" "≣")
    done
    separator
    info "$(L "Log trovati:" "Logs found:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun log testuale trovato." "No text log found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Log analizzati" "Logs analysed")" "$FOUND" "info")
    finish_report "linux_syslog" "Linux System Logs" "LOG" "/var/log" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 2 — systemd Journal ---
module_linux_journal() {
    section_header "Linux — systemd Journal" "$GREEN"
    check_target_root || return 1
    local JDIR; JDIR=$(ci_find_dir "$WIN_ROOT" "var/log/journal")
    [[ -z "$JDIR" ]] && { warn "$(L "Nessun journal persistente (var/log/journal assente)." "No persistent journal (var/log/journal missing).")"; return 0; }
    mapfile -t JFILES < <(find "$JDIR" -type f -iname "*.journal" 2>/dev/null)
    [[ ${#JFILES[@]} -eq 0 ]] && { warn "$(L "Nessun file .journal trovato." "No .journal file found.")"; return 0; }
    ok "$(L "File journal trovati:" "Journal files found:") ${BOLD}${#JFILES[@]}"

    local KW="fail|error|denied|refused|sudo|root|segfault|authentication"
    local BODY="" PARSED=0 HAVE_JCTL=0
    command -v journalctl >/dev/null 2>&1 && HAVE_JCTL=1
    for JF in "${JFILES[@]}"; do
        local TXT=""
        if [[ $HAVE_JCTL -eq 1 ]]; then
            TXT=$(journalctl --no-pager --file "$JF" -o short-iso 2>/dev/null | head -100000)
        fi
        if [[ -z "$TXT" ]]; then
            # fallback: estrazione stringhe leggibili dal binario
            TXT=$(strings -n 8 "$JF" 2>/dev/null | grep -iE 'MESSAGE=|_COMM=|_EXE=|_HOSTNAME=' | sed 's/^MESSAGE=//' | head -100000)
        fi
        [[ -z "$TXT" ]] && continue
        PARSED=$((PARSED + 1))
        echo -e "  ${DIM}• $(basename "$JF")${RESET}"
        local TMPF; TMPF=$(mktemp); printf '%s\n' "$TXT" > "$TMPF"
        BODY+=$(file_card_html "$TMPF" "$KW" "◷")
        rm -f "$TMPF"
    done
    [[ $HAVE_JCTL -eq 0 ]] && warn "$(L "journalctl non disponibile: usato fallback 'strings' (output parziale)." "journalctl unavailable: used 'strings' fallback (partial output).")"
    separator
    info "$(L "Journal elaborati:" "Journals processed:") ${BOLD}$PARSED"
    [[ $PARSED -eq 0 ]] && { warn "$(L "Impossibile estrarre contenuti dal journal." "Could not extract journal contents.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File journal" "Journal files")" "${#JFILES[@]}" "info")
    finish_report "linux_journal" "systemd Journal" "JRN" "/var/log/journal" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 3 — Login History (wtmp/btmp/lastlog) ---
module_linux_logins() {
    section_header "Linux — Login History" "$CYAN"
    check_target_root || return 1
    local LOGDIR; LOGDIR=$(ci_find_dir "$WIN_ROOT" "var/log")
    local VARRUN; VARRUN=$(ci_find_dir "$WIN_ROOT" "var/run")
    local WTMP BTMP LASTLOG
    [[ -n "$LOGDIR" ]] && { WTMP=$(ci_find_file "$LOGDIR" "wtmp"); BTMP=$(ci_find_file "$LOGDIR" "btmp"); LASTLOG=$(ci_find_file "$LOGDIR" "lastlog"); }
    if [[ -z "$WTMP" && -z "$BTMP" ]]; then
        warn "$(L "wtmp/btmp non trovati." "wtmp/btmp not found.")"; return 0
    fi
    local BODY="" TOTAL=0
    for PAIR in "wtmp|$WTMP|$(L "Login riusciti" "Successful logins")" "btmp|$BTMP|$(L "Login FALLITI" "FAILED logins")"; do
        IFS='|' read -r TAG FILE LABEL <<< "$PAIR"
        [[ -z "$FILE" || ! -s "$FILE" ]] && continue
        echo -e "  ${BOLD}${LABEL}${RESET}  ${DIM}($FILE)${RESET}"
        local OUT; OUT=$(_parse_utmp "$FILE")
        local N; N=$(printf '%s\n' "$OUT" | grep -c . || true)
        TOTAL=$((TOTAL + N))
        printf '%s\n' "$OUT" | head -40 | while IFS= read -r R; do echo -e "      ${DIM}$R${RESET}"; done
        local TMPF; TMPF=$(mktemp); printf '%s\n' "$OUT" > "$TMPF"
        local KW; [[ "$TAG" == "btmp" ]] && KW=".*" || KW="root|0\\.0\\.0\\.0"
        BODY+=$(generic_card_html "$LABEL" "$FILE" "$N record" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "$KW")</pre>" "⇆")
        rm -f "$TMPF"
        echo ""
    done
    separator
    info "$(L "Record totali:" "Total records:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun record di login leggibile." "No readable login record.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Record login" "Login records")" "$TOTAL" "info")
    finish_report "linux_logins" "Linux Login History" "LOG" "wtmp / btmp / lastlog" "$STATS" "<div class='cards'>$BODY</div>"
}

# Parser binario utmp/wtmp/btmp (struct standard glibc, 384 byte/record)
_parse_utmp() {
    "$PY3" - "$1" << 'PYEOF'
import sys, struct, datetime
# struct utmp (Linux x86-64): ut_type(h) pad(2x) ut_pid(i) ut_line(32s) ut_id(4s)
# ut_user(32s) ut_host(256s) ut_exit(4s) ut_session(i) tv_sec(i) tv_usec(i) ut_addr_v6(16s) unused(20s)
REC = 384
TYPES = {0:'EMPTY',1:'RUNLVL',2:'BOOT',3:'NEWTIME',4:'OLDTIME',5:'INIT',6:'LOGIN',7:'USER',8:'DEAD'}
try:
    with open(sys.argv[1],'rb') as f:
        data=f.read()
    rows=[]
    for off in range(0,len(data)-REC+1,REC):
        rec=data[off:off+REC]
        try:
            ut_type=struct.unpack('<h',rec[0:2])[0]
            pid=struct.unpack('<i',rec[4:8])[0]
            line=rec[8:40].split(b'\x00')[0].decode('utf-8','replace')
            user=rec[44:76].split(b'\x00')[0].decode('utf-8','replace')
            host=rec[76:332].split(b'\x00')[0].decode('utf-8','replace')
            tv_sec=struct.unpack('<i',rec[340:344])[0]
        except Exception:
            continue
        if tv_sec<=0 and not user: continue
        try: ts=datetime.datetime.utcfromtimestamp(tv_sec).strftime('%Y-%m-%d %H:%M:%S')
        except Exception: ts='?'
        t=TYPES.get(ut_type,str(ut_type))
        if not user and t in ('EMPTY','DEAD'): continue
        rows.append(f"{ts}  {t:7s}  {user:16s}  {line:12s}  {host}")
    print('\n'.join(rows))
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF
}

# --- LINUX 4 — Shell History (per utente) ---
module_linux_shell_history() {
    section_header "Linux — Shell History" "$MAGENTA"
    check_target_root || return 1
    local FILES=(.bash_history .zsh_history .sh_history .history .python_history .mysql_history .psql_history .lesshst .node_repl_history)
    local KW="password|passwd|secret|credential|token|ssh|scp|curl|wget|nc |ncat|base64|sudo|chmod \\+x|/dev/tcp|reverse|nmap|chattr"
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        for HF in "${FILES[@]}"; do
            local F="$HOME_DIR/$HF"
            [[ -f "$F" && -s "$F" ]] || continue
            UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
            # I REPL readline (python3, node, psql, mysql) codificano spazi e
            # backslash in ottale: senza decodifica il report mostra
            # "\040\040value = ..." al posto del codice.
            local MODE="histts"
            case "$HF" in
                .python_history|.node_repl_history|.psql_history|.mysql_history) MODE="histrl" ;;
            esac
            CARDS+=$(file_card_html "$F" "$KW" "\$" "$MODE")
        done
        # fish history
        local FISH; FISH=$(ci_find_dir "$HOME_DIR" ".local/share/fish")
        if [[ -n "$FISH" && -s "$FISH/fish_history" ]]; then
            UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
            CARDS+=$(file_card_html "$FISH/fish_history" "$KW" "\$" "histts")
        fi
        [[ $UCOUNT -eq 0 ]] && { dim_msg "$UNAME — $(L "nessuna history" "no history")"; continue; }
        USERS=$((USERS + 1))
        ok "$UNAME — ${BOLD}$UCOUNT file"
        for HF in "${FILES[@]}"; do
            local F="$HOME_DIR/$HF"; [[ -f "$F" && -s "$F" ]] || continue
            echo -e "  ${DIM}• $HF${RESET}"; print_file_lines "$F" "$KW" 30; echo ""
        done
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT file" "$CARDS" "◢")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Utenti con history:" "Users with history:") ${BOLD}$USERS${RESET}  |  File: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history trovata." "No history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "File history" "History files")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_shell_history" "Linux Shell History" "SH" "~/.bash_history · .zsh_history · ..." "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 5 — AI CLI History ---
module_linux_ai_history() {
    section_header "Linux — AI CLI History" "$MAGENTA"
    check_target_root || return 1
    # path relativi alla home da scandire (file o directory)
    local REL=(.claude .config/claude .aider.chat.history.md .aider.input.history .config/aichat .ollama/history
               .config/io.datasette.llm .codeium .config/github-copilot .continue .cursor)
    local KW="password|secret|token|api_key|apikey|key=|credential"
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        for R in "${REL[@]}"; do
            local P="$HOME_DIR/$R"
            if [[ -f "$P" && -s "$P" ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$P" "$KW" "◈")
            elif [[ -d "$P" ]]; then
                while IFS= read -r AF; do
                    [[ -s "$AF" ]] || continue
                    UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$AF" "$KW" "◈")
                done < <(find "$P" -maxdepth 3 -type f \( -iname "*.json" -o -iname "*.jsonl" -o -iname "*history*" -o -iname "*.md" \) 2>/dev/null | head -25)
            fi
        done
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); ok "$UNAME — ${BOLD}$UCOUNT file AI"
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT file" "$CARDS" "◈")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Artefatti AI trovati:" "AI artifacts found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history di assistenti AI trovata." "No AI assistant history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "File AI" "AI files")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_ai_history" "Linux AI CLI History" "AI" "claude · aider · aichat · ollama · copilot" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 6 — Browser History ---
module_linux_browser() {
    section_header "Linux — Browser History" "$CYAN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        # Chromium-like: History (SQLite) sotto vari profili
        local CHROME_ROOTS=(".config/google-chrome" ".config/chromium" ".config/BraveSoftware/Brave-Browser"
                            ".config/microsoft-edge" ".config/vivaldi"
                            "snap/chromium/common/chromium" ".var/app/com.google.Chrome/config/google-chrome")
        for CR in "${CHROME_ROOTS[@]}"; do
            local BASE="$HOME_DIR/$CR"; [[ -d "$BASE" ]] || continue
            while IFS= read -r HISTDB; do
                [[ -f "$HISTDB" ]] || continue
                local ROWS; ROWS=$(query_sqlite "$HISTDB" "SELECT datetime(last_visit_time/1000000-11644473600,'unixepoch'), url, title FROM urls ORDER BY last_visit_time DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
                ok "$UNAME — $(basename "$(dirname "$HISTDB")") (Chromium)"
                CARDS+=$(_browser_table_card "$HISTDB" "$ROWS")
            done < <(find "$BASE" -maxdepth 2 -name "History" -type f 2>/dev/null)
        done
        # Firefox: places.sqlite
        for FR in ".mozilla/firefox" "snap/firefox/common/.mozilla/firefox" ".var/app/org.mozilla.firefox/.mozilla/firefox"; do
            local FBASE="$HOME_DIR/$FR"; [[ -d "$FBASE" ]] || continue
            while IFS= read -r PLACES; do
                local ROWS; ROWS=$(query_sqlite "$PLACES" "SELECT datetime(last_visit_date/1000000,'unixepoch'), url, title FROM moz_places WHERE last_visit_date IS NOT NULL ORDER BY last_visit_date DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
                ok "$UNAME — $(basename "$(dirname "$PLACES")") (Firefox)"
                CARDS+=$(_browser_table_card "$PLACES" "$ROWS")
            done < <(find "$FBASE" -maxdepth 2 -name "places.sqlite" -type f 2>/dev/null)
        done
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1))
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT profili" "$CARDS" "◐")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Profili browser con history:" "Browser profiles with history:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history browser trovata." "No browser history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Profili" "Profiles")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_browser" "Linux Browser History" "WEB" "Firefox · Chrome · Chromium · Brave · Edge" "$STATS" "<div class='cards'>$BODY</div>"
}

# Tabella HTML per le righe (data\turl\ttitolo) di un DB browser
_browser_table_card() {
    local DB="$1" ROWS="$2"
    local _RTMP; _RTMP=$(mktemp); printf '%s\n' "$ROWS" > "$_RTMP"
    local TABLE; TABLE=$("$PY3" - "$_RTMP" << 'PYEOF'
import sys, html
print("<table><tr><th>Data</th><th>URL</th><th>Titolo</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    parts=line.rstrip('\n').split('\t')
    if len(parts)<2: continue
    d=html.escape(parts[0]); u=html.escape(parts[1]); t=html.escape(parts[2] if len(parts)>2 else '')
    print(f"<tr><td class='mono dim'>{d}</td><td class='mono'>{u}</td><td>{t}</td></tr>")
print("</table>")
PYEOF
)
    rm -f "$_RTMP"
    local N; N=$(printf '%s\n' "$ROWS" | grep -c . || true)
    generic_card_html "$(basename "$(dirname "$DB")")" "$DB" "$N URL" "$TABLE" "◐"
}

# --- LINUX 7 — User Accounts ---
module_linux_accounts() {
    section_header "Linux — User Accounts" "$RED"
    check_target_root || return 1
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    [[ -z "$ETC" ]] && { warn "$(L "Directory etc non trovata." "etc directory not found.")"; return 0; }
    local BODY="" FOUND=0
    # passwd: evidenzia UID 0 e shell interattive
    local PASSWD; PASSWD=$(ci_find_file "$ETC" "passwd")
    if [[ -n "$PASSWD" ]]; then
        FOUND=$((FOUND + 1))
        ok "passwd"
        local TABLE; TABLE=$("$PY3" - "$PASSWD" << 'PYEOF'
import sys, html
print("<table><tr><th>User</th><th>UID</th><th>GID</th><th>Home</th><th>Shell</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    line=line.rstrip('\n')
    if not line or line.startswith('#'): continue
    p=line.split(':')
    if len(p)<7: continue
    uid=p[2]; sh=p[6]
    bad = uid=='0' or (sh and not sh.endswith(('nologin','false','sync')))
    cls=" class='bad'" if uid=='0' else ""
    print(f"<tr{cls}><td class='mono'>{html.escape(p[0])}</td><td class='mono'>{uid}</td><td class='mono'>{p[3]}</td><td class='mono dim'>{html.escape(p[5])}</td><td class='mono'>{html.escape(sh)}</td></tr>")
print("</table>")
PYEOF
)
        awk -F: '$3==0{print "      \033[0;31m[UID 0] "$1" → "$7"\033[0m"}' "$PASSWD"
        BODY+=$(generic_card_html "passwd" "$PASSWD" "$(grep -cv '^#' "$PASSWD" 2>/dev/null) account" "$TABLE" "◉")
    fi
    # shadow: stato password (hash presente / vuoto / lock)
    local SHADOW; SHADOW=$(ci_find_file "$ETC" "shadow")
    if [[ -n "$SHADOW" && -r "$SHADOW" ]]; then
        FOUND=$((FOUND + 1)); ok "shadow"
        local STAB; STAB=$("$PY3" - "$SHADOW" << 'PYEOF'
import sys, html
print("<table><tr><th>User</th><th>Stato password</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    line=line.rstrip('\n')
    if not line or line.startswith('#'): continue
    p=line.split(':')
    if len(p)<2: continue
    h=p[1]
    if h in ('!','*','!!',''): st='bloccata / nessuna'
    elif h.startswith('!'): st='bloccata (hash presente)'
    else: st='hash impostato'
    cls=" class='bad'" if h=='' else ""
    print(f"<tr{cls}><td class='mono'>{html.escape(p[0])}</td><td>{st}</td></tr>")
print("</table>")
PYEOF
)
        BODY+=$(generic_card_html "shadow" "$SHADOW" "" "$STAB" "◉")
    elif [[ -n "$SHADOW" ]]; then
        warn "shadow $(L "presente ma non leggibile" "present but unreadable")"
    fi
    # group, sudoers, sudoers.d
    for NF in group sudoers; do
        local F; F=$(ci_find_file "$ETC" "$NF")
        [[ -n "$F" && -s "$F" ]] || continue
        FOUND=$((FOUND + 1)); ok "$NF"
        BODY+=$(file_card_html "$F" "wheel|sudo|admin|ALL|NOPASSWD|root" "◉")
    done
    local SUDOERSD; SUDOERSD=$(ci_find_dir "$ETC" "sudoers.d")
    if [[ -n "$SUDOERSD" ]]; then
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1))
            BODY+=$(file_card_html "$F" "NOPASSWD|ALL|root" "◉")
        done < <(find "$SUDOERSD" -maxdepth 1 -type f 2>/dev/null)
    fi
    separator
    info "$(L "File analizzati:" "Files analysed:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File account" "Account files")" "$FOUND" "info")
    finish_report "linux_accounts" "Linux User Accounts" "USR" "/etc/passwd · shadow · group · sudoers" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 8 — Persistence ---
module_linux_persistence() {
    section_header "Linux — Persistence" "$ORANGE"
    check_target_root || return 1
    local KW="curl|wget|/tmp/|/dev/shm|base64|nc |ncat|python -c|bash -i|/dev/tcp|chmod|reverse|\\.onion|http"
    local BODY="" FOUND=0
    _persist_add() { # label, path(file o dir), glob
        local LABEL="$1" P="$2"
        if [[ -f "$P" && -s "$P" ]]; then
            FOUND=$((FOUND + 1)); ok "$LABEL — $(basename "$P")"
            BODY+=$(file_card_html "$P" "$KW" "⟳")
        elif [[ -d "$P" ]]; then
            while IFS= read -r F; do
                [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1))
                BODY+=$(file_card_html "$F" "$KW" "⟳")
            done < <(find "$P" -maxdepth 2 -type f 2>/dev/null | head -60)
            [[ -n "$(find "$P" -maxdepth 2 -type f 2>/dev/null | head -1)" ]] && ok "$LABEL ($(basename "$P"))"
        fi
    }
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    if [[ -n "$ETC" ]]; then
        _persist_add "cron" "$(ci_find_file "$ETC" "crontab")"
        for D in cron.d cron.daily cron.hourly cron.weekly cron.monthly; do
            _persist_add "cron" "$(ci_find_dir "$ETC" "$D")"
        done
        _persist_add "init" "$(ci_find_file "$ETC" "rc.local")"
        _persist_add "init.d" "$(ci_find_dir "$ETC" "init.d")"
        _persist_add "systemd" "$(ci_find_dir "$ETC" "systemd/system")"
        _persist_add "xdg-autostart" "$(ci_find_dir "$ETC" "xdg/autostart")"
        _persist_add "ld.so.preload" "$(ci_find_file "$ETC" "ld.so.preload")"
    fi
    # user crontabs
    local SPOOL; SPOOL=$(ci_find_dir "$WIN_ROOT" "var/spool/cron")
    [[ -n "$SPOOL" ]] && _persist_add "user-cron" "$SPOOL"
    # systemd vendor units
    local USRLIB; USRLIB=$(ci_find_dir "$WIN_ROOT" "usr/lib/systemd/system")
    # per-user autostart, systemd user, shell init
    while IFS= read -r HOME_DIR; do
        _persist_add "autostart" "$(ci_find_dir "$HOME_DIR" ".config/autostart")"
        _persist_add "systemd-user" "$(ci_find_dir "$HOME_DIR" ".config/systemd/user")"
        for RC in .bashrc .bash_profile .profile .zshrc .zprofile; do
            _persist_add "shell-init" "$HOME_DIR/$RC"
        done
    done < <(get_linux_user_homes)
    separator
    info "$(L "Artefatti di persistenza:" "Persistence artifacts:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun artefatto di persistenza trovato." "No persistence artifact found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Artefatti" "Artifacts")" "$FOUND" "info")
    finish_report "linux_persistence" "Linux Persistence" "PER" "cron · systemd · autostart · rc · shell-init" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 9 — SSH Artifacts ---
module_linux_ssh() {
    section_header "Linux — SSH Artifacts" "$YELLOW"
    check_target_root || return 1
    local KW="ssh-rsa|ssh-ed25519|ecdsa|PermitRootLogin|PasswordAuthentication|AuthorizedKeys|ForwardAgent"
    local BODY="" FOUND=0
    # sshd_config di sistema
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc/ssh")
    if [[ -n "$ETC" ]]; then
        local SC; SC=$(ci_find_file "$ETC" "sshd_config")
        [[ -n "$SC" && -s "$SC" ]] && { FOUND=$((FOUND + 1)); ok "sshd_config"; print_file_lines "$SC" "PermitRootLogin|PasswordAuthentication|AllowUsers" 40; echo ""; BODY+=$(file_card_html "$SC" "$KW" "⚿"); }
    fi
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local SSHD; SSHD=$(ci_find_dir "$HOME_DIR" ".ssh")
        [[ -z "$SSHD" ]] && continue
        local CARDS="" UCOUNT=0
        for KF in authorized_keys authorized_keys2 known_hosts config; do
            local F="$SSHD/$KF"
            [[ -f "$F" && -s "$F" ]] || continue
            UCOUNT=$((UCOUNT + 1)); FOUND=$((FOUND + 1)); CARDS+=$(file_card_html "$F" "$KW" "⚿")
        done
        # chiavi private presenti (solo presenza, non contenuto)
        local PRIV; PRIV=$(find "$SSHD" -maxdepth 1 -type f -name "id_*" ! -name "*.pub" 2>/dev/null | wc -l)
        [[ "$PRIV" -gt 0 ]] && CARDS+="<div class='card' style='margin-bottom:.8rem'><div class='card-header'><div class='uicon'>⚿</div><div><div class='uname'>$(L "Chiavi private" "Private keys")</div><div class='upath'>$SSHD/id_*</div></div><div class='badge warn'>$PRIV $(L "chiavi" "keys")</div></div></div>"
        [[ $UCOUNT -eq 0 && "$PRIV" -eq 0 ]] && continue
        ok "$UNAME — ${BOLD}$UCOUNT file"
        BODY+=$(generic_card_html "$UNAME" "$SSHD" "$UCOUNT file" "$CARDS" "⚿")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Artefatti SSH:" "SSH artifacts:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun artefatto SSH trovato." "No SSH artifact found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File SSH" "SSH files")" "$FOUND" "info")
    finish_report "linux_ssh" "Linux SSH Artifacts" "SSH" "authorized_keys · known_hosts · sshd_config" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 10 — Network Config ---
module_linux_network() {
    section_header "Linux — Network Config" "$CYAN"
    check_target_root || return 1
    local KW="psk|password|key=|wpa|pre-shared|nameserver|DROP|ACCEPT|REJECT|-j |MASQUERADE"
    local BODY="" FOUND=0
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    [[ -z "$ETC" ]] && { warn "$(L "Directory etc non trovata." "etc directory not found.")"; return 0; }
    for NF in hosts resolv.conf hostname; do
        local F; F=$(ci_find_file "$ETC" "$NF")
        [[ -n "$F" && -s "$F" ]] || continue
        FOUND=$((FOUND + 1)); ok "$NF"; BODY+=$(file_card_html "$F" "$KW" "⇄")
    done
    # NetworkManager connections (contengono PSK Wi-Fi)
    local NM; NM=$(ci_find_dir "$ETC" "NetworkManager/system-connections")
    if [[ -n "$NM" ]]; then
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); ok "NM: $(basename "$F")"
            BODY+=$(file_card_html "$F" "$KW" "⇄")
        done < <(find "$NM" -maxdepth 1 -type f 2>/dev/null)
    fi
    # netplan
    local NP; NP=$(ci_find_dir "$ETC" "netplan")
    [[ -n "$NP" ]] && while IFS= read -r F; do
        [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); BODY+=$(file_card_html "$F" "$KW" "⇄")
    done < <(find "$NP" -maxdepth 1 -type f 2>/dev/null)
    # iptables/nftables salvate
    for RF in "iptables/rules.v4" "iptables/rules.v6" "nftables.conf"; do
        local F; F=$(ci_find_file "$ETC" "$(basename "$RF")")
        [[ -n "$F" && -s "$F" ]] || continue; FOUND=$((FOUND + 1)); ok "$(basename "$RF")"
        BODY+=$(file_card_html "$F" "$KW" "⇄")
    done
    separator
    info "$(L "File di rete:" "Network files:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun file di configurazione di rete." "No network configuration file.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File rete" "Network files")" "$FOUND" "info")
    finish_report "linux_network" "Linux Network Config" "NET" "hosts · NetworkManager · netplan · iptables" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 11 — Installed Packages ---
module_linux_packages() {
    section_header "Linux — Installed Packages" "$GREEN"
    check_target_root || return 1
    local BODY="" FOUND=0
    # dpkg
    local DPKG; DPKG=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "var/lib/dpkg")" "status")
    if [[ -n "$DPKG" && -s "$DPKG" ]]; then
        FOUND=$((FOUND + 1))
        local PKGS; PKGS=$(grep -c '^Package:' "$DPKG" 2>/dev/null)
        ok "dpkg — ${BOLD}$PKGS pacchetti"
        local TAB; TAB=$(awk '/^Package:/{p=$2} /^Version:/{v=$2} /^Status:/{s=$0} /^$/{if(p)printf "<tr><td class=mono>%s</td><td class=\"mono dim\">%s</td></tr>\n",p,v; p=v=""}' "$DPKG" | head -100000)
        BODY+=$(generic_card_html "dpkg packages" "$DPKG" "$PKGS pkg" "<table><tr><th>Package</th><th>Version</th></tr>$TAB</table>" "▦")
    fi
    # apt history (timeline installazioni)
    local APTLOG; APTLOG=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "var/log/apt")" "history.log")
    if [[ -n "$APTLOG" && -s "$APTLOG" ]]; then
        FOUND=$((FOUND + 1)); ok "apt history.log"
        print_file_lines "$APTLOG" "Install:|Remove:|Purge:" 40; echo ""
        BODY+=$(file_card_html "$APTLOG" "Install:|Remove:|Purge:|Commandline:" "▦")
    fi
    # rpm (db binario: solo presenza + tentativo via rpm se disponibile)
    local RPMDB; RPMDB=$(ci_find_dir "$WIN_ROOT" "var/lib/rpm")
    if [[ -n "$RPMDB" ]]; then
        FOUND=$((FOUND + 1)); ok "rpm db $(L "presente" "present")"
        local RTAB=""
        if command -v rpm >/dev/null 2>&1; then
            RTAB=$(rpm -qa --dbpath "$RPMDB" 2>/dev/null | sort | head -100000 | "$PY3" -c 'import sys,html;[print(f"<tr><td class=mono>{html.escape(l.strip())}</td></tr>") for l in sys.stdin]')
        fi
        BODY+=$(generic_card_html "rpm packages" "$RPMDB" "" "<table><tr><th>Package</th></tr>${RTAB:-<tr><td class=dim>rpm CLI non disponibile per dump offline</td></tr>}</table>" "▦")
    fi
    # snap / flatpak (elenco directory)
    local SNAP; SNAP=$(ci_find_dir "$WIN_ROOT" "var/lib/snapd/snaps")
    [[ -n "$SNAP" ]] && { FOUND=$((FOUND + 1)); ok "snap"; local STAB; STAB=$(find "$SNAP" -maxdepth 1 -name "*.snap" 2>/dev/null | sort | "$PY3" -c 'import sys,html,os;[print(f"<tr><td class=mono>{html.escape(os.path.basename(l.strip()))}</td></tr>") for l in sys.stdin]'); BODY+=$(generic_card_html "snap" "$SNAP" "" "<table><tr><th>Snap</th></tr>$STAB</table>" "▦"); }
    separator
    info "$(L "Fonti pacchetti:" "Package sources:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun database pacchetti trovato." "No package database found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Fonti" "Sources")" "$FOUND" "info")
    finish_report "linux_packages" "Linux Installed Packages" "PKG" "dpkg · rpm · apt history · snap" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 12 — Trash & Recent ---
module_linux_trash() {
    section_header "Linux — Trash & Recent" "$GREEN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        # Trash info files (.trashinfo: path + DeletionDate)
        local TINFO; TINFO=$(ci_find_dir "$HOME_DIR" ".local/share/Trash/info")
        if [[ -n "$TINFO" ]]; then
            local TTAB="" TN=0
            while IFS= read -r TF; do
                [[ -s "$TF" ]] || continue; TN=$((TN + 1))
                local OPATH DDATE
                OPATH=$(grep -m1 '^Path=' "$TF" 2>/dev/null | cut -d= -f2-)
                DDATE=$(grep -m1 '^DeletionDate=' "$TF" 2>/dev/null | cut -d= -f2-)
                TTAB+="<tr><td class='mono dim'>$(html_esc "$DDATE")</td><td class='mono'>$(html_esc "$OPATH")</td></tr>"
            done < <(find "$TINFO" -maxdepth 1 -name "*.trashinfo" 2>/dev/null)
            if [[ $TN -gt 0 ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + TN))
                ok "$UNAME — ${BOLD}$TN $(L "file cestinati" "trashed files")"
                CARDS+=$(generic_card_html "Trash" "$TINFO" "$TN file" "<table><tr><th>$(L "Cancellato il" "Deleted on")</th><th>$(L "Percorso originale" "Original path")</th></tr>$TTAB</table>" "♺")
            fi
        fi
        # recently-used.xbel
        local XBEL; XBEL=$(ci_find_file "$HOME_DIR" ".local/share/recently-used.xbel")
        [[ -z "$XBEL" ]] && XBEL=$(ci_find_file "$HOME_DIR" ".recently-used.xbel")
        if [[ -n "$XBEL" && -s "$XBEL" ]]; then
            UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
            CARDS+=$(file_card_html "$XBEL" "href|exec|modified|visited" "♺")
        fi
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1))
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT" "$CARDS" "♺")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Voci trovate:" "Items found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun elemento nel cestino / recenti." "No trash / recent items.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Voci" "Items")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_trash" "Linux Trash & Recent" "TRH" "~/.local/share/Trash · recently-used.xbel" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 13 — Filesystem Timeline ---
module_linux_timeline() {
    section_header "Linux — Filesystem Timeline" "$YELLOW"
    check_target_root || return 1
    info "$(L "Raccolta timestamp MAC nelle aree sensibili (può richiedere tempo)..." "Collecting MAC timestamps in sensitive areas (may take a while)...")"
    # aree di interesse forense
    local AREAS=(tmp var/tmp dev/shm root home etc/cron.d etc/systemd/system usr/local/bin var/www)
    local TMPF; TMPF=$(mktemp)
    local SCANNED=0
    for A in "${AREAS[@]}"; do
        local D; D=$(ci_find_dir "$WIN_ROOT" "$A")
        [[ -z "$D" ]] && continue
        SCANNED=$((SCANNED + 1))
        find "$D" -xdev -type f -printf '%T+\t%A+\t%C+\t%s\t%p\n' 2>/dev/null | head -100000 >> "$TMPF"
    done
    local N; N=$(wc -l < "$TMPF" 2>/dev/null || echo 0)
    if [[ "$N" -eq 0 ]]; then warn "$(L "Nessun file nelle aree sensibili." "No files in sensitive areas.")"; rm -f "$TMPF"; return 0; fi
    ok "$(L "File raccolti:" "Files collected:") ${BOLD}$N"
    # ordina per mtime desc, mostra anteprima
    sort -r "$TMPF" -o "$TMPF"
    head -25 "$TMPF" | while IFS=$'\t' read -r MT AT CT SZ P; do
        echo -e "      ${DIM}$MT${RESET}  ${P#$WIN_ROOT}"
    done
    separator
    ask_yn "Generare report HTML?" || { rm -f "$TMPF"; return 0; }
    local TABLE; TABLE=$("$PY3" - "$TMPF" "$WIN_ROOT" << 'PYEOF'
import sys, html
root=sys.argv[2]
print("<table><tr><th>Modified</th><th>Accessed</th><th>Changed</th><th>Size</th><th>Path</th></tr>")
with open(sys.argv[1]) as f:
    for line in f:
        c=line.rstrip('\n').split('\t')
        if len(c)<5: continue
        mt,at,ct,sz,p=c[0],c[1],c[2],c[3],c[4]
        rel=p[len(root):] if p.startswith(root) else p
        print(f"<tr><td class='mono dim'>{html.escape(mt)}</td><td class='mono dim'>{html.escape(at)}</td><td class='mono dim'>{html.escape(ct)}</td><td class='mono'>{sz}</td><td class='mono'>{html.escape(rel)}</td></tr>")
print("</table>")
PYEOF
)
    rm -f "$TMPF"
    local BODY; BODY=$(generic_card_html "$(L "Timeline aree sensibili" "Sensitive areas timeline")" "$WIN_ROOT" "$N file" "$TABLE" "◷")
    local STATS; STATS="$(stat_box "$(L "File" "Files")" "$N")$(stat_box "$(L "Aree" "Areas")" "$SCANNED" "info")"
    finish_report "linux_timeline" "Linux Filesystem Timeline" "TML" "MAC times (find/stat)" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- LINUX 15 — auditd (/var/log/audit) ---
#
# Su RHEL/CentOS/Fedora e su ogni sistema hardenizzato auditd e' la fonte
# primaria di intrusion detection: registra syscall, esecuzioni, autenticazioni
# e violazioni di policy con un dettaglio che syslog non ha. Il modulo "System
# Logs" copre /var/log testuali ma non tocca audit.log, che ha un formato
# proprio (campi chiave=valore, stringhe in esadecimale, timestamp epoch).
module_linux_auditd() {
    section_header "Linux — auditd" "$RED"
    check_target_root || return 1

    local AUDITDIR; AUDITDIR=$(ci_find_dir "$WIN_ROOT" "var/log/audit")
    if [[ -z "$AUDITDIR" ]]; then
        warn "$(L "Directory /var/log/audit non trovata (auditd non installato o log altrove)." \
                 "/var/log/audit directory not found (auditd not installed or logs elsewhere).")"
        return 0
    fi
    mapfile -t LOGS < <(find "$AUDITDIR" -maxdepth 1 -type f -name 'audit.log*' -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
    if [[ ${#LOGS[@]} -eq 0 ]]; then
        warn "$(L "Nessun audit.log presente." "No audit.log present.")"
        return 0
    fi
    info "$(L "File di audit trovati:" "Audit files found:") ${BOLD}${#LOGS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local SUM; SUM=$(mktemp); register_tmp "$SUM"
    "$PY3" - "$OUT" "$SUM" "${LOGS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, datetime, collections, binascii

out_path, sum_path = sys.argv[1], sys.argv[2]
files = sys.argv[3:]

HDR = re.compile(r'type=(?P<type>\S+)\s+msg=audit\((?P<epoch>\d+)\.(?P<ms>\d+):(?P<serial>\d+)\):\s*(?P<rest>.*)')
KV  = re.compile(r'(\w+)=("([^"]*)"|\S+)')

# Campi che auditd puo' emettere in esadecimale quando contengono spazi o
# caratteri speciali: senza decodifica il comando dell'attaccante resta
# illeggibile proprio nei casi piu' interessanti.
HEXFIELDS = {"proctitle", "exe", "cmd", "comm", "name", "cwd", "a0", "a1", "a2", "a3"}

# Tipi che pesano in un'indagine.
NOTABLE = {
    "USER_AUTH", "USER_LOGIN", "USER_ACCT", "USER_START", "USER_CMD",
    "ADD_USER", "DEL_USER", "ADD_GROUP", "DEL_GROUP", "USER_CHAUTHTOK",
    "USER_ROLE_CHANGE", "ROLE_ASSIGN", "ANOM_ABEND", "AVC", "SECCOMP",
    "EXECVE", "ANOM_PROMISCUOUS", "CONFIG_CHANGE", "MAC_POLICY_LOAD",
}

def unhex(val):
    """auditd codifica alcuni valori come esadecimale puro (senza virgolette)."""
    if len(val) >= 4 and len(val) % 2 == 0 and re.fullmatch(r'[0-9A-Fa-f]+', val):
        try:
            return binascii.unhexlify(val).decode('utf-8', 'replace').replace('\x00', ' ').strip()
        except Exception:
            return val
    return val

counts = collections.Counter()
rows = []

for fp in files:
    try:
        fh = open(fp, encoding='utf-8', errors='replace')
    except Exception:
        continue
    with fh:
        for line in fh:
            m = HDR.match(line.strip())
            if not m:
                continue
            typ = m.group('type')
            counts[typ] += 1
            try:
                ts = datetime.datetime.utcfromtimestamp(int(m.group('epoch'))).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                ts = ''
            fields = {}
            for km in KV.finditer(m.group('rest')):
                key = km.group(1)
                if km.group(3) is not None:
                    val = km.group(3)          # valore fra virgolette
                else:
                    # Valore non quotato. auditd annida i sottocampi dentro
                    # msg='...': l'ultimo di questi si porta dietro l'apice di
                    # chiusura (res=failed'), che va tolto o il confronto salta.
                    val = km.group(2).strip().rstrip("',")
                    if key in HEXFIELDS:
                        val = unhex(val)
                fields[key] = val

            # Ricompone il comando dagli argomenti a0..aN di EXECVE.
            if typ == 'EXECVE':
                args = []
                i = 0
                while f'a{i}' in fields:
                    args.append(fields[f'a{i}'])
                    i += 1
                detail = ' '.join(args) if args else m.group('rest')[:300]
            else:
                parts = []
                for key in ('res', 'acct', 'uid', 'auid', 'exe', 'comm', 'terminal',
                            'hostname', 'addr', 'op', 'key', 'proctitle', 'cmd', 'name'):
                    if key in fields and fields[key] not in ('', '?', '(none)'):
                        parts.append(f"{key}={fields[key]}")
                detail = ' '.join(parts) or m.group('rest')[:300]

            user = fields.get('acct') or fields.get('auid') or fields.get('uid') or ''
            failed = '1' if fields.get('res') in ('failed', 'fail', '0') else '0'
            notable = '1' if (typ in NOTABLE or failed == '1') else '0'
            rows.append((ts, typ, str(user), detail[:500], notable, failed))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write("\t".join(x.replace("\t", " ").replace("\n", " ") for x in r) + "\n")

with open(sum_path, 'w', encoding='utf-8') as fh:
    for typ, n in counts.most_common():
        fh.write(f"{typ}\t{n}\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun record auditd interpretabile." "No parsable auditd record.")"
        return 0
    fi
    local NFAIL NNOTE
    NFAIL=$(awk -F'\t' '$6=="1"' "$OUT" | wc -l)
    NNOTE=$(awk -F'\t' '$5=="1"' "$OUT" | wc -l)
    local NEXEC; NEXEC=$(awk -F'\t' '$2=="EXECVE"' "$OUT" | wc -l)

    ok "$(L "Record auditd:" "auditd records:") ${BOLD}$TOTAL"
    info "EXECVE: ${BOLD}${NEXEC}${RESET}  ·  $(L "esiti negativi:" "failed outcomes:") ${BOLD}${NFAIL}"
    echo ""
    info "$(L "Distribuzione per tipo:" "Breakdown by type:")"
    head -12 "$SUM" | while IFS=$'\t' read -r T N; do
        printf "      ${DIM}%-24s %6s${RESET}\n" "$T" "$N"
    done

    ask_yn "Generare report HTML?" || return 0

    # Prima gli eventi rilevanti, poi il resto.
    local ROWS; ROWS=$( { awk -F'\t' '$5=="1"{print $1"\t"$2"\t"$3"\t"$4}' "$OUT";
                          awk -F'\t' '$5=="0"{print $1"\t"$2"\t"$3"\t"$4}' "$OUT"; } | head -20000 )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Data (UTC)" "Date (UTC)")" "$(L "Tipo" "Type")" "$(L "Utente" "User")" "$(L "Dettaglio" "Detail")")
    local SUMTABLE; SUMTABLE=$(_rows_to_table "$(cat "$SUM")" "$(L "Tipo" "Type")" "$(L "Occorrenze" "Occurrences")")

    local BODY="<div class='cards'>"
    BODY+=$(generic_card_html "$(L "Distribuzione per tipo" "Breakdown by type")" "$AUDITDIR" "$TOTAL" "$SUMTABLE" "∑")
    BODY+=$(generic_card_html "$(L "Eventi" "Events")" "$(L "rilevanti in testa" "notable first")" "$TOTAL" "$TABLE" "⚑")
    BODY+="</div>"

    local STATS
    STATS="$(stat_box "$(L "Record" "Records")" "$TOTAL")"
    STATS+="$(stat_box "EXECVE" "$NEXEC" "info")"
    STATS+="$(stat_box "$(L "Esiti negativi" "Failed")" "$NFAIL" "warn")"
    STATS+="$(stat_box "$(L "Rilevanti" "Notable")" "$NNOTE" "warn")"
    finish_report "linux_auditd" "Linux auditd" "AUD" "/var/log/audit/audit.log" "$STATS" "$BODY"
}

# --- LINUX 16 — Container (Docker / Podman) ---
#
# Un host con container ha un intero piano di esecuzione che gli altri moduli
# non vedono: processi, filesystem e persistenza vivono dentro le immagini.
# Qui si ricostruisce l'inventario offline dai metadati sul disco e si
# evidenziano le configurazioni che permettono la fuga dal container verso
# l'host — privileged, mount di / o del socket Docker, CAP_SYS_ADMIN,
# condivisione del namespace PID/rete dell'host.
module_linux_containers() {
    section_header "Linux — Container (Docker / Podman)" "$BLUE"
    check_target_root || return 1

    local DOCKER PODMAN
    DOCKER=$(ci_find_dir "$WIN_ROOT" "var/lib/docker")
    PODMAN=$(ci_find_dir "$WIN_ROOT" "var/lib/containers/storage")
    if [[ -z "$DOCKER" && -z "$PODMAN" ]]; then
        warn "$(L "Nessun runtime container trovato (/var/lib/docker, /var/lib/containers)." \
                 "No container runtime found (/var/lib/docker, /var/lib/containers).")"
        return 0
    fi
    [[ -n "$DOCKER" ]] && info "Docker: ${BOLD}${DOCKER}"
    [[ -n "$PODMAN" ]] && info "Podman: ${BOLD}${PODMAN}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${DOCKER:-}" "${PODMAN:-}" << 'PYEOF' 2>/dev/null
import sys, os, json, glob

out_path = sys.argv[1]
docker   = sys.argv[2] if len(sys.argv) > 2 else ''
podman   = sys.argv[3] if len(sys.argv) > 3 else ''

rows = []

def risk_of(hostconfig, config):
    """Indicatori di fuga dal container verso l'host."""
    risks = []
    hc = hostconfig or {}
    if hc.get('Privileged'):
        risks.append('PRIVILEGED')
    for b in (hc.get('Binds') or []):
        src = str(b).split(':')[0]
        if src == '/':
            risks.append('MOUNT_ROOT_HOST')
        elif 'docker.sock' in str(b):
            risks.append('DOCKER_SOCKET')
        elif src in ('/etc', '/var/run', '/proc', '/sys', '/boot'):
            risks.append(f'MOUNT_{src.strip("/").upper()}')
    caps = hc.get('CapAdd') or []
    for c in caps:
        if str(c).upper().replace('CAP_', '') in ('SYS_ADMIN', 'SYS_PTRACE', 'SYS_MODULE', 'ALL'):
            risks.append(f'CAP_{str(c).upper().replace("CAP_", "")}')
    if str(hc.get('PidMode', '')) == 'host':
        risks.append('PID_HOST')
    if str(hc.get('NetworkMode', '')) == 'host':
        risks.append('NET_HOST')
    if str(hc.get('IpcMode', '')) == 'host':
        risks.append('IPC_HOST')
    return risks

# ---- Docker -------------------------------------------------------------
if docker and os.path.isdir(docker):
    for cdir in sorted(glob.glob(os.path.join(docker, 'containers', '*'))):
        cfg_path = os.path.join(cdir, 'config.v2.json')
        hc_path  = os.path.join(cdir, 'hostconfig.json')
        if not os.path.isfile(cfg_path):
            continue
        try:
            cfg = json.load(open(cfg_path, encoding='utf-8', errors='replace'))
        except Exception:
            continue
        hc = {}
        if os.path.isfile(hc_path):
            try:
                hc = json.load(open(hc_path, encoding='utf-8', errors='replace'))
            except Exception:
                hc = {}
        name  = str(cfg.get('Name', '')).lstrip('/')
        image = cfg.get('Config', {}).get('Image') or cfg.get('Image', '')
        created = str(cfg.get('Created', ''))[:19].replace('T', ' ')
        state = cfg.get('State', {}) or {}
        started  = str(state.get('StartedAt', ''))[:19].replace('T', ' ')
        finished = str(state.get('FinishedAt', ''))[:19].replace('T', ' ')
        running = 'running' if state.get('Running') else 'stopped'
        cmd = ' '.join((cfg.get('Config', {}) or {}).get('Cmd') or [])
        entry = ' '.join((cfg.get('Config', {}) or {}).get('Entrypoint') or [])
        risks = risk_of(hc, cfg)
        rows.append((
            'docker', created or started, name, str(image), running,
            (entry + ' ' + cmd).strip()[:300], ';'.join(risks), os.path.basename(cdir)[:12],
            started, finished,
        ))

# ---- Podman -------------------------------------------------------------
if podman and os.path.isdir(podman):
    cjson = os.path.join(podman, 'overlay-containers', 'containers.json')
    if os.path.isfile(cjson):
        try:
            for c in json.load(open(cjson, encoding='utf-8', errors='replace')):
                rows.append((
                    'podman', str(c.get('created', ''))[:19].replace('T', ' '),
                    str(c.get('names', [''])[0] if c.get('names') else ''),
                    str(c.get('image', '')), '', '', '', str(c.get('id', ''))[:12], '', '',
                ))
        except Exception:
            pass

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write("\t".join(str(x).replace("\t", " ").replace("\n", " ") for x in r) + "\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun container ricostruibile dai metadati." "No container reconstructable from metadata.")"
        return 0
    fi
    local NRISK; NRISK=$(awk -F'\t' '$7!=""' "$OUT" | wc -l)

    ok "$(L "Container trovati:" "Containers found:") ${BOLD}$TOTAL"
    if [[ "$NRISK" -gt 0 ]]; then
        warn "$(L "Container con configurazione a rischio di fuga:" "Containers with escape-prone configuration:") ${BOLD}$NRISK"
        awk -F'\t' '$7!=""{printf "      %s  [%s]  %s\n", $3, $7, $4}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    else
        info "$(L "Nessun indicatore di fuga rilevato." "No escape indicator detected.")"
    fi

    # Log stdout dei container: spesso contengono l'attivita' dell'attaccante.
    local NLOGS=0
    if [[ -n "$DOCKER" ]]; then
        NLOGS=$(find "$DOCKER/containers" -maxdepth 2 -name '*-json.log' -size +0 2>/dev/null | wc -l)
        [[ "$NLOGS" -gt 0 ]] && info "$(L "Log stdout disponibili:" "stdout logs available:") ${BOLD}${NLOGS}"
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7"\t"$8}' "$OUT")
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "Runtime" "$(L "Creato" "Created")" "$(L "Nome" "Name")" "Image" "$(L "Stato" "State")" \
        "$(L "Comando" "Command")" "$(L "Rischi" "Risks")" "ID")

    local NOTE=""
    if [[ "$NRISK" -gt 0 ]]; then
        NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
        NOTE+="<b>$(L "Indicatori di fuga dal container" "Container escape indicators")</b><br>"
        NOTE+="$(L "PRIVILEGED e MOUNT_ROOT_HOST danno di fatto accesso completo all'host. DOCKER_SOCKET consente di creare nuovi container privilegiati. CAP_SYS_ADMIN e CAP_SYS_MODULE permettono di caricare moduli kernel. PID_HOST espone i processi dell'host." \
            "PRIVILEGED and MOUNT_ROOT_HOST effectively grant full host access. DOCKER_SOCKET allows spawning new privileged containers. CAP_SYS_ADMIN and CAP_SYS_MODULE allow loading kernel modules. PID_HOST exposes host processes.")"
        NOTE+="</div></div>"
    fi

    local STATS
    STATS="$(stat_box "Container" "$TOTAL")"
    STATS+="$(stat_box "$(L "A rischio" "At risk")" "$NRISK" "$([[ "$NRISK" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Log stdout" "stdout logs")" "$NLOGS" "info")"
    finish_report "linux_containers" "Linux Container Forensics" "CNT" "/var/lib/docker · /var/lib/containers" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Inventario container" "Container inventory")" "${DOCKER:-$PODMAN}" "$TOTAL" "$TABLE" "▣")</div>"
}

# ================================================================
#  LINUX 17 — PAM (Pluggable Authentication Modules)
#
#  PAM decide chi puo' autenticarsi e come. Una backdoor PAM e' fra le
#  persistenze piu' efficaci su Linux: aggiungendo una riga a /etc/pam.d/sshd
#  o sostituendo pam_unix.so con una versione modificata, l'attaccante ottiene
#  accesso con una password universale che nessun log segnala come anomala —
#  l'autenticazione risulta semplicemente riuscita.
#
#  Il modulo Persistence copre cron, systemd e autostart, ma non tocca PAM.
#  Qui si guardano tre cose: righe di configurazione sospette, moduli .so che
#  non appartengono a nessun pacchetto, e moduli con data di modifica
#  disallineata rispetto ai vicini.
# ================================================================
module_linux_pam() {
    section_header "Linux — PAM" "$RED"
    check_target_root || return 1

    local PAMD; PAMD=$(ci_find_dir "$WIN_ROOT" "etc/pam.d")
    local PAMCONF; PAMCONF=$(ci_find_file "$WIN_ROOT" "etc/pam.conf")
    if [[ -z "$PAMD" && -z "$PAMCONF" ]]; then
        warn "$(L "Configurazione PAM non trovata." "PAM configuration not found.")"
        return 0
    fi

    # Righe che meritano attenzione in una configurazione PAM.
    #  pam_exec      esegue un comando esterno a ogni autenticazione
    #  pam_permit    concede l'accesso incondizionatamente
    #  nullok        accetta password vuote
    #  pam_python /
    #  percorsi non standard  moduli caricati da fuori la directory di sistema
    local KW="pam_exec|pam_permit|nullok|pam_python|pam_script|/tmp/|/dev/shm|/home/|pam_succeed_if.*quiet.*uid|debug"

    local BODY="" NFILE=0 NSUSP=0 SUSPROWS=""
    local F
    if [[ -n "$PAMD" ]]; then
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue
            NFILE=$((NFILE + 1))
            local N; N=$(grep -cE "$KW" "$F" 2>/dev/null || echo 0)
            if [[ "$N" -gt 0 ]]; then
                NSUSP=$((NSUSP + 1))
                ok "$(basename "$F") — ${BOLD}${N}${RESET} $(L "righe da verificare" "lines to check")"
                local LN
                while IFS= read -r LN; do
                    SUSPROWS+="$(basename "$F")	${LN}
"
                done < <(grep -nE "$KW" "$F" 2>/dev/null | head -20)
            fi
            BODY+=$(file_card_html "$F" "$KW" "PAM")
        done < <(find "$PAMD" -maxdepth 1 -type f 2>/dev/null | sort)
    fi
    if [[ -n "$PAMCONF" && -s "$PAMCONF" ]]; then
        NFILE=$((NFILE + 1))
        BODY+=$(file_card_html "$PAMCONF" "$KW" "PAM")
    fi

    # Moduli PAM sul disco: un .so non pacchettizzato o con mtime fuori linea
    # rispetto ai vicini e' il segnale piu' concreto di manomissione.
    local SOROWS="" NSO=0 NORPH=0
    local D
    for D in "lib/x86_64-linux-gnu/security" "lib64/security" "lib/security" \
             "usr/lib/x86_64-linux-gnu/security" "usr/lib64/security" "usr/lib/security"; do
        local SD; SD=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$SD" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue
            NSO=$((NSO + 1))
            local MT SZ OWN
            MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
            SZ=$(stat -c %s "$F" 2>/dev/null)
            # Appartenenza a un pacchetto: si verifica offline nei database
            # di dpkg (list dei file) e rpm non e' interrogabile senza il DB,
            # quindi qui si usa solo dpkg quando presente.
            OWN="?"
            local DPKGINFO; DPKGINFO=$(ci_find_dir "$WIN_ROOT" "var/lib/dpkg/info")
            if [[ -n "$DPKGINFO" ]]; then
                local REL="${F#$WIN_ROOT}"
                if grep -rqsF "$REL" "$DPKGINFO"/*.list 2>/dev/null; then
                    OWN="$(L "pacchetto" "package")"
                else
                    OWN="$(L "NON PACCHETTIZZATO" "NOT PACKAGED")"
                    NORPH=$((NORPH + 1))
                fi
            fi
            SOROWS+="${MT}	$(basename "$F")	${SZ}	${OWN}	${F}
"
        done < <(find "$SD" -maxdepth 1 -type f -name '*.so' 2>/dev/null | sort)
    done

    separator
    info "$(L "File di configurazione PAM:" "PAM configuration files:") ${BOLD}$NFILE"
    [[ $NSUSP -gt 0 ]] && warn "$(L "File con righe da verificare:" "Files with lines to check:") ${BOLD}$NSUSP"
    info "$(L "Moduli PAM (.so):" "PAM modules (.so):") ${BOLD}$NSO"
    if [[ "$NORPH" -gt 0 ]]; then
        warn "$(L "Moduli non appartenenti ad alcun pacchetto:" "Modules not owned by any package:") ${BOLD}$NORPH"
    fi
    [[ $NFILE -eq 0 && $NSO -eq 0 ]] && return 0

    ask_yn "Generare report HTML?" || return 0

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Come leggere le segnalazioni" "How to read the flags")</b><br>"
    NOTE+="$(L "Le parole chiave evidenziate non sono di per se' malevole: pam_exec e nullok hanno usi legittimi. Vanno lette come 'questa riga cambia chi puo' entrare e come, quindi va giustificata'. Il segnale piu' forte e' invece un modulo .so che non appartiene a nessun pacchetto, o la cui data di modifica e' isolata rispetto agli altri moduli della stessa directory." \
        "The highlighted keywords are not malicious in themselves: pam_exec and nullok have legitimate uses. Read them as 'this line changes who can log in and how, so it needs justifying'. The stronger signal is a .so module owned by no package, or whose modification date stands alone among the other modules in the same directory.")<br><br>"
    NOTE+="$(L "Una backdoor PAM non lascia tracce nei log di autenticazione: l'accesso risulta semplicemente riuscito." \
        "A PAM backdoor leaves no trace in authentication logs: the login simply succeeds.")"
    NOTE+="</div></div>"

    local CARDS=""
    if [[ -n "$SUSPROWS" ]]; then
        CARDS+="<div class='cards'>$(generic_card_html "$(L "Righe da verificare" "Lines to check")" "/etc/pam.d" "$NSUSP" \
            "$(_rows_to_table "$SUSPROWS" "$(L "File" "File")" "$(L "Riga" "Line")")" "⚑")</div>"
    fi
    if [[ -n "$SOROWS" ]]; then
        CARDS+="<div class='cards'>$(generic_card_html "$(L "Moduli PAM installati" "Installed PAM modules")" "security/*.so" "$NSO" \
            "$(_rows_to_table "$SOROWS" "$(L "Ultima modifica" "Last modified")" "$(L "Modulo" "Module")" "$(L "Byte" "Bytes")" "$(L "Origine" "Origin")" "$(L "Percorso" "Path")")" "◧")</div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Config" "Config")" "$NFILE")"
    STATS+="$(stat_box "$(L "Da verificare" "To check")" "$NSUSP" "$([[ $NSUSP -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Moduli .so" "Modules .so")" "$NSO" "info")"
    STATS+="$(stat_box "$(L "Non pacchettizzati" "Not packaged")" "$NORPH" "$([[ $NORPH -gt 0 ]] && echo warn || echo info)")"
    finish_report "linux_pam" "Linux PAM" "PAM" "/etc/pam.d · security/*.so" "$STATS" \
        "${NOTE}${CARDS}<div class='cards'>${BODY}</div>"
}

# ================================================================
#  LINUX 18 — Moduli kernel e rootkit
#
#  Un LKM rootkit gira in kernel space: puo' nascondere processi, file,
#  connessioni e se stesso agli strumenti in user space. Nessun modulo che
#  legge /proc o esegue comandi sulla macchina viva puo' vederlo — l'analisi
#  offline del filesystem, invece, si'.
#
#  Il modulo Persistence copre solo ld.so.preload. Qui si guardano i punti da
#  cui un modulo viene caricato all'avvio e i .ko presenti sul disco che non
#  appartengono a nessun pacchetto.
# ================================================================
module_linux_kernel_modules() {
    section_header "Linux — $(L "Moduli kernel e rootkit" "Kernel modules and rootkits")" "$RED"
    check_target_root || return 1

    local BODY="" ROWS="" NCFG=0 NSUSP=0
    local F D

    # --- Punti di caricamento automatico ---------------------------------
    local KW="^[a-zA-Z0-9_-]+"
    for D in "etc/modules-load.d" "etc/modprobe.d" "usr/lib/modules-load.d" "lib/modules-load.d"; do
        local DIR; DIR=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$DIR" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue
            NCFG=$((NCFG + 1))
            BODY+=$(file_card_html "$F" "install |^blacklist|/tmp/|/dev/shm|insmod|modprobe" "KO")
        done < <(find "$DIR" -maxdepth 1 -type f 2>/dev/null | sort)
    done
    F=$(ci_find_file "$WIN_ROOT" "etc/modules")
    if [[ -s "$F" ]]; then
        NCFG=$((NCFG + 1))
        BODY+=$(file_card_html "$F" "$KW" "KO")
    fi

    # "install <modulo> <comando>" in modprobe.d esegue un comando arbitrario
    # al caricamento del modulo: e' una persistenza poco nota e molto efficace.
    local INSTROWS=""
    for D in "etc/modprobe.d" "usr/lib/modprobe.d" "lib/modprobe.d"; do
        local DIR; DIR=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$DIR" ]] && continue
        while IFS= read -r LN; do
            [[ -n "$LN" ]] && { INSTROWS+="${LN}
"; NSUSP=$((NSUSP + 1)); }
        done < <(grep -rhnE '^[[:space:]]*install[[:space:]]+\S+[[:space:]]+\S' "$DIR" 2>/dev/null | grep -vE '/bin/(true|false)[[:space:]]*$' | head -50)
    done

    # --- Moduli .ko sul disco --------------------------------------------
    local NKO=0 NORPH=0 KOROWS=""
    local MODDIR; MODDIR=$(ci_find_dir "$WIN_ROOT" "lib/modules")
    [[ -z "$MODDIR" ]] && MODDIR=$(ci_find_dir "$WIN_ROOT" "usr/lib/modules")
    if [[ -n "$MODDIR" ]]; then
        local DPKGINFO; DPKGINFO=$(ci_find_dir "$WIN_ROOT" "var/lib/dpkg/info")
        # Un rootkit installato a mano finisce quasi sempre in extra/ o misc/,
        # o direttamente nella radice della versione del kernel.
        while IFS= read -r F; do
            NKO=$((NKO + 1))
            local MT SZ OWN REL
            MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
            SZ=$(stat -c %s "$F" 2>/dev/null)
            REL="${F#$WIN_ROOT}"
            OWN="?"
            if [[ -n "$DPKGINFO" ]]; then
                if grep -rqsF "$REL" "$DPKGINFO"/*.list 2>/dev/null; then
                    OWN="$(L "pacchetto" "package")"
                else
                    OWN="$(L "NON PACCHETTIZZATO" "NOT PACKAGED")"
                    NORPH=$((NORPH + 1))
                    KOROWS+="${MT}	$(basename "$F")	${SZ}	${OWN}	${REL}
"
                fi
            fi
        done < <(find "$MODDIR" -type f \( -name '*.ko' -o -name '*.ko.xz' -o -name '*.ko.zst' -o -name '*.ko.gz' \) 2>/dev/null)
    fi

    # --- initramfs --------------------------------------------------------
    local NINIT=0 INITROWS=""
    local BOOT; BOOT=$(ci_find_dir "$WIN_ROOT" "boot")
    if [[ -n "$BOOT" ]]; then
        while IFS= read -r F; do
            NINIT=$((NINIT + 1))
            INITROWS+="$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)	$(basename "$F")	$(stat -c %s "$F" 2>/dev/null)
"
        done < <(find "$BOOT" -maxdepth 1 -type f \( -iname 'initrd*' -o -iname 'initramfs*' \) 2>/dev/null | sort)
    fi

    separator
    info "$(L "Configurazioni di caricamento:" "Load configurations:") ${BOLD}$NCFG"
    [[ "$NSUSP" -gt 0 ]] && warn "$(L "Direttive 'install' con comando:" "'install' directives with a command:") ${BOLD}$NSUSP"
    info "$(L "Moduli .ko sul disco:" "Kernel modules on disk:") ${BOLD}$NKO"
    if [[ "$NORPH" -gt 0 ]]; then
        warn "$(L "Moduli non appartenenti ad alcun pacchetto:" "Modules owned by no package:") ${BOLD}$NORPH"
        printf '%s' "$KOROWS" | head -15 | while IFS=$'\t' read -r MT NM SZ OWN P; do
            echo -e "      ${RED}${MT}  ${NM}  ${P}${RESET}"
        done
    fi
    info "initramfs: ${BOLD}$NINIT"

    if [[ $NCFG -eq 0 && $NKO -eq 0 && $NINIT -eq 0 ]]; then
        warn "$(L "Nessun artefatto relativo ai moduli kernel trovato." "No kernel-module artefact found.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' l'analisi offline conta qui" "Why offline analysis matters here")</b><br>"
    NOTE+="$(L "Un rootkit LKM gira in kernel space e puo' nascondersi a qualunque strumento eseguito sulla macchina compromessa, incluso lsmod. Sul filesystem montato da fuori non ha modo di mentire: i .ko ci sono o non ci sono." \
        "An LKM rootkit runs in kernel space and can hide from any tool executed on the compromised machine, lsmod included. On a filesystem mounted from outside it has no way to lie: the .ko files are either there or not.")<br><br>"
    NOTE+="<b>$(L "La direttiva install" "The install directive")</b><br>"
    NOTE+="$(L "In modprobe.d, 'install &lt;modulo&gt; &lt;comando&gt;' esegue quel comando quando il modulo viene caricato. E' una persistenza poco nota: non compare in cron, systemd o autostart, e scatta al primo uso di un dispositivo qualsiasi." \
        "In modprobe.d, 'install &lt;module&gt; &lt;command&gt;' runs that command when the module is loaded. It is a little-known persistence: it appears in no cron, systemd or autostart list, and triggers on the first use of any device.")"
    NOTE+="</div></div>"

    local CARDS=""
    [[ -n "$INSTROWS" ]] && CARDS+="<div class='cards'>$(generic_card_html "$(L "Direttive install" "install directives")" "modprobe.d" "$NSUSP" "$(_rows_to_table "$INSTROWS" "$(L "Riga" "Line")")" "⚑")</div>"
    [[ -n "$KOROWS" ]] && CARDS+="<div class='cards'>$(generic_card_html "$(L "Moduli non pacchettizzati" "Modules not packaged")" "$MODDIR" "$NORPH" "$(_rows_to_table "$KOROWS" "$(L "Ultima modifica" "Last modified")" "$(L "Modulo" "Module")" "$(L "Byte" "Bytes")" "$(L "Origine" "Origin")" "$(L "Percorso" "Path")")" "◧")</div>"
    [[ -n "$INITROWS" ]] && CARDS+="<div class='cards'>$(generic_card_html "initramfs" "$BOOT" "$NINIT" "$(_rows_to_table "$INITROWS" "$(L "Ultima modifica" "Last modified")" "$(L "File" "File")" "$(L "Byte" "Bytes")")" "▤")</div>"

    local STATS
    STATS="$(stat_box "$(L "Moduli .ko" "Modules .ko")" "$NKO")"
    STATS+="$(stat_box "$(L "Non pacchettizzati" "Not packaged")" "$NORPH" "$([[ "$NORPH" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "install" "$NSUSP" "$([[ "$NSUSP" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "initramfs" "$NINIT" "info")"
    finish_report "linux_kernel_modules" "Linux Kernel Modules" "KRN" "lib/modules · modprobe.d · initramfs" "$STATS" \
        "${NOTE}${CARDS}<div class='cards'>${BODY}</div>"
}

# ================================================================
#  LINUX 19 — Log dei web server
#
#  Su un server esposto la compromissione iniziale passa quasi sempre da qui.
#  Il modulo System Logs legge /var/log testuali generici, ma non tocca gli
#  access log di nginx e Apache, che hanno un formato proprio e un volume tale
#  da richiedere un'analisi mirata.
#
#  Cosa cerca: richieste verso webshell note, path traversal, tentativi di
#  SQL injection, upload, User-Agent di strumenti offensivi, e i codici di
#  risposta che distinguono un tentativo fallito da uno riuscito.
# ================================================================
module_linux_webserver_logs() {
    section_header "Linux — $(L "Log web server" "Web server logs")" "$ORANGE"
    check_target_root || return 1

    local -a LOGS=()
    local D F
    for D in "var/log/nginx" "var/log/apache2" "var/log/httpd" "var/log/lighttpd" "var/log/caddy"; do
        local DIR; DIR=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$DIR" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && LOGS+=("$F")
        done < <(find "$DIR" -maxdepth 2 -type f \( -name '*access*' -o -name '*error*' \) ! -name '*.gz' 2>/dev/null | sort)
    done
    if [[ ${#LOGS[@]} -eq 0 ]]; then
        warn "$(L "Nessun log di web server trovato." "No web server log found.")"
        return 0
    fi
    info "$(L "Log trovati:" "Logs found:") ${BOLD}${#LOGS[@]}"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local SUM; SUM=$(mktemp); register_tmp "$SUM"

    "$PY3" - "$OUT" "$SUM" "$IOCTMP" "${LOGS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os, collections

out_path, sum_path, ioc_path = sys.argv[1], sys.argv[2], sys.argv[3]
logs = sys.argv[4:]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

# Combined log format: IP - user [data] "METODO path proto" status size "ref" "ua"
CLF = re.compile(
    r'^(?P<ip>\S+)\s+\S+\s+(?P<user>\S+)\s+\[(?P<ts>[^\]]+)\]\s+'
    r'"(?P<method>[A-Z]+)\s+(?P<path>[^"\s]*)[^"]*"\s+(?P<status>\d{3})\s+(?P<size>\S+)'
    r'(?:\s+"(?P<ref>[^"]*)"\s+"(?P<ua>[^"]*)")?')

MONTHS = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
          'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}

def iso(ts):
    # 10/Oct/2024:13:55:36 +0000
    try:
        d, t = ts.split(':', 1)
        day, mon, year = d.split('/')
        return f"{year}-{MONTHS.get(mon,'01')}-{int(day):02d} {t.split()[0]}"
    except Exception:
        return ''

# Pattern di attacco: il nome del gruppo diventa l'etichetta nel report.
ATTACKS = [
    ('WEBSHELL',   re.compile(r'(?i)(c99|r57|b374k|wso\.php|shell\.php|cmd\.php|adminer\.php|alfa\.php|tiny\.php|/\.well-known/[^ ]*\.php)')),
    ('TRAVERSAL',  re.compile(r'(?i)(\.\./|%2e%2e[/%]|/etc/passwd|/proc/self/environ|\\\.\\\.)')),
    ('SQLI',       re.compile(r"(?i)(union[\s+]+select|' or '1'='1|sleep\(\d|benchmark\(|information_schema|xp_cmdshell)")),
    ('RCE',        re.compile(r'(?i)(\bwget\b|\bcurl\b.+http|/bin/(ba)?sh|nc\s+-e|python\s+-c|base64\s+-d|\$\(.*\)|%24%28)')),
    ('UPLOAD',     re.compile(r'(?i)(multipart/form-data|\.php[3457]?(\?|$)|\.jsp(\?|$)|\.aspx?(\?|$)).*(POST|PUT)')),
    ('LOG4J',      re.compile(r'(?i)\$\{jndi:')),
    ('SCANNER_UA', re.compile(r'(?i)(sqlmap|nikto|nmap|masscan|dirbuster|gobuster|wpscan|acunetix|nuclei|feroxbuster|zgrab)')),
]

rows = []
per_ip = collections.Counter()
per_attack = collections.Counter()
status_of_attack = collections.Counter()

for path in logs:
    src = os.path.basename(path)
    try:
        fh = open(path, encoding='utf-8', errors='replace')
    except Exception:
        continue
    with fh:
        for line in fh:
            line = line.rstrip('\n')
            if not line:
                continue
            m = CLF.match(line)
            ip = m.group('ip') if m else ''
            status = m.group('status') if m else ''
            when = iso(m.group('ts')) if m else ''
            req = (m.group('path') if m else line)
            ua = (m.group('ua') or '') if m else ''
            hay = line
            hits = [name for name, rx in ATTACKS if rx.search(hay)]
            if iocs and any(i in hay.lower() for i in iocs):
                hits.append('IOC')
            if not hits:
                continue
            for h in hits:
                per_attack[h] += 1
            if ip:
                per_ip[ip] += 1
            # Uno status 200/500 su una richiesta di attacco vale molto piu' di
            # un 404: distingue il tentativo dall'esito.
            if status.startswith(('2', '5')) or status == '301':
                status_of_attack['riuscita_o_errore_server'] += 1
            rows.append((when, ip, status, ';'.join(hits), req[:300], ua[:150], src))

rows.sort(key=lambda r: r[0])
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')

with open(sum_path, 'w', encoding='utf-8') as fh:
    for name, n in per_attack.most_common():
        fh.write(f"attacco\t{name}\t{n}\n")
    for ip, n in per_ip.most_common(40):
        fh.write(f"ip\t{ip}\t{n}\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessuna richiesta sospetta rilevata nei log analizzati." "No suspicious request found in the analysed logs.")"
        return 0
    fi
    local NHIT2XX; NHIT2XX=$(awk -F'\t' '$3 ~ /^[25]/' "$OUT" | wc -l)

    warn "$(L "Richieste sospette:" "Suspicious requests:") ${BOLD}$TOTAL"
    warn "$(L "Di cui con risposta 2xx/5xx (possibile successo):" "Of which answered 2xx/5xx (possible success):") ${BOLD}$NHIT2XX"
    echo ""
    info "$(L "Per tipo:" "By type:")"
    awk -F'\t' '$1=="attacco"{printf "      %-14s %6s\n", $2, $3}' "$SUM" | head -10
    echo ""
    info "$(L "IP piu' attivi:" "Most active IPs:")"
    awk -F'\t' '$1=="ip"{printf "      %-40s %6s\n", $2, $3}' "$SUM" | head -8

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$3 ~ /^[25]/' "$OUT"; awk -F'\t' '$3 !~ /^[25]/' "$OUT"; } | head -20000 )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data" "Date")" "IP" "$(L "Stato" "Status")" "$(L "Tipo" "Type")" \
        "$(L "Richiesta" "Request")" "User-Agent" "Log")
    local ATABLE; ATABLE=$(_rows_to_table "$(awk -F'\t' '$1=="attacco"{print $2"\t"$3}' "$SUM")" "$(L "Tipo" "Type")" "$(L "Occorrenze" "Occurrences")")
    local ITABLE; ITABLE=$(_rows_to_table "$(awk -F'\t' '$1=="ip"{print $2"\t"$3}' "$SUM")" "IP" "$(L "Richieste sospette" "Suspicious requests")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Lo stato HTTP e' la parte che conta" "The HTTP status is the part that matters")</b><br>"
    NOTE+="$(L "Un log esposto su Internet contiene sempre migliaia di tentativi automatici: la loro presenza non e' un incidente. Cio' che distingue un tentativo da una compromissione e' la risposta: un 404 su una webshell e' rumore di fondo, un 200 sulla stessa richiesta significa che la webshell esiste. Le righe con risposta 2xx e 5xx sono messe in cima per questo motivo." \
        "A log exposed to the Internet always contains thousands of automated attempts: their presence is not an incident. What separates an attempt from a compromise is the response: a 404 on a webshell is background noise, a 200 on the same request means the webshell is there. Rows answered 2xx and 5xx are listed first for this reason.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Richieste sospette" "Suspicious requests")" "$TOTAL" "warn")"
    STATS+="$(stat_box "2xx/5xx" "$NHIT2XX" "$([[ "$NHIT2XX" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Log" "${#LOGS[@]}" "info")"
    finish_report "linux_webserver" "Linux Web Server Logs" "WWW" "nginx · apache · lighttpd" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Per tipo di attacco" "By attack type")" "$(L "riepilogo" "summary")" "$TOTAL" "$ATABLE" "∑")$(generic_card_html "$(L "IP piu' attivi" "Most active IPs")" "$(L "riepilogo" "summary")" "$TOTAL" "$ITABLE" "◉")</div><div class='cards'>$(generic_card_html "$(L "Richieste" "Requests")" "$(L "risposte 2xx/5xx in testa" "2xx/5xx responses first")" "$TOTAL" "$TABLE" "⚑")</div>"
}

# ================================================================
#  LINUX 20 — Credenziali cloud e di sviluppo
#
#  Su una workstation di sviluppo o su un server di build, il bottino non e'
#  la password dell'utente: sono le chiavi che danno accesso all'infrastruttura.
#  Un file ~/.aws/credentials o un ~/.kube/config permettono di passare dal
#  singolo host compromesso all'intero ambiente cloud, e non sono coperti da
#  nessun altro modulo.
#
#  Il modulo rileva la PRESENZA e i metadati di questi file, e mostra gli
#  identificativi non segreti (nome del profilo, access key ID, cluster,
#  registry). NON stampa mai il materiale segreto: la sua esistenza e' il dato
#  che serve all'analista, il valore in chiaro sarebbe solo un rischio in piu'
#  dentro il report.
# ================================================================
module_linux_cloud_credentials() {
    section_header "Linux — $(L "Credenziali cloud e sviluppo" "Cloud and development credentials")" "$RED"
    check_target_root || return 1

    local ROWS="" NFILE=0 NSECRET=0
    local HOME_DIR

    # spec: percorso relativo alla home | etichetta | tipo
    local -a SPECS=(
        ".aws/credentials|AWS|secret"
        ".aws/config|AWS|config"
        ".kube/config|Kubernetes|secret"
        ".docker/config.json|Docker Registry|secret"
        ".config/gcloud/credentials.db|Google Cloud|secret"
        ".config/gcloud/application_default_credentials.json|Google Cloud|secret"
        ".azure/accessTokens.json|Azure|secret"
        ".azure/azureProfile.json|Azure|config"
        ".netrc|netrc|secret"
        ".git-credentials|Git|secret"
        ".npmrc|npm|secret"
        ".pypirc|PyPI|secret"
        ".config/rclone/rclone.conf|rclone|secret"
        ".s3cfg|S3|secret"
        ".terraformrc|Terraform|secret"
        ".config/gh/hosts.yml|GitHub CLI|secret"
        ".ssh/id_rsa|SSH|key"
        ".ssh/id_ed25519|SSH|key"
        ".ssh/id_ecdsa|SSH|key"
    )

    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local SPEC REL LABEL KIND
        for SPEC in "${SPECS[@]}"; do
            IFS='|' read -r REL LABEL KIND <<< "$SPEC"
            local F; F=$(ci_find_file "$HOME_DIR" "$REL")
            [[ -s "$F" ]] || continue
            NFILE=$((NFILE + 1))
            [[ "$KIND" != "config" ]] && NSECRET=$((NSECRET + 1))

            local MT PERM SZ IDENT
            MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
            PERM=$(stat -c %a "$F" 2>/dev/null)
            SZ=$(stat -c %s "$F" 2>/dev/null)

            # Identificativi NON segreti, utili a capire a cosa da' accesso.
            IDENT=""
            case "$LABEL" in
                AWS)
                    IDENT=$(grep -aoE '^\[[^]]+\]|aws_access_key_id[[:space:]]*=[[:space:]]*[A-Z0-9]{16,}' "$F" 2>/dev/null \
                            | sed 's/.*=[[:space:]]*//' | tr '\n' ' ' | cut -c1-160) ;;
                Kubernetes)
                    IDENT=$(grep -aoE 'server:[[:space:]]*\S+|name:[[:space:]]*\S+' "$F" 2>/dev/null | head -6 | tr '\n' ' ' | cut -c1-160) ;;
                "Docker Registry")
                    IDENT=$(grep -aoE '"[a-z0-9.-]+\.[a-z]{2,}(:[0-9]+)?"' "$F" 2>/dev/null | sort -u | tr '\n' ' ' | cut -c1-160) ;;
                Git|netrc)
                    # Solo host e utente: la password viene esclusa.
                    IDENT=$(grep -aoE 'https://[^:/@]+@[^/]+|machine[[:space:]]+\S+|login[[:space:]]+\S+' "$F" 2>/dev/null \
                            | sed -E 's#https://([^:@]+):[^@]*@#https://\1@#' | head -6 | tr '\n' ' ' | cut -c1-160) ;;
                SSH)
                    IDENT=$(head -1 "$F" 2>/dev/null | grep -aoE 'BEGIN [A-Z ]+PRIVATE KEY' || echo "$(L "chiave privata" "private key")")
                    grep -aq 'ENCRYPTED' "$F" 2>/dev/null && IDENT+=" ($(L "protetta da passphrase" "passphrase protected"))" \
                        || IDENT+=" ($(L "SENZA passphrase" "NO passphrase"))" ;;
                *)
                    IDENT=$(grep -aoE '^\[[^]]+\]|^[a-z_]+[[:space:]]*=' "$F" 2>/dev/null | head -5 | tr '\n' ' ' | cut -c1-120) ;;
            esac

            local WARNP=""
            # Un file di credenziali leggibile da altri utenti e' un problema
            # a prescindere dall'incidente in corso.
            [[ -n "$PERM" && "${PERM: -2}" != "00" ]] && WARNP="$(L "permessi larghi" "loose permissions")"

            ROWS+="${MT}	${U}	${LABEL}	${REL}	${PERM} ${WARNP}	${SZ}	${IDENT}
"
            if [[ -n "$WARNP" ]]; then
                warn "$U — $REL ($PERM) — $WARNP"
            else
                ok "$U — ${BOLD}${LABEL}${RESET} ${DIM}${REL}${RESET}"
            fi
        done
    done < <(get_target_user_homes)

    separator
    if [[ "$NFILE" -eq 0 ]]; then
        info "$(L "Nessun file di credenziali cloud o di sviluppo trovato." "No cloud or development credential file found.")"
        return 0
    fi
    ok "$(L "File di credenziali trovati:" "Credential files found:") ${BOLD}$NFILE"
    warn "$(L "Di cui contenenti materiale segreto:" "Of which holding secret material:") ${BOLD}$NSECRET"

    ask_yn "Generare report HTML?" || return 0

    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Utente" "User")" "$(L "Servizio" "Service")" \
        "$(L "File" "File")" "$(L "Permessi" "Permissions")" "$(L "Byte" "Bytes")" "$(L "Identificativi" "Identifiers")")

    local NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Cosa mostra e cosa no" "What is shown and what is not")</b><br>"
    NOTE+="$(L "La tabella riporta identificativi non segreti — profilo, access key ID, endpoint del cluster, registry, host — perche' servono a capire a cosa quelle credenziali davano accesso. Il materiale segreto (secret key, token, password, contenuto delle chiavi private) NON viene mai stampato: e' la sua esistenza a essere il dato utile, riprodurlo qui aggiungerebbe solo un altro posto da cui puo' trapelare." \
        "The table shows non-secret identifiers — profile, access key ID, cluster endpoint, registry, host — because they establish what those credentials granted access to. Secret material (secret keys, tokens, passwords, private key contents) is never printed: its existence is the useful fact, reproducing it here would only add one more place it can leak from.")<br><br>"
    NOTE+="<b>$(L "Priorita' di risposta" "Response priority")</b><br>"
    NOTE+="$(L "Ogni credenziale elencata va considerata compromessa e ruotata. Le chiavi SSH senza passphrase e i file con permessi larghi vanno per primi." \
        "Every credential listed must be treated as compromised and rotated. SSH keys without a passphrase and files with loose permissions come first.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "File" "Files")" "$NFILE")"
    STATS+="$(stat_box "$(L "Con segreti" "With secrets")" "$NSECRET" "warn")"
    finish_report "linux_cloud_credentials" "Linux Cloud & Dev Credentials" "KEY" "~/.aws · ~/.kube · ~/.docker · ~/.ssh" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Credenziali rilevate" "Credentials found")" "$(L "home utenti" "user homes")" "$NFILE" "$TABLE" "⚿")</div>"
}

# ================================================================
#  LINUX 21 — SUID/SGID, capabilities e file scrivibili da tutti
#
#  E' la superficie di privilege escalation del sistema. Conta in due momenti
#  diversi dell'indagine:
#
#  - come CAUSA: un binario SUID inatteso (o un /bin/bash con il bit SUID)
#    spiega come l'attaccante e' passato da utente a root;
#  - come EFFETTO: molti attaccanti lasciano dietro di se' un SUID come
#    backdoor di riserva, piu' discreta di un utente aggiuntivo.
#
#  L'analisi offline e' l'unico modo affidabile di farla: su una macchina
#  compromessa find puo' essere sostituito o un rootkit puo' nascondere i file.
# ================================================================
module_linux_suid_caps() {
    section_header "Linux — SUID/SGID $(L "e capabilities" "and capabilities")" "$ORANGE"
    check_target_root || return 1

    info "$(L "Scansione del filesystem in corso..." "Scanning the filesystem...")"

    # Binari SUID/SGID considerati normali su una distribuzione: servono a
    # separare il rumore dal segnale, NON a dichiarare sicuro cio' che vi
    # corrisponde (un /usr/bin/passwd modificato resta in questa lista).
    local EXPECTED="/(passwd|chsh|chfn|newgrp|gpasswd|su|sudo|mount|umount|ping|ping6|fusermount|fusermount3|pkexec|crontab|at|ssh-agent|unix_chkpwd|expiry|chage|wall|write|dotlockfile|dbus-daemon-launch-helper|polkit-agent-helper-1|sg|staprun|Xorg|snap-confine|utempter|screen|mount\.nfs|pam_timestamp_check|cockpit-session|vmware-user-suid-wrapper)$"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    # -xdev: resta sul volume montato, non segue mount annidati o /proc.
    find "$WIN_ROOT" -xdev \( -perm -4000 -o -perm -2000 \) -type f -printf '%M\t%u\t%g\t%s\t%TY-%Tm-%Td %TH:%TM:%TS\t%p\n' 2>/dev/null \
        | sed 's/\.[0-9]*\t/\t/' > "$OUT" || true

    # NB: `grep -c` stampa 0 ma esce con stato 1 quando non trova nulla, quindi
    # un `|| echo 0` produrrebbe due zeri. Si usa wc -l.
    local NSUID NUNEXP
    NSUID=$(wc -l < "$OUT" 2>/dev/null || echo 0)
    local UNEXP; UNEXP=$(mktemp); register_tmp "$UNEXP"
    grep -vE "$EXPECTED" "$OUT" > "$UNEXP" 2>/dev/null || true
    NUNEXP=$(wc -l < "$UNEXP" 2>/dev/null || echo 0)

    # File con capabilities: alternativa moderna al SUID, spesso trascurata.
    # getcap non funziona su un volume montato senza supporto xattr, quindi si
    # tenta e si dichiara l'esito invece di far finta di aver guardato.
    local CAPOUT; CAPOUT=$(mktemp); register_tmp "$CAPOUT"
    local CAPS_OK=false NCAP=0
    if command -v getcap > /dev/null 2>&1; then
        CAPS_OK=true
        getcap -r "$WIN_ROOT" 2>/dev/null | head -500 > "$CAPOUT" || true
        NCAP=$(wc -l < "$CAPOUT" 2>/dev/null || echo 0)
    fi

    # Directory e file scrivibili da tutti fuori dalle aree temporanee attese.
    local WWOUT; WWOUT=$(mktemp); register_tmp "$WWOUT"
    find "$WIN_ROOT" -xdev -perm -0002 ! -type l \
         ! -path "*/tmp/*" ! -path "*/var/tmp/*" ! -path "*/dev/shm/*" ! -path "*/proc/*" \
         -printf '%M\t%u\t%s\t%TY-%Tm-%Td %TH:%TM\t%p\n' 2>/dev/null | head -400 > "$WWOUT" || true
    local NWW; NWW=$(wc -l < "$WWOUT" 2>/dev/null || echo 0)

    separator
    ok "$(L "Binari SUID/SGID:" "SUID/SGID binaries:") ${BOLD}$NSUID"
    if [[ "$NUNEXP" -gt 0 ]]; then
        warn "$(L "Fuori dall'elenco atteso:" "Outside the expected set:") ${BOLD}$NUNEXP"
        awk -F'\t' '{printf "      %s  %-8s %s\n", $1, $2, $6}' "$UNEXP" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi
    if $CAPS_OK; then
        info "$(L "File con capabilities:" "Files with capabilities:") ${BOLD}$NCAP"
    else
        info "$(L "getcap non disponibile: capabilities non verificate." "getcap unavailable: capabilities not checked.")"
    fi
    info "$(L "Scrivibili da tutti (fuori da tmp):" "World-writable (outside tmp):") ${BOLD}$NWW"

    if [[ "$NSUID" -eq 0 && "$NCAP" -eq 0 && "$NWW" -eq 0 ]]; then
        warn "$(L "Nessun risultato: il volume potrebbe essere montato senza supporto ai permessi." \
                 "No result: the volume may be mounted without permission support.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Come leggere l'elenco atteso" "How to read the expected set")</b><br>"
    BODY+="$(L "I binari SUID di sistema (passwd, sudo, mount...) sono separati dagli altri solo per ridurre il rumore. Non sono dichiarati sicuri: un /usr/bin/passwd sostituito resta nell'elenco atteso, e va confrontato con l'hash del pacchetto. Il segnale forte e' un SUID FUORI dall'elenco — soprattutto una shell, un interprete o un binario in /tmp, /home o /var." \
        "System SUID binaries (passwd, sudo, mount...) are separated from the rest only to cut noise. They are not declared safe: a replaced /usr/bin/passwd still sits in the expected set and must be checked against the package hash. The strong signal is a SUID OUTSIDE the set — especially a shell, an interpreter, or a binary under /tmp, /home or /var.")<br><br>"
    BODY+="$(L "Le capabilities sono l'alternativa moderna al SUID e vengono spesso dimenticate in fase di audit: CAP_SETUID, CAP_SYS_ADMIN e CAP_DAC_OVERRIDE su un binario arbitrario equivalgono di fatto a root." \
        "Capabilities are the modern alternative to SUID and are often overlooked during audits: CAP_SETUID, CAP_SYS_ADMIN and CAP_DAC_OVERRIDE on an arbitrary binary are effectively root.")"
    if ! $CAPS_OK; then
        BODY+="<br><br><b>$(L "Capabilities non verificate" "Capabilities not checked")</b><br>"
        BODY+="$(L "getcap non era disponibile sull'host di analisi: questa parte del controllo non e' stata eseguita e va rifatta a parte." \
            "getcap was unavailable on the analysis host: this part of the check was not performed and must be repeated separately.")"
    fi
    BODY+="</div></div>"

    if [[ "$NUNEXP" -gt 0 ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "SUID/SGID fuori dall'elenco atteso" "SUID/SGID outside the expected set")" "$WIN_ROOT" "$NUNEXP" \
            "$(_rows_to_table "$(cat "$UNEXP")" "$(L "Permessi" "Mode")" "$(L "Utente" "Owner")" "$(L "Gruppo" "Group")" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")" "$(L "Percorso" "Path")")" "⚑")</div>"
    fi
    BODY+="<div class='cards'>$(generic_card_html "$(L "Tutti i SUID/SGID" "All SUID/SGID")" "$WIN_ROOT" "$NSUID" \
        "$(_rows_to_table "$(head -2000 "$OUT")" "$(L "Permessi" "Mode")" "$(L "Utente" "Owner")" "$(L "Gruppo" "Group")" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")" "$(L "Percorso" "Path")")" "◧")</div>"
    if [[ "$NCAP" -gt 0 ]]; then
        BODY+="<div class='cards'>$(generic_card_html "Capabilities" "getcap -r" "$NCAP" \
            "$(_rows_to_table "$(sed 's/ /\t/' "$CAPOUT")" "$(L "Percorso" "Path")" "Capabilities")" "⚙")</div>"
    fi
    if [[ "$NWW" -gt 0 ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "Scrivibili da tutti" "World-writable")" "$(L "fuori da tmp" "outside tmp")" "$NWW" \
            "$(_rows_to_table "$(cat "$WWOUT")" "$(L "Permessi" "Mode")" "$(L "Utente" "Owner")" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")" "$(L "Percorso" "Path")")" "◔")</div>"
    fi

    local STATS
    STATS="$(stat_box "SUID/SGID" "$NSUID")"
    STATS+="$(stat_box "$(L "Inattesi" "Unexpected")" "$NUNEXP" "$([[ "$NUNEXP" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Capabilities" "$($CAPS_OK && echo "$NCAP" || echo "n/d")" "info")"
    STATS+="$(stat_box "$(L "Scrivibili da tutti" "World-writable")" "$NWW" "$([[ "$NWW" -gt 0 ]] && echo warn || echo info)")"
    finish_report "linux_suid_caps" "Linux SUID & Capabilities" "SUID" "$(L "superficie di privilege escalation" "privilege escalation surface")" "$STATS" "$BODY"
}

# --- macOS 1 — System Logs ---
module_macos_logs() {
    section_header "macOS — System Logs" "$GREEN"
    check_target_root || return 1
    local KW="fail|error|denied|invalid|sudo|root|unauthorized|jailbreak|malware"
    local BODY="" FOUND=0
    local LOGDIR; LOGDIR=$(ci_find_dir "$WIN_ROOT" "var/log")
    if [[ -n "$LOGDIR" ]]; then
        for NAME in system.log install.log secure.log wifi.log appfirewall.log; do
            local F; F=$(ci_find_file "$LOGDIR" "$NAME")
            [[ -n "$F" && -s "$F" ]] || continue
            FOUND=$((FOUND + 1)); ok "$NAME"; print_file_lines "$F" "$KW" 40; echo ""
            BODY+=$(file_card_html "$F" "$KW" "≣")
        done
        # ASL (formato binario: estrazione stringhe)
        local ASL; ASL=$(ci_find_dir "$LOGDIR" "asl")
        if [[ -n "$ASL" ]]; then
            local ACOUNT; ACOUNT=$(find "$ASL" -maxdepth 1 -type f 2>/dev/null | wc -l)
            [[ "$ACOUNT" -gt 0 ]] && { FOUND=$((FOUND + 1)); ok "asl ($ACOUNT file)"; BODY+="<div class='card'><div class='card-header'><div class='uicon'>≣</div><div><div class='uname'>ASL logs</div><div class='upath'>$ASL</div></div><div class='badge'>$ACOUNT file</div></div><div style='padding:1rem 1.5rem'><p class='dim mono' style='font-size:.72rem'>$(L "Formato binario ASL — analisi approfondita fuori scope offline." "Binary ASL format — deep parsing out of offline scope.")</p></div></div>"; }
        fi
    fi
    # Nota: unified logs .tracev3 esplicitamente fuori scope
    warn "$(L "Unified logs (.tracev3) non analizzati: richiedono il tool 'log' o parser dedicati (fuori scope)." "Unified logs (.tracev3) not parsed: require the 'log' tool or dedicated parsers (out of scope).")"
    separator
    info "$(L "Log trovati:" "Logs found:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun log testuale trovato." "No text log found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Log" "Logs")" "$FOUND" "info")
    finish_report "macos_logs" "macOS System Logs" "LOG" "system.log · install.log · asl" "$STATS" "<div class='cards'>$BODY</div>"
}

# Estrae i campi salienti da un plist utente dslocal
_parse_dslocal_user() {
    "$PY3" - "$1" << 'PYEOF'
import sys, plistlib
try:
    d = plistlib.load(open(sys.argv[1], 'rb'))
    def g(k):
        v = d.get(k)
        return str(v[0]) if isinstance(v, list) and v else (str(v) if v is not None else '')
    has_hash = 'yes' if d.get('ShadowHashData') else 'no'
    print(f"{g('name')}\t{g('uid')}\t{g('gid')}\t{g('home')}\t{g('shell')}\t{has_hash}")
except Exception as e:
    print(f"ERROR\t{e}\t\t\t\t")
PYEOF
}

# --- macOS 2 — User Accounts (dslocal) ---
module_macos_accounts() {
    section_header "macOS — User Accounts" "$RED"
    check_target_root || return 1
    local UDIR
    UDIR=$(ci_find_dir "$WIN_ROOT" "var/db/dslocal/nodes/Default/users")
    [[ -z "$UDIR" ]] && UDIR=$(ci_find_dir "$WIN_ROOT" "private/var/db/dslocal/nodes/Default/users")
    [[ -z "$UDIR" ]] && { warn "$(L "Database dslocal non trovato." "dslocal database not found.")"; return 0; }
    local ROWS="" COUNT=0
    while IFS= read -r PL; do
        local R; R=$(_parse_dslocal_user "$PL")
        [[ "$R" == ERROR* ]] && continue
        local UNAME; UNAME=$(printf '%s' "$R" | cut -f1)
        [[ "$UNAME" == _* ]] && continue   # account di servizio
        COUNT=$((COUNT + 1)); ROWS+="$R"$'\n'
        ok "$UNAME $(printf '%s' "$R" | awk -F'\t' '{print "(uid "$2", hash:"$6")"}')"
    done < <(find "$UDIR" -maxdepth 1 -type f -iname "*.plist" 2>/dev/null)
    separator
    info "$(L "Account utente:" "User accounts:") ${BOLD}$COUNT"
    [[ $COUNT -eq 0 ]] && { warn "$(L "Nessun account utente reale trovato." "No real user account found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local TABLE; TABLE=$(_rows_to_table "${ROWS%$'\n'}" "User" "UID" "GID" "Home" "Shell" "Hash")
    local BODY; BODY=$(generic_card_html "dslocal users" "$UDIR" "$COUNT account" "$TABLE" "◉")
    local STATS; STATS=$(stat_box "$(L "Account" "Accounts")" "$COUNT" "info")
    finish_report "macos_accounts" "macOS User Accounts" "USR" "/var/db/dslocal/.../users/*.plist" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 3 — Persistence (LaunchAgents/Daemons/cron) ---
module_macos_persistence() {
    section_header "macOS — Persistence" "$ORANGE"
    check_target_root || return 1
    local KW="curl|wget|/tmp/|/var/tmp|base64|python|bash -i|nc |/dev/tcp|RunAtLoad|KeepAlive|http"
    local BODY="" FOUND=0
    _mac_persist_dir() { # label, dir
        local LABEL="$1" D="$2"; [[ -z "$D" || ! -d "$D" ]] && return
        local n=0
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); n=$((n + 1))
            if [[ "$F" == *.plist ]]; then
                # rende leggibili anche i plist binari
                local TMPF; TMPF=$(mktemp); read_plist "$F" > "$TMPF" 2>/dev/null
                [[ -s "$TMPF" ]] || cat "$F" > "$TMPF" 2>/dev/null
                BODY+=$(generic_card_html "$(basename "$F")" "$F" "$LABEL" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "$KW")</pre>" "⟳")
                rm -f "$TMPF"
            else
                BODY+=$(file_card_html "$F" "$KW" "⟳")
            fi
        done < <(find "$D" -maxdepth 1 -type f \( -iname "*.plist" -o -iname "*.conf" \) 2>/dev/null)
        [[ $n -gt 0 ]] && ok "$LABEL — $n plist ($D)"
    }
    _mac_persist_dir "LaunchDaemons" "$(ci_find_dir "$WIN_ROOT" "Library/LaunchDaemons")"
    _mac_persist_dir "LaunchAgents"  "$(ci_find_dir "$WIN_ROOT" "Library/LaunchAgents")"
    _mac_persist_dir "System LaunchDaemons" "$(ci_find_dir "$WIN_ROOT" "System/Library/LaunchDaemons")"
    # cron + periodic
    local CRON; CRON=$(ci_find_dir "$WIN_ROOT" "var/at/tabs"); _mac_persist_dir "cron" "$CRON"
    local PERIODIC; PERIODIC=$(ci_find_dir "$WIN_ROOT" "etc/periodic")
    # per-user LaunchAgents
    while IFS= read -r HOME_DIR; do
        _mac_persist_dir "$(basename "$HOME_DIR") LaunchAgents" "$(ci_find_dir "$HOME_DIR" "Library/LaunchAgents")"
    done < <(get_macos_user_homes)
    separator
    info "$(L "Artefatti di persistenza:" "Persistence artifacts:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun artefatto di persistenza trovato." "No persistence artifact found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Artefatti" "Artifacts")" "$FOUND" "info")
    finish_report "macos_persistence" "macOS Persistence" "PER" "LaunchAgents · LaunchDaemons · cron" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 4 — Login Items (BTM) ---
module_macos_loginitems() {
    section_header "macOS — Login Items (BTM)" "$RED"
    check_target_root || return 1
    local BODY="" FOUND=0
    # backgrounditems.btm (vari percorsi a seconda della versione)
    mapfile -t BTM < <(find "$WIN_ROOT" -maxdepth 8 -iname "backgrounditems.btm" -type f 2>/dev/null; find "$WIN_ROOT" -maxdepth 8 -path "*backgroundtaskmanagementagent*" -type f 2>/dev/null | head -20)
    for F in "${BTM[@]}"; do
        [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); ok "$(basename "$F")"
        # BTM è un blob binario: estrai riferimenti a path/app leggibili
        local STR; STR=$(strings -n 5 "$F" 2>/dev/null | grep -iE '\.app|/Users/|/Library/|/usr/|\.plist|\.sh$' | sort -u | head -100)
        local TMPF; TMPF=$(mktemp); printf '%s\n' "$STR" > "$TMPF"
        BODY+=$(generic_card_html "$(basename "$F")" "$F" "$(printf '%s\n' "$STR" | grep -c .) ref" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "tmp|/private|curl|\\.sh")</pre>" "⚑")
        rm -f "$TMPF"
    done
    separator
    info "$(L "File BTM trovati:" "BTM files found:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun login item (BTM) trovato." "No login item (BTM) found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File BTM" "BTM files")" "$FOUND" "info")
    finish_report "macos_loginitems" "macOS Login Items (BTM)" "BTM" "backgrounditems.btm" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 5 — Quarantine / Downloads ---
module_macos_quarantine() {
    section_header "macOS — Quarantine / Downloads" "$CYAN"
    check_target_root || return 1
    local BODY="" TOTAL=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        mapfile -t QDB < <(find "$HOME_DIR" -maxdepth 4 -iname "com.apple.LaunchServices.QuarantineEventsV2*" -type f 2>/dev/null)
        for DB in "${QDB[@]}"; do
            # L'URL diretto del file (LSQuarantineDataURLString) è spesso vuoto su Chrome:
            # in tal caso si usa l'URL della pagina di origine (LSQuarantineOriginURLString).
            # Mostra anche la pagina di origine come colonna separata.
            local ROWS; ROWS=$(query_sqlite "$DB" "SELECT datetime(LSQuarantineTimeStamp+978307200,'unixepoch'), LSQuarantineAgentName, COALESCE(NULLIF(LSQuarantineDataURLString,''), LSQuarantineOriginURLString, ''), COALESCE(LSQuarantineOriginURLString,'') FROM LSQuarantineEvent ORDER BY LSQuarantineTimeStamp DESC LIMIT 100000")
            local HEADERS_EXTRA=1
            if [[ -z "$ROWS" || "$ROWS" == ERROR* ]]; then
                # Fallback per schema più vecchi/parziali
                ROWS=$(query_sqlite "$DB" "SELECT datetime(LSQuarantineTimeStamp+978307200,'unixepoch'), LSQuarantineAgentName, LSQuarantineDataURLString FROM LSQuarantineEvent ORDER BY LSQuarantineTimeStamp DESC LIMIT 100000")
                HEADERS_EXTRA=0
            fi
            [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
            local N; N=$(printf '%s\n' "$ROWS" | grep -c .); TOTAL=$((TOTAL + N))
            ok "$UNAME — ${BOLD}$N $(L "download tracciati" "tracked downloads")"
            local TABLE
            if [[ $HEADERS_EXTRA -eq 1 ]]; then
                TABLE=$(_rows_to_table "$ROWS" "$(L "Data" "Date")" "Agent" "URL" "$(L "Pagina origine" "Origin page")")
            else
                TABLE=$(_rows_to_table "$ROWS" "$(L "Data" "Date")" "Agent" "URL")
            fi
            BODY+=$(generic_card_html "$UNAME" "$DB" "$N" "$TABLE" "⤓")
        done
    done < <(get_macos_user_homes)
    separator
    info "$(L "Eventi di quarantena:" "Quarantine events:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun evento di quarantena trovato." "No quarantine event found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Eventi" "Events")" "$TOTAL" "info")
    finish_report "macos_quarantine" "macOS Quarantine / Downloads" "DL" "QuarantineEventsV2" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 6 — TCC Privacy ---
module_macos_tcc() {
    section_header "macOS — TCC Privacy" "$RED"
    check_target_root || return 1
    local BODY="" TOTAL=0
    _tcc_card() { # db, label
        local DB="$1" LBL="$2"
        local ROWS; ROWS=$(query_sqlite "$DB" "SELECT service, client, CASE auth_value WHEN 0 THEN 'denied' WHEN 2 THEN 'allowed' ELSE auth_value END FROM access ORDER BY service")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && ROWS=$(query_sqlite "$DB" "SELECT service, client, allowed FROM access ORDER BY service")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && return
        local N; N=$(printf '%s\n' "$ROWS" | grep -c .); TOTAL=$((TOTAL + N))
        ok "$LBL — ${BOLD}$N $(L "permessi" "permissions")"
        local TABLE; TABLE=$(_rows_to_table "$ROWS" "Service" "Client" "Auth")
        BODY+=$(generic_card_html "$LBL" "$DB" "$N" "$TABLE" "⊘")
    }
    local SYS; SYS=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "Library/Application Support/com.apple.TCC")" "TCC.db")
    [[ -n "$SYS" ]] && _tcc_card "$SYS" "System TCC"
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local DB; DB=$(ci_find_file "$(ci_find_dir "$HOME_DIR" "Library/Application Support/com.apple.TCC")" "TCC.db")
        [[ -n "$DB" ]] && _tcc_card "$DB" "$U TCC"
    done < <(get_macos_user_homes)
    separator
    info "$(L "Permessi TCC:" "TCC permissions:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun database TCC leggibile." "No readable TCC database.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Permessi" "Permissions")" "$TOTAL" "info")
    finish_report "macos_tcc" "macOS TCC Privacy" "TCC" "TCC.db (cam/mic/disco)" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 7 — KnowledgeC ---
module_macos_knowledgec() {
    section_header "macOS — KnowledgeC" "$BLUE"
    check_target_root || return 1
    local BODY="" TOTAL=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local DB; DB=$(ci_find_file "$(ci_find_dir "$HOME_DIR" "Library/Application Support/Knowledge")" "knowledgeC.db")
        [[ -z "$DB" ]] && continue
        local ROWS; ROWS=$(query_sqlite "$DB" "SELECT datetime(ZCREATIONDATE+978307200,'unixepoch'), ZSTREAMNAME, ZVALUESTRING FROM ZOBJECT WHERE ZVALUESTRING IS NOT NULL ORDER BY ZCREATIONDATE DESC LIMIT 100000")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
        local N; N=$(printf '%s\n' "$ROWS" | grep -c .); TOTAL=$((TOTAL + N))
        ok "$U — ${BOLD}$N $(L "eventi" "events")"
        local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Data" "Date")" "Stream" "Value")
        BODY+=$(generic_card_html "$U" "$DB" "$N" "$TABLE" "◴")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Eventi KnowledgeC:" "KnowledgeC events:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun knowledgeC.db leggibile." "No readable knowledgeC.db.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Eventi" "Events")" "$TOTAL" "info")
    finish_report "macos_knowledgec" "macOS KnowledgeC" "KC" "knowledgeC.db (app usage)" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 8 — Browser History ---
module_macos_browser() {
    section_header "macOS — Browser History" "$CYAN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR"); local CARDS="" UCOUNT=0
        # Safari
        local SAF; SAF=$(ci_find_file "$(ci_find_dir "$HOME_DIR" "Library/Safari")" "History.db")
        if [[ -n "$SAF" ]]; then
            local ROWS; ROWS=$(query_sqlite "$SAF" "SELECT datetime(v.visit_time+978307200,'unixepoch'), i.url, v.title FROM history_visits v JOIN history_items i ON v.history_item=i.id ORDER BY v.visit_time DESC LIMIT 100000")
            if [[ -n "$ROWS" && "$ROWS" != ERROR* ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); ok "$U — Safari"
                CARDS+=$(generic_card_html "Safari" "$SAF" "$(printf '%s\n' "$ROWS" | grep -c .) URL" "$(_rows_to_table "$ROWS" "$(L "Data" "Date")" "URL" "$(L "Titolo" "Title")")" "◐")
            fi
        fi
        # Chrome/Brave/Edge
        for CR in "Library/Application Support/Google/Chrome" "Library/Application Support/BraveSoftware/Brave-Browser" "Library/Application Support/Microsoft Edge"; do
            local BASE; BASE=$(ci_find_dir "$HOME_DIR" "$CR"); [[ -z "$BASE" ]] && continue
            while IFS= read -r HISTDB; do
                local ROWS; ROWS=$(query_sqlite "$HISTDB" "SELECT datetime(last_visit_time/1000000-11644473600,'unixepoch'), url, title FROM urls ORDER BY last_visit_time DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); ok "$U — $(basename "$(dirname "$HISTDB")") (Chromium)"
                CARDS+=$(_browser_table_card "$HISTDB" "$ROWS")
            done < <(find "$BASE" -maxdepth 2 -name "History" -type f 2>/dev/null)
        done
        # Firefox
        local FB; FB=$(ci_find_dir "$HOME_DIR" "Library/Application Support/Firefox/Profiles")
        if [[ -n "$FB" ]]; then
            while IFS= read -r PLACES; do
                local ROWS; ROWS=$(query_sqlite "$PLACES" "SELECT datetime(last_visit_date/1000000,'unixepoch'), url, title FROM moz_places WHERE last_visit_date IS NOT NULL ORDER BY last_visit_date DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); ok "$U — Firefox"
                CARDS+=$(_browser_table_card "$PLACES" "$ROWS")
            done < <(find "$FB" -maxdepth 2 -name "places.sqlite" -type f 2>/dev/null)
        fi
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); BODY+=$(generic_card_html "$U" "$HOME_DIR" "$UCOUNT" "$CARDS" "◐")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Profili browser:" "Browser profiles:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history browser trovata." "No browser history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Profili" "Profiles")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "macos_browser" "macOS Browser History" "WEB" "Safari · Chrome · Firefox" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 9 — Shell & AI History ---
module_macos_shell_ai_history() {
    section_header "macOS — Shell & AI History" "$MAGENTA"
    check_target_root || return 1
    local FILES=(.zsh_history .bash_history .sh_history .python_history .psql_history .node_repl_history
                 .claude .aider.chat.history.md .config/aichat .ollama/history)
    local KW="password|secret|token|api_key|apikey|credential|ssh|curl|base64|security|keychain|sudo"
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR"); local CARDS="" UCOUNT=0
        for HF in "${FILES[@]}"; do
            local P="$HOME_DIR/$HF"
            if [[ -f "$P" && -s "$P" ]]; then
                local MODE="histts"
                case "$(basename "$P")" in
                    .python_history|.node_repl_history|.psql_history|.mysql_history) MODE="histrl" ;;
                esac
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$P" "$KW" "\$" "$MODE")
            elif [[ -d "$P" ]]; then
                while IFS= read -r AF; do
                    [[ -s "$AF" ]] || continue; UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$AF" "$KW" "◈")
                done < <(find "$P" -maxdepth 3 -type f \( -iname "*.json" -o -iname "*.jsonl" -o -iname "*history*" -o -iname "*.md" \) 2>/dev/null | head -20)
            fi
        done
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); ok "$U — ${BOLD}$UCOUNT file"
        BODY+=$(generic_card_html "$U" "$HOME_DIR" "$UCOUNT file" "$CARDS" "◢")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Artefatti history:" "History artifacts:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history trovata." "No history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "File" "Files")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "macos_shell_ai_history" "macOS Shell & AI History" "SH" ".zsh_history · .bash_history · AI CLI" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 10 — Recent Items ---
module_macos_recent() {
    section_header "macOS — Recent Items" "$GREEN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR"); local CARDS="" UCOUNT=0
        # SFL / SFL2 (shared file list — recent apps/docs/servers)
        local SFLD; SFLD=$(ci_find_dir "$HOME_DIR" "Library/Application Support/com.apple.sharedfilelist")
        if [[ -n "$SFLD" ]]; then
            while IFS= read -r SF; do
                [[ -s "$SF" ]] || continue
                local STR; STR=$(strings -n 5 "$SF" 2>/dev/null | grep -iE '/Users/|/Volumes/|\.app|\.|smb://|afp://|ftp://' | sort -u | head -60)
                [[ -z "$STR" ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
                local TMPF; TMPF=$(mktemp); printf '%s\n' "$STR" > "$TMPF"
                CARDS+=$(generic_card_html "$(basename "$SF")" "$SF" "$(printf '%s\n' "$STR" | grep -c .) ref" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "smb://|afp://|/Volumes")</pre>" "◇")
                rm -f "$TMPF"
            done < <(find "$SFLD" -maxdepth 2 -type f \( -iname "*.sfl" -o -iname "*.sfl2" -o -iname "*.sfl3" \) 2>/dev/null)
        fi
        # Trash
        local TRASH; TRASH=$(ci_find_dir "$HOME_DIR" ".Trash")
        if [[ -n "$TRASH" ]]; then
            local TN; TN=$(find "$TRASH" -mindepth 1 -maxdepth 1 2>/dev/null | wc -l)
            if [[ "$TN" -gt 0 ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + TN))
                local TLIST; TLIST=$(find "$TRASH" -mindepth 1 -maxdepth 1 -printf '%T+\t%p\n' 2>/dev/null | sort -r | head -200)
                CARDS+=$(generic_card_html ".Trash" "$TRASH" "$TN" "$(_rows_to_table "$TLIST" "$(L "Modificato" "Modified")" "$(L "Percorso" "Path")")" "♺")
            fi
        fi
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); ok "$U — ${BOLD}$UCOUNT $(L "sorgenti" "sources")"
        BODY+=$(generic_card_html "$U" "$HOME_DIR" "$UCOUNT" "$CARDS" "◇")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Voci recenti:" "Recent items:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna voce recente trovata." "No recent item found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Voci" "Items")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "macos_recent" "macOS Recent Items" "RCN" "SFL · .Trash · recent items" "$STATS" "<div class='cards'>$BODY</div>"
}

# --- macOS 11 — FSEvents (/.fseventsd) ---
#
# FSEvents e' il registro delle modifiche al filesystem tenuto da macOS: l'analogo
# dello USN Journal di NTFS, e la fonte piu' ricca per ricostruire creazione,
# rinomina e cancellazione di file — comprese quelle di file non piu' presenti.
#
# Formato: file gzip in /.fseventsd, ciascuno con una o piu' pagine DLS1/DLS2.
# Ogni record e' path NUL-terminato + event id (u64 LE) + flag (u32 LE), piu'
# un node id (u64 LE) nelle pagine DLS2/DLS3.
#
# ATTENZIONE sui tempi: i record NON contengono un timestamp. L'event id e' un
# contatore monotono. L'unico riferimento temporale e' l'intervallo coperto dal
# file che li contiene, quindi le date qui sono un LIMITE SUPERIORE approssimato
# (mtime del file di log), non l'istante dell'evento. Il report lo dichiara.
module_macos_fsevents() {
    section_header "macOS — FSEvents" "$MAGENTA"
    check_target_root || return 1

    local FSEDIR; FSEDIR=$(ci_find_dir "$WIN_ROOT" ".fseventsd")
    [[ -z "$FSEDIR" ]] && { warn "$(L "Directory /.fseventsd non trovata." "/.fseventsd directory not found.")"; return 0; }

    mapfile -t FSEFILES < <(find "$FSEDIR" -maxdepth 1 -type f ! -name 'fseventsd-uuid' -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
    if [[ ${#FSEFILES[@]} -eq 0 ]]; then
        warn "$(L "Nessun log FSEvents presente." "No FSEvents log present.")"
        return 0
    fi
    info "$(L "File FSEvents trovati:" "FSEvents logs found:") ${BOLD}${#FSEFILES[@]}"

    local KW="${IOC_LIST[*]:-}"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${FSEFILES[@]}" << 'PYEOF' 2>/dev/null
import sys, gzip, struct, os, datetime

out_path = sys.argv[1]
files = sys.argv[2:]

# Bitmask degli eventi FSEvents.
FLAGS = [
    (0x00000001, "FolderEvent"),      (0x00000002, "Mount"),
    (0x00000004, "Unmount"),          (0x00000020, "EndOfTransaction"),
    (0x00000800, "LastHardLinkRemoved"), (0x00001000, "HardLink"),
    (0x00004000, "SymbolicLink"),     (0x00008000, "FileEvent"),
    (0x00010000, "PermissionChange"), (0x00020000, "XattrModified"),
    (0x00040000, "XattrRemoved"),     (0x00100000, "DocumentRevision"),
    (0x00400000, "ItemCloned"),       (0x01000000, "Created"),
    (0x02000000, "Removed"),          (0x04000000, "InodeMetaMod"),
    (0x08000000, "Renamed"),          (0x10000000, "Modified"),
    (0x20000000, "Exchange"),         (0x40000000, "FinderInfoMod"),
    (0x80000000, "FolderCreated"),
]

# Eventi che meritano attenzione in un'indagine: cancellazioni e rinomine
# sono il segnale tipico di anti-forensics e di stage/esfiltrazione.
NOTABLE = {"Removed", "Renamed", "LastHardLinkRemoved", "PermissionChange", "XattrRemoved"}

def decode_flags(v):
    names = [n for bit, n in FLAGS if v & bit]
    return ";".join(names) if names else f"0x{v:08x}"

def parse(data):
    """Genera (path, event_id, flags) per ogni record delle pagine del file."""
    off = 0
    total = len(data)
    while off + 12 <= total:
        magic = data[off:off + 4]
        if magic not in (b"1SLD", b"2SLD", b"3SLD"):
            break
        # header: magic(4) + unknown(4) + page_size(4)
        page_size = struct.unpack_from("<I", data, off + 8)[0]
        if page_size < 12 or off + page_size > total:
            break
        # DLS2 e DLS3 aggiungono un node id da 8 byte per record
        extra = 8 if magic in (b"2SLD", b"3SLD") else 0
        p = off + 12
        end = off + page_size
        while p < end:
            nul = data.find(b"\x00", p, end)
            if nul < 0:
                break
            path = data[p:nul].decode("utf-8", "replace")
            p = nul + 1
            if p + 12 + extra > end:
                break
            event_id, flags = struct.unpack_from("<QI", data, p)
            p += 12 + extra
            if path:
                yield path, event_id, flags
        off = end

rows = []
for fp in files:
    try:
        with gzip.open(fp, "rb") as fh:
            data = fh.read()
    except Exception:
        continue
    # Unico ancoraggio temporale disponibile: il mtime del file di log.
    try:
        approx = datetime.datetime.utcfromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        approx = ""
    src = os.path.basename(fp)
    for path, event_id, flags in parse(data):
        names = decode_flags(flags)
        notable = "1" if any(n in NOTABLE for n in names.split(";")) else "0"
        rows.append((approx, path, names, str(event_id), src, notable))

with open(out_path, "w", encoding="utf-8") as fh:
    for r in rows:
        fh.write("\t".join(x.replace("\t", " ") for x in r) + "\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun record FSEvents decodificato (log vuoti o formato non riconosciuto)." \
                 "No FSEvents record decoded (empty logs or unrecognised format).")"
        return 0
    fi

    local NOTABLE; NOTABLE=$(awk -F'\t' '$6=="1"' "$OUT" | wc -l)
    ok "$(L "Record FSEvents decodificati:" "FSEvents records decoded:") ${BOLD}$TOTAL"
    info "$(L "Eventi di cancellazione/rinomina:" "Delete/rename events:") ${BOLD}$NOTABLE"

    # Anteprima a console: solo gli eventi rilevanti, altrimenti e' illeggibile.
    awk -F'\t' '$6=="1"{print $1"  "$3"  "$2}' "$OUT" | head -40 | while IFS= read -r LN; do
        echo -e "      ${DIM}${LN}${RESET}"
    done

    ask_yn "Generare report HTML?" || return 0

    # Il report completo puo' contenere milioni di record: si limita a un
    # volume gestibile dal browser, privilegiando gli eventi rilevanti.
    local LIMIT=20000
    local ROWS; ROWS=$( { awk -F'\t' '$6=="1"{print $1"\t"$3"\t"$2"\t"$4"\t"$5}' "$OUT";
                          awk -F'\t' '$6=="0"{print $1"\t"$3"\t"$2"\t"$4"\t"$5}' "$OUT"; } | head -$LIMIT )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data (approx.)" "Date (approx.)")" "$(L "Evento" "Event")" "Path" "Event ID" "$(L "File di log" "Log file")")

    local NOTE
    NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Nota sui tempi" "Note on timestamps")</b><br>"
    NOTE+="$(L "I record FSEvents non contengono un timestamp: l'event ID e' un contatore monotono. La colonna data riporta il mtime del file di log che contiene il record, cioe' un LIMITE SUPERIORE approssimato dell'istante dell'evento — non la sua ora esatta. Per una datazione precisa va correlato con altri artefatti." \
        "FSEvents records carry no timestamp: the event ID is a monotonic counter. The date column shows the mtime of the log file containing the record, i.e. an approximate UPPER BOUND of when the event happened — not its exact time. Precise dating requires correlation with other artefacts.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Record" "Records")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Cancellazioni/rinomine" "Deletes/renames")" "$NOTABLE" "warn")"
    STATS+="$(stat_box "$(L "File di log" "Log files")" "${#FSEFILES[@]}" "info")"
    [[ "$TOTAL" -gt "$LIMIT" ]] && STATS+="$(stat_box "$(L "Mostrati" "Shown")" "$LIMIT" "info")"

    finish_report "macos_fsevents" "macOS FSEvents" "FSE" "/.fseventsd" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "FSEvents" "$FSEDIR" "$TOTAL" "$TABLE" "⟳")</div>"
}

# --- macOS 12 — Spotlight (provenienza download e metadati) ---
#
# Lo store Spotlight indicizza attributi che sopravvivono al file indicizzato,
# fra cui kMDItemWhereFroms — l'URL da cui un file e' stato scaricato — e i nomi
# di file poi cancellati.
#
# Il formato di store.db e' proprietario, compresso a blocchi, e un parser
# completo esula da un tool senza dipendenze: qui si fa un'estrazione EURISTICA
# delle stringhe leggibili (URL e percorsi). Il report lo dichiara apertamente e
# rimanda a spotlight_parser per l'analisi strutturata.
module_macos_spotlight() {
    section_header "macOS — Spotlight" "$CYAN"
    check_target_root || return 1

    local -a STORES=()
    local D
    while IFS= read -r D; do
        [[ -n "$D" ]] && STORES+=("$D")
    done < <(find "$WIN_ROOT" -maxdepth 6 -type d -name "Store-V2" 2>/dev/null
             find "$WIN_ROOT" -maxdepth 6 -type d -name ".Spotlight-V100" 2>/dev/null)

    mapfile -t DBS < <(find "$WIN_ROOT" -maxdepth 8 -type f \( -name "store.db" -o -name ".store.db" \) 2>/dev/null)
    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessuno store Spotlight trovato." "No Spotlight store found.")"
        return 0
    fi
    info "$(L "Store Spotlight trovati:" "Spotlight stores found:") ${BOLD}${#DBS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os

out_path = sys.argv[1]
dbs = sys.argv[2:]

# Estrazione euristica: si cercano URL e percorsi nelle stringhe leggibili.
URL = re.compile(rb'(?:https?|ftp)://[!-~]{4,300}')
# Percorsi utente: evitano il rumore dei path di sistema.
PATH = re.compile(rb'/Users/[A-Za-z0-9._-]{1,40}/[!-~]{3,200}')

rows = []
seen = set()
for db in dbs:
    try:
        size = os.path.getsize(db)
        with open(db, 'rb') as fh:
            data = fh.read(256 * 1024 * 1024)   # tetto di sicurezza: 256 MB
    except Exception:
        continue
    src = db
    for kind, rx in (("URL", URL), ("path", PATH)):
        for m in rx.finditer(data):
            try:
                val = m.group(0).decode('utf-8', 'strict')
            except Exception:
                continue
            # Scarta stringhe con caratteri di controllo residui
            if any(ord(c) < 32 for c in val):
                continue
            key = (kind, val)
            if key in seen:
                continue
            seen.add(key)
            rows.append((kind, val, src))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write("\t".join(x.replace("\t", " ") for x in r) + "\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    local NURL=0 NPATH=0
    if [[ "$TOTAL" -gt 0 ]]; then
        NURL=$(awk -F'\t' '$1=="URL"' "$OUT" | wc -l)
        NPATH=$(awk -F'\t' '$1=="path"' "$OUT" | wc -l)
    fi

    # Inventario degli store, utile anche quando l'estrazione non produce nulla.
    local INV="" 
    local DB
    for DB in "${DBS[@]}"; do
        local SZ MT
        SZ=$(stat -c %s "$DB" 2>/dev/null || echo "?")
        MT=$(stat -c %y "$DB" 2>/dev/null | cut -d. -f1 || echo "?")
        INV+="$(html_esc "$DB")	${SZ}	${MT}
"
    done
    local INVTABLE; INVTABLE=$(_rows_to_table "$INV" "store.db" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")")

    separator
    ok "$(L "Stringhe estratte:" "Strings extracted:") ${BOLD}$TOTAL"
    info "URL: ${BOLD}${NURL}${RESET}  ·  path: ${BOLD}${NPATH}"

    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna stringa utile estratta dagli store Spotlight." "No useful string extracted from the Spotlight stores.")"
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Metodo" "Method")</b><br>"
    BODY+="$(L "store.db ha un formato proprietario compresso a blocchi. Questo modulo NON lo decodifica: esegue un'estrazione euristica delle stringhe leggibili (URL e percorsi utente), utile per recuperare provenienza dei download e nomi di file poi cancellati. I risultati vanno considerati indizi da confermare, non un dump strutturato dell'indice." \
        "store.db uses a proprietary block-compressed format. This module does NOT decode it: it performs a heuristic extraction of readable strings (URLs and user paths), useful to recover download provenance and names of later-deleted files. Treat results as leads to confirm, not as a structured index dump.")<br><br>"
    BODY+="$(L "Per un'analisi strutturata (kMDItemWhereFroms, date, attributi completi) usa" "For structured analysis (kMDItemWhereFroms, dates, full attributes) use") "
    BODY+="<code>spotlight_parser</code> (Yogesh Khatri)."
    BODY+="</div></div>"
    BODY+="<div class='cards'>$(generic_card_html "$(L "Store trovati" "Stores found")" "$WIN_ROOT" "${#DBS[@]}" "$INVTABLE" "▤")</div>"

    if [[ "$TOTAL" -gt 0 ]]; then
        local ROWS; ROWS=$( { awk -F'\t' '$1=="URL"' "$OUT"; awk -F'\t' '$1=="path"' "$OUT"; } | head -20000 )
        local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Tipo" "Type")" "$(L "Valore" "Value")" "store.db")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Stringhe estratte" "Extracted strings")" "$(L "estrazione euristica" "heuristic extraction")" "$TOTAL" "$TABLE" "⌕")</div>"
    fi

    local STATS
    STATS="$(stat_box "URL" "$NURL" "warn")"
    STATS+="$(stat_box "$(L "Percorsi" "Paths")" "$NPATH")"
    STATS+="$(stat_box "Store" "${#DBS[@]}" "info")"
    finish_report "macos_spotlight" "macOS Spotlight" "SPT" ".Spotlight-V100 / store.db" "$STATS" "$BODY"
}

# ================================================================
#  macOS 14 — Messages (chat.db)
#
#  iMessage e SMS inoltrati dall'iPhone finiscono in chat.db sul Mac. In
#  un'indagine contano per lo smishing (il link arrivato via SMS che ha
#  iniziato la compromissione), per i codici 2FA intercettati e per la
#  corrispondenza scambiata fuori dai canali aziendali.
#
#  Il database conserva anche i messaggi con `is_deleted`, e gli allegati
#  restano su disco in ~/Library/Messages/Attachments anche dopo la
#  cancellazione della conversazione.
# ================================================================
module_macos_messages() {
    section_header "macOS — Messages" "$CYAN"
    check_target_root || return 1

    local BODY="" TOTAL=0 NATT=0 NSUSP=0
    local HOME_DIR
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local DB; DB=$(ci_find_file "$HOME_DIR" "Library/Messages/chat.db")
        [[ -s "$DB" ]] || continue

        # La colonna date e' in nanosecondi dal 2001 su macOS recenti e in
        # secondi su quelli vecchi: si normalizza in SQL guardando l'ordine di
        # grandezza, altrimenti le date risultano nel futuro remoto.
        local ROWS
        ROWS=$(query_sqlite "$DB" "
            SELECT datetime(CASE WHEN m.date > 100000000000
                                 THEN m.date/1000000000 + 978307200
                                 ELSE m.date + 978307200 END, 'unixepoch'),
                   COALESCE(h.id,'?'),
                   CASE m.is_from_me WHEN 1 THEN 'inviato' ELSE 'ricevuto' END,
                   COALESCE(m.text,''),
                   COALESCE(m.service,'')
            FROM message m LEFT JOIN handle h ON m.handle_id = h.ROWID
            WHERE m.text IS NOT NULL AND m.text <> ''
            ORDER BY m.date DESC LIMIT 50000")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue

        local N; N=$(printf '%s\n' "$ROWS" | grep -c . || true)
        TOTAL=$((TOTAL + N))
        ok "$U — ${BOLD}${N}${RESET} $(L "messaggi" "messages")"

        # Messaggi con link o riferimenti a credenziali: sono quelli che
        # spiegano un accesso iniziale o una compromissione di account.
        local SUSP
        SUSP=$(printf '%s\n' "$ROWS" | grep -iE 'https?://|bit\.ly|tinyurl|codice|code|verifica|verif|otp|password|accedi|login|urgente|urgent|bloccat|blocked' || true)
        local NS; NS=$(printf '%s\n' "$SUSP" | grep -c . || true)
        NSUSP=$((NSUSP + NS))

        local TABLE; TABLE=$(_rows_to_table "$ROWS" \
            "$(L "Data" "Date")" "$(L "Interlocutore" "Handle")" "$(L "Direzione" "Direction")" \
            "$(L "Testo" "Text")" "$(L "Servizio" "Service")")
        BODY+=$(generic_card_html "$U" "$DB" "$N" "$TABLE" "✉")

        if [[ "$NS" -gt 0 ]]; then
            local STABLE; STABLE=$(_rows_to_table "$SUSP" \
                "$(L "Data" "Date")" "$(L "Interlocutore" "Handle")" "$(L "Direzione" "Direction")" \
                "$(L "Testo" "Text")" "$(L "Servizio" "Service")")
            BODY="<div class='cards'>$(generic_card_html "$U — $(L "messaggi con link o credenziali" "messages with links or credentials")" "$DB" "$NS" "$STABLE" "⚑")</div>${BODY}"
        fi

        # Allegati: restano su disco anche dopo la cancellazione della chat.
        local ATTDIR; ATTDIR=$(ci_find_dir "$HOME_DIR" "Library/Messages/Attachments")
        if [[ -n "$ATTDIR" ]]; then
            local NA; NA=$(find "$ATTDIR" -type f 2>/dev/null | wc -l)
            NATT=$((NATT + NA))
            [[ "$NA" -gt 0 ]] && info "  $(L "allegati su disco:" "attachments on disk:") ${BOLD}${NA}"
        fi
    done < <(get_macos_user_homes)

    separator
    info "$(L "Messaggi totali:" "Total messages:") ${BOLD}$TOTAL"
    [[ "$NSUSP" -gt 0 ]] && warn "$(L "Con link o riferimenti a credenziali:" "With links or credential references:") ${BOLD}$NSUSP"
    info "$(L "Allegati su disco:" "Attachments on disk:") ${BOLD}$NATT"
    [[ "$TOTAL" -eq 0 ]] && { warn "$(L "Nessun database Messages leggibile." "No readable Messages database.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="$(L "Gli allegati restano in ~/Library/Messages/Attachments anche dopo la cancellazione della conversazione: vanno esaminati a parte, non compaiono in questa tabella. La colonna data e' normalizzata dal formato Apple (riferimento 2001), che su macOS recenti e' in nanosecondi." \
        "Attachments remain in ~/Library/Messages/Attachments even after the conversation is deleted: examine them separately, they are not in this table. The date column is normalised from the Apple epoch (2001 reference), which on recent macOS is in nanoseconds.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Messaggi" "Messages")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Segnalati" "Flagged")" "$NSUSP" "$([[ "$NSUSP" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Allegati" "Attachments")" "$NATT" "info")"
    finish_report "macos_messages" "macOS Messages" "MSG" "Library/Messages/chat.db" "$STATS" "${NOTE}<div class='cards'>$BODY</div>"
}

# ================================================================
#  macOS 15 — Cookie Safari e cronologia download
#
#  Due artefatti che il modulo Browser History non copre:
#
#  - Cookies.binarycookies: formato proprietario Safari. I cookie di sessione
#    dicono a quali servizi l'utente era autenticato e quando; un cookie di un
#    servizio mai usato dall'utente e' un indizio di sessione altrui.
#  - Downloads.plist: la cronologia dei download di Safari, con l'URL DI
#    ORIGINE oltre al nome del file. Il modulo Quarantine copre gli eventi di
#    quarantena, ma non tutti i download vi finiscono.
# ================================================================
module_macos_cookies_downloads() {
    section_header "macOS — Cookie & Download" "$YELLOW"
    check_target_root || return 1

    local -a COOKIES=() PLISTS=()
    local HOME_DIR F
    while IFS= read -r HOME_DIR; do
        while IFS= read -r F; do
            [[ -s "$F" ]] && COOKIES+=("$F")
        done < <(find "$HOME_DIR" -maxdepth 6 -type f -name 'Cookies.binarycookies' 2>/dev/null)
        while IFS= read -r F; do
            [[ -s "$F" ]] && PLISTS+=("$F")
        done < <(find "$HOME_DIR" -maxdepth 6 -type f -name 'Downloads.plist' 2>/dev/null)
    done < <(get_macos_user_homes)

    if [[ ${#COOKIES[@]} -eq 0 && ${#PLISTS[@]} -eq 0 ]]; then
        warn "$(L "Nessun cookie Safari o Downloads.plist trovato." "No Safari cookies or Downloads.plist found.")"
        return 0
    fi

    # ---------- Cookie ----------
    local COUT; COUT=$(mktemp); register_tmp "$COUT"
    local NCOOK=0
    if [[ ${#COOKIES[@]} -gt 0 ]]; then
        "$PY3" - "$COUT" "${COOKIES[@]}" << 'PYEOF' 2>/dev/null
import sys, struct, os, datetime

out_path = sys.argv[1]
files = sys.argv[2:]

def mac_time(v):
    """Secondi dal 2001-01-01 (epoca Apple) -> stringa UTC."""
    try:
        if not v or v <= 0:
            return ''
        return (datetime.datetime(2001, 1, 1) + datetime.timedelta(seconds=float(v))).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ''

def cstr(buf, off):
    end = buf.find(b'\x00', off)
    return buf[off:end if end >= 0 else len(buf)].decode('utf-8', 'replace')

rows = []
for path in files:
    src = os.path.basename(os.path.dirname(path)) or os.path.basename(path)
    try:
        with open(path, 'rb') as fh:
            data = fh.read()
    except Exception:
        continue
    if data[:4] != b'cook':
        continue
    try:
        npages = struct.unpack_from('>I', data, 4)[0]
        sizes = [struct.unpack_from('>I', data, 8 + 4 * i)[0] for i in range(npages)]
    except Exception:
        continue
    off = 8 + 4 * npages
    for psize in sizes:
        page = data[off:off + psize]
        off += psize
        if len(page) < 12:
            continue
        try:
            ncook = struct.unpack_from('<I', page, 4)[0]
            offsets = [struct.unpack_from('<I', page, 8 + 4 * i)[0] for i in range(ncook)]
        except Exception:
            continue
        for co in offsets:
            if co + 56 > len(page):
                continue
            try:
                (_size, _u1, flags, _u2, url_o, name_o, path_o, val_o) = struct.unpack_from('<8I', page, co)
                expiry, creation = struct.unpack_from('<dd', page, co + 40)
            except Exception:
                continue
            try:
                dom  = cstr(page, co + url_o)
                name = cstr(page, co + name_o)
                cpath = cstr(page, co + path_o)
                val  = cstr(page, co + val_o)
            except Exception:
                continue
            if not dom:
                continue
            attrs = []
            if flags & 0x1: attrs.append('Secure')
            if flags & 0x4: attrs.append('HttpOnly')
            # Il VALORE del cookie e' materiale di sessione: se ne riporta solo
            # la lunghezza, perche' un cookie valido nel report e' una credenziale
            # riutilizzabile da chiunque legga il report.
            rows.append((mac_time(creation), dom, name, f"{len(val)} byte",
                         mac_time(expiry), ';'.join(attrs), cpath, src))

rows.sort(key=lambda r: r[0], reverse=True)
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')
PYEOF
        [[ -s "$COUT" ]] && NCOOK=$(wc -l < "$COUT")
        ok "$(L "Cookie Safari:" "Safari cookies:") ${BOLD}$NCOOK"
    fi

    # ---------- Download ----------
    local DOUT; DOUT=$(mktemp); register_tmp "$DOUT"
    local NDL=0
    local P
    for P in "${PLISTS[@]}"; do
        local TXT; TXT=$(read_plist "$P" 2>/dev/null)
        [[ -z "$TXT" ]] && continue
        # Dalla rappresentazione testuale si estraggono le coppie URL/percorso.
        printf '%s\n' "$TXT" | "$PY3" -c "
import sys, re
txt = sys.stdin.read()
urls  = re.findall(r'DownloadEntryURL\W+([^\n<]+)', txt) or re.findall(r'(https?://[^\s<\"]+)', txt)
paths = re.findall(r'DownloadEntryPath\W+([^\n<]+)', txt)
for i, u in enumerate(urls):
    p = paths[i] if i < len(paths) else ''
    print(f\"{u.strip()}\t{p.strip()}\")
" >> "$DOUT" 2>/dev/null || true
    done
    [[ -s "$DOUT" ]] && NDL=$(wc -l < "$DOUT")
    [[ "$NDL" -gt 0 ]] && ok "$(L "Voci di download:" "Download entries:") ${BOLD}$NDL"

    separator
    if [[ "$NCOOK" -eq 0 && "$NDL" -eq 0 ]]; then
        warn "$(L "Nessun dato estratto." "No data extracted.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Il valore dei cookie non viene riportato" "Cookie values are not reported")</b><br>"
    BODY+="$(L "Della parte segreta si indica solo la lunghezza. Un cookie di sessione ancora valido e' una credenziale a tutti gli effetti: stamparlo nel report significherebbe consegnare a chiunque lo legga la possibilita' di impersonare l'utente. Dominio, nome, date e attributi bastano a stabilire a quali servizi l'utente era autenticato e quando." \
        "Only the length of the secret part is shown. A still-valid session cookie is a credential: printing it in the report would hand anyone reading it the ability to impersonate the user. Domain, name, dates and attributes are enough to establish which services the user was authenticated to, and when.")"
    BODY+="</div></div>"

    if [[ "$NCOOK" -gt 0 ]]; then
        local CT; CT=$(_rows_to_table "$(head -20000 "$COUT")" \
            "$(L "Creato" "Created")" "$(L "Dominio" "Domain")" "$(L "Nome" "Name")" \
            "$(L "Valore" "Value")" "$(L "Scadenza" "Expiry")" "$(L "Attributi" "Attributes")" "Path" "$(L "Origine" "Source")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Cookie Safari" "Safari cookies")" "Cookies.binarycookies" "$NCOOK" "$CT" "🍪")</div>"
    fi
    if [[ "$NDL" -gt 0 ]]; then
        local DT; DT=$(_rows_to_table "$(cat "$DOUT")" "$(L "URL di origine" "Source URL")" "$(L "Percorso locale" "Local path")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Download Safari" "Safari downloads")" "Downloads.plist" "$NDL" "$DT" "⤓")</div>"
    fi

    local STATS
    STATS="$(stat_box "Cookie" "$NCOOK")"
    STATS+="$(stat_box "Download" "$NDL" "info")"
    finish_report "macos_cookies_downloads" "macOS Cookie & Download" "CKD" "Cookies.binarycookies · Downloads.plist" "$STATS" "$BODY"
}

# ================================================================
#  macOS 16 — XProtect, Gatekeeper e policy di sistema
#
#  Sono le difese native di macOS. In un'indagine servono a rispondere a due
#  domande diverse:
#
#  - le difese hanno visto qualcosa? (XProtect Remediator, versione delle
#    firme al momento dei fatti)
#  - le difese sono state indebolite? Un'app approvata a mano in SystemPolicy,
#    un kext di terze parti autorizzato in KextPolicy o Gatekeeper disattivato
#    sono azioni deliberate che spesso precedono l'esecuzione del malware.
#
#  La versione delle firme XProtect e' rilevante di per se': se e' molto
#  anteriore alla data dei fatti, l'assenza di rilevamenti non significa nulla.
# ================================================================
module_macos_xprotect() {
    section_header "macOS — XProtect / Gatekeeper" "$RED"
    check_target_root || return 1

    local BODY="" ROWS="" NFIND=0

    # ---------- Versione delle firme XProtect ----------
    local XPV="" XPD=""
    local META
    for META in "Library/Apple/System/Library/CoreServices/XProtect.bundle/Contents/Info.plist" \
                "System/Library/CoreServices/XProtect.bundle/Contents/Info.plist"; do
        local F; F=$(ci_find_file "$WIN_ROOT" "$META")
        [[ -s "$F" ]] || continue
        XPV=$(read_plist "$F" 2>/dev/null | grep -aoE 'CFBundleShortVersionString[^0-9]*([0-9]+)' | grep -oE '[0-9]+$' | head -1)
        XPD=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
        [[ -n "$XPV" ]] && break
    done
    if [[ -n "$XPV" ]]; then
        ok "XProtect $(L "versione firme:" "signature version:") ${BOLD}${XPV}${RESET} ${DIM}(${XPD})${RESET}"
        ROWS+="XProtect	$(L "versione firme" "signature version")	${XPV}	${XPD}
"
    else
        warn "$(L "Versione delle firme XProtect non determinabile." "XProtect signature version not determinable.")"
    fi

    # ---------- Rilevamenti di XProtect Remediator ----------
    local XPDB
    for XPDB in "private/var/protected/xprotect/XPdb" "var/protected/xprotect/XPdb"; do
        local F; F=$(ci_find_file "$WIN_ROOT" "$XPDB")
        [[ -s "$F" ]] || continue
        local R; R=$(query_sqlite "$F" "SELECT * FROM sqlite_master WHERE type='table'")
        [[ -z "$R" || "$R" == ERROR* ]] && continue
        info "XPdb: ${DIM}${F}${RESET}"
        ROWS+="XProtect Remediator	$(L "database presente" "database present")	$(stat -c %s "$F" 2>/dev/null) B	$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
"
    done

    # ---------- SystemPolicy: eseguibili approvati a mano ----------
    local NAPPROVED=0 APPROWS=""
    local SP; SP=$(ci_find_file "$WIN_ROOT" "private/var/db/SystemPolicy")
    [[ -z "$SP" ]] && SP=$(ci_find_file "$WIN_ROOT" "var/db/SystemPolicy")
    if [[ -s "$SP" ]]; then
        local R
        R=$(query_sqlite "$SP" "SELECT datetime(ctime,'unixepoch'), type, COALESCE(requirement,''), COALESCE(remarks,'') FROM authority WHERE allow=1 AND (remarks IS NOT NULL AND remarks <> '') ORDER BY ctime DESC LIMIT 500")
        if [[ -n "$R" && "$R" != ERROR* ]]; then
            NAPPROVED=$(printf '%s\n' "$R" | grep -c . || true)
            APPROWS="$R"
            [[ "$NAPPROVED" -gt 0 ]] && warn "$(L "Autorizzazioni Gatekeeper con annotazione:" "Gatekeeper authorities with remarks:") ${BOLD}$NAPPROVED"
        fi
    fi

    # ---------- KextPolicy: estensioni kernel di terze parti approvate ----------
    local NKEXT=0 KEXTROWS=""
    local KP; KP=$(ci_find_file "$WIN_ROOT" "private/var/db/SystemPolicyConfiguration/KextPolicy")
    [[ -z "$KP" ]] && KP=$(ci_find_file "$WIN_ROOT" "var/db/SystemPolicyConfiguration/KextPolicy")
    if [[ -s "$KP" ]]; then
        local R
        R=$(query_sqlite "$KP" "SELECT team_id, bundle_id, allowed, COALESCE(developer_name,'') FROM kext_policy ORDER BY bundle_id")
        if [[ -n "$R" && "$R" != ERROR* ]]; then
            NKEXT=$(printf '%s\n' "$R" | grep -c . || true)
            KEXTROWS="$R"
            [[ "$NKEXT" -gt 0 ]] && warn "$(L "Estensioni kernel di terze parti registrate:" "Third-party kernel extensions registered:") ${BOLD}$NKEXT"
        fi
    fi

    # ---------- Stato di Gatekeeper ----------
    local GK; GK=$(ci_find_file "$WIN_ROOT" "private/var/db/.LastGKReject")
    [[ -n "$GK" ]] && ROWS+="Gatekeeper	$(L "ultimo rifiuto registrato" "last recorded rejection")	-	$(stat -c %y "$GK" 2>/dev/null | cut -d. -f1)
"
    local SIP; SIP=$(ci_find_file "$WIN_ROOT" "private/var/db/.AppleSetupDone")
    [[ -n "$SIP" ]] && ROWS+="$(L "Setup completato" "Setup completed")	.AppleSetupDone	-	$(stat -c %y "$SIP" 2>/dev/null | cut -d. -f1)
"

    NFIND=$(( NAPPROVED + NKEXT ))
    separator
    if [[ -z "$ROWS" && "$NFIND" -eq 0 ]]; then
        warn "$(L "Nessun artefatto XProtect/Gatekeeper leggibile." "No readable XProtect/Gatekeeper artefact.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    BODY="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "La versione delle firme cambia il significato dell'assenza di rilevamenti" "The signature version changes what 'no detections' means")</b><br>"
    BODY+="$(L "XProtect rileva solo cio' che le sue firme conoscono. Se la versione qui riportata e' molto anteriore alla data dei fatti, il fatto che non abbia segnalato nulla non e' un'assoluzione: e' semplicemente un dato privo di valore probatorio." \
        "XProtect only detects what its signatures know. If the version reported here long predates the events, the absence of detections is not an acquittal: it is simply a fact with no evidential weight.")<br><br>"
    BODY+="<b>$(L "Autorizzazioni concesse a mano" "Manually granted authorisations")</b><br>"
    BODY+="$(L "Una voce in SystemPolicy con annotazione indica un'app che qualcuno ha approvato esplicitamente superando Gatekeeper; una in KextPolicy indica un'estensione kernel di terze parti autorizzata. Entrambe richiedono un'azione deliberata dell'utente e spesso precedono immediatamente l'esecuzione di codice indesiderato." \
        "An entry in SystemPolicy with remarks indicates an app someone explicitly approved past Gatekeeper; one in KextPolicy indicates an authorised third-party kernel extension. Both require a deliberate user action and often immediately precede the execution of unwanted code.")"
    BODY+="</div></div>"

    [[ -n "$ROWS" ]] && BODY+="<div class='cards'>$(generic_card_html "$(L "Stato delle difese" "Defence status")" "$WIN_ROOT" "-" \
        "$(_rows_to_table "$ROWS" "$(L "Componente" "Component")" "$(L "Dato" "Item")" "$(L "Valore" "Value")" "$(L "Data" "Date")")" "🛡")</div>"
    [[ -n "$APPROWS" ]] && BODY+="<div class='cards'>$(generic_card_html "$(L "Autorizzazioni Gatekeeper" "Gatekeeper authorities")" "SystemPolicy" "$NAPPROVED" \
        "$(_rows_to_table "$APPROWS" "$(L "Data" "Date")" "$(L "Tipo" "Type")" "Requirement" "$(L "Annotazione" "Remarks")")" "⚑")</div>"
    [[ -n "$KEXTROWS" ]] && BODY+="<div class='cards'>$(generic_card_html "$(L "Estensioni kernel" "Kernel extensions")" "KextPolicy" "$NKEXT" \
        "$(_rows_to_table "$KEXTROWS" "Team ID" "Bundle ID" "$(L "Consentita" "Allowed")" "$(L "Sviluppatore" "Developer")")" "⚙")</div>"

    local STATS
    STATS="$(stat_box "$(L "Firme XProtect" "XProtect signatures")" "${XPV:-n/d}" "info")"
    STATS+="$(stat_box "$(L "Autorizzazioni" "Authorities")" "$NAPPROVED" "$([[ "$NAPPROVED" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "kext" "$NKEXT" "$([[ "$NKEXT" -gt 0 ]] && echo warn || echo info)")"
    finish_report "macos_xprotect" "macOS XProtect / Gatekeeper" "XPR" "XProtect · SystemPolicy · KextPolicy" "$STATS" "$BODY"
}

# ================================================================
#  macOS 17 — Inventario applicazioni
#
#  Serve a rispondere a "cosa era installato e da dove veniva". Su macOS un
#  bundle .app e' una directory: puo' essere copiato ovunque ed eseguito senza
#  installazione, quindi le applicazioni fuori da /Applications meritano
#  attenzione — soprattutto quelle in ~/Downloads, /tmp o /Users/Shared.
#
#  Nota sulla firma: verificarla richiede `codesign`, disponibile solo su un
#  host macOS. Da Linux si puo' stabilire se il bundle CONTIENE una firma
#  (_CodeSignature) e se e' notarizzato in modo evidente, non se la firma sia
#  valida. Il report distingue le due cose invece di lasciarlo intendere.
# ================================================================
module_macos_applications() {
    section_header "macOS — $(L "Inventario applicazioni" "Application inventory")" "$GREEN"
    check_target_root || return 1

    local -a ROOTS=()
    local D
    for D in "Applications" "Applications/Utilities" "System/Applications" "Library/Application Support"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ROOTS+=("$R")
    done
    local HOME_DIR
    while IFS= read -r HOME_DIR; do
        ROOTS+=("$HOME_DIR")
    done < <(get_macos_user_homes)

    if [[ ${#ROOTS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna directory di applicazioni trovata." "No application directory found.")"
        return 0
    fi

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    info "$(L "Enumerazione dei bundle..." "Enumerating bundles...")"

    local R APP
    for R in "${ROOTS[@]}"; do
        while IFS= read -r APP; do
            [[ -d "$APP" ]] || continue
            local NAME PLIST BID VER SIGNED NOTAR MT LOC
            NAME=$(basename "$APP")
            MT=$(stat -c %y "$APP" 2>/dev/null | cut -d. -f1)
            PLIST="$APP/Contents/Info.plist"
            BID=""; VER=""
            if [[ -s "$PLIST" ]]; then
                local TXT; TXT=$(read_plist "$PLIST" 2>/dev/null | head -400)
                BID=$(printf '%s' "$TXT" | grep -aoE 'CFBundleIdentifier[^A-Za-z0-9]*[A-Za-z0-9.-]+' | grep -oE '[A-Za-z0-9.-]+$' | head -1)
                VER=$(printf '%s' "$TXT" | grep -aoE 'CFBundleShortVersionString[^0-9]*[0-9][0-9A-Za-z._-]*' | grep -oE '[0-9][0-9A-Za-z._-]*$' | head -1)
            fi
            # Presenza della firma, NON sua validita'.
            if [[ -d "$APP/Contents/_CodeSignature" ]]; then
                SIGNED="$(L "firma presente" "signature present")"
            else
                SIGNED="$(L "NESSUNA FIRMA" "NO SIGNATURE")"
            fi
            # La ricevuta App Store distingue l'origine ufficiale.
            if [[ -f "$APP/Contents/_MASReceipt/receipt" ]]; then
                NOTAR="App Store"
            else
                NOTAR="-"
            fi
            # Posizione: fuori dalle directory di sistema e' il dato che conta.
            # NB: */Applications/* copre gia' anche */System/Applications/*.
            case "$APP" in
                */Applications/*) LOC="$(L "standard" "standard")" ;;
                *)                LOC="$(L "FUORI DA /Applications" "OUTSIDE /Applications")" ;;
            esac
            printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
                "$MT" "$NAME" "${BID:-?}" "${VER:-?}" "$SIGNED" "$NOTAR" "$LOC" "$APP" >> "$OUT"
        done < <(find "$R" -maxdepth 4 -type d -name '*.app' -prune 2>/dev/null)
    done

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun bundle applicativo trovato." "No application bundle found.")"
        return 0
    fi
    local NUNSIGNED NOUTSIDE
    NUNSIGNED=$(awk -F'\t' '$5 ~ /NESSUNA|NO SIGNATURE/' "$OUT" | wc -l)
    NOUTSIDE=$(awk -F'\t' '$7 ~ /FUORI|OUTSIDE/' "$OUT" | wc -l)

    ok "$(L "Applicazioni trovate:" "Applications found:") ${BOLD}$TOTAL"
    [[ "$NUNSIGNED" -gt 0 ]] && warn "$(L "Senza firma nel bundle:" "Without a signature in the bundle:") ${BOLD}$NUNSIGNED"
    if [[ "$NOUTSIDE" -gt 0 ]]; then
        warn "$(L "Fuori dalle directory standard:" "Outside standard directories:") ${BOLD}$NOUTSIDE"
        awk -F'\t' '$7 ~ /FUORI|OUTSIDE/{printf "      %s  %s\n", $1, $8}' "$OUT" | head -15 | while IFS= read -r LN; do
            echo -e "      ${YELLOW}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$7 ~ /FUORI|OUTSIDE/' "$OUT"; awk -F'\t' '$5 ~ /NESSUNA|NO SIGNATURE/ && $7 !~ /FUORI|OUTSIDE/' "$OUT"; awk -F'\t' '$5 !~ /NESSUNA|NO SIGNATURE/ && $7 !~ /FUORI|OUTSIDE/' "$OUT"; } )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Applicazione" "Application")" "Bundle ID" \
        "$(L "Versione" "Version")" "$(L "Firma" "Signature")" "$(L "Origine" "Origin")" \
        "$(L "Posizione" "Location")" "$(L "Percorso" "Path")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Cosa dice davvero la colonna Firma" "What the Signature column actually says")</b><br>"
    NOTE+="$(L "Indica se il bundle contiene una directory _CodeSignature, non se la firma sia valida: la verifica richiede codesign, disponibile solo su un host macOS. Un'app puo' avere una firma presente ma non valida, revocata o di uno sviluppatore qualsiasi. Per la verifica: codesign -dv --verbose=4 /percorso/App.app e spctl -a -vv /percorso/App.app." \
        "It indicates whether the bundle contains a _CodeSignature directory, not whether the signature is valid: verification requires codesign, available only on a macOS host. An app may carry a present but invalid or revoked signature, or one from any developer. To verify: codesign -dv --verbose=4 /path/App.app and spctl -a -vv /path/App.app.")<br><br>"
    NOTE+="$(L "Un bundle .app e' una directory: puo' essere copiato ed eseguito da qualunque posizione senza installazione. Le applicazioni fuori dalle directory standard sono elencate per prime." \
        "An .app bundle is a directory: it can be copied and run from anywhere without installation. Applications outside the standard directories are listed first.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Applicazioni" "Applications")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Senza firma" "Unsigned")" "$NUNSIGNED" "$([[ "$NUNSIGNED" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Fuori standard" "Non-standard")" "$NOUTSIDE" "$([[ "$NOUTSIDE" -gt 0 ]] && echo warn || echo info)")"
    finish_report "macos_applications" "macOS Applications" "APP" "/Applications · home utenti" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Applicazioni installate" "Installed applications")" "$WIN_ROOT" "$TOTAL" "$TABLE" "▣")</div>"
}

# ================================================================
#  macOS 18 — Time Machine e snapshot
#
#  Come per le shadow copy su Windows, gli snapshot contengono versioni
#  precedenti dei file: contenuto poi cancellato o cifrato, e configurazioni
#  anteriori alla compromissione.
#
#  Valgono anche al contrario: il ransomware su macOS cancella gli snapshot
#  locali (tmutil deletelocalsnapshots) prima di cifrare, quindi la loro
#  scomparsa su una macchina che li aveva e' un indicatore.
#
#  Limite dichiarato: l'elenco degli snapshot APFS vive nei metadati del
#  container, non nel filesystem montato. Da qui si leggono la configurazione
#  di Time Machine, le destinazioni note e le tracce degli snapshot, non
#  l'elenco autoritativo — che richiede `tmutil listlocalsnapshots` o
#  `diskutil apfs listSnapshots` sul volume vivo.
# ================================================================
module_macos_backups() {
    section_header "macOS — Time Machine / Snapshot" "$BLUE"
    check_target_root || return 1

    local ROWS="" NDEST=0 NTRACE=0

    # ---------- Configurazione Time Machine ----------
    local TMP_PLIST
    for TMP_PLIST in "Library/Preferences/com.apple.TimeMachine.plist" \
                     "private/var/db/com.apple.xpc.launchd/disabled.plist"; do
        local F; F=$(ci_find_file "$WIN_ROOT" "$TMP_PLIST")
        [[ -s "$F" ]] || continue
        local TXT; TXT=$(read_plist "$F" 2>/dev/null)
        [[ -z "$TXT" ]] && continue
        if [[ "$TMP_PLIST" == *TimeMachine* ]]; then
            # Destinazioni, ultimo backup, esclusioni.
            local DESTS; DESTS=$(printf '%s\n' "$TXT" | grep -aoE '(BackupAlias|DestinationID|LastKnownVolumeName|LastDestinationID|SnapshotDates?|LastKnownEncryptionState|AutoBackup)[^<]*<?[^<>]{0,120}' | head -40)
            local LN
            while IFS= read -r LN; do
                [[ -z "$LN" ]] && continue
                NDEST=$((NDEST + 1))
                ROWS+="Time Machine	${LN}	$(basename "$F")
"
            done <<< "$DESTS"
            ok "$(L "Configurazione Time Machine trovata:" "Time Machine configuration found:") ${DIM}${F}${RESET}"
        fi
    done

    # ---------- Tracce di snapshot locali ----------
    local D
    for D in ".MobileBackups" "Volumes/.timemachine" "private/var/db/com.apple.TimeMachine.SnapshotCount"; do
        local X; X=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$X" ]] && X=$(ci_find_file "$WIN_ROOT" "$D")
        [[ -z "$X" ]] && continue
        NTRACE=$((NTRACE + 1))
        ROWS+="Snapshot	${D}	$(stat -c %y "$X" 2>/dev/null | cut -d. -f1)
"
        info "$(L "Traccia di snapshot:" "Snapshot trace:") ${DIM}${X}${RESET}"
    done

    # ---------- Backup montati ----------
    local NBK=0
    local BK; BK=$(ci_find_dir "$WIN_ROOT" "Backups.backupdb")
    if [[ -n "$BK" ]]; then
        while IFS= read -r F; do
            NBK=$((NBK + 1))
            ROWS+="Backup	$(basename "$F")	$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
"
        done < <(find "$BK" -mindepth 2 -maxdepth 2 -type d 2>/dev/null | head -100)
        ok "Backups.backupdb — ${BOLD}${NBK}${RESET} $(L "backup" "backups")"
    fi

    separator
    if [[ -z "$ROWS" ]]; then
        warn "$(L "Nessun artefatto Time Machine o snapshot trovato." "No Time Machine or snapshot artefact found.")"
        echo ""
        warn "$(L "L'assenza va interpretata: su macOS il ransomware cancella gli snapshot locali prima di cifrare (tmutil deletelocalsnapshots). Verificare la history della shell e i log per l'uso di tmutil." \
                 "Absence needs interpreting: on macOS, ransomware deletes local snapshots before encrypting (tmutil deletelocalsnapshots). Check shell history and logs for tmutil usage.")"
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Quello che questo modulo NON puo' dirti" "What this module cannot tell you")</b><br>"
    BODY+="$(L "L'elenco autoritativo degli snapshot APFS sta nei metadati del container, non nel filesystem montato: da un'analisi offline si vedono la configurazione di Time Machine e le tracce lasciate sul volume, non gli snapshot effettivamente esistenti. Per l'elenco reale servono, sul volume vivo o sull'immagine del container: tmutil listlocalsnapshots / e diskutil apfs listSnapshots." \
        "The authoritative list of APFS snapshots lives in the container metadata, not in the mounted filesystem: offline analysis shows the Time Machine configuration and the traces left on the volume, not the snapshots that actually exist. For the real list, on the live volume or the container image: tmutil listlocalsnapshots / and diskutil apfs listSnapshots.")<br><br>"
    BODY+="<b>$(L "Se ci sono, sfruttali" "If they exist, use them")</b><br>"
    BODY+="$(L "Uno snapshot montato e' un volume analizzabile: rilancia FIUTO su di esso e confronta i report per isolare cosa e' cambiato nella finestra di compromissione." \
        "A mounted snapshot is an analysable volume: re-run FIUTO on it and compare the reports to isolate what changed during the compromise window.")"
    BODY+="</div></div>"
    BODY+="<div class='cards'>$(generic_card_html "$(L "Artefatti di backup" "Backup artefacts")" "$WIN_ROOT" "$((NDEST + NTRACE + NBK))" \
        "$(_rows_to_table "$ROWS" "$(L "Tipo" "Type")" "$(L "Dato" "Item")" "$(L "Origine / data" "Source / date")")" "◫")"

    local STATS
    STATS="$(stat_box "$(L "Config TM" "TM config")" "$NDEST" "info")"
    STATS+="$(stat_box "$(L "Tracce snapshot" "Snapshot traces")" "$NTRACE" "$([[ "$NTRACE" -eq 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Backup" "$NBK" "info")"
    finish_report "macos_backups" "macOS Time Machine / Snapshot" "TM" "com.apple.TimeMachine.plist · Backups.backupdb" "$STATS" "$BODY"
}

# ================================================================
#  macOS 19 — Unified Logs (.tracev3)
#
#  E' il registro centrale di macOS: esecuzioni, autenticazioni, rete,
#  installazioni, XProtect. Nelle versioni precedenti FIUTO lo dichiarava
#  fuori scope, ed era il gap piu' evidente dell'analisi macOS.
#
#  ONESTA' SUL LIVELLO DI SUPPORTO. Un parser completo di .tracev3 deve
#  ricostruire il catalogo, risolvere i riferimenti alle stringhe nel
#  dyld_shared_cache e nei file .uuidtext, e reidratare i formati: e' un
#  progetto a se' (cfr. mandiant/macos-UnifiedLogs). Questo modulo NON lo fa.
#
#  Fa una cosa piu' modesta e verificabile: decomprime i chunk LZ4 (bv41) di
#  cui il file e' composto ed estrae le stringhe leggibili che ne emergono —
#  percorsi di processo, nomi di bundle, frammenti di messaggi gia' in chiaro.
#  Senza la decompressione un `strings` sul file grezzo non restituirebbe
#  quasi nulla, perche' il contenuto e' compresso.
#
#  Risultato: indizi datati approssimativamente e cercabili, non una timeline
#  di log ricostruita. Per quella serve `log show --archive` su un Mac.
# ================================================================
module_macos_unified_logs() {
    section_header "macOS — Unified Logs (.tracev3)" "$MAGENTA"
    check_target_root || return 1

    local -a FILES=()
    local D F
    for D in "private/var/db/diagnostics" "var/db/diagnostics"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$R" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && FILES+=("$F")
        done < <(find "$R" -type f -name '*.tracev3' 2>/dev/null | sort)
    done

    if [[ ${#FILES[@]} -eq 0 ]]; then
        warn "$(L "Nessun file .tracev3 trovato." "No .tracev3 file found.")"
        return 0
    fi
    local TOTB=0
    for F in "${FILES[@]}"; do
        TOTB=$(( TOTB + $(stat -c %s "$F" 2>/dev/null || echo 0) ))
    done
    info "$(L "File .tracev3:" ".tracev3 files:") ${BOLD}${#FILES[@]}${RESET} — $(numfmt --to=iec "$TOTB" 2>/dev/null || echo "$TOTB B")"
    info "$(L "Decompressione dei chunk in corso..." "Decompressing chunks...")"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local STATSF; STATSF=$(mktemp); register_tmp "$STATSF"

    "$PY3" - "$OUT" "$STATSF" "$IOCTMP" "${FILES[@]}" << 'PYEOF' 2>/dev/null
import sys, os, re, struct, datetime

out_path, stats_path, ioc_path = sys.argv[1], sys.argv[2], sys.argv[3]
files = sys.argv[4:]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []


def lz4_block(src, expected):
    """Decompressore LZ4 block format in puro Python (nessuna dipendenza).

    Apple incapsula i chunk in 'bv41' + dimensioni + blocco LZ4. Senza questo
    passaggio il contenuto del file resta compresso e illeggibile.
    """
    out = bytearray()
    i, n = 0, len(src)
    while i < n:
        token = src[i]; i += 1
        lit = token >> 4
        if lit == 15:
            while i < n:
                b = src[i]; i += 1
                lit += b
                if b != 255:
                    break
        out += src[i:i + lit]
        i += lit
        if i + 2 > n:
            break
        off = src[i] | (src[i + 1] << 8); i += 2
        if off == 0:
            break
        mlen = token & 0x0F
        if mlen == 15:
            while i < n:
                b = src[i]; i += 1
                mlen += b
                if b != 255:
                    break
        mlen += 4
        start = len(out) - off
        if start < 0:
            break
        for k in range(mlen):
            out.append(out[start + k])
        if expected and len(out) >= expected:
            break
    return bytes(out)


def decompress(data):
    """Concatena il contenuto di tutti i chunk bv41/bv4- del file."""
    out = bytearray()
    pos = 0
    n = len(data)
    while True:
        k = data.find(b'bv4', pos)
        if k < 0 or k + 12 > n:
            break
        tag = data[k:k + 4]
        if tag == b'bv41':
            try:
                usize, csize = struct.unpack_from('<II', data, k + 4)
            except Exception:
                pos = k + 4
                continue
            blob = data[k + 12:k + 12 + csize]
            if 0 < csize <= n and usize < 64 * 1024 * 1024:
                out += lz4_block(blob, usize)
            pos = k + 12 + max(csize, 1)
        elif tag == b'bv4-':
            try:
                usize = struct.unpack_from('<I', data, k + 4)[0]
            except Exception:
                pos = k + 4
                continue
            out += data[k + 8:k + 8 + usize]
            pos = k + 8 + max(usize, 1)
        else:
            pos = k + 3
        if len(out) > 256 * 1024 * 1024:
            break
    return bytes(out)


# Stringhe che vale la pena tenere: percorsi, bundle id, comandi, URL.
KEEP = re.compile(
    rb'(?:/(?:usr|bin|sbin|Applications|Library|System|Users|private|tmp|var|opt)/[!-~]{3,180}'
    rb'|[a-z][a-z0-9]+(?:\.[a-z0-9-]+){2,6}'          # bundle id / reverse-dns
    rb'|https?://[!-~]{4,200}'
    rb'|[A-Za-z0-9_.-]+\.(?:app|kext|plist|dylib|sh|py|pkg|dmg)\b)')

# Rumore ricorrente dei log di sistema: escluderlo rende il resto leggibile.
NOISE = re.compile(rb'(?i)(com\.apple\.(?:coreanimation|uikit|coretext|metal|springboard|coreui))')

rows = []
seen = set()
stats = []

for path in files:
    src = os.path.basename(path)
    try:
        with open(path, 'rb') as fh:
            raw = fh.read(512 * 1024 * 1024)
    except Exception:
        continue
    # Unico ancoraggio temporale disponibile senza ricostruire il catalogo.
    try:
        approx = datetime.datetime.utcfromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        approx = ''
    plain = decompress(raw)
    stats.append((src, str(len(raw)), str(len(plain))))
    if not plain:
        continue
    for m in KEEP.finditer(plain):
        s = m.group(0)
        if NOISE.search(s):
            continue
        try:
            t = s.decode('utf-8', 'strict')
        except Exception:
            continue
        if len(t) < 6:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        flag = 'IOC' if (iocs and any(i in key for i in iocs)) else ''
        rows.append((approx, t[:300], flag, src))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ') for x in r) + '\n')
with open(stats_path, 'w', encoding='utf-8') as fh:
    for s in stats:
        fh.write('\t'.join(s) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    local DECOMP=0
    [[ -s "$STATSF" ]] && DECOMP=$(awk -F'\t' '{s+=$3} END{print s+0}' "$STATSF")

    separator
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna stringa estratta: i chunk non sono stati decompressi (formato non riconosciuto)." \
                 "No string extracted: chunks were not decompressed (unrecognised format).")"
        info "$(L "Per l'analisi completa serve un Mac: log show --archive <cartella>.logarchive" \
                 "Full analysis requires a Mac: log show --archive <folder>.logarchive")"
        return 0
    fi
    ok "$(L "Byte decompressi:" "Bytes decompressed:") ${BOLD}$(numfmt --to=iec "$DECOMP" 2>/dev/null || echo "$DECOMP")"
    ok "$(L "Stringhe rilevanti estratte:" "Relevant strings extracted:") ${BOLD}$TOTAL"
    local NIOC; NIOC=$(awk -F'\t' '$3=="IOC"' "$OUT" | wc -l)
    [[ "$NIOC" -gt 0 ]] && warn "$(L "Con match IoC:" "With IoC match:") ${BOLD}$NIOC"

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$3=="IOC"' "$OUT"; awk -F'\t' '$3!="IOC"' "$OUT"; } | head -30000 )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data file (approx.)" "File date (approx.)")" "$(L "Stringa" "String")" "IoC" "$(L "Origine" "Source")")
    local STABLE; STABLE=$(_rows_to_table "$(cat "$STATSF")" "$(L "File" "File")" "$(L "Byte su disco" "Bytes on disk")" "$(L "Byte decompressi" "Bytes decompressed")")

    local NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,166,87,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Livello di supporto: parziale, e dichiarato" "Support level: partial, and stated")</b><br>"
    NOTE+="$(L "Questo modulo NON ricostruisce i messaggi di log. Un parser completo di .tracev3 deve interpretare il catalogo e risolvere i riferimenti alle stringhe nei file .uuidtext e nel dyld_shared_cache: e' un progetto a se'. Qui i chunk LZ4 vengono decompressi e se ne estraggono le stringhe gia' leggibili — percorsi, bundle id, URL, nomi di file." \
        "This module does NOT reconstruct log messages. A complete .tracev3 parser must interpret the catalogue and resolve string references in .uuidtext files and the dyld_shared_cache: that is a project of its own. Here the LZ4 chunks are decompressed and the already-readable strings are extracted — paths, bundle ids, URLs, file names.")<br><br>"
    NOTE+="$(L "Le date sono quelle di modifica del file che contiene la stringa: un limite superiore approssimato, non l'istante dell'evento." \
        "Dates are the modification times of the file containing the string: an approximate upper bound, not the moment of the event.")<br><br>"
    NOTE+="<b>$(L "Per l'analisi completa" "For full analysis")</b><br>"
    NOTE+="<code>log show --archive /percorso/diagnostics.logarchive --info --debug</code> ($(L "richiede un Mac" "requires a Mac")) — "
    NOTE+="$(L "oppure" "or") <code>mandiant/macos-UnifiedLogs</code>."
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Stringhe" "Strings")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Decompressi" "Decompressed")" "$(numfmt --to=iec "$DECOMP" 2>/dev/null || echo "$DECOMP")" "info")"
    STATS+="$(stat_box "File" "${#FILES[@]}" "info")"
    STATS+="$(stat_box "IoC" "$NIOC" "$([[ "$NIOC" -gt 0 ]] && echo warn || echo info)")"
    finish_report "macos_unified_logs" "macOS Unified Logs" "ULG" "private/var/db/diagnostics/*.tracev3" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Resa della decompressione" "Decompression yield")" "diagnostics" "${#FILES[@]}" "$STABLE" "∑")</div><div class='cards'>$(generic_card_html "$(L "Stringhe estratte" "Extracted strings")" "$(L "estrazione parziale" "partial extraction")" "$TOTAL" "$TABLE" "⌕")</div>"
}

# ================================================================
#  MASTER TIMELINE CROSS-MODULO (Linux/macOS)
#  Aggrega tutte le evidenze con timestamp dai report generati in sessione.
#  Pensato per girare per ULTIMO (è l'ultima voce dei registri Linux/macOS):
#  legge GENERATED_REPORTS, che a quel punto contiene gli altri moduli.
# ================================================================
module_xplat_master_timeline() {
    section_header "$(L "Master Timeline — Aggregazione Cross-Moduli" "Master Timeline — Cross-Module Aggregation")" "$YELLOW"
    check_target_root || return 1

    local -a SRC=()
    local r
    for r in "${GENERATED_REPORTS[@]}"; do
        [[ "$r" == *master_timeline* ]] && continue
        [[ "$r" == */index.html ]] && continue
        [[ -f "$r" ]] && SRC+=("$r")
    done
    if [[ ${#SRC[@]} -eq 0 ]]; then
        warn "$(L "Nessun report generato in questa sessione." "No reports generated in this session.")"
        info "$(L "Esegui prima gli altri moduli (o usa 'Esegui TUTTI'), poi la Master Timeline." "Run the other modules first (or 'Run ALL'), then the Master Timeline.")"
        return 0
    fi
    info "$(L "Report da aggregare:" "Reports to aggregate:") ${BOLD}${#SRC[@]}"

    # Estrae gli eventi con timestamp da tabelle (<tr>) e righe di testo (<span class="line">).
    local YEAR; YEAR=$(date +%Y)
    local TL_RAW
    TL_RAW=$("$PY3" - "$YEAR" "${SRC[@]}" << 'PYEOF' 2>/dev/null || true
import sys, re, html as H
year = sys.argv[1]
MONTHS = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
          'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
TS_ISO = re.compile(r'\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}')
TS_SYS = re.compile(r'\b(' + '|'.join(MONTHS) + r')\s+(\d{1,2})\s+(\d{2}:\d{2}:\d{2})')
TD     = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL | re.I)
PRE    = re.compile(r'<pre[^>]*>(.*?)</pre>', re.DOTALL | re.I)
TAG    = re.compile(r'<[^>]+>')
def strip(s): return TAG.sub('', H.unescape(s)).strip()
def iso(m):   return m.group(0).replace('T', ' ')[:19]
def sysfmt(m):return f"{year}-{MONTHS[m.group(1)]}-{int(m.group(2)):02d} {m.group(3)}"

seen = set(); out = []
def add(ts, mod, desc):
    desc = ' '.join(desc.split())[:200]
    k = (ts, mod, desc[:50])
    if k in seen: return
    seen.add(k); out.append(f"{ts}\t{mod}\t{desc}")

for rp in sys.argv[2:]:
    parts = rp.rstrip('/').split('/')
    mod = re.sub(r'_\d{8}_\d{6}$', '', parts[-2]) if len(parts) >= 2 else parts[-1]
    try:
        content = open(rp, encoding='utf-8', errors='replace').read()
    except Exception:
        continue
    # 1) Righe di tabella con un timestamp in una cella
    for tr in re.finditer(r'<tr[^>]*>(.*?)</tr>', content, re.DOTALL | re.I):
        cells = [strip(m.group(1)) for m in TD.finditer(tr.group(1))]
        if not cells: continue
        ts = ''
        for c in cells:
            m = TS_ISO.search(c)
            if m: ts = iso(m); break
        if not ts: continue
        desc = ' | '.join(c[:90] for c in cells
                          if c and not TS_ISO.fullmatch(c.strip()) and not re.fullmatch(r'\d+', c.strip()))
        add(ts, mod, desc)
    # 2) Blocchi <pre> (log, history, config): tolgo i tag mantenendo i newline,
    #    poi valuto ogni riga (con timestamp ISO, [ISO] o syslog "Mon DD HH:MM:SS").
    for pm in PRE.finditer(content):
        block = H.unescape(TAG.sub('', pm.group(1)))
        for raw in block.split('\n'):
            txt = re.sub(r'^\s*\d+\s+', '', raw).strip()   # toglie il numero di riga
            if not txt: continue
            m = TS_ISO.search(txt)
            if m:
                ts = iso(m)
            else:
                m = TS_SYS.search(txt)
                if not m: continue
                ts = sysfmt(m)
            add(ts, mod, txt)

print('\n'.join(out))
PYEOF
)

    local TOTAL; TOTAL=$(printf '%s' "$TL_RAW" | grep -c $'\t' 2>/dev/null || echo 0)
    local _TMP; _TMP=$(mktemp)
    printf '%s\n' "$TL_RAW" | awk 'NF' | sort -t$'\t' -k1 -r | head -100000 > "$_TMP"
    local SHOW; SHOW=$(wc -l < "$_TMP" 2>/dev/null || echo 0)

    separator
    info "$(L "Eventi con timestamp:" "Events with timestamp:") ${BOLD}$TOTAL${RESET} ($(L "mostrati" "showing") $SHOW)"
    if [[ "$TOTAL" -eq 0 ]]; then
        rm -f "$_TMP"
        warn "$(L "Nessun timestamp trovato nei report aggregati." "No timestamp found in the aggregated reports.")"
        return 0
    fi
    ask_yn "$(L "Generare report HTML Master Timeline?" "Generate Master Timeline HTML report?")" || { rm -f "$_TMP"; return 0; }

    local TABLE; TABLE=$("$PY3" - "$_TMP" << 'PYEOF'
import sys, html
print("<table><tr><th style='width:15%'>Timestamp</th><th style='width:16%'>Modulo</th><th>Descrizione</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    c = line.rstrip('\n').split('\t')
    if len(c) < 3: continue
    print(f"<tr><td class='mono ok' style='white-space:nowrap'>{html.escape(c[0])}</td>"
          f"<td class='mono'>{html.escape(c[1])}</td>"
          f"<td class='mono' style='color:var(--text)'>{html.escape(c[2])}</td></tr>")
print("</table>")
PYEOF
)
    rm -f "$_TMP"
    local BODY; BODY=$(generic_card_html "$(L "Timeline cronologica (decrescente)" "Chronological timeline (descending)")" "$WIN_ROOT" "$SHOW $(L "eventi" "events")" "$TABLE" "◷")
    local STATS; STATS="$(stat_box "$(L "Report aggregati" "Aggregated reports")" "${#SRC[@]}")$(stat_box "$(L "Eventi" "Events")" "$TOTAL" "info")$(stat_box "$(L "Mostrati" "Shown")" "$SHOW" "ok")"
    finish_report "master_timeline" "Master Timeline" "TL" "$(L "Aggregazione cross-moduli" "Cross-module aggregation")" "$STATS" "<div class='cards'>$BODY</div>"
}

# ================================================================
#  CROSS-OS — Recupero da spazio libero SQLite
#
#  Praticamente ogni artefatto moderno e' un database SQLite: cronologia dei
#  browser, TCC, KnowledgeC, quarantena, Messages, cookie di Chrome, chat.
#  Tutti i moduli che li leggono vedono pero' solo i record VIVI.
#
#  Un record cancellato non sparisce dal file: la pagina finisce nella
#  freelist, o lo spazio che occupava diventa un freeblock. Il contenuto resta
#  leggibile finche' non viene sovrascritto. Quando l'utente "ha cancellato la
#  cronologia", questo e' spesso l'unico posto dove quella cronologia esiste
#  ancora.
#
#  Il modulo attraversa i database del volume e ne estrae le stringhe dallo
#  spazio non allocato. Vale per tutti e tre i sistemi operativi.
# ================================================================
module_xplat_sqlite_recovery() {
    section_header "$(L "Recupero record cancellati — SQLite" "Deleted record recovery — SQLite")" "$MAGENTA"
    check_target_root || return 1

    # Database che valgono la pena in un'indagine, per OS.
    local -a PATTERNS=(
        "History" "places.sqlite" "cookies.sqlite" "Cookies" "Web Data" "Login Data"
        "chat.db" "TCC.db" "knowledgeC.db" "QuarantineEventsV2*" "Downloads.sqlite"
        "ActivitiesCache.db" "SRUDB.dat" "index.sqlite" "Extension Cookies"
        "History.db" "*.sqlite" "*.db"
    )

    info "$(L "Ricerca dei database SQLite..." "Searching for SQLite databases...")"
    local MANIFEST; MANIFEST=$(mktemp); register_tmp "$MANIFEST"

    # Si limita alle aree utente e alle directory applicative: una scansione
    # dell'intero volume su un disco reale richiederebbe ore e restituirebbe
    # soprattutto database di sistema senza interesse.
    local -a ROOTS=()
    local HOME_DIR
    while IFS= read -r HOME_DIR; do
        [[ -n "$HOME_DIR" ]] && ROOTS+=("$HOME_DIR")
    done < <(get_target_user_homes)
    local D
    for D in "ProgramData" "private/var/db" "var/db" "var/lib"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ROOTS+=("$R")
    done
    [[ ${#ROOTS[@]} -eq 0 ]] && ROOTS=("$WIN_ROOT")

    local R F
    for R in "${ROOTS[@]}"; do
        while IFS= read -r F; do
            # Il magic e' l'unico criterio affidabile: molti database SQLite
            # non hanno estensione .db o .sqlite (Chrome li chiama "History").
            [[ -s "$F" ]] || continue
            if [[ "$(head -c 15 "$F" 2>/dev/null)" == "SQLite format 3" ]]; then
                printf '%s\n' "$F" >> "$MANIFEST"
            fi
        done < <(find "$R" -maxdepth 8 -type f -size +1k -size -512M 2>/dev/null | head -4000)
    done

    if [[ ! -s "$MANIFEST" ]]; then
        warn "$(L "Nessun database SQLite trovato." "No SQLite database found.")"
        return 0
    fi
    sort -u "$MANIFEST" -o "$MANIFEST"
    local NDB; NDB=$(wc -l < "$MANIFEST")
    info "$(L "Database SQLite individuati:" "SQLite databases found:") ${BOLD}$NDB"
    info "$(L "Carving dello spazio non allocato..." "Carving unallocated space...")"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local SUM; SUM=$(mktemp); register_tmp "$SUM"

    run_py_with_lib pylib_sqlite_recover "$MANIFEST" "$IOCTMP" "$OUT" "$SUM" << 'PYEOF' 2>/dev/null
import sys, os, re

manifest, ioc_path, out_path, sum_path = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

# Un flag binario "rilevante" non aiuta: su una cronologia cancellata ogni
# stringa contiene un URL e finirebbe marcata, azzerando il valore di triage.
# Si classifica invece per tipo, e si ordina mettendo per prime le categorie
# rare (credenziali, eseguibili) rispetto a quelle abbondanti (URL).
CAT_CRED = re.compile(r'(?i)(password|passwd|token|secret|api[_-]?key|bearer |authorization|BEGIN [A-Z ]*PRIVATE KEY)')
CAT_EXE  = re.compile(r'(?i)\.(exe|dll|ps1|vbs|js|bat|cmd|scr|sh|jar|apk|dmg|pkg)\b')
CAT_DOC  = re.compile(r'(?i)\.(docx?|xlsx?|pptx?|pdf|zip|7z|rar|csv|sql|bak)\b')
CAT_PATH = re.compile(r'(?i)(/Users/|/home/|[A-Z]:\\\\)')
CAT_URL  = re.compile(r'(?i)(https?://|ftp://)')

def classify(t):
    if CAT_CRED.search(t):
        return 'credenziale'
    if CAT_EXE.search(t):
        return 'eseguibile'
    if CAT_DOC.search(t):
        return 'documento'
    if CAT_PATH.search(t):
        return 'percorso'
    if CAT_URL.search(t):
        return 'url'
    return ''

# Ordine di priorita' per il triage: le categorie rare prima.
PRIORITY = {'credenziale': 0, 'eseguibile': 1, 'documento': 2, 'percorso': 3, 'url': 4, '': 5}

rows = []
per_db = []
for line in open(manifest, encoding='utf-8', errors='replace'):
    path = line.rstrip('\n')
    if not path:
        continue
    try:
        found = sqlite_carve(path, cap=4000)
    except Exception:
        continue
    if not found:
        continue
    name = os.path.basename(path)
    nnot = 0
    for origin, text in found:
        cat = classify(text)
        if iocs and any(i in text.lower() for i in iocs):
            cat = (cat + ';IOC') if cat else 'IOC'
        if cat.split(';')[0] in ('credenziale', 'eseguibile', 'documento') or 'IOC' in cat:
            nnot += 1
        rows.append((name, origin, cat, text, path))
    per_db.append((name, str(len(found)), str(nnot), path))

rows.sort(key=lambda r: (0 if 'IOC' in r[2] else 1,
                         PRIORITY.get(r[2].split(';')[0], 5), r[0]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
with open(sum_path, 'w', encoding='utf-8') as fh:
    for r in sorted(per_db, key=lambda x: -int(x[1])):
        fh.write('\t'.join(r) + '\n')
PYEOF

    local TOTAL=0 NDBHIT=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    [[ -s "$SUM" ]] && NDBHIT=$(wc -l < "$SUM")
    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessun contenuto recuperabile dallo spazio libero dei database." "No recoverable content in the databases' free space.")"
        return 0
    fi
    local NNOT NIOC
    # "Prioritarie" = credenziali, eseguibili, documenti. Gli URL sono esclusi
    # di proposito: su una cronologia cancellata sono la norma, non il segnale.
    NNOT=$(awk -F'\t' '$3 ~ /credenziale|eseguibile|documento/' "$OUT" | wc -l)
    NIOC=$(awk -F'\t' '$3 ~ /IOC/' "$OUT" | wc -l)

    ok "$(L "Stringhe recuperate:" "Strings recovered:") ${BOLD}$TOTAL"
    info "$(L "Da" "From") ${BOLD}${NDBHIT}${RESET} $(L "database su" "databases out of") ${NDB}"
    [[ "$NNOT" -gt 0 ]] && warn "$(L "Prioritarie (credenziali, eseguibili, documenti):" "Priority (credentials, executables, documents):") ${BOLD}$NNOT"
    [[ "$NIOC" -gt 0 ]] && warn "$(L "Con match IoC:" "With IoC match:") ${BOLD}$NIOC"
    echo ""
    awk -F'\t' '$3 ~ /credenziale|eseguibile|documento|IOC/{printf "      %-14s %-12s %s\n", $1, $3, substr($4,1,70)}' "$OUT" | head -15 | while IFS= read -r LN; do
        echo -e "      ${MAGENTA}${LN}${RESET}"
    done

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(head -25000 "$OUT" | awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4}')
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "Database" "$(L "Origine" "Origin")" "$(L "Categoria" "Category")" "$(L "Contenuto recuperato" "Recovered content")")
    local STABLE; STABLE=$(_rows_to_table "$(cat "$SUM")" \
        "Database" "$(L "Stringhe" "Strings")" "$(L "Prioritarie" "Priority")" "$(L "Percorso" "Path")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Che cosa sono queste stringhe" "What these strings are")</b><br>"
    NOTE+="$(L "Provengono dallo spazio che SQLite considera libero: pagine finite nella freelist (colonna 'freelist') e spazio non allocato dentro pagine ancora in uso (colonna 'unallocated'). In pratica sono i resti di record CANCELLATI, ancora presenti perche' non sovrascritti." \
        "They come from space SQLite considers free: pages moved to the freelist (origin 'freelist') and unallocated space inside pages still in use (origin 'unallocated'). In practice they are the remains of DELETED records, still present because not yet overwritten.")<br><br>"
    NOTE+="<b>$(L "Limite: sono frammenti, non record" "Limitation: fragments, not records")</b><br>"
    NOTE+="$(L "Il modulo estrae stringhe, non ricostruisce righe di tabella: campi adiacenti possono comparire concatenati e all'inizio puo' esserci qualche byte di intestazione del record. Ricostruire un record cancellato richiede lo schema e l'interpretazione dei serial type, ed e' un'operazione fragile: un record ricomposto male in una perizia e' peggio di nessun record. Va quindi trattato come indizio da corroborare, non come contenuto autoritativo." \
        "The module extracts strings, it does not rebuild table rows: adjacent fields may appear concatenated and a few record-header bytes may prefix the text. Rebuilding a deleted record requires the schema and serial-type interpretation, and is fragile: a badly reassembled record in an expert report is worse than none. Treat this as a lead to corroborate, not as authoritative content.")<br><br>"
    NOTE+="<b>$(L "Le categorie" "The categories")</b><br>"
    NOTE+="$(L "Le stringhe sono classificate per tipo e ordinate mettendo per prime le categorie rare. Gli URL sono deliberatamente in fondo: recuperare una cronologia cancellata produce migliaia di URL, quindi marcarli tutti come rilevanti non aiuterebbe a decidere da dove iniziare." \
        "Strings are classified by type and ordered with the rare categories first. URLs are deliberately last: recovering a deleted history yields thousands of URLs, so flagging them all as notable would not help decide where to start.")<br><br>"
    NOTE+="$(L "Non trovare nulla non significa che l'utente non abbia cancellato: significa che lo spazio e' stato riutilizzato, o che il database e' stato compattato con VACUUM — operazione che azzera proprio questo tipo di recupero ed e' essa stessa degna di nota." \
        "Finding nothing does not mean the user deleted nothing: it means the space was reused, or the database was compacted with VACUUM — an operation that wipes exactly this kind of recovery and is itself worth noting.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Stringhe" "Strings")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Prioritarie" "Priority")" "$NNOT" "$([[ "$NNOT" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Database" "${NDBHIT}/${NDB}" "info")"
    STATS+="$(stat_box "IoC" "$NIOC" "$([[ "$NIOC" -gt 0 ]] && echo warn || echo info)")"
    finish_report "xplat_sqlite_recovery" "SQLite — $(L "record cancellati" "deleted records")" "SQL" \
        "$(L "freelist e spazio non allocato" "freelist and unallocated space")" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Resa per database" "Yield per database")" "$(L "riepilogo" "summary")" "$NDBHIT" "$STABLE" "∑")</div><div class='cards'>$(generic_card_html "$(L "Contenuto recuperato" "Recovered content")" "$(L "segnalati in testa" "flagged first")" "$TOTAL" "$TABLE" "♺")</div>"
}

# ================================================================
#  CROSS-OS — EFI System Partition e bootkit
#
#  La ESP e' una piccola partizione FAT che il firmware legge all'accensione
#  per trovare il bootloader. Il codice che sta li' viene eseguito PRIMA del
#  sistema operativo, del kernel e di qualunque agente EDR: e' la posizione di
#  persistenza piu' ambita, e sopravvive alla reinstallazione del sistema e
#  alla formattazione della partizione di sistema.
#
#  I bootkit noti (ESPecter, BlackLotus, Bootkitty, CosmicStrand) agiscono
#  sostituendo o affiancando i loader legittimi. Il modulo inventaria la ESP,
#  calcola gli hash di tutto cio' che vi trova e segnala le anomalie
#  strutturali: file non-EFI, loader in posizioni inattese, date isolate.
#
#  Vale per tutti e tre i sistemi: la ESP e' condivisa fra gli OS installati.
# ================================================================
module_xplat_esp_bootkit() {
    section_header "EFI System Partition — bootkit" "$RED"
    check_target_root || return 1

    # La ESP puo' essere il volume stesso (montata a parte) oppure trovarsi
    # sotto /boot/efi, /efi o /Volumes/EFI del volume in analisi.
    local -a ESPS=()
    local D
    for D in "EFI" "boot/efi/EFI" "efi/EFI" "Volumes/EFI/EFI" "boot/EFI"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ESPS+=("$R")
    done

    if [[ ${#ESPS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna EFI System Partition raggiungibile da questo volume." \
                 "No EFI System Partition reachable from this volume.")"
        info "$(L "La ESP e' una partizione separata: se non e' montata, va montata a parte (di norma la prima partizione FAT32 del disco) e analizzata indicandola come root." \
                 "The ESP is a separate partition: if not mounted, mount it separately (usually the disk's first FAT32 partition) and analyse it as the root.")"
        return 0
    fi

    local ESP; ESP="${ESPS[0]}"
    info "ESP: ${BOLD}${ESP}"

    # Loader legittimi attesi. Un nome fuori da questo elenco non e' di per se'
    # malevolo (molte distribuzioni ne aggiungono di propri), ma va giustificato.
    local KNOWN="bootx64.efi|bootia32.efi|bootaa64.efi|bootmgfw.efi|bootmgr.efi|memtest.efi|shimx64.efi|shimaa64.efi|shim.efi|grubx64.efi|grubaa64.efi|mmx64.efi|fbx64.efi|mokmanager.efi|systemd-bootx64.efi|fwupdx64.efi|kernel.efi|vmlinuz.efi|refind_x64.efi|BOOT.EFI|apfs.efi|boot.efi|firmware.scap|immutablekernel"

    # Nomi ricorrenti nei bootkit documentati e nei loro payload.
    local BADNAMES="grubx64_real|bootmgfw_original|bootmgfw\.efi\.bak|winload\.efi|bootkit|especter|blacklotus|bootlicker|cosmicstrand|\.sys$|\.dll$|\.ps1$|\.bat$|\.vbs$|\.exe$"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local NFILE=0 NEFI=0 NNONEFI=0 NUNKNOWN=0 NSUSP=0
    local F
    while IFS= read -r F; do
        [[ -f "$F" ]] || continue
        NFILE=$((NFILE + 1))
        local BASE REL SZ MT HASH FLAGS
        BASE=$(basename "$F")
        REL="${F#$ESP}"
        SZ=$(stat -c %s "$F" 2>/dev/null || echo 0)
        MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
        HASH=$(sha256_file "$F")
        FLAGS=""

        # Un binario EFI e' un PE: inizia con MZ. Un file nella ESP che non lo
        # e' — e non e' un file di configurazione noto — non ha motivo di stare li'.
        local MAGIC; MAGIC=$(head -c 2 "$F" 2>/dev/null)
        if [[ "$MAGIC" == "MZ" ]]; then
            NEFI=$((NEFI + 1))
        else
            case "${BASE,,}" in
                *.efi) FLAGS+="EFI_SENZA_HEADER_PE;"; NSUSP=$((NSUSP + 1)) ;;
                *.cfg|*.conf|*.ini|*.json|*.txt|bcd|*.crt|*.cer|*.esl|*.auth|*.scap|*.dat) : ;;
                *) FLAGS+="NON_EFI;"; NNONEFI=$((NNONEFI + 1)) ;;
            esac
        fi
        if ! printf '%s' "${BASE,,}" | grep -qiE "^(${KNOWN})$"; then
            case "${BASE,,}" in
                *.efi) FLAGS+="LOADER_NON_STANDARD;"; NUNKNOWN=$((NUNKNOWN + 1)) ;;
            esac
        fi
        if printf '%s' "${BASE,,}" | grep -qiE "$BADNAMES"; then
            FLAGS+="NOME_SOSPETTO;"
            NSUSP=$((NSUSP + 1))
        fi
        printf '%s\t%s\t%s\t%s\t%s\n' "$MT" "${REL#/}" "$SZ" "${FLAGS%;}" "$HASH" >> "$OUT"
    done < <(find "$ESP" -type f 2>/dev/null | sort)

    separator
    ok "$(L "File nella ESP:" "Files in the ESP:") ${BOLD}$NFILE"
    info "$(L "Binari EFI (header PE):" "EFI binaries (PE header):") ${BOLD}$NEFI"
    [[ "$NUNKNOWN" -gt 0 ]] && warn "$(L "Loader non standard:" "Non-standard loaders:") ${BOLD}$NUNKNOWN"
    [[ "$NNONEFI" -gt 0 ]] && warn "$(L "File non-EFI nella ESP:" "Non-EFI files in the ESP:") ${BOLD}$NNONEFI"
    [[ "$NSUSP" -gt 0 ]] && warn "$(L "Segnalazioni forti:" "Strong flags:") ${BOLD}$NSUSP"
    if [[ "$NUNKNOWN" -gt 0 || "$NNONEFI" -gt 0 || "$NSUSP" -gt 0 ]]; then
        echo ""
        awk -F'\t' '$4!=""{printf "      [%s] %s\n", $4, $2}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi
    [[ "$NFILE" -eq 0 ]] && { warn "$(L "ESP vuota o illeggibile." "ESP empty or unreadable.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$4!=""' "$OUT"; awk -F'\t' '$4==""' "$OUT"; } )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Percorso nella ESP" "Path in ESP")" \
        "$(L "Byte" "Bytes")" "$(L "Segnalazioni" "Flags")" "SHA256")

    local NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' la ESP conta piu' di quanto sembri" "Why the ESP matters more than it looks")</b><br>"
    NOTE+="$(L "Il codice nella ESP viene eseguito prima del sistema operativo, del kernel e di qualunque agente EDR. Una persistenza qui sopravvive alla reinstallazione del sistema e alla formattazione della partizione di sistema: se un incidente si ripresenta dopo un ripristino completo, questo e' il primo posto da guardare." \
        "Code in the ESP runs before the operating system, the kernel and any EDR agent. Persistence here survives OS reinstallation and formatting of the system partition: if an incident recurs after a full rebuild, this is the first place to look.")<br><br>"
    NOTE+="<b>$(L "Come usare gli hash" "How to use the hashes")</b><br>"
    NOTE+="$(L "Gli SHA256 in tabella vanno confrontati con quelli dei loader legittimi della distribuzione o della versione di Windows installata, e cercati nelle basi di reputazione. Un bootx64.efi con un hash che non corrisponde a nessuna build ufficiale e' il reperto." \
        "The SHA256 values in the table should be compared against the legitimate loaders of the installed distribution or Windows build, and looked up in reputation databases. A bootx64.efi whose hash matches no official build is the finding.")<br><br>"
    NOTE+="<b>$(L "Le segnalazioni" "The flags")</b><br>"
    NOTE+="$(L "LOADER_NON_STANDARD significa 'nome non nell'elenco dei loader noti': molte distribuzioni ne aggiungono di legittimi, quindi va giustificato, non temuto. NON_EFI e EFI_SENZA_HEADER_PE sono piu' pesanti: un file nella ESP che non e' un binario EFI ne' una configurazione nota non ha una ragione ovvia per stare li'." \
        "NON_STANDARD_LOADER means 'name not in the known-loader list': many distributions add legitimate ones, so it needs justifying, not fearing. NON_EFI and EFI_WITHOUT_PE_HEADER are heavier: a file in the ESP that is neither an EFI binary nor a known configuration has no obvious reason to be there.")<br><br>"
    NOTE+="$(L "Per un controllo mirato sulle famiglie note serve YARA con regole aggiornate: questo modulo non ne include, perche' regole ferme al momento del rilascio darebbero una falsa sensazione di copertura." \
        "For targeted checks against known families use YARA with current rules: this module ships none, because rules frozen at release time would give a false sense of coverage.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "File" "Files")" "$NFILE")"
    STATS+="$(stat_box "$(L "Binari EFI" "EFI binaries")" "$NEFI" "info")"
    STATS+="$(stat_box "$(L "Non standard" "Non-standard")" "$NUNKNOWN" "$([[ "$NUNKNOWN" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Segnalati" "Flagged")" "$((NNONEFI + NSUSP))" "$([[ $((NNONEFI + NSUSP)) -gt 0 ]] && echo warn || echo info)")"
    finish_report "xplat_esp_bootkit" "EFI System Partition" "ESP" "$(L "bootkit e persistenza pre-boot" "bootkits and pre-boot persistence")" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Contenuto della ESP" "ESP contents")" "$ESP" "$NFILE" "$TABLE" "⏻")</div>"
}

# ================================================================
#  CROSS-OS — Scansione YARA
#
#  YARA e' il formato con cui l'industria distribuisce le firme: un feed di
#  threat intelligence, l'IR di un vendor o il CERT nazionale mandano regole
#  .yar, e finora FIUTO non aveva modo di applicarle. Con --yara le applica.
#
#  IL PUNTO DELICATO E' L'AMBITO. Scansionare un volume da un terabyte file per
#  file non e' praticabile su una workstation forense, e un report che non
#  dichiarasse cosa ha guardato sarebbe peggio di nessun report: "nessun match"
#  verrebbe letto come "il disco e' pulito". Il modulo scansiona quindi un
#  insieme limitato di posizioni — quelle scrivibili senza privilegi, dove il
#  codice non installato dai pacchetti finisce quasi sempre — e ELENCA NEL
#  REPORT esattamente cosa ha scansionato, cosa ha saltato e perche'.
#
#  Con --yara-scan si indica un percorso diverso e l'ambito diventa quello.
# ================================================================

# Guardia: senza regole non c'e' niente da fare, e in batch il modulo va
# saltato con un motivo invece di produrre un report vuoto.
_guard_yara() {
    if [[ -z "${YARA_RULES:-}" ]]; then
        L "nessuna regola (--yara)" "no rules (--yara)"
        return 1
    fi
    return 0
}

module_xplat_yara() {
    section_header "YARA" "$RED"
    check_target_root || return 1

    if [[ -z "${YARA_RULES:-}" ]]; then
        warn "$(L "Nessuna regola indicata." "No rules given.")"
        info "$(L "Uso: --yara /percorso/regole.yar oppure --yara /percorso/directory/" \
                 "Usage: --yara /path/rules.yar or --yara /path/directory/")"
        return 0
    fi
    if [[ ! -e "$YARA_RULES" ]]; then
        err "$(L "Percorso regole inesistente:" "Rules path does not exist:") $YARA_RULES"
        return 1
    fi
    if ! "$PY3" -c "import yara" 2>/dev/null; then
        err "$(L "yara-python non disponibile." "yara-python unavailable.")"
        info "$(L "Installa con: pip install yara-python" "Install with: pip install yara-python")"
        info "$(L "Senza il motore non esiste alcun ripiego: una scansione YARA senza YARA non e' una scansione." \
                 "Without the engine there is no fallback: a YARA scan without YARA is not a scan.")"
        return 1
    fi

    # --- ambito ------------------------------------------------------------
    # Posizioni scrivibili senza privilegi, piu' i punti in cui il sistema
    # deposita eseguibili scaricati. Non e' l'intero volume: e' dichiarato.
    local -a TARGETS=()
    local D
    if [[ -n "${YARA_SCAN_PATH:-}" ]]; then
        [[ -d "$YARA_SCAN_PATH" ]] || { err "$(L "Percorso da scansionare inesistente:" "Scan path does not exist:") $YARA_SCAN_PATH"; return 1; }
        TARGETS=("$YARA_SCAN_PATH")
    else
        local -a CAND=()
        case "$OS_TYPE" in
            windows)
                CAND=("Windows/Temp" "ProgramData" "Users" "Windows/Tasks"
                      "Windows/System32/Tasks" "PerfLogs" "EFI"
                      "ProgramData/Microsoft/Windows Defender/Quarantine")
                ;;
            linux)
                CAND=("tmp" "var/tmp" "dev/shm" "home" "root" "opt"
                      "usr/local" "var/www" "etc/cron.d" "etc/systemd/system" "EFI")
                ;;
            macos)
                CAND=("tmp" "private/tmp" "private/var/tmp" "Users" "Library/LaunchAgents"
                      "Library/LaunchDaemons" "Library/Application Support" "EFI")
                ;;
        esac
        for D in "${CAND[@]}"; do
            local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
            [[ -n "$R" && -d "$R" ]] && TARGETS+=("$R")
        done
    fi

    if [[ ${#TARGETS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna delle posizioni previste esiste su questo volume." \
                 "None of the expected locations exists on this volume.")"
        return 0
    fi

    info "$(L "Regole:" "Rules:") ${BOLD}${YARA_RULES}"
    info "$(L "Posizioni da scansionare:" "Locations to scan:") ${BOLD}${#TARGETS[@]}${RESET}  ·  $(L "tetto per file:" "per-file cap:") ${BOLD}${YARA_MAX_MB} MB"
    info "$(L "Scansione in corso (puo' richiedere molto tempo)..." "Scanning (this can take a long time)...")"

    local TLIST; TLIST=$(mktemp); register_tmp "$TLIST"
    printf '%s\n' "${TARGETS[@]}" > "$TLIST"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local STATS; STATS=$(mktemp); register_tmp "$STATS"

    "$PY3" - "$YARA_RULES" "$TLIST" "$OUT" "$STATS" "$YARA_MAX_MB" "$YARA_MAX_FILES" << 'PYEOF'
import sys, os, json, hashlib

rules_path, tlist, out_path, stats_path = sys.argv[1:5]
max_bytes = int(sys.argv[5]) * 1024 * 1024
max_files = int(sys.argv[6])

import yara

# --- compilazione ----------------------------------------------------------
# Un file di regole con un errore di sintassi non deve far fallire tutto il
# resto: si compila file per file e si dichiara quali sono stati scartati.
sources = {}
bad_rules = []
if os.path.isdir(rules_path):
    for root, _dirs, files in os.walk(rules_path):
        for fn in sorted(files):
            if fn.lower().endswith(('.yar', '.yara')):
                sources[os.path.relpath(os.path.join(root, fn), rules_path)] = \
                    os.path.join(root, fn)
else:
    sources[os.path.basename(rules_path)] = rules_path

compiled = {}
for ns, path in sources.items():
    try:
        compiled[ns] = yara.compile(filepath=path)
    except Exception as exc:
        bad_rules.append({'file': ns, 'error': str(exc)[:300]})

stats = {
    'rule_files': len(sources),
    'rule_files_compiled': len(compiled),
    'rule_files_rejected': bad_rules,
    'scanned': 0,
    'skipped_too_big': 0,
    'skipped_unreadable': 0,
    'stopped_at_cap': False,
    'targets': [],
}

if not compiled:
    json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
    sys.exit(0)

targets = [t.strip() for t in open(tlist, encoding='utf-8') if t.strip()]
rows = []
seen_files = 0


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


for target in targets:
    per_target = 0
    for root, dirs, files in os.walk(target, followlinks=False):
        for fn in files:
            if seen_files >= max_files:
                stats['stopped_at_cap'] = True
                break
            p = os.path.join(root, fn)
            try:
                if os.path.islink(p):
                    continue
                sz = os.path.getsize(p)
            except OSError:
                stats['skipped_unreadable'] += 1
                continue
            if sz == 0:
                continue
            if sz > max_bytes:
                # Dichiarato, non omesso: un payload dentro un file da 2 GB
                # non verrebbe visto, e chi legge deve saperlo.
                stats['skipped_too_big'] += 1
                continue
            seen_files += 1
            per_target += 1
            stats['scanned'] += 1
            for ns, rl in compiled.items():
                try:
                    matches = rl.match(p, timeout=30)
                except Exception:
                    stats['skipped_unreadable'] += 1
                    continue
                for m in matches:
                    # Le stringhe che hanno fatto match sono la prova del
                    # riscontro: senza, resta solo un nome di regola.
                    frag = []
                    try:
                        for s in m.strings[:4]:
                            for inst in s.instances[:2]:
                                frag.append('%s@%d:%s' % (
                                    s.identifier, inst.offset,
                                    inst.matched_data[:60].decode('utf-8', 'replace')))
                    except Exception:
                        pass
                    try:
                        digest = sha256(p)
                    except Exception:
                        digest = ''
                    rows.append((m.rule, ns, p, str(sz), digest,
                                 ' | '.join(frag)[:400],
                                 ','.join(m.tags)[:80]))
        if stats['stopped_at_cap']:
            break
    stats['targets'].append({'path': target, 'files': per_target})
    if stats['stopped_at_cap']:
        break

rows.sort()
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
PYEOF

    local NSCAN=0 NBIG=0 NBAD=0 NRULES=0 NREJ=0 CAPPED=false
    if [[ -s "$STATS" ]]; then
        read -r NSCAN NBIG NBAD NRULES NREJ CAPPED < <("$PY3" -c '
import json, sys
s = json.load(open(sys.argv[1]))
print(s["scanned"], s["skipped_too_big"], s["skipped_unreadable"],
      s["rule_files_compiled"], len(s["rule_files_rejected"]),
      str(s["stopped_at_cap"]).lower())' "$STATS" 2>/dev/null)
    fi

    if [[ "${NRULES:-0}" -eq 0 ]]; then
        err "$(L "Nessun file di regole compilato." "No rule file compiled.")"
        [[ "${NREJ:-0}" -gt 0 ]] && "$PY3" -c '
import json, sys
for r in json.load(open(sys.argv[1]))["rule_files_rejected"]:
    print("      %s: %s" % (r["file"], r["error"]))' "$STATS" 2>/dev/null
        return 1
    fi
    [[ "${NREJ:-0}" -gt 0 ]] && warn "$(L "File di regole scartati (errore di sintassi):" "Rule files rejected (syntax error):") ${BOLD}${NREJ}"

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    ok "$(L "File scansionati:" "Files scanned:") ${BOLD}${NSCAN}"
    [[ "${NBIG:-0}" -gt 0 ]] && info "$(L "Saltati perche' oltre il tetto:" "Skipped as over the cap:") ${BOLD}${NBIG}"
    [[ "$CAPPED" == "true" ]] && warn "$(L "Raggiunto il tetto di file: la scansione e' PARZIALE." \
                                          "File cap reached: the scan is PARTIAL.")"

    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessun match." "No match.")"
        info "$(L "Nessun match sulle posizioni scansionate: non equivale a un volume pulito." \
                 "No match in the scanned locations: this is not equivalent to a clean volume.")"
    else
        warn "$(L "Match:" "Matches:") ${BOLD}${TOTAL}"
        awk -F'\t' '{printf "      %s  %s\n", $1, $3}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local TABLE
    if [[ "$TOTAL" -gt 0 ]]; then
        TABLE=$(_rows_to_table "$(head -5000 "$OUT")" \
            "$(L "Regola" "Rule")" "$(L "File regole" "Rule file")" "$(L "Percorso" "Path")" \
            "$(L "Byte" "Bytes")" "SHA-256" "$(L "Stringhe" "Strings")" "Tag")
    else
        TABLE="<div style='padding:.6rem 0;font-size:.85rem'>$(L "Nessun match." "No match.")</div>"
    fi

    # Il cartiglio dell'ambito non e' un dettaglio: e' cio' che rende il
    # risultato interpretabile.
    local SCOPE; SCOPE=$("$PY3" - "$STATS" "${LANG:-en}" << 'PYEOF' 2>/dev/null
import json, sys, html

s = json.load(open(sys.argv[1]))
it = sys.argv[2] == 'it'


def L(i, e):
    return i if it else e


out = ['<div class="card" style="margin-bottom:1rem">'
       '<div style="padding:1rem 1.5rem;font-size:.8rem;line-height:1.7">']
out.append('<b>' + L('Ambito effettivo della scansione', 'Actual scope of the scan') + '</b><br>')
out.append(L("Non e' stato scansionato l'intero volume. Un \"nessun match\" vale solo per le "
             "posizioni elencate qui sotto.",
             'The whole volume was not scanned. "No match" only holds for the locations '
             'listed below.') + '<br><br>')
out.append('<table><tr><th>' + L('Posizione', 'Location') + '</th><th>'
           + L('File scansionati', 'Files scanned') + '</th></tr>')
for t in s['targets']:
    out.append('<tr><td class="mono">%s</td><td class="mono">%d</td></tr>'
               % (html.escape(t['path']), t['files']))
out.append('</table><br>')
nums = (s['skipped_too_big'], s['skipped_unreadable'], s['rule_files_compiled'])
out.append(L("File saltati perche' oltre il tetto per file: <b>%d</b>. Illeggibili o in "
             "errore: <b>%d</b>. File di regole compilati: <b>%d</b>." % nums,
             'Files skipped as over the per-file cap: <b>%d</b>. Unreadable or errored: '
             '<b>%d</b>. Rule files compiled: <b>%d</b>.' % nums))
if s['stopped_at_cap']:
    out.append('<br><b>' + L("La scansione si e' fermata al tetto massimo di file: e' PARZIALE.",
                             'The scan stopped at the maximum file cap: it is PARTIAL.') + '</b>')
for r in s['rule_files_rejected']:
    out.append('<br>' + L('Regole scartate: ', 'Rules rejected: ')
               + html.escape(r['file']) + ' — ' + html.escape(r['error']))
out.append('</div></div>')
print(''.join(out))
PYEOF
    )

    local SB
    SB="$(stat_box "Match" "$TOTAL" "$([[ "$TOTAL" -gt 0 ]] && echo warn || echo ok)")"
    SB+="$(stat_box "$(L "File scansionati" "Files scanned")" "${NSCAN:-0}" "info")"
    SB+="$(stat_box "$(L "Saltati" "Skipped")" "$(( ${NBIG:-0} + ${NBAD:-0} ))" "info")"
    SB+="$(stat_box "$(L "File regole" "Rule files")" "${NRULES:-0}" "info")"
    finish_report "yara" "YARA" "YAR" "$(basename "$YARA_RULES")" "$SB" \
        "${SCOPE}<div class='cards'>$(generic_card_html "$(L "Riscontri" "Matches")" "$YARA_RULES" "$TOTAL" "$TABLE" "⚑")</div>"
}

# ================================================================
#  LIBRERIA PYTHON CONDIVISA — LevelDB / Snappy
#
#  Le app Electron (ChatGPT Desktop, Slack, Discord, Teams) memorizzano i
#  messaggi in LevelDB: SSTable .ldb e write-ahead log .log, con i data-block
#  compressi in Snappy e i singoli valori IndexedDB compressi una seconda
#  volta da Blink.
#
#  Questo codice nasceva dentro il modulo 39 (AI Chat). Serve identico al
#  modulo 48 (chat desktop), quindi vive qui: duplicarlo avrebbe significato
#  correggere ogni bug in due punti.
#
#  Uso da un modulo:
#      OUT=$(run_py_with_lib pylib_leveldb "$ARG" << 'PYEOF'
#      import sys
#      for _, role, text in carve_leveldb(sys.argv[1]):
#          ...
#      PYEOF
#      )
# ================================================================

# Emette il sorgente della libreria sullo stdout.
pylib_leveldb() {
    cat << 'FIUTO_PYLIB_EOF'
import re

# Lunghezza massima del testo estratto per messaggio. Il modulo chiamante puo'
# ridefinirla: la sua assegnazione viene dopo questa e quindi prevale.
MAX_TXT = 6000

# Carving di stringhe da file LevelDB (.ldb/.log) usati da ChatGPT Desktop
# (IndexedDB / Local Storage). I file SSTable non sono parsabili senza la
# libreria leveldb, ma i messaggi delle conversazioni sono memorizzati come
# stringhe in chiaro UTF-8/UTF-16. Estraiamo le run leggibili e teniamo SOLO
# quelle che sembrano linguaggio naturale (prosa), scartando i frammenti
# JSON/serializzati di configurazione (feature flag, chiavi JWK, metadati) che
# popolano soprattutto il Local Storage e non hanno valore investigativo.
MIN_LDB = 6

# Marcatori di dato serializzato: "key":  ,"  :{  :[  }] ,  ecc.
_JSON_FRAG = re.compile(r'["\}\]]\s*[:,]|[:,]\s*["\{\[]')
_WORD      = re.compile(r'[A-Za-zÀ-ÿ]{2,}')
# Prefisso V8/IndexedDB: chiave nota + " + 1 byte-tag (rumore di serializzazione)
_VPFX = re.compile(
    r'^(?:text|parts|content|content_type|message|title|name|author|value|role)".')
# Byte di controllo C0/C1 (tranne tab/CR/LF): tag e varint di lunghezza V8 che
# capitano stampabili e sporcano il testo (es. \x02 davanti a un messaggio).
_CTRL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')
# Segnali di risposta AI nel contenuto (markdown/struttura): grassetto, heading,
# liste, code fence, tabelle, link/citazioni. Più affidabile dell'encoding per
# distinguere assistente da utente (un prompt utente con € o emoji è comunque UTF-16).
_AI_SIG = re.compile(
    r'\*\*.+?\*\*|^#{1,6}\s|^\s*[-*]\s|^\s*\d+\.\s|```|\|.+\||\]\(https?://',
    re.MULTILINE)
# Marcatori dei record di metadati IndexedDB (id, utenti, flag) → mai nella prosa
_META = re.compile(
    r'accountUserId|authUserId|isArchived|client-created-root|conversationId|'
    r'currentNode|asyncStatus|updateTime[A-Z]|messagesa[A-Z]|gizmoId|'
    r'id"\$?[0-9a-f]{8}-[0-9a-f]{4}-')

def _ldb_interesting(s):
    if len(s) < 12:
        return False
    # Una frase reale contiene spazi → scarta blob base64/hex/cert/token
    if ' ' not in s:
        return False
    # Scarta i record di metadati IndexedDB (id conversazione, userId, flag…)
    if _META.search(s):
        return False
    # Scarta frammenti JSON / config / JWK / feature-flag
    if _JSON_FRAG.search(s):
        return False
    if s.count('"') >= 2 and (':' in s or ',' in s):
        return False
    # Richiede almeno 3 parole "vere" → frase, non token isolato
    if len(_WORD.findall(s)) < 3:
        return False
    # Prevalenza di caratteri alfabetici → testo, non blob/codice
    alpha = sum(c.isalpha() for c in s)
    if alpha / max(len(s), 1) < 0.55:
        return False
    return True

def _carve_bytes(data):
    """Estrae le run di testo stampabile da un buffer, restituendole in ordine
    di posizione. Nei valori V8/IndexedDB i prompt utente sono spesso one-byte
    (Latin-1) e le risposte AI two-byte (UTF-16, per markdown/caratteri speciali):
    mescolandole per offset la conversazione si legge in ordine cronologico."""
    # Caratteri base di una run: stampabili ASCII + supplemento Latin-1
    # (accenti à,è,é,ù…) + tab/CR/LF (i messaggi multi-riga restano interi).
    # One-byte = Latin-1, two-byte = UTF-16LE (high byte 0x00).
    def _dec(s, e, enc):
        return data[s:e].decode('latin-1' if enc == 'l' else 'utf-16-le', errors='replace')
    runs = []  # [start, end, enc, text, is_meta]
    for m in re.finditer(rb'[\x09\x0a\x0d\x20-\x7e\xa0-\xff]{%d,}' % MIN_LDB, data):
        t = _dec(m.start(), m.end(), 'l')
        runs.append([m.start(), m.end(), 'l', t, bool(_META.search(t))])
    for m in re.finditer(rb'(?:[\x09\x0a\x0d\x20-\x7e\xa0-\xff]\x00){%d,}' % MIN_LDB, data):
        t = _dec(m.start(), m.end(), 'u')
        runs.append([m.start(), m.end(), 'u', t, bool(_META.search(t))])
    runs.sort(key=lambda x: x[0])
    # Fusione di run di CONTENUTO adiacenti, stesso encoding, gap ≤12 byte: è
    # quasi sempre un simbolo Unicode (→ • ✓ —, emoji) o un a-capo che spezza il
    # testo; ri-decodificando lo span viene assorbito. NON si fonde mai attraverso
    # un record di metadati (id/userId/flag), per non trascinare via il testo utente.
    merged = []
    for r in runs:
        if merged:
            last = merged[-1]
            gap = r[0] - last[1]
            if (not r[4] and not last[4] and r[2] == last[2] and 0 <= gap <= 12
                    and (r[2] == 'l' or (gap % 2 == 0 and last[0] % 2 == r[0] % 2))):
                last[1] = r[1]
                last[3] = _dec(last[0], last[1], last[2])
                continue
        merged.append(r)
    # Ritorna (testo, encoding): one-byte 'l' ≈ prompt utente (testo semplice),
    # two-byte 'u' ≈ risposta AI (markdown/simboli/emoji → UTF-16). Euristica.
    return [(r[3], r[2]) for r in merged]

# ── Decompressione Snappy (opzionale) ──────────────────────────────
# I blocchi dati delle SSTable LevelDB sono spesso compressi con Snappy
# (raw block format). Senza decomprimerli, gran parte del testo delle
# conversazioni resta non carvabile. Proviamo python-snappy, poi cramjam.
_SNAPPY = None
try:
    import snappy as _sn
    _SNAPPY = ('snappy', _sn)
except Exception:
    try:
        import cramjam as _cj
        _SNAPPY = ('cramjam', _cj)
    except Exception:
        _SNAPPY = None

# Decompressore Snappy "raw" in puro Python (nessuna dipendenza): preambolo
# varint con la lunghezza non compressa, poi tag literal/copy LZ77. Usato come
# fallback quando python-snappy/cramjam non sono installati nel Python in uso.
def _snappy_py(buf):
    try:
        n = len(buf); pos = 0
        shift = 0
        while pos < n:                       # preambolo: lunghezza (scartata)
            b = buf[pos]; pos += 1
            if not (b & 0x80):
                break
            shift += 7
        out = bytearray()
        while pos < n:
            tag = buf[pos]; pos += 1
            t = tag & 0x03
            if t == 0:                       # literal
                length = (tag >> 2) + 1
                if length > 60:
                    nb = length - 60
                    length = 1
                    for i in range(nb):
                        length += buf[pos + i] << (8 * i)
                    pos += nb
                out += buf[pos:pos + length]
                pos += length
            else:
                if t == 1:                   # copy, offset 1 byte
                    length = ((tag >> 2) & 0x07) + 4
                    offset = ((tag >> 5) << 8) | buf[pos]; pos += 1
                elif t == 2:                 # copy, offset 2 byte
                    length = (tag >> 2) + 1
                    offset = buf[pos] | (buf[pos + 1] << 8); pos += 2
                else:                        # copy, offset 4 byte
                    length = (tag >> 2) + 1
                    offset = (buf[pos] | (buf[pos + 1] << 8) |
                              (buf[pos + 2] << 16) | (buf[pos + 3] << 24)); pos += 4
                if offset == 0 or offset > len(out):
                    break
                start = len(out) - offset
                for i in range(length):      # copia byte-a-byte (gestisce overlap)
                    out.append(out[start + i])
        return bytes(out)
    except Exception:
        return b''

def _snappy_raw(buf):
    if _SNAPPY is not None:
        kind, mod = _SNAPPY
        try:
            if kind == 'snappy':
                return mod.uncompress(buf)
            return bytes(mod.snappy.decompress_raw(buf))
        except Exception:
            pass
    return _snappy_py(buf)

def _uvarint(buf, pos):
    result = shift = 0
    while pos < len(buf):
        b = buf[pos]; pos += 1
        result |= (b & 0x7f) << shift
        if not (b & 0x80):
            return result, pos
        shift += 7
    return result, pos

# Legge un blocco SSTable a (offset,size); il trailer è 1 byte tipo + 4 byte CRC.
def _read_block(data, off, size):
    if off < 0 or off + size > len(data):
        return b''
    raw = data[off:off + size]
    ctype = data[off + size] if off + size < len(data) else 0
    if ctype == 0:
        return raw
    if ctype == 1:                       # Snappy
        return _snappy_raw(raw)
    return b''                            # zstd/altro non gestito

# Estrae i valori (record) da un blocco SSTable (formato entry + restart array).
def _block_values(block):
    if len(block) < 4:
        return []
    num_restarts = int.from_bytes(block[-4:], 'little')
    restart_start = len(block) - 4 - num_restarts * 4
    if restart_start < 0:
        return []
    pos = 0; last_key = b''; vals = []
    while pos < restart_start and len(vals) < 100000:
        shared, pos = _uvarint(block, pos)
        nonshared, pos = _uvarint(block, pos)
        vlen, pos = _uvarint(block, pos)
        key = last_key[:shared] + block[pos:pos + nonshared]; pos += nonshared
        value = block[pos:pos + vlen]; pos += vlen
        last_key = key
        vals.append(value)
    return vals

_LDB_MAGIC = 0xdb4775248b80fb57

def _printable_ratio(b, n=4096):
    s = b[:n]
    if not s:
        return 0.0
    return sum(0x20 <= c < 0x7f or c in (9, 10, 13) for c in s) / len(s)

# Chromium/Blink comprime i singoli valori IndexedDB con Snappy, dietro un
# breve header wrapper (lunghezza variabile). Proviamo a decomprimere saltando
# 0..15 byte iniziali; se otteniamo testo plausibile e più lungo, lo usiamo.
def _decompress_value(v):
    if len(v) < 32:
        return v
    for k in range(0, 16):
        d = _snappy_raw(v[k:])
        # Soglia bassa: i valori IndexedDB V8 contengono stringhe UTF-16
        # (byte \x00 alternati) che abbassano la quota di stampabili.
        if d and len(d) > len(v) * 1.2 and _printable_ratio(d) > 0.45:
            return d
    return v

# Estrae il contenuto applicativo da una SSTable (.ldb): decomprime i data-block
# (livello LevelDB) e poi decomprime i singoli valori dei record (livello Blink).
# Best-effort: b'' se il file non è una SSTable valida.
def _sstable_data(data):
    if len(data) < 48:
        return b''
    footer = data[-48:]
    if int.from_bytes(footer[40:48], 'little') != _LDB_MAGIC:
        return b''
    try:
        pos = 0
        _mi_off, pos = _uvarint(footer, pos); _mi_sz, pos = _uvarint(footer, pos)
        ix_off, pos = _uvarint(footer, pos); ix_sz, pos = _uvarint(footer, pos)
        index_block = _read_block(data, ix_off, ix_sz)
        out = bytearray()
        for h in _block_values(index_block):
            p = 0
            off, p = _uvarint(h, p); sz, p = _uvarint(h, p)
            blk = _read_block(data, off, sz)        # livello LevelDB (Snappy/none)
            if not blk:
                continue
            for v in _block_values(blk):            # valori dei record
                out += _decompress_value(v) + b'\x00'   # livello Blink (Snappy)
                if len(out) > 96 * 1024 * 1024:
                    return bytes(out)
        return bytes(out)
    except Exception:
        return b''

# Estrae i valori scritti da un write-ahead log LevelDB (.log): blocchi fisici
# da 32 KB con record header crc(4)+len(2)+type(1); i record logici (FULL o
# FIRST/MIDDLE/LAST) sono WriteBatch (seq 8 + count 4 + voci put/delete). Per
# ogni PUT decomprimiamo il valore (livello Blink). Best-effort.
_WAL_BLOCK = 32768

def _wal_data(data):
    try:
        logical, cur, pos, n = [], bytearray(), 0, len(data)
        while pos + 7 <= n:
            off = pos % _WAL_BLOCK
            if _WAL_BLOCK - off < 7:          # trailer di blocco → blocco succ.
                pos += _WAL_BLOCK - off
                continue
            length = data[pos + 4] | (data[pos + 5] << 8)
            rtype = data[pos + 6]
            payload = data[pos + 7:pos + 7 + length]
            pos += 7 + length
            if rtype == 1:                     # FULL
                logical.append(bytes(payload))
            elif rtype == 2:                   # FIRST
                cur = bytearray(payload)
            elif rtype == 3:                   # MIDDLE
                cur += payload
            elif rtype == 4:                   # LAST
                cur += payload; logical.append(bytes(cur)); cur = bytearray()
            elif length == 0:                  # padding → blocco successivo
                pos = (pos // _WAL_BLOCK + 1) * _WAL_BLOCK
        out = bytearray()
        for rec in logical:
            p, L = 12, len(rec)                # salta seq(8)+count(4)
            while p < L:
                t = rec[p]; p += 1
                if t not in (0, 1):
                    break                      # formato inatteso
                klen, p = _uvarint(rec, p); p += klen      # salta key
                if t == 1:                     # kTypeValue
                    vlen, p = _uvarint(rec, p)
                    out += _decompress_value(rec[p:p + vlen]) + b'\x00'
                    p += vlen
                if len(out) > 96 * 1024 * 1024:
                    return bytes(out)
        return bytes(out)
    except Exception:
        return b''

def carve_leveldb(path, cap=2500):
    try:
        with open(path, 'rb') as f:
            data = f.read(64 * 1024 * 1024)
    except Exception as e:
        return [('', 'error', str(e))]
    # Decomprimiamo il contenuto e carviamo SOLO quello (pulito): per le SSTable
    # (.ldb) i data-block, per i write-ahead log (.log) i record WriteBatch. In
    # entrambi i casi i valori IndexedDB vengono poi de-comprimati (livello Blink).
    # Il carving del file grezzo darebbe frammenti spezzati → solo come fallback.
    low = path.lower()
    if low.endswith('.ldb'):
        blocks = _sstable_data(data)
    elif low.endswith('.log'):
        blocks = _wal_data(data)
    else:
        # File blob esterno IndexedDB (valore grande wrappato+Snappy da Blink):
        # un singolo valore → proviamo a decomprimerlo direttamente.
        dv = _decompress_value(data)
        blocks = dv if len(dv) > len(data) else b''
    cand = _carve_bytes(blocks) if blocks else _carve_bytes(data)
    out, seen = [], set()
    for s, enc in cand:
        s = s.strip()
        # Strip conservativo del prefisso V8: chiave IndexedDB nota + " + 1
        # byte-tag (es. text"T… → non…). Solo se resta del testo.
        s2 = _VPFX.sub('', s)
        if s2:
            s = s2.strip()
        # Rimuove i byte di controllo residui (tag/varint V8) ovunque nel testo.
        s = _CTRL.sub('', s).strip()
        # Rimuove il tag V8 di fine-oggetto '{' (kEndJSObject) in coda al messaggio.
        s = re.sub(r'\s*\{+\s*$', '', s)
        if not _ldb_interesting(s):
            continue
        key = s.lower()[:80]
        if key in seen:
            continue
        seen.add(key)
        # Ruolo dal contenuto: la presenza di markdown/struttura indica una
        # risposta AI; altrimenti è un prompt utente (più robusto dell'encoding).
        role = 'assistant' if _AI_SIG.search(s) else 'user'
        out.append(('', role, s[:MAX_TXT]))
        if len(out) >= cap:
            break
    if not out:
        out.append(('', 'info',
                    'Nessuna stringa leggibile estratta (file vuoto, binario o solo metadati).'))
    return out
FIUTO_PYLIB_EOF
}

# Esegue uno script Python letto da stdin, anteponendogli una libreria condivisa.
#
# Concatenare invece di importare tiene il programma finale identico a com'era
# quando la libreria stava dentro il modulo: stesso spazio dei nomi, nessuna
# differenza di comportamento da dimostrare.
#
#   run_py_with_lib <funzione_libreria> [argomenti...] << 'PYEOF' ... PYEOF
run_py_with_lib() {
    local LIBF="$1"; shift
    local TMP; TMP=$(mktemp) || return 1
    { "$LIBF"; printf '\n'; cat; } > "$TMP"
    "$PY3" "$TMP" "$@"
    local RC=$?
    rm -f "$TMP"
    return $RC
}

# ================================================================
#  LIBRERIA PYTHON CONDIVISA — recupero da spazio libero SQLite
#
#  query_sqlite() legge i record vivi (e copia il -wal per vedere le
#  transazioni non consolidate), ma un record CANCELLATO non sparisce dal
#  file: la sua pagina finisce nella freelist, o lo spazio che occupava
#  diventa un freeblock dentro la pagina. Il contenuto resta li' finche' non
#  viene sovrascritto.
#
#  Su un artefatto come la cronologia del browser, TCC.db o chat.db questo e'
#  spesso l'unico modo di vedere cio' che l'utente ha cancellato.
#
#  Livello di supporto: si estraggono le STRINGHE dallo spazio non allocato,
#  non i record ricostruiti. Ricostruire un record cancellato richiede lo
#  schema e l'interpretazione dei serial type: e' possibile ma fragile, e un
#  record ricomposto male in un report forense e' peggio di nessun record.
#
#  Uso:
#      run_py_with_lib pylib_sqlite_recover "$DB" << 'PYEOF'
#      import sys
#      for kind, text in sqlite_carve(sys.argv[1]):
#          ...
#      PYEOF
# ================================================================

pylib_sqlite_recover() {
    cat << 'FIUTO_SQLPY_EOF'
import re
import struct

# Lunghezza minima di una stringa perche' valga la pena riportarla.
SQLITE_MIN_STR = 5

_PRINTABLE = re.compile(rb'[\x20-\x7e\xc0-\xf4][\x20-\x7e\x80-\xbf]{%d,}' % (SQLITE_MIN_STR - 1))


def _pages(data):
    """Genera (numero_pagina, byte) per ogni pagina del database."""
    if len(data) < 100 or data[:15] != b'SQLite format 3':
        return
    page_size = struct.unpack_from('>H', data, 16)[0]
    # 1 significa 65536 (il campo e' a 16 bit).
    if page_size == 1:
        page_size = 65536
    if page_size < 512 or page_size & (page_size - 1):
        return
    total = len(data) // page_size
    for i in range(total):
        yield i + 1, data[i * page_size:(i + 1) * page_size]


def _freelist_pages(data):
    """Numeri delle pagine nella freelist, seguendo la catena dei trunk."""
    out = set()
    if len(data) < 100:
        return out
    page_size = struct.unpack_from('>H', data, 16)[0]
    if page_size == 1:
        page_size = 65536
    try:
        trunk = struct.unpack_from('>I', data, 32)[0]
        count = struct.unpack_from('>I', data, 36)[0]
    except Exception:
        return out
    seen = set()
    guard = 0
    while trunk and trunk not in seen and guard < 100000:
        guard += 1
        seen.add(trunk)
        off = (trunk - 1) * page_size
        if off < 0 or off + 8 > len(data):
            break
        out.add(trunk)
        try:
            nxt, n = struct.unpack_from('>II', data, off)
        except Exception:
            break
        n = min(n, (page_size - 8) // 4)
        for i in range(n):
            try:
                leaf = struct.unpack_from('>I', data, off + 8 + 4 * i)[0]
            except Exception:
                break
            if leaf:
                out.add(leaf)
        trunk = nxt
    return out


def _unallocated(page, is_first):
    """Spazio non allocato e freeblock di una pagina b-tree.

    Layout: header, array dei puntatori alle celle, spazio libero, contenuto
    delle celle. Cio' che sta fra la fine dell'array e l'inizio del contenuto
    non e' in uso — ed e' li' che restano i record cancellati.
    """
    base = 100 if is_first else 0
    if len(page) < base + 8:
        return b''
    ptype = page[base]
    if ptype not in (0x02, 0x05, 0x0a, 0x0d):
        return b''
    hdr = 12 if ptype in (0x02, 0x05) else 8
    try:
        ncells = struct.unpack_from('>H', page, base + 3)[0]
        content = struct.unpack_from('>H', page, base + 5)[0]
        freeblk = struct.unpack_from('>H', page, base + 1)[0]
    except Exception:
        return b''
    if content == 0:
        content = 65536
    start = base + hdr + 2 * ncells
    out = bytearray()
    if 0 < start < content <= len(page):
        out += page[start:content]
    # Catena dei freeblock: spazio liberato dentro l'area del contenuto.
    guard = 0
    while freeblk and freeblk + 4 <= len(page) and guard < 10000:
        guard += 1
        try:
            nxt, size = struct.unpack_from('>HH', page, freeblk)
        except Exception:
            break
        if size and freeblk + size <= len(page):
            out += page[freeblk:freeblk + size]
        if nxt <= freeblk:
            break
        freeblk = nxt
    return bytes(out)


def sqlite_carve(path, cap=20000):
    """Ritorna [(origine, stringa)] recuperate dallo spazio non allocato.

    origine: 'freelist' per le pagine liberate per intero, 'unallocated' per
    lo spazio non usato dentro pagine ancora in uso.
    """
    try:
        with open(path, 'rb') as fh:
            data = fh.read(512 * 1024 * 1024)
    except Exception:
        return []
    if data[:15] != b'SQLite format 3':
        return []

    free = _freelist_pages(data)
    out = []
    seen = set()

    def harvest(blob, origin):
        for m in _PRINTABLE.finditer(blob):
            # errors='ignore' e non 'strict': il pattern ammette i byte di
            # continuazione UTF-8 (\x80-\xbf) per non spezzare i caratteri
            # accentati, ma in una pagina SQLite l'header del record segue il
            # payload, quindi il match finisce quasi sempre con un byte di
            # continuazione isolato. Con 'strict' l'intera stringa veniva
            # scartata in silenzio — e sono proprio quelle interessanti.
            s = m.group(0).decode('utf-8', 'ignore')
            # Via i residui di controllo dell'header del record.
            s = ''.join(ch for ch in s if ch >= ' ' and ch != '\x7f').strip()
            if len(s) < SQLITE_MIN_STR:
                continue
            # Scarta le sequenze senza alcuna lettera: sono quasi sempre
            # residui binari che capitano nell'intervallo stampabile.
            if not re.search(r'[A-Za-z]{3}', s):
                continue
            k = s.lower()
            if k in seen:
                continue
            seen.add(k)
            out.append((origin, s[:500]))
            if len(out) >= cap:
                return True
        return False

    for num, page in _pages(data):
        if len(out) >= cap:
            break
        if num in free:
            if harvest(page, 'freelist'):
                break
        else:
            blob = _unallocated(page, num == 1)
            if blob and harvest(blob, 'unallocated'):
                break
    return out
FIUTO_SQLPY_EOF
}

# ================================================================
#  LIBRERIA PYTHON CONDIVISA — estrazione eventi dai report HTML
#
#  L'HTML gia' prodotto dai moduli e' l'unico punto in cui il formato e'
#  omogeneo per tutti e 91: ogni modulo emette tabelle e blocchi <pre>, non
#  esiste una struttura dati comune a monte. Estrarre gli eventi da li' e'
#  quello che permette all'export JSONL e all'executive summary di coprire
#  l'intero toolkit senza toccare i moduli uno per uno.
#
#  Il codice nasceva dentro export_report_jsonl. Serve identico all'executive
#  summary, che deve poter correlare gli eventi anche quando --jsonl non e'
#  stato chiesto: vive qui perche' due copie divergerebbero al primo bugfix.
#
#  Uso da un modulo:
#      run_py_with_lib pylib_timeline "$HTML" << 'PYEOF'
#      import sys
#      for ev in extract_events(sys.argv[1]):
#          print(ev['datetime'], ev['message'])
#      PYEOF
# ================================================================

# Emette il sorgente della libreria sullo stdout.
pylib_timeline() {
    cat << 'FIUTO_PYLIB_EOF'
import re as _re
import html as _H
import datetime as _dt

_MONTHS = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
           'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
_TS_ISO = _re.compile(r'\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}')
_TS_SYS = _re.compile(r'\b(' + '|'.join(_MONTHS) + r')\s+(\d{1,2})\s+(\d{2}:\d{2}:\d{2})')
_TD = _re.compile(r'<td[^>]*>(.*?)</td>', _re.DOTALL | _re.I)
_TR = _re.compile(r'<tr[^>]*>(.*?)</tr>', _re.DOTALL | _re.I)
_PRE = _re.compile(r'<pre[^>]*>(.*?)</pre>', _re.DOTALL | _re.I)
_TAG = _re.compile(r'<[^>]+>')


def strip_tags(s):
    return _TAG.sub('', _H.unescape(s)).strip()


_SENS = _re.compile(r'<span class="line sensitive">(.*?)</span>', _re.DOTALL)


def read_report_data(path):
    """Solo il DATO del report: righe di tabella e righe dei blocchi <pre>.

    Non il testo completo della pagina. I report contengono cartigli
    esplicativi scritti da FIUTO ("...tipico del malware che...") e un motore
    di detection che cercasse parole nell'intera pagina scatterebbe sulla
    propria prosa. Qui la prosa non entra proprio.

    L'unita' e' la RIGA, non la cella: un conteggio di celle non significa
    niente per chi legge ("3 riscontri" su una riga sola), e una regola che
    deve vedere data e percorso insieme non funzionerebbe.
    """
    try:
        content = open(path, encoding='utf-8', errors='replace').read()
    except Exception:
        return []
    items = []
    for tr in _TR.finditer(content):
        cells = [strip_tags(m.group(1)) for m in _TD.finditer(tr.group(1))]
        row = ' | '.join(c for c in cells if c)
        if row:
            items.append(row)
    for pm in _PRE.finditer(content):
        block = _H.unescape(_TAG.sub('', pm.group(1)))
        for line in block.split('\n'):
            line = _re.sub(r'^\s*\d+\s+', '', line).strip()
            if line:
                items.append(line)
    return items


def read_sensitive_lines(path):
    """Righe che il modulo stesso ha marcato come degne di attenzione.

    render_pre_block applica la classe 'sensitive' alle righe che contengono
    le parole chiave passate dal modulo: e' un giudizio gia' espresso da chi
    conosce l'artefatto, molto piu' affidabile di una regola generica scritta
    a valle.
    """
    try:
        content = open(path, encoding='utf-8', errors='replace').read()
    except Exception:
        return []
    out = []
    for m in _SENS.finditer(content):
        s = _re.sub(r'^\s*\d+\s+', '', strip_tags(m.group(1))).strip()
        if s:
            out.append(s)
    return out


def extract_events(path):
    """Eventi datati di un report HTML, deduplicati e ordinati.

    Ritorna [{'datetime','message','year_inferred'}]. Le date NON sono
    riportate a un fuso comune: vedi src/lib/16-time.sh per il perche'.
    """
    try:
        content = open(path, encoding='utf-8', errors='replace').read()
    except Exception:
        return []

    # I log syslog non portano l'anno: si usa quello corrente (l'analisi e'
    # contestuale all'acquisizione). Approssimazione esplicita, non silenziosa.
    year = str(_dt.date.today().year)
    out = []
    seen = set()

    def emit(dt, message, assumed):
        message = ' '.join(message.split())[:2000]
        if not message:
            return
        key = (dt, message[:120])
        if key in seen:
            return
        seen.add(key)
        out.append({'datetime': dt, 'message': message, 'year_inferred': assumed})

    # 1) Righe di tabella in cui una cella contiene un timestamp
    for tr in _TR.finditer(content):
        cells = [strip_tags(m.group(1)) for m in _TD.finditer(tr.group(1))]
        if not cells:
            continue
        dt = None
        for c in cells:
            m = _TS_ISO.search(c)
            if m:
                dt = m.group(0).replace('T', ' ')[:19].replace(' ', 'T')
                break
        if not dt:
            continue
        msg = ' | '.join(c for c in cells
                         if c and not _TS_ISO.fullmatch(c.strip())
                         and not _re.fullmatch(r'\d+', c.strip()))
        emit(dt, msg, False)

    # 2) Blocchi <pre> (log, history, config): una riga per evento
    for pm in _PRE.finditer(content):
        block = _H.unescape(_TAG.sub('', pm.group(1)))
        for raw in block.split('\n'):
            txt = _re.sub(r'^\s*\d+\s+', '', raw).strip()   # via il numero di riga
            if not txt:
                continue
            m = _TS_ISO.search(txt)
            if m:
                emit(m.group(0).replace('T', ' ')[:19].replace(' ', 'T'), txt, False)
                continue
            m = _TS_SYS.search(txt)
            if m:
                emit(f"{year}-{_MONTHS[m.group(1)]}-{int(m.group(2)):02d}T{m.group(3)}",
                     txt, True)

    out.sort(key=lambda r: r['datetime'])
    return out
FIUTO_PYLIB_EOF
}

# ================================================================
#  LIBRERIA PYTHON CONDIVISA — compilatore di regole Sigma
#
#  Sigma e' un linguaggio ampio e questo e' un motore parziale. La parte
#  rischiosa non e' leggere gli EVTX, e' decidere cosa una regola significa:
#  un modificatore interpretato male, o una condizione valutata al contrario,
#  produce un falso negativo che nessuno vede. Per questo il compilatore vive
#  qui e non dentro il modulo: cosi' e' esercitabile dai test con eventi
#  sintetici, senza bisogno di un .evtx.
#
#  SOTTOINSIEME SUPPORTATO — dichiarato, non implicito:
#    selezioni     mappa campo/valore, liste di valori (OR), liste di mappe (OR)
#    modificatori  contains, startswith, endswith, re, all, cased
#    condizioni    "sel", "a and b", "a or b", "a and not b", "not a",
#                  "1 of x*", "all of x*", "1 of them", "all of them"
#    valori        null = campo assente o vuoto
#
#  Tutto il resto solleva Unsupported con il motivo, e il chiamante e' tenuto
#  a ELENCARE le regole scartate: una regola non valutata non e' una regola
#  non scattata.
#
#  Uso:
#      run_py_with_lib pylib_sigma "$RULES" << 'PYEOF'
#      rules, rejected = load_sigma_rules(sys.argv[1], {'security.evtx': '/x/Security.evtx'})
#      PYEOF
# ================================================================

pylib_sigma() {
    cat << 'FIUTO_PYLIB_EOF'
import os as _os
import re as _re
import glob as _glob

import yaml as _yaml


class Unsupported(Exception):
    """La regola usa una costruzione che questo motore non implementa."""


SUPPORTED_MODS = {'contains', 'startswith', 'endswith', 're', 'all', 'cased'}

# Sigma descrive la sorgente in astratto; sul disco ci sono file con nomi
# precisi. Una logsource che non sappiamo mappare NON viene fatta girare a
# tappeto su tutti i log: la regola verrebbe valutata su campi che in quel
# canale non esistono, e il "non scattata" sarebbe privo di significato.
SERVICE_MAP = {
    'security': ['Security.evtx'],
    'system': ['System.evtx'],
    'application': ['Application.evtx'],
    'sysmon': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'powershell': ['Microsoft-Windows-PowerShell%4Operational.evtx'],
    'powershell-classic': ['Windows PowerShell.evtx'],
    'taskscheduler': ['Microsoft-Windows-TaskScheduler%4Operational.evtx'],
    'windefend': ['Microsoft-Windows-Windows Defender%4Operational.evtx'],
    'terminalservices-localsessionmanager':
        ['Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx'],
    'wmi': ['Microsoft-Windows-WMI-Activity%4Operational.evtx'],
    'bits-client': ['Microsoft-Windows-Bits-Client%4Operational.evtx'],
    'smbclient-security': ['Microsoft-Windows-SmbClient%4Security.evtx'],
    'ntlm': ['Microsoft-Windows-NTLM%4Operational.evtx'],
}
# process_creation sta sia in Security (4688) sia in Sysmon (EID 1): una regola
# scritta per l'uno vale spesso per l'altro, quindi si applica a entrambi.
CATEGORY_MAP = {
    'process_creation': ['Security.evtx', 'Microsoft-Windows-Sysmon%4Operational.evtx'],
    'network_connection': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'image_load': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'file_event': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'registry_event': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'registry_set': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'registry_add': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'process_access': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'pipe_created': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'dns_query': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'ps_script': ['Microsoft-Windows-PowerShell%4Operational.evtx'],
    'ps_module': ['Microsoft-Windows-PowerShell%4Operational.evtx'],
    'ps_classic_start': ['Windows PowerShell.evtx'],
}


def compile_field(key, raw):
    """(nome_campo, predicato) da una chiave 'Campo|modificatore'."""
    parts = key.split('|')
    field = parts[0]
    mods = [m.lower() for m in parts[1:]]
    for m in mods:
        if m not in SUPPORTED_MODS:
            raise Unsupported('modificatore |%s' % m)
    cased = 'cased' in mods
    values = raw if isinstance(raw, list) else [raw]

    def norm(v):
        return None if v is None else (str(v) if cased else str(v).lower())

    vals = [norm(v) for v in values]

    def actual_of(a):
        a = '' if a is None else str(a)
        return a if cased else a.lower()

    if 're' in mods:
        rxs = [_re.compile(str(v), 0 if cased else _re.I) for v in values]

        def pred(a):
            return any(r.search('' if a is None else str(a)) for r in rxs)
    elif 'contains' in mods:
        if 'all' in mods:
            def pred(a):
                s = actual_of(a)
                return all(v is not None and v in s for v in vals)
        else:
            def pred(a):
                s = actual_of(a)
                return any(v is not None and v in s for v in vals)
    elif 'startswith' in mods:
        def pred(a):
            s = actual_of(a)
            return any(v is not None and s.startswith(v) for v in vals)
    elif 'endswith' in mods:
        def pred(a):
            s = actual_of(a)
            return any(v is not None and s.endswith(v) for v in vals)
    else:
        def pred(a):
            s = actual_of(a)
            # null in Sigma significa "campo assente o vuoto".
            return any((v is None and not s) or (v is not None and s == v) for v in vals)

    return field, pred


def compile_selection(sel):
    """funzione(evento)->bool. Una lista di mappe e' un OR fra le mappe."""
    if isinstance(sel, list):
        subs = [compile_selection(s) for s in sel]
        return lambda ev: any(s(ev) for s in subs)
    if not isinstance(sel, dict):
        raise Unsupported('selezione non e\' una mappa')
    checks = [compile_field(k, v) for k, v in sel.items()]

    def run(ev):
        return all(pred(ev.get(field, '')) for field, pred in checks)
    return run


_COND_RX = _re.compile(
    r'^(?:(1|all)\s+of\s+(\S+)|not\s+(\S+)|(\S+))'
    r'(?:\s+(and|or)\s+(?:(not)\s+)?(?:(1|all)\s+of\s+)?(\S+))?$', _re.I)


def compile_condition(cond, sels):
    """Sottoinsieme documentato della grammatica delle condizioni Sigma."""
    if not isinstance(cond, str):
        raise Unsupported('condizione non testuale')
    c = ' '.join(cond.split())
    if '(' in c or '|' in c:
        raise Unsupported('condizione con parentesi o aggregazione')

    def resolve(name):
        n = name.strip()
        if n.lower() == 'them':
            got = list(sels.values())
        elif n.endswith('*'):
            got = [v for k, v in sels.items() if k.startswith(n[:-1])]
        else:
            got = [sels[n]] if n in sels else []
        if not got:
            raise Unsupported('selezione %s inesistente' % n)
        return got

    m = _COND_RX.match(c)
    if not m:
        raise Unsupported('condizione non riconosciuta: %s' % c[:60])
    quant, quant_name, not_name, plain, op, op_not, rquant, rhs = m.groups()

    def side(quantifier, name, negated):
        group = resolve(name)
        if (quantifier or '').lower() == 'all':
            base = lambda ev: all(f(ev) for f in group)   # noqa: E731
        else:
            base = lambda ev: any(f(ev) for f in group)   # noqa: E731
        return (lambda ev: not base(ev)) if negated else base

    if quant:
        left = side(quant, quant_name, False)
    elif not_name:
        left = side(None, not_name, True)
    else:
        left = side(None, plain, False)

    if not op:
        return left

    right = side(rquant, rhs, bool(op_not))
    if op.lower() == 'and':
        return lambda ev: left(ev) and right(ev)
    return lambda ev: left(ev) or right(ev)


def _rule_files(path):
    if _os.path.isdir(path):
        out = []
        for ext in ('yml', 'yaml'):
            out += _glob.glob(_os.path.join(path, '**', '*.' + ext), recursive=True)
        return sorted(out)
    return [path]


def load_sigma_rules(rules_path, available):
    """(rules, rejected).

    `available` mappa nome-file-evtx-minuscolo -> percorso reale sul volume.
    Ogni regola scartata finisce in `rejected` con il motivo: e' quella lista
    che impedisce di leggere "nessun match" come "nessuna minaccia".
    """
    rules, rejected = [], []
    for rf in _rule_files(rules_path):
        try:
            with open(rf, encoding='utf-8', errors='replace') as fh:
                docs = [d for d in _yaml.safe_load_all(fh) if isinstance(d, dict)]
        except Exception as exc:
            rejected.append({'rule': _os.path.basename(rf),
                             'reason': 'YAML illeggibile: %s' % str(exc)[:120]})
            continue
        for doc in docs:
            name = doc.get('title') or _os.path.basename(rf)
            det = doc.get('detection')
            if not isinstance(det, dict) or 'condition' not in det:
                rejected.append({'rule': name, 'reason': 'detection assente o senza condition'})
                continue
            ls = doc.get('logsource') or {}
            svc = str(ls.get('service', '')).lower()
            cat = str(ls.get('category', '')).lower()
            targets = SERVICE_MAP.get(svc) or CATEGORY_MAP.get(cat) or []
            if not targets:
                rejected.append({'rule': name, 'reason': 'logsource non mappata (%s/%s)'
                                 % (ls.get('product', '-'), svc or cat or '-')})
                continue
            present = [available[t.lower()] for t in targets if t.lower() in available]
            if not present:
                rejected.append({'rule': name, 'reason': 'canale non presente sul volume (%s)'
                                 % ', '.join(targets)})
                continue
            try:
                sels = {k: compile_selection(v) for k, v in det.items() if k != 'condition'}
                pred = compile_condition(det['condition'], sels)
            except Unsupported as exc:
                rejected.append({'rule': name, 'reason': str(exc)})
                continue
            except Exception as exc:
                rejected.append({'rule': name,
                                 'reason': 'compilazione fallita: %s' % str(exc)[:120]})
                continue
            rules.append({
                'title': name,
                'id': doc.get('id', ''),
                'level': str(doc.get('level', 'medium')),
                'tags': [t for t in (doc.get('tags') or []) if isinstance(t, str)],
                'files': present,
                'pred': pred,
            })
    return rules, rejected
FIUTO_PYLIB_EOF
}

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

# ================================================================
#  MAIN
# ================================================================
main() {
    SCAN_DATE=$(date "+%d/%m/%Y %H:%M:%S")

    # La directory degli hive ricostruiti viene creata da subshell, che non
    # possono registrarla per il cleanup: la si registra qui, nel processo
    # padre. rm -rf ignora i percorsi inesistenti, quindi registrarla anche
    # quando non verra' mai creata e' innocuo.
    register_tmp "${TMPDIR:-/tmp}/fiuto_hives_$$"
    register_tmp "${TMPDIR:-/tmp}/fiuto_custody_$$"
    register_tmp "${TMPDIR:-/tmp}/fiuto_time_$$"

    # Contesto della sessione, congelato all'avvio per il manifesto.
    CUSTODY_START_UTC=$(date -u "+%Y-%m-%dT%H:%M:%SZ")
    CUSTODY_CMDLINE="$0 $*"
    CUSTODY_OPERATOR="$(id -un 2>/dev/null || echo '?')"
    CUSTODY_HOST="$(hostname 2>/dev/null || echo '?')"

    # Always ask for language at the very start (unless --help is passed)
    if [[ "${1:-}" != "-h" && "${1:-}" != "--help" ]]; then
        select_language
    fi

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
                if [[ "$LANG" == "it" ]]; then
                    echo -e "${CYAN}${BOLD}fiuto.sh${RESET} — Toolkit DFIR per analisi di disco Windows offline"
                    echo ""
                    echo -e "  ${BOLD}Uso:${RESET}"
                    echo -e "    ./fiuto.sh                            # menu interattivo"
                    echo -e "    ./fiuto.sh /mnt/windows               # imposta root e apre il menu"
                    echo -e "    ./fiuto.sh /mnt/windows --all         # esegui tutti i moduli"
                    echo -e "    ./fiuto.sh /mnt/windows --module 3    # esegui modulo specifico"
                    echo -e "    ./fiuto.sh /mnt/windows --modules 1,3,5-8  # esegui selezione"
                    echo -e "    ./fiuto.sh /mnt/windows --all --ioc /path/to/ioc.txt  # con IoC"
                    echo -e "    ./fiuto.sh /mnt/windows --all --no-log-replay  # non applicare i .LOG1/.LOG2"
                    echo -e "    ./fiuto.sh /mnt/windows --all --jsonl  # esporta anche JSONL per Timesketch"
                    echo -e "    ./fiuto.sh /mnt/windows --all --no-hash  # manifesto senza SHA256 (piu' veloce)"
                    echo -e "    ./fiuto.sh /mnt/windows --all --since 2026-03-01 --until 2026-03-08  # solo la finestra"
                    echo -e "    ./fiuto.sh /mnt/windows --all --since -7d     # ultimi 7 giorni"
                    echo -e "    ./fiuto.sh /mnt/disk --all --yara /regole/     # applica regole YARA"
                    echo -e "    ./fiuto.sh /mnt/disk --all --yara r.yar --yara-scan /mnt/disk/Users  # ambito esplicito"
                    echo -e "    ./fiuto.sh /mnt/windows --all --sigma /sigma/rules/  # regole Sigma sugli EVTX"
                    echo -e "    ./fiuto.sh /mnt/disk --all --jobs 4            # 4 moduli in parallelo"
                    echo -e "    ./fiuto.sh /mnt/disk --all --redact            # copia condivisibile senza segreti"
                    echo -e "    ./fiuto.sh --image disco.E01 --list-partitions # elenca le partizioni (senza root)"
                    echo -e "    sudo ./fiuto.sh --image disco.E01 --partition 2 --all"
                    echo -e "    sudo ./fiuto.sh --image disco.raw --unlock chiave.txt --all  # BitLocker/LUKS"
                    echo -e "    sudo ./fiuto.sh / --all --report-dir /tmp/analisi  # sistema vivo, report fuori dal volume"
                    echo ""
                    echo -e "  ${DIM}--yara non scansiona l'intero volume: si limita alle posizioni"
                    echo -e "    scrivibili senza privilegi e le ELENCA nel report. Usa --yara-scan"
                    echo -e "    per indicare un ambito diverso.${RESET}"
                    echo ""
                    echo -e "  ${DIM}--since/--until confrontano le date come compaiono nell'artefatto,"
                    echo -e "    senza riportarle a un fuso comune: gli artefatti dello stesso volume"
                    echo -e "    mescolano UTC e ora locale. Per finestre di poche ore, allargale"
                    echo -e "    dell'offset del volume (dichiarato all'avvio).${RESET}"
                    echo ""
                    echo -e "  ${DIM}Di default i transaction log del registro (.LOG1/.LOG2) vengono"
                    echo -e "    riapplicati su una copia temporanea: senza questo passaggio le"
                    echo -e "    scritture piu' recenti dell'hive non sono visibili.${RESET}"
                    echo ""
                    echo -e "  ${BOLD}Moduli disponibili per sistema operativo:${RESET}"
                else
                    echo -e "${CYAN}${BOLD}fiuto.sh${RESET} — DFIR Toolkit for offline Windows disk analysis"
                    echo ""
                    echo -e "  ${BOLD}Usage:${RESET}"
                    echo -e "    ./fiuto.sh                            # interactive menu"
                    echo -e "    ./fiuto.sh /mnt/windows               # set root and open menu"
                    echo -e "    ./fiuto.sh /mnt/windows --all         # run all modules"
                    echo -e "    ./fiuto.sh /mnt/windows --module 3    # run specific module"
                    echo -e "    ./fiuto.sh /mnt/windows --modules 1,3,5-8  # run selection"
                    echo -e "    ./fiuto.sh /mnt/windows --all --ioc /path/to/ioc.txt  # with IoCs"
                    echo -e "    ./fiuto.sh /mnt/windows --all --no-log-replay  # skip .LOG1/.LOG2 replay"
                    echo -e "    ./fiuto.sh /mnt/windows --all --jsonl  # also export JSONL for Timesketch"
                    echo -e "    ./fiuto.sh /mnt/windows --all --no-hash  # manifest without SHA256 (faster)"
                    echo -e "    ./fiuto.sh /mnt/windows --all --since 2026-03-01 --until 2026-03-08  # window only"
                    echo -e "    ./fiuto.sh /mnt/windows --all --since -7d     # last 7 days"
                    echo -e "    ./fiuto.sh /mnt/disk --all --yara /rules/     # apply YARA rules"
                    echo -e "    ./fiuto.sh /mnt/disk --all --yara r.yar --yara-scan /mnt/disk/Users  # explicit scope"
                    echo -e "    ./fiuto.sh /mnt/windows --all --sigma /sigma/rules/  # Sigma rules over EVTX"
                    echo -e "    ./fiuto.sh /mnt/disk --all --jobs 4            # 4 modules in parallel"
                    echo -e "    ./fiuto.sh /mnt/disk --all --redact            # shareable copy without secrets"
                    echo -e "    ./fiuto.sh --image disk.E01 --list-partitions  # list partitions (no root needed)"
                    echo -e "    sudo ./fiuto.sh --image disk.E01 --partition 2 --all"
                    echo -e "    sudo ./fiuto.sh --image disk.raw --unlock key.txt --all      # BitLocker/LUKS"
                    echo -e "    sudo ./fiuto.sh / --all --report-dir /tmp/case  # live system, reports outside the volume"
                    echo ""
                    echo -e "  ${DIM}--yara does not scan the whole volume: it covers the locations"
                    echo -e "    writable without privileges and LISTS them in the report. Use"
                    echo -e "    --yara-scan to point it somewhere else.${RESET}"
                    echo ""
                    echo -e "  ${DIM}--since/--until compare dates as they appear in the artefact, without"
                    echo -e "    normalising them to a common zone: artefacts on the same volume mix"
                    echo -e "    UTC and local time. For windows of a few hours, widen them by the"
                    echo -e "    volume offset (declared at startup).${RESET}"
                    echo ""
                    echo -e "  ${DIM}By default registry transaction logs (.LOG1/.LOG2) are replayed"
                    echo -e "    onto a temporary copy: without this step the most recent hive"
                    echo -e "    writes are not visible.${RESET}"
                    echo ""
                    echo -e "  ${BOLD}Available modules, by operating system:${RESET}"
                fi
                print_module_list
                echo ""
                echo -e "  ${DIM}$(L "I numeri di --module e --modules si riferiscono al sistema rilevato sul volume." \
                                     "The numbers used by --module and --modules refer to the OS detected on the volume.")${RESET}"
                echo ""
                exit 0
                ;;
            --all)       ARG_ALL=true ;;
            --module)    ARG_MODULE="$2"; shift ;;
            --modules)   ARG_MODULES="$2"; shift ;;
            --ioc)       ARG_IOC="$2"; shift ;;
            --no-log-replay) HIVE_REPLAY=false ;;
            --jsonl)     EXPORT_JSONL=true ;;
            --no-custody)  CUSTODY=false ;;
            --no-hash)     CUSTODY_HASH=false ;;
            --hash-limit)  CUSTODY_HASH_LIMIT_MB="${2:-1024}"; shift ;;
            --yara)        YARA_RULES="${2:-}"; shift ;;
            --yara-scan)   YARA_SCAN_PATH="${2:-}"; shift ;;
            --yara-max-mb) YARA_MAX_MB="${2:-64}"; shift ;;
            --sigma)       SIGMA_RULES="${2:-}"; shift ;;
            --image)       IMAGE_PATH="${2:-}"; shift ;;
            --partition)   IMAGE_PARTITION="${2:-}"; shift ;;
            --unlock)      IMAGE_UNLOCK="${2:-}"; shift ;;
            --list-partitions) IMAGE_LIST_ONLY=true ;;
            --report-dir)  REPORT_DIR_FIXED="${2:-}"; shift ;;
            --redact)      REDACT=true ;;
            --defang)      REDACT=true; REDACT_DEFANG=true ;;
            --jobs)
                if [[ "${2:-}" =~ ^[1-9][0-9]*$ ]]; then
                    JOBS="$2"
                else
                    err "$(L "--jobs richiede un intero positivo:" "--jobs requires a positive integer:") '${2:-}'"
                    exit 1
                fi
                shift ;;
            --since|--until)
                # Un limite scritto male non deve passare in silenzio: filtrerebbe
                # tutto o niente, e in entrambi i casi il report sarebbe falso.
                local _BOUND _KIND
                [[ "$1" == "--since" ]] && _KIND=start || _KIND=end
                if ! _BOUND=$(parse_time_bound "${2:-}" "$_KIND"); then
                    err "$(L "Data non valida per" "Invalid date for") $1: '${2:-}'"
                    info "$(L "Formati ammessi: 2026-03-01 · '2026-03-01 14:30' · 2026-03-01T14:30:00 · -7d · -36h · -90m" \
                             "Accepted formats: 2026-03-01 · '2026-03-01 14:30' · 2026-03-01T14:30:00 · -7d · -36h · -90m")"
                    exit 1
                fi
                [[ "$1" == "--since" ]] && TIME_SINCE="$_BOUND" || TIME_UNTIL="$_BOUND"
                shift ;;
            --format)    [[ "${2:-}" == "jsonl" ]] && EXPORT_JSONL=true; shift ;;
            -*)          local UNKNOWN_OPT="$([ "$LANG" = "it" ] && echo "Opzione sconosciuta:" || echo "Unknown option:")"; warn "$UNKNOWN_OPT $1" ;;
            *)           [[ -z "$ARG_ROOT" ]] && ARG_ROOT="$1" ;;
        esac
        shift
    done

    # La cartella dei report si valida SUBITO, prima di leggere qualunque cosa:
    # scoprire a fine analisi che non era scrivibile significa aver letto un
    # disco per niente. Un percorso indicato e non utilizzabile e' un errore,
    # non un motivo per ripiegare in silenzio sul default.
    if [[ -n "$REPORT_DIR_FIXED" ]]; then
        REPORT_DIR_FIXED=$(realpath -m "$REPORT_DIR_FIXED" 2>/dev/null || echo "$REPORT_DIR_FIXED")
        if ! mkdir -p "$REPORT_DIR_FIXED" 2>/dev/null; then
            err "$(L "Impossibile creare la cartella dei report:" "Cannot create the report directory:") $REPORT_DIR_FIXED"
            exit 1
        fi
        if [[ ! -w "$REPORT_DIR_FIXED" ]]; then
            err "$(L "Cartella dei report non scrivibile:" "Report directory is not writable:") $REPORT_DIR_FIXED"
            exit 1
        fi
    fi

    if [[ -n "$TIME_SINCE" && -n "$TIME_UNTIL" && "$TIME_SINCE" > "$TIME_UNTIL" ]]; then
        err "$(L "Finestra temporale vuota:" "Empty time window:") --since ${TIME_SINCE/T/ } > --until ${TIME_UNTIL/T/ }"
        exit 1
    fi
    if time_window_active; then
        info "$(L "Finestra di analisi:" "Analysis window:") ${BOLD}$(time_window_label)"
        info "$(L "Le righe datate fuori dalla finestra saranno escluse dai report e dall'export." \
                 "Dated rows outside the window will be excluded from reports and export.")"
    fi

    # L'immagine si apre prima di tutto: da qui in poi e' una root come le
    # altre, e nessun modulo deve sapere di stare leggendo dentro una E01.
    if [[ -n "$IMAGE_PATH" ]]; then
        register_tmp "${TMPDIR:-/tmp}/fiuto_image_$$"
        local _MP _RC
        _MP=$(image_open); _RC=$?
        if [[ "$_RC" -eq 2 ]]; then
            exit 0            # --list-partitions: elenco stampato, niente da montare
        elif [[ "$_RC" -ne 0 || -z "$_MP" ]]; then
            exit 1
        fi
        ARG_ROOT="$_MP"
        evidence_note "$IMAGE_PATH" "immagine forense analizzata"
    fi

    if [[ -n "$ARG_ROOT" ]]; then
        if [[ ! -d "$ARG_ROOT" ]]; then
            err "$(t dir_not_found) $ARG_ROOT"; exit 1
        fi
        _apply_win_root "$ARG_ROOT"
    fi

    # Non-interactive mode
    if $ARG_ALL; then
        print_banner
        [[ -z "$WIN_ROOT" ]] && { err "$(t specify_root_all)"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        run_all_from_registry "$(active_registry_name)"
        exit 0
    fi
    if [[ -n "$ARG_MODULE" ]]; then
        [[ -z "$WIN_ROOT" ]] && { err "$(t specify_root_module)"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        dispatch_from_registry "$(active_registry_name)" "$ARG_MODULE"
        exit 0
    fi
    if [[ -n "$ARG_MODULES" ]]; then
        [[ -z "$WIN_ROOT" ]] && { err "Specifica la root: $0 /mnt/windows --modules $ARG_MODULES"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        local MOD_NUMS
        mapfile -t MOD_NUMS < <(expand_module_list "$ARG_MODULES")
        local _rn; _rn=$(active_registry_name)
        for N in "${MOD_NUMS[@]}"; do
            dispatch_from_registry "$_rn" "$N"
        done
        exit 0
    fi

    # Carica IoC se specificato (per --all o modalità interattiva)
    [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"

    # Modalità interattiva — chiedi prima ROOT, poi REPORT dir
    if [[ -z "$WIN_ROOT" ]]; then
        print_banner
        echo -e "  ${CYAN}[*]${RESET} $(L "Ricerca automatica di volumi montati (Windows/Linux/macOS)..." "Automatically searching for mounted volumes (Windows/Linux/macOS)...")"
        if ! autodetect_win_root silent; then
            echo ""
            warn "$(L "Nessun volume Windows rilevato automaticamente." "No Windows volume detected automatically.")"
            echo ""
            echo -ne "  ${YELLOW}[?]${RESET} $(L "Inserisci il path della root da analizzare (o INVIO per saltare):" "Enter analysis root path (or ENTER to skip):") "
            local _MR; read -r _MR || true
            if [[ -n "$_MR" ]]; then
                _MR=$(realpath -m "$_MR" 2>/dev/null || echo "$_MR")
                if [[ -d "$_MR" ]]; then
                    _apply_win_root "$_MR"
                else
                    err "$(L "Directory non trovata:" "Directory not found:") $_MR"
                    info "$(L "Usa [R] dal menu per impostare la root." "Use [R] from menu to set the root.")"
                    sleep 2
                fi
            else
                info "$(L "Root non impostata. Usa [R] dal menu." "Root not set. Use [R] from menu.")"
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
        local _RN; _RN=$(active_registry_name)
        if [[ -n "$_RN" ]]; then
            render_menu_from_registry "$_RN"
        else
            warn "$(L "Nessun volume valido selezionato. Usa [R] per impostarlo." "No valid volume selected. Use [R] to set it.")"
            echo ""
            echo -e "  ${WHITE}[R]${RESET}  $(L "Imposta root da analizzare" "Set analysis root")"
            echo -e "  ${YELLOW}[D]${RESET}  $(L "Debug mount attivi" "Debug active mounts")"
            echo -e "  ${RED}[Q]${RESET}  $(L "Esci" "Quit")"
            echo ""
            echo -ne "  ${YELLOW}$(L "Scelta" "Choice"):${RESET} "
        fi
        read -r CHOICE
        echo ""

        case "${CHOICE^^}" in
            P)  setup_report_dir || true; sleep 1 ;;
            R)  set_win_root; sleep 1 ;;
            D)  debug_mounts ;;
            0)  run_all_from_registry "$(active_registry_name)"
                return_to_menu ;;
            S)  if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
                    generate_executive_summary
                else
                    warn "$(L "Nessun report in questa sessione: esegui prima almeno un modulo." \
                             "No report in this session: run at least one module first.")"
                fi
                return_to_menu ;;
            Q)  echo ""
                if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
                    clear
                    echo -e "${CYAN}${BOLD}"
                    echo "  ╔══════════════════════════════════════════════════════════╗"
                    local _bt; _bt="$(L "FIUTO — Report generati in questa sessione" "FIUTO — Reports generated in this session")"
                    local _btl=$(( (58 - ${#_bt}) / 2 )) _btr=$(( 58 - ${#_bt} - (58 - ${#_bt}) / 2 ))
                    printf "  ║%*s%s%*s║\n" "$_btl" "" "$_bt" "$_btr" ""
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
                    # Il comando suggerito e' quello che esiste davvero su
                    # questa macchina: su macOS xdg-open non c'e'. Se non ce
                    # n'e' nessuno non se ne inventa uno: suggerire un comando
                    # assente manda l'utente a sbattere.
                    if _OPENER=$(report_opener); then
                        echo -e "  ${DIM}$(L "Apri con:" "Open with:") ${_OPENER} \"<$(L "percorso" "path")>\"${RESET}"
                    else
                        echo -e "  ${DIM}$(L "Apri i report con il tuo browser." "Open the reports with your browser.")${RESET}"
                    fi
                    echo ""
                fi
                echo -e "  ${DIM}$(L "Uscita." "Exiting.")${RESET}"; echo ""; exit 0 ;;
            *)  if [[ "$CHOICE" =~ ^[0-9]+$ ]]; then
                    local _rn; _rn=$(active_registry_name)
                    if [[ -n "$_rn" ]]; then
                        dispatch_from_registry "$_rn" "$CHOICE"; return_to_menu
                    else
                        warn "$(L "Nessun volume valido selezionato. Usa [R]." "No valid volume selected. Use [R].")"; sleep 1
                    fi
                else
                    warn "$(L "Scelta non valida:" "Invalid choice:") '$CHOICE'"; sleep 1
                fi ;;
        esac
    done
}

# FIUTO_LIB_ONLY=1 carica le funzioni senza avviare l'interfaccia:
# usato dalla suite di test per fare unit test degli helper.
[[ -n "${FIUTO_LIB_ONLY:-}" ]] || main "$@"
