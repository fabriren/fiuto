_global_cleanup() {
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
