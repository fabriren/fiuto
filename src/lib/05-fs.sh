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
