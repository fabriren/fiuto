
# ================================================================
#  AUTODETECT ROOT WINDOWS
# ================================================================

debug_mounts() {
    echo ""
    section_header "$(L "DEBUG - Mount attivi su questo sistema" "DEBUG - Active Mounts on This System")" "$YELLOW"
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
            _RW_MSG="${RED}[$(L "SOLA LETTURA - i report NON potranno essere salvati!" "READ ONLY - reports CANNOT be saved!")]${RESET}"
        fi
    else
        local _P; _P=$(dirname "$_CHOSEN")
        if [[ -w "$_P" ]]; then
            _RW_OK=true
            _RW_MSG="${GREEN}[$(L "verrà creata - parent scrivibile" "will be created - parent writable")]${RESET}"
        else
            _RW_MSG="${RED}[$(L "parent '${_P}' NON scrivibile - i report NON potranno essere salvati!" "parent '${_P}' NOT writable - reports CANNOT be saved!")]${RESET}"
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
