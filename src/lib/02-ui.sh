
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
    echo -e "  ║                    ${MAGENTA}${BOLD}v2.1 - zi®iginal${RESET}${CYAN}                      ║"
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
    [[ "${RESP,,}" != "n" ]] && xdg-open "$RPATH" 2>/dev/null &
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
