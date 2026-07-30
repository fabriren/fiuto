
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
                    echo -e "  ${BOLD}Moduli disponibili (1-50):${RESET}"
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
                    echo -e "  ${BOLD}Available modules (1-50):${RESET}"
                fi
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
                echo -e "   39  AI Chat History"
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

    if [[ -n "$TIME_SINCE" && -n "$TIME_UNTIL" && "$TIME_SINCE" > "$TIME_UNTIL" ]]; then
        err "$(L "Finestra temporale vuota:" "Empty time window:") --since ${TIME_SINCE/T/ } > --until ${TIME_UNTIL/T/ }"
        exit 1
    fi
    if time_window_active; then
        info "$(L "Finestra di analisi:" "Analysis window:") ${BOLD}$(time_window_label)"
        info "$(L "Le righe datate fuori dalla finestra saranno escluse dai report e dall'export." \
                 "Dated rows outside the window will be excluded from reports and export.")"
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
                    echo -e "  ${DIM}$(L "Apri con:" "Open with:") xdg-open \"<$(L "percorso" "path")>\"${RESET}"
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
