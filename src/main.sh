
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
                    echo ""
                    echo -e "  ${DIM}Di default i transaction log del registro (.LOG1/.LOG2) vengono"
                    echo -e "    riapplicati su una copia temporanea: senza questo passaggio le"
                    echo -e "    scritture piu' recenti dell'hive non sono visibili.${RESET}"
                    echo ""
                    echo -e "  ${BOLD}Moduli disponibili (1-39):${RESET}"
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
                    echo ""
                    echo -e "  ${DIM}By default registry transaction logs (.LOG1/.LOG2) are replayed"
                    echo -e "    onto a temporary copy: without this step the most recent hive"
                    echo -e "    writes are not visible.${RESET}"
                    echo ""
                    echo -e "  ${BOLD}Available modules (1-39):${RESET}"
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
            --format)    [[ "${2:-}" == "jsonl" ]] && EXPORT_JSONL=true; shift ;;
            -*)          local UNKNOWN_OPT="$([ "$LANG" = "it" ] && echo "Opzione sconosciuta:" || echo "Unknown option:")"; warn "$UNKNOWN_OPT $1" ;;
            *)           [[ -z "$ARG_ROOT" ]] && ARG_ROOT="$1" ;;
        esac
        shift
    done

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
        if [[ "$OS_TYPE" == "windows" ]]; then
            run_all_modules
        else
            run_all_from_registry "$(active_registry_name)"
        fi
        exit 0
    fi
    if [[ -n "$ARG_MODULE" ]]; then
        [[ -z "$WIN_ROOT" ]] && { err "$(t specify_root_module)"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        if [[ "$OS_TYPE" != "windows" ]]; then
            dispatch_from_registry "$(active_registry_name)" "$ARG_MODULE"
            exit 0
        fi
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
            39) module_ai_chat ;;
            *)  err "$(L "Modulo sconosciuto:" "Unknown module:") $ARG_MODULE" ;;
        esac
        exit 0
    fi
    if [[ -n "$ARG_MODULES" ]]; then
        [[ -z "$WIN_ROOT" ]] && { err "Specifica la root: $0 /mnt/windows --modules $ARG_MODULES"; exit 1; }
        [[ -n "$ARG_IOC" ]] && load_ioc_file "$ARG_IOC"
        local MOD_NUMS
        mapfile -t MOD_NUMS < <(expand_module_list "$ARG_MODULES")
        local _rn; _rn=$(active_registry_name)
        for N in "${MOD_NUMS[@]}"; do
            if [[ "$OS_TYPE" == "windows" ]]; then
                run_module_by_number "$N"
            else
                dispatch_from_registry "$_rn" "$N"
            fi
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
            print_menu
        fi
        read -r CHOICE
        echo ""

        case "${CHOICE^^}" in
            P)  setup_report_dir || true; sleep 1 ;;
            R)  set_win_root; sleep 1 ;;
            D)  debug_mounts ;;
            0)  if [[ "$OS_TYPE" == "windows" ]]; then
                    run_all_modules
                else
                    run_all_from_registry "$(active_registry_name)"
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
                    if [[ "$OS_TYPE" == "windows" ]]; then
                        run_module_by_number "$CHOICE"; return_to_menu
                    else
                        local _rn; _rn=$(active_registry_name)
                        if [[ -n "$_rn" ]]; then
                            dispatch_from_registry "$_rn" "$CHOICE"; return_to_menu
                        else
                            warn "$(L "Nessun volume valido selezionato. Usa [R]." "No valid volume selected. Use [R].")"; sleep 1
                        fi
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
