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

        "en:batch_started") echo "=== FIUTO started in batch mode - WIN_ROOT=" ;;
        "it:batch_started") echo "=== FIUTO avviato in modalità batch - WIN_ROOT=" ;;

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
