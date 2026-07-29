
# Carica un file IoC (una entry per riga, righe # ignorate)
load_ioc_file() {
    local IOCFILE="$1"
    if [[ ! -f "$IOCFILE" ]]; then
        warn "$(t ioc_not_found) $IOCFILE"
        return 1
    fi
    while IFS= read -r LINE || [[ -n "$LINE" ]]; do
        [[ -z "$LINE" || "$LINE" == \#* ]] && continue
        IOC_LIST+=("$LINE")
    done < "$IOCFILE"
    ok "$(t ioc_loaded) ${#IOC_LIST[@]} from $IOCFILE"
    log_msg "[IOC] Loaded ${#IOC_LIST[@]} IoCs from $IOCFILE"
}

# Controlla se una stringa contiene un IoC caricato; ritorna 0 se trovato
check_ioc() {
    local TEXT="${1,,}"
    for IOC in "${IOC_LIST[@]}"; do
        [[ "${TEXT}" == *"${IOC,,}"* ]] && return 0
    done
    return 1
}
