
# --- macOS 5 — Quarantine / Downloads ---
module_macos_quarantine() {
    section_header "macOS - Quarantine / Downloads" "$CYAN"
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
            ok "$UNAME - ${BOLD}$N $(L "download tracciati" "tracked downloads")"
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
