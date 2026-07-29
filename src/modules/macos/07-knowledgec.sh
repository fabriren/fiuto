
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
