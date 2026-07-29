
# --- macOS 8 — Browser History ---
module_macos_browser() {
    section_header "macOS — Browser History" "$CYAN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR"); local CARDS="" UCOUNT=0
        # Safari
        local SAF; SAF=$(ci_find_file "$(ci_find_dir "$HOME_DIR" "Library/Safari")" "History.db")
        if [[ -n "$SAF" ]]; then
            local ROWS; ROWS=$(query_sqlite "$SAF" "SELECT datetime(v.visit_time+978307200,'unixepoch'), i.url, v.title FROM history_visits v JOIN history_items i ON v.history_item=i.id ORDER BY v.visit_time DESC LIMIT 100000")
            if [[ -n "$ROWS" && "$ROWS" != ERROR* ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); ok "$U — Safari"
                CARDS+=$(generic_card_html "Safari" "$SAF" "$(printf '%s\n' "$ROWS" | grep -c .) URL" "$(_rows_to_table "$ROWS" "$(L "Data" "Date")" "URL" "$(L "Titolo" "Title")")" "◐")
            fi
        fi
        # Chrome/Brave/Edge
        for CR in "Library/Application Support/Google/Chrome" "Library/Application Support/BraveSoftware/Brave-Browser" "Library/Application Support/Microsoft Edge"; do
            local BASE; BASE=$(ci_find_dir "$HOME_DIR" "$CR"); [[ -z "$BASE" ]] && continue
            while IFS= read -r HISTDB; do
                local ROWS; ROWS=$(query_sqlite "$HISTDB" "SELECT datetime(last_visit_time/1000000-11644473600,'unixepoch'), url, title FROM urls ORDER BY last_visit_time DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); ok "$U — $(basename "$(dirname "$HISTDB")") (Chromium)"
                CARDS+=$(_browser_table_card "$HISTDB" "$ROWS")
            done < <(find "$BASE" -maxdepth 2 -name "History" -type f 2>/dev/null)
        done
        # Firefox
        local FB; FB=$(ci_find_dir "$HOME_DIR" "Library/Application Support/Firefox/Profiles")
        if [[ -n "$FB" ]]; then
            while IFS= read -r PLACES; do
                local ROWS; ROWS=$(query_sqlite "$PLACES" "SELECT datetime(last_visit_date/1000000,'unixepoch'), url, title FROM moz_places WHERE last_visit_date IS NOT NULL ORDER BY last_visit_date DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); ok "$U — Firefox"
                CARDS+=$(_browser_table_card "$PLACES" "$ROWS")
            done < <(find "$FB" -maxdepth 2 -name "places.sqlite" -type f 2>/dev/null)
        fi
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); BODY+=$(generic_card_html "$U" "$HOME_DIR" "$UCOUNT" "$CARDS" "◐")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Profili browser:" "Browser profiles:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history browser trovata." "No browser history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Profili" "Profiles")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "macos_browser" "macOS Browser History" "WEB" "Safari · Chrome · Firefox" "$STATS" "<div class='cards'>$BODY</div>"
}
