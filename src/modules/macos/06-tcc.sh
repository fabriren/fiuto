
# --- macOS 6 — TCC Privacy ---
module_macos_tcc() {
    section_header "macOS - TCC Privacy" "$RED"
    check_target_root || return 1
    local BODY="" TOTAL=0
    _tcc_card() { # db, label
        local DB="$1" LBL="$2"
        local ROWS; ROWS=$(query_sqlite "$DB" "SELECT service, client, CASE auth_value WHEN 0 THEN 'denied' WHEN 2 THEN 'allowed' ELSE auth_value END FROM access ORDER BY service")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && ROWS=$(query_sqlite "$DB" "SELECT service, client, allowed FROM access ORDER BY service")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && return
        local N; N=$(printf '%s\n' "$ROWS" | grep -c .); TOTAL=$((TOTAL + N))
        ok "$LBL - ${BOLD}$N $(L "permessi" "permissions")"
        local TABLE; TABLE=$(_rows_to_table "$ROWS" "Service" "Client" "Auth")
        BODY+=$(generic_card_html "$LBL" "$DB" "$N" "$TABLE" "⊘")
    }
    local SYS; SYS=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "Library/Application Support/com.apple.TCC")" "TCC.db")
    [[ -n "$SYS" ]] && _tcc_card "$SYS" "System TCC"
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local DB; DB=$(ci_find_file "$(ci_find_dir "$HOME_DIR" "Library/Application Support/com.apple.TCC")" "TCC.db")
        [[ -n "$DB" ]] && _tcc_card "$DB" "$U TCC"
    done < <(get_macos_user_homes)
    separator
    info "$(L "Permessi TCC:" "TCC permissions:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun database TCC leggibile." "No readable TCC database.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Permessi" "Permissions")" "$TOTAL" "info")
    finish_report "macos_tcc" "macOS TCC Privacy" "TCC" "TCC.db (cam/mic/disco)" "$STATS" "<div class='cards'>$BODY</div>"
}
