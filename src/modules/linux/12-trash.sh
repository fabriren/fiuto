
# --- LINUX 12 — Trash & Recent ---
module_linux_trash() {
    section_header "Linux — Trash & Recent" "$GREEN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        # Trash info files (.trashinfo: path + DeletionDate)
        local TINFO; TINFO=$(ci_find_dir "$HOME_DIR" ".local/share/Trash/info")
        if [[ -n "$TINFO" ]]; then
            local TTAB="" TN=0
            while IFS= read -r TF; do
                [[ -s "$TF" ]] || continue; TN=$((TN + 1))
                local OPATH DDATE
                OPATH=$(grep -m1 '^Path=' "$TF" 2>/dev/null | cut -d= -f2-)
                DDATE=$(grep -m1 '^DeletionDate=' "$TF" 2>/dev/null | cut -d= -f2-)
                TTAB+="<tr><td class='mono dim'>$(html_esc "$DDATE")</td><td class='mono'>$(html_esc "$OPATH")</td></tr>"
            done < <(find "$TINFO" -maxdepth 1 -name "*.trashinfo" 2>/dev/null)
            if [[ $TN -gt 0 ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + TN))
                ok "$UNAME — ${BOLD}$TN $(L "file cestinati" "trashed files")"
                CARDS+=$(generic_card_html "Trash" "$TINFO" "$TN file" "<table><tr><th>$(L "Cancellato il" "Deleted on")</th><th>$(L "Percorso originale" "Original path")</th></tr>$TTAB</table>" "♺")
            fi
        fi
        # recently-used.xbel
        local XBEL; XBEL=$(ci_find_file "$HOME_DIR" ".local/share/recently-used.xbel")
        [[ -z "$XBEL" ]] && XBEL=$(ci_find_file "$HOME_DIR" ".recently-used.xbel")
        if [[ -n "$XBEL" && -s "$XBEL" ]]; then
            UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
            CARDS+=$(file_card_html "$XBEL" "href|exec|modified|visited" "♺")
        fi
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1))
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT" "$CARDS" "♺")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Voci trovate:" "Items found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun elemento nel cestino / recenti." "No trash / recent items.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Voci" "Items")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_trash" "Linux Trash & Recent" "TRH" "~/.local/share/Trash · recently-used.xbel" "$STATS" "<div class='cards'>$BODY</div>"
}
