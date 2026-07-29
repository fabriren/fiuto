
# --- macOS 10 — Recent Items ---
module_macos_recent() {
    section_header "macOS — Recent Items" "$GREEN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR"); local CARDS="" UCOUNT=0
        # SFL / SFL2 (shared file list — recent apps/docs/servers)
        local SFLD; SFLD=$(ci_find_dir "$HOME_DIR" "Library/Application Support/com.apple.sharedfilelist")
        if [[ -n "$SFLD" ]]; then
            while IFS= read -r SF; do
                [[ -s "$SF" ]] || continue
                local STR; STR=$(strings -n 5 "$SF" 2>/dev/null | grep -iE '/Users/|/Volumes/|\.app|\.|smb://|afp://|ftp://' | sort -u | head -60)
                [[ -z "$STR" ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
                local TMPF; TMPF=$(mktemp); printf '%s\n' "$STR" > "$TMPF"
                CARDS+=$(generic_card_html "$(basename "$SF")" "$SF" "$(printf '%s\n' "$STR" | grep -c .) ref" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "smb://|afp://|/Volumes")</pre>" "◇")
                rm -f "$TMPF"
            done < <(find "$SFLD" -maxdepth 2 -type f \( -iname "*.sfl" -o -iname "*.sfl2" -o -iname "*.sfl3" \) 2>/dev/null)
        fi
        # Trash
        local TRASH; TRASH=$(ci_find_dir "$HOME_DIR" ".Trash")
        if [[ -n "$TRASH" ]]; then
            local TN; TN=$(find "$TRASH" -mindepth 1 -maxdepth 1 2>/dev/null | wc -l)
            if [[ "$TN" -gt 0 ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + TN))
                local TLIST; TLIST=$(find "$TRASH" -mindepth 1 -maxdepth 1 -printf '%T+\t%p\n' 2>/dev/null | sort -r | head -200)
                CARDS+=$(generic_card_html ".Trash" "$TRASH" "$TN" "$(_rows_to_table "$TLIST" "$(L "Modificato" "Modified")" "$(L "Percorso" "Path")")" "♺")
            fi
        fi
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); ok "$U — ${BOLD}$UCOUNT $(L "sorgenti" "sources")"
        BODY+=$(generic_card_html "$U" "$HOME_DIR" "$UCOUNT" "$CARDS" "◇")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Voci recenti:" "Recent items:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna voce recente trovata." "No recent item found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Voci" "Items")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "macos_recent" "macOS Recent Items" "RCN" "SFL · .Trash · recent items" "$STATS" "<div class='cards'>$BODY</div>"
}
