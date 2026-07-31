
# --- LINUX 4 — Shell History (per utente) ---
module_linux_shell_history() {
    section_header "Linux - Shell History" "$MAGENTA"
    check_target_root || return 1
    local FILES=(.bash_history .zsh_history .sh_history .history .python_history .mysql_history .psql_history .lesshst .node_repl_history)
    local KW="password|passwd|secret|credential|token|ssh|scp|curl|wget|nc |ncat|base64|sudo|chmod \\+x|/dev/tcp|reverse|nmap|chattr"
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        for HF in "${FILES[@]}"; do
            local F="$HOME_DIR/$HF"
            [[ -f "$F" && -s "$F" ]] || continue
            UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
            # I REPL readline (python3, node, psql, mysql) codificano spazi e
            # backslash in ottale: senza decodifica il report mostra
            # "\040\040value = ..." al posto del codice.
            local MODE="histts"
            case "$HF" in
                .python_history|.node_repl_history|.psql_history|.mysql_history) MODE="histrl" ;;
            esac
            CARDS+=$(file_card_html "$F" "$KW" "\$" "$MODE")
        done
        # fish history
        local FISH; FISH=$(ci_find_dir "$HOME_DIR" ".local/share/fish")
        if [[ -n "$FISH" && -s "$FISH/fish_history" ]]; then
            UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
            CARDS+=$(file_card_html "$FISH/fish_history" "$KW" "\$" "histts")
        fi
        [[ $UCOUNT -eq 0 ]] && { dim_msg "$UNAME - $(L "nessuna history" "no history")"; continue; }
        USERS=$((USERS + 1))
        ok "$UNAME - ${BOLD}$UCOUNT file"
        for HF in "${FILES[@]}"; do
            local F="$HOME_DIR/$HF"; [[ -f "$F" && -s "$F" ]] || continue
            echo -e "  ${DIM}• $HF${RESET}"; print_file_lines "$F" "$KW" 30; echo ""
        done
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT file" "$CARDS" "◢")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Utenti con history:" "Users with history:") ${BOLD}$USERS${RESET}  |  File: ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history trovata." "No history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "File history" "History files")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_shell_history" "Linux Shell History" "SH" "~/.bash_history · .zsh_history · ..." "$STATS" "<div class='cards'>$BODY</div>"
}
