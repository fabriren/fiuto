
# --- macOS 9 — Shell & AI History ---
module_macos_shell_ai_history() {
    section_header "macOS — Shell & AI History" "$MAGENTA"
    check_target_root || return 1
    local FILES=(.zsh_history .bash_history .sh_history .python_history .psql_history .node_repl_history
                 .claude .aider.chat.history.md .config/aichat .ollama/history)
    local KW="password|secret|token|api_key|apikey|credential|ssh|curl|base64|security|keychain|sudo"
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR"); local CARDS="" UCOUNT=0
        for HF in "${FILES[@]}"; do
            local P="$HOME_DIR/$HF"
            if [[ -f "$P" && -s "$P" ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$P" "$KW" "\$" "histts")
            elif [[ -d "$P" ]]; then
                while IFS= read -r AF; do
                    [[ -s "$AF" ]] || continue; UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$AF" "$KW" "◈")
                done < <(find "$P" -maxdepth 3 -type f \( -iname "*.json" -o -iname "*.jsonl" -o -iname "*history*" -o -iname "*.md" \) 2>/dev/null | head -20)
            fi
        done
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); ok "$U — ${BOLD}$UCOUNT file"
        BODY+=$(generic_card_html "$U" "$HOME_DIR" "$UCOUNT file" "$CARDS" "◢")
    done < <(get_macos_user_homes)
    separator
    info "$(L "Artefatti history:" "History artifacts:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history trovata." "No history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "File" "Files")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "macos_shell_ai_history" "macOS Shell & AI History" "SH" ".zsh_history · .bash_history · AI CLI" "$STATS" "<div class='cards'>$BODY</div>"
}
