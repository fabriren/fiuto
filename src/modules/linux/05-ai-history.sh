
# --- LINUX 5 — AI CLI History ---
module_linux_ai_history() {
    section_header "Linux — AI CLI History" "$MAGENTA"
    check_target_root || return 1
    # path relativi alla home da scandire (file o directory)
    local REL=(.claude .config/claude .aider.chat.history.md .aider.input.history .config/aichat .ollama/history
               .config/io.datasette.llm .codeium .config/github-copilot .continue .cursor)
    local KW="password|secret|token|api_key|apikey|key=|credential"
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        for R in "${REL[@]}"; do
            local P="$HOME_DIR/$R"
            if [[ -f "$P" && -s "$P" ]]; then
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$P" "$KW" "◈")
            elif [[ -d "$P" ]]; then
                while IFS= read -r AF; do
                    [[ -s "$AF" ]] || continue
                    UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1)); CARDS+=$(file_card_html "$AF" "$KW" "◈")
                done < <(find "$P" -maxdepth 3 -type f \( -iname "*.json" -o -iname "*.jsonl" -o -iname "*history*" -o -iname "*.md" \) 2>/dev/null | head -25)
            fi
        done
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1)); ok "$UNAME — ${BOLD}$UCOUNT file AI"
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT file" "$CARDS" "◈")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Artefatti AI trovati:" "AI artifacts found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history di assistenti AI trovata." "No AI assistant history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "File AI" "AI files")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_ai_history" "Linux AI CLI History" "AI" "claude · aider · aichat · ollama · copilot" "$STATS" "<div class='cards'>$BODY</div>"
}
