
# --- macOS 4 — Login Items (BTM) ---
module_macos_loginitems() {
    section_header "macOS — Login Items (BTM)" "$RED"
    check_target_root || return 1
    local BODY="" FOUND=0
    # backgrounditems.btm (vari percorsi a seconda della versione)
    mapfile -t BTM < <(find "$WIN_ROOT" -maxdepth 8 -iname "backgrounditems.btm" -type f 2>/dev/null; find "$WIN_ROOT" -maxdepth 8 -path "*backgroundtaskmanagementagent*" -type f 2>/dev/null | head -20)
    for F in "${BTM[@]}"; do
        [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); ok "$(basename "$F")"
        # BTM è un blob binario: estrai riferimenti a path/app leggibili
        local STR; STR=$(strings -n 5 "$F" 2>/dev/null | grep -iE '\.app|/Users/|/Library/|/usr/|\.plist|\.sh$' | sort -u | head -100)
        local TMPF; TMPF=$(mktemp); printf '%s\n' "$STR" > "$TMPF"
        BODY+=$(generic_card_html "$(basename "$F")" "$F" "$(printf '%s\n' "$STR" | grep -c .) ref" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "tmp|/private|curl|\\.sh")</pre>" "⚑")
        rm -f "$TMPF"
    done
    separator
    info "$(L "File BTM trovati:" "BTM files found:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun login item (BTM) trovato." "No login item (BTM) found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File BTM" "BTM files")" "$FOUND" "info")
    finish_report "macos_loginitems" "macOS Login Items (BTM)" "BTM" "backgrounditems.btm" "$STATS" "<div class='cards'>$BODY</div>"
}
