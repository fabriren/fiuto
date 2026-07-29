
# --- LINUX 2 — systemd Journal ---
module_linux_journal() {
    section_header "Linux — systemd Journal" "$GREEN"
    check_target_root || return 1
    local JDIR; JDIR=$(ci_find_dir "$WIN_ROOT" "var/log/journal")
    [[ -z "$JDIR" ]] && { warn "$(L "Nessun journal persistente (var/log/journal assente)." "No persistent journal (var/log/journal missing).")"; return 0; }
    mapfile -t JFILES < <(find "$JDIR" -type f -iname "*.journal" 2>/dev/null)
    [[ ${#JFILES[@]} -eq 0 ]] && { warn "$(L "Nessun file .journal trovato." "No .journal file found.")"; return 0; }
    ok "$(L "File journal trovati:" "Journal files found:") ${BOLD}${#JFILES[@]}"

    local KW="fail|error|denied|refused|sudo|root|segfault|authentication"
    local BODY="" PARSED=0 HAVE_JCTL=0
    command -v journalctl >/dev/null 2>&1 && HAVE_JCTL=1
    for JF in "${JFILES[@]}"; do
        local TXT=""
        if [[ $HAVE_JCTL -eq 1 ]]; then
            TXT=$(journalctl --no-pager --file "$JF" -o short-iso 2>/dev/null | head -100000)
        fi
        if [[ -z "$TXT" ]]; then
            # fallback: estrazione stringhe leggibili dal binario
            TXT=$(strings -n 8 "$JF" 2>/dev/null | grep -iE 'MESSAGE=|_COMM=|_EXE=|_HOSTNAME=' | sed 's/^MESSAGE=//' | head -100000)
        fi
        [[ -z "$TXT" ]] && continue
        PARSED=$((PARSED + 1))
        echo -e "  ${DIM}• $(basename "$JF")${RESET}"
        local TMPF; TMPF=$(mktemp); printf '%s\n' "$TXT" > "$TMPF"
        BODY+=$(file_card_html "$TMPF" "$KW" "◷")
        rm -f "$TMPF"
    done
    [[ $HAVE_JCTL -eq 0 ]] && warn "$(L "journalctl non disponibile: usato fallback 'strings' (output parziale)." "journalctl unavailable: used 'strings' fallback (partial output).")"
    separator
    info "$(L "Journal elaborati:" "Journals processed:") ${BOLD}$PARSED"
    [[ $PARSED -eq 0 ]] && { warn "$(L "Impossibile estrarre contenuti dal journal." "Could not extract journal contents.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File journal" "Journal files")" "${#JFILES[@]}" "info")
    finish_report "linux_journal" "systemd Journal" "JRN" "/var/log/journal" "$STATS" "<div class='cards'>$BODY</div>"
}
