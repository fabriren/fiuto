
# --- macOS 1 — System Logs ---
module_macos_logs() {
    section_header "macOS — System Logs" "$GREEN"
    check_target_root || return 1
    local KW="fail|error|denied|invalid|sudo|root|unauthorized|jailbreak|malware"
    local BODY="" FOUND=0
    local LOGDIR; LOGDIR=$(ci_find_dir "$WIN_ROOT" "var/log")
    if [[ -n "$LOGDIR" ]]; then
        for NAME in system.log install.log secure.log wifi.log appfirewall.log; do
            local F; F=$(ci_find_file "$LOGDIR" "$NAME")
            [[ -n "$F" && -s "$F" ]] || continue
            FOUND=$((FOUND + 1)); ok "$NAME"; print_file_lines "$F" "$KW" 40; echo ""
            BODY+=$(file_card_html "$F" "$KW" "≣")
        done
        # ASL (formato binario: estrazione stringhe)
        local ASL; ASL=$(ci_find_dir "$LOGDIR" "asl")
        if [[ -n "$ASL" ]]; then
            local ACOUNT; ACOUNT=$(find "$ASL" -maxdepth 1 -type f 2>/dev/null | wc -l)
            [[ "$ACOUNT" -gt 0 ]] && { FOUND=$((FOUND + 1)); ok "asl ($ACOUNT file)"; BODY+="<div class='card'><div class='card-header'><div class='uicon'>≣</div><div><div class='uname'>ASL logs</div><div class='upath'>$ASL</div></div><div class='badge'>$ACOUNT file</div></div><div style='padding:1rem 1.5rem'><p class='dim mono' style='font-size:.72rem'>$(L "Formato binario ASL — analisi approfondita fuori scope offline." "Binary ASL format — deep parsing out of offline scope.")</p></div></div>"; }
        fi
    fi
    # Nota: unified logs .tracev3 esplicitamente fuori scope
    warn "$(L "Unified logs (.tracev3) non analizzati: richiedono il tool 'log' o parser dedicati (fuori scope)." "Unified logs (.tracev3) not parsed: require the 'log' tool or dedicated parsers (out of scope).")"
    separator
    info "$(L "Log trovati:" "Logs found:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun log testuale trovato." "No text log found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Log" "Logs")" "$FOUND" "info")
    finish_report "macos_logs" "macOS System Logs" "LOG" "system.log · install.log · asl" "$STATS" "<div class='cards'>$BODY</div>"
}
