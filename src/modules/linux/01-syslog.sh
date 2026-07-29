
# ================================================================
#  MODULI LINUX
# ================================================================

# --- LINUX 1 — System Logs (/var/log testuali) ---
module_linux_syslog() {
    section_header "Linux — System Logs" "$GREEN"
    check_target_root || return 1
    local LOGDIR; LOGDIR=$(ci_find_dir "$WIN_ROOT" "var/log")
    [[ -z "$LOGDIR" ]] && { warn "$(L "Directory var/log non trovata." "var/log directory not found.")"; return 0; }

    local TARGETS=(syslog messages kern.log dmesg auth.log secure boot.log faillog dpkg.log yum.log)
    local KW="fail|error|denied|refused|invalid|segfault|root|sudo|su:|authentication failure|break-in|illegal"
    local BODY="" FOUND=0
    for NAME in "${TARGETS[@]}"; do
        local F; F=$(ci_find_file "$LOGDIR" "$NAME")
        [[ -z "$F" || ! -s "$F" ]] && continue
        FOUND=$((FOUND + 1))
        ok "$NAME — ${BOLD}$(stat -c %s "$F" 2>/dev/null) B"
        print_file_lines "$F" "$KW" 60
        echo ""
        BODY+=$(file_card_html "$F" "$KW" "≣")
    done
    separator
    info "$(L "Log trovati:" "Logs found:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun log testuale trovato." "No text log found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Log analizzati" "Logs analysed")" "$FOUND" "info")
    finish_report "linux_syslog" "Linux System Logs" "LOG" "/var/log" "$STATS" "<div class='cards'>$BODY</div>"
}
