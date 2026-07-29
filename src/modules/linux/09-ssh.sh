
# --- LINUX 9 — SSH Artifacts ---
module_linux_ssh() {
    section_header "Linux — SSH Artifacts" "$YELLOW"
    check_target_root || return 1
    local KW="ssh-rsa|ssh-ed25519|ecdsa|PermitRootLogin|PasswordAuthentication|AuthorizedKeys|ForwardAgent"
    local BODY="" FOUND=0
    # sshd_config di sistema
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc/ssh")
    if [[ -n "$ETC" ]]; then
        local SC; SC=$(ci_find_file "$ETC" "sshd_config")
        [[ -n "$SC" && -s "$SC" ]] && { FOUND=$((FOUND + 1)); ok "sshd_config"; print_file_lines "$SC" "PermitRootLogin|PasswordAuthentication|AllowUsers" 40; echo ""; BODY+=$(file_card_html "$SC" "$KW" "⚿"); }
    fi
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local SSHD; SSHD=$(ci_find_dir "$HOME_DIR" ".ssh")
        [[ -z "$SSHD" ]] && continue
        local CARDS="" UCOUNT=0
        for KF in authorized_keys authorized_keys2 known_hosts config; do
            local F="$SSHD/$KF"
            [[ -f "$F" && -s "$F" ]] || continue
            UCOUNT=$((UCOUNT + 1)); FOUND=$((FOUND + 1)); CARDS+=$(file_card_html "$F" "$KW" "⚿")
        done
        # chiavi private presenti (solo presenza, non contenuto)
        local PRIV; PRIV=$(find "$SSHD" -maxdepth 1 -type f -name "id_*" ! -name "*.pub" 2>/dev/null | wc -l)
        [[ "$PRIV" -gt 0 ]] && CARDS+="<div class='card' style='margin-bottom:.8rem'><div class='card-header'><div class='uicon'>⚿</div><div><div class='uname'>$(L "Chiavi private" "Private keys")</div><div class='upath'>$SSHD/id_*</div></div><div class='badge warn'>$PRIV $(L "chiavi" "keys")</div></div></div>"
        [[ $UCOUNT -eq 0 && "$PRIV" -eq 0 ]] && continue
        ok "$UNAME — ${BOLD}$UCOUNT file"
        BODY+=$(generic_card_html "$UNAME" "$SSHD" "$UCOUNT file" "$CARDS" "⚿")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Artefatti SSH:" "SSH artifacts:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun artefatto SSH trovato." "No SSH artifact found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File SSH" "SSH files")" "$FOUND" "info")
    finish_report "linux_ssh" "Linux SSH Artifacts" "SSH" "authorized_keys · known_hosts · sshd_config" "$STATS" "<div class='cards'>$BODY</div>"
}
