
# --- LINUX 10 — Network Config ---
module_linux_network() {
    section_header "Linux — Network Config" "$CYAN"
    check_target_root || return 1
    local KW="psk|password|key=|wpa|pre-shared|nameserver|DROP|ACCEPT|REJECT|-j |MASQUERADE"
    local BODY="" FOUND=0
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    [[ -z "$ETC" ]] && { warn "$(L "Directory etc non trovata." "etc directory not found.")"; return 0; }
    for NF in hosts resolv.conf hostname; do
        local F; F=$(ci_find_file "$ETC" "$NF")
        [[ -n "$F" && -s "$F" ]] || continue
        FOUND=$((FOUND + 1)); ok "$NF"; BODY+=$(file_card_html "$F" "$KW" "⇄")
    done
    # NetworkManager connections (contengono PSK Wi-Fi)
    local NM; NM=$(ci_find_dir "$ETC" "NetworkManager/system-connections")
    if [[ -n "$NM" ]]; then
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); ok "NM: $(basename "$F")"
            BODY+=$(file_card_html "$F" "$KW" "⇄")
        done < <(find "$NM" -maxdepth 1 -type f 2>/dev/null)
    fi
    # netplan
    local NP; NP=$(ci_find_dir "$ETC" "netplan")
    [[ -n "$NP" ]] && while IFS= read -r F; do
        [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); BODY+=$(file_card_html "$F" "$KW" "⇄")
    done < <(find "$NP" -maxdepth 1 -type f 2>/dev/null)
    # iptables/nftables salvate
    for RF in "iptables/rules.v4" "iptables/rules.v6" "nftables.conf"; do
        local F; F=$(ci_find_file "$ETC" "$(basename "$RF")")
        [[ -n "$F" && -s "$F" ]] || continue; FOUND=$((FOUND + 1)); ok "$(basename "$RF")"
        BODY+=$(file_card_html "$F" "$KW" "⇄")
    done
    separator
    info "$(L "File di rete:" "Network files:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun file di configurazione di rete." "No network configuration file.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File rete" "Network files")" "$FOUND" "info")
    finish_report "linux_network" "Linux Network Config" "NET" "hosts · NetworkManager · netplan · iptables" "$STATS" "<div class='cards'>$BODY</div>"
}
