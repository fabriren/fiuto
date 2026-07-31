
# --- LINUX 8 — Persistence ---
module_linux_persistence() {
    section_header "Linux - Persistence" "$ORANGE"
    check_target_root || return 1
    local KW="curl|wget|/tmp/|/dev/shm|base64|nc |ncat|python -c|bash -i|/dev/tcp|chmod|reverse|\\.onion|http"
    local BODY="" FOUND=0
    _persist_add() { # label, path(file o dir), glob
        local LABEL="$1" P="$2"
        if [[ -f "$P" && -s "$P" ]]; then
            FOUND=$((FOUND + 1)); ok "$LABEL - $(basename "$P")"
            BODY+=$(file_card_html "$P" "$KW" "⟳")
        elif [[ -d "$P" ]]; then
            while IFS= read -r F; do
                [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1))
                BODY+=$(file_card_html "$F" "$KW" "⟳")
            done < <(find "$P" -maxdepth 2 -type f 2>/dev/null | head -60)
            [[ -n "$(find "$P" -maxdepth 2 -type f 2>/dev/null | head -1)" ]] && ok "$LABEL ($(basename "$P"))"
        fi
    }
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    if [[ -n "$ETC" ]]; then
        _persist_add "cron" "$(ci_find_file "$ETC" "crontab")"
        for D in cron.d cron.daily cron.hourly cron.weekly cron.monthly; do
            _persist_add "cron" "$(ci_find_dir "$ETC" "$D")"
        done
        _persist_add "init" "$(ci_find_file "$ETC" "rc.local")"
        _persist_add "init.d" "$(ci_find_dir "$ETC" "init.d")"
        _persist_add "systemd" "$(ci_find_dir "$ETC" "systemd/system")"
        _persist_add "xdg-autostart" "$(ci_find_dir "$ETC" "xdg/autostart")"
        _persist_add "ld.so.preload" "$(ci_find_file "$ETC" "ld.so.preload")"
    fi
    # user crontabs
    local SPOOL; SPOOL=$(ci_find_dir "$WIN_ROOT" "var/spool/cron")
    [[ -n "$SPOOL" ]] && _persist_add "user-cron" "$SPOOL"
    # systemd vendor units
    local USRLIB; USRLIB=$(ci_find_dir "$WIN_ROOT" "usr/lib/systemd/system")
    # per-user autostart, systemd user, shell init
    while IFS= read -r HOME_DIR; do
        _persist_add "autostart" "$(ci_find_dir "$HOME_DIR" ".config/autostart")"
        _persist_add "systemd-user" "$(ci_find_dir "$HOME_DIR" ".config/systemd/user")"
        for RC in .bashrc .bash_profile .profile .zshrc .zprofile; do
            _persist_add "shell-init" "$HOME_DIR/$RC"
        done
    done < <(get_linux_user_homes)
    separator
    info "$(L "Artefatti di persistenza:" "Persistence artifacts:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun artefatto di persistenza trovato." "No persistence artifact found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Artefatti" "Artifacts")" "$FOUND" "info")
    finish_report "linux_persistence" "Linux Persistence" "PER" "cron · systemd · autostart · rc · shell-init" "$STATS" "<div class='cards'>$BODY</div>"
}
