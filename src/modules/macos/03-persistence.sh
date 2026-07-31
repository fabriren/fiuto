
# --- macOS 3 — Persistence (LaunchAgents/Daemons/cron) ---
module_macos_persistence() {
    section_header "macOS - Persistence" "$ORANGE"
    check_target_root || return 1
    local KW="curl|wget|/tmp/|/var/tmp|base64|python|bash -i|nc |/dev/tcp|RunAtLoad|KeepAlive|http"
    local BODY="" FOUND=0
    _mac_persist_dir() { # label, dir
        local LABEL="$1" D="$2"; [[ -z "$D" || ! -d "$D" ]] && return
        local n=0
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1)); n=$((n + 1))
            if [[ "$F" == *.plist ]]; then
                # rende leggibili anche i plist binari
                local TMPF; TMPF=$(mktemp); read_plist "$F" > "$TMPF" 2>/dev/null
                [[ -s "$TMPF" ]] || cat "$F" > "$TMPF" 2>/dev/null
                BODY+=$(generic_card_html "$(basename "$F")" "$F" "$LABEL" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "$KW")</pre>" "⟳")
                rm -f "$TMPF"
            else
                BODY+=$(file_card_html "$F" "$KW" "⟳")
            fi
        done < <(find "$D" -maxdepth 1 -type f \( -iname "*.plist" -o -iname "*.conf" \) 2>/dev/null)
        [[ $n -gt 0 ]] && ok "$LABEL - $n plist ($D)"
    }
    _mac_persist_dir "LaunchDaemons" "$(ci_find_dir "$WIN_ROOT" "Library/LaunchDaemons")"
    _mac_persist_dir "LaunchAgents"  "$(ci_find_dir "$WIN_ROOT" "Library/LaunchAgents")"
    _mac_persist_dir "System LaunchDaemons" "$(ci_find_dir "$WIN_ROOT" "System/Library/LaunchDaemons")"
    # cron + periodic
    local CRON; CRON=$(ci_find_dir "$WIN_ROOT" "var/at/tabs"); _mac_persist_dir "cron" "$CRON"
    local PERIODIC; PERIODIC=$(ci_find_dir "$WIN_ROOT" "etc/periodic")
    # per-user LaunchAgents
    while IFS= read -r HOME_DIR; do
        _mac_persist_dir "$(basename "$HOME_DIR") LaunchAgents" "$(ci_find_dir "$HOME_DIR" "Library/LaunchAgents")"
    done < <(get_macos_user_homes)
    separator
    info "$(L "Artefatti di persistenza:" "Persistence artifacts:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun artefatto di persistenza trovato." "No persistence artifact found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Artefatti" "Artifacts")" "$FOUND" "info")
    finish_report "macos_persistence" "macOS Persistence" "PER" "LaunchAgents · LaunchDaemons · cron" "$STATS" "<div class='cards'>$BODY</div>"
}
