
# --- LINUX 11 — Installed Packages ---
module_linux_packages() {
    section_header "Linux — Installed Packages" "$GREEN"
    check_target_root || return 1
    local BODY="" FOUND=0
    # dpkg
    local DPKG; DPKG=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "var/lib/dpkg")" "status")
    if [[ -n "$DPKG" && -s "$DPKG" ]]; then
        FOUND=$((FOUND + 1))
        local PKGS; PKGS=$(grep -c '^Package:' "$DPKG" 2>/dev/null)
        ok "dpkg — ${BOLD}$PKGS pacchetti"
        local TAB; TAB=$(awk '/^Package:/{p=$2} /^Version:/{v=$2} /^Status:/{s=$0} /^$/{if(p)printf "<tr><td class=mono>%s</td><td class=\"mono dim\">%s</td></tr>\n",p,v; p=v=""}' "$DPKG" | head -100000)
        BODY+=$(generic_card_html "dpkg packages" "$DPKG" "$PKGS pkg" "<table><tr><th>Package</th><th>Version</th></tr>$TAB</table>" "▦")
    fi
    # apt history (timeline installazioni)
    local APTLOG; APTLOG=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "var/log/apt")" "history.log")
    if [[ -n "$APTLOG" && -s "$APTLOG" ]]; then
        FOUND=$((FOUND + 1)); ok "apt history.log"
        print_file_lines "$APTLOG" "Install:|Remove:|Purge:" 40; echo ""
        BODY+=$(file_card_html "$APTLOG" "Install:|Remove:|Purge:|Commandline:" "▦")
    fi
    # rpm (db binario: solo presenza + tentativo via rpm se disponibile)
    local RPMDB; RPMDB=$(ci_find_dir "$WIN_ROOT" "var/lib/rpm")
    if [[ -n "$RPMDB" ]]; then
        FOUND=$((FOUND + 1)); ok "rpm db $(L "presente" "present")"
        local RTAB=""
        if command -v rpm >/dev/null 2>&1; then
            RTAB=$(rpm -qa --dbpath "$RPMDB" 2>/dev/null | sort | head -100000 | "$PY3" -c 'import sys,html;[print(f"<tr><td class=mono>{html.escape(l.strip())}</td></tr>") for l in sys.stdin]')
        fi
        BODY+=$(generic_card_html "rpm packages" "$RPMDB" "" "<table><tr><th>Package</th></tr>${RTAB:-<tr><td class=dim>rpm CLI non disponibile per dump offline</td></tr>}</table>" "▦")
    fi
    # snap / flatpak (elenco directory)
    local SNAP; SNAP=$(ci_find_dir "$WIN_ROOT" "var/lib/snapd/snaps")
    [[ -n "$SNAP" ]] && { FOUND=$((FOUND + 1)); ok "snap"; local STAB; STAB=$(find "$SNAP" -maxdepth 1 -name "*.snap" 2>/dev/null | sort | "$PY3" -c 'import sys,html,os;[print(f"<tr><td class=mono>{html.escape(os.path.basename(l.strip()))}</td></tr>") for l in sys.stdin]'); BODY+=$(generic_card_html "snap" "$SNAP" "" "<table><tr><th>Snap</th></tr>$STAB</table>" "▦"); }
    separator
    info "$(L "Fonti pacchetti:" "Package sources:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && { warn "$(L "Nessun database pacchetti trovato." "No package database found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Fonti" "Sources")" "$FOUND" "info")
    finish_report "linux_packages" "Linux Installed Packages" "PKG" "dpkg · rpm · apt history · snap" "$STATS" "<div class='cards'>$BODY</div>"
}
