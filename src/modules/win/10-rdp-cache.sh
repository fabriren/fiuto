
# ================================================================
#  MODULO 4 — Cache RDP (Terminal Server Client)
# ================================================================
module_rdp_cache() {
    section_header "$(L "Cache RDP - Terminal Server Client" "RDP Cache - Terminal Server Client")" "$CYAN"
    check_win_root || return 1

    local RDP_REL="AppData/Local/Microsoft/Terminal Server Client/Cache"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_DIRS=() RES_FILES_LIST=()

    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local CACHE_DIR
        CACHE_DIR=$(ci_find_dir "$USER_DIR" "$RDP_REL")
        if [[ -z "$CACHE_DIR" || ! -d "$CACHE_DIR" ]]; then
            dim_msg "$USERNAME - $(L "Cache RDP non trovata" "RDP cache not found")"
            continue
        fi
        mapfile -t CACHE_FILES < <(find "$CACHE_DIR" -maxdepth 1 -type f \( -iname "*.bmc" -o -iname "*.bin" \) -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        if [[ ${#CACHE_FILES[@]} -eq 0 ]]; then
            mapfile -t CACHE_FILES < <(find "$CACHE_DIR" -maxdepth 1 -type f \( -iname "*.bmc" -o -iname "*.bin" \) -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
        fi
        local COUNT=${#CACHE_FILES[@]}
        if [[ $COUNT -eq 0 ]]; then
            warn "$USERNAME - $(L "Directory cache trovata ma vuota" "Cache directory found but empty")"
            continue
        fi
        ok "$USERNAME - $COUNT $(L "file cache trovati in:" "cache files found in:") $CACHE_DIR"
        local FILES_INFO=""
        for F in "${CACHE_FILES[@]}"; do
            local FNAME; FNAME=$(basename "$F")
            local FSIZE; FSIZE=$(stat -c "%s" "$F" 2>/dev/null || echo "?")
            local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
            echo -e "      ${DIM}• $FNAME  (${FSIZE} bytes - mod: $FMTIME)${RESET}"
            FILES_INFO="${FILES_INFO}${FNAME}:${FSIZE}:${FMTIME}|"
        done
        RES_USERS+=("$USERNAME")
        RES_DIRS+=("$CACHE_DIR")
        RES_FILES_LIST+=("${FILES_INFO%|}")
        TOTAL_FILES=$((TOTAL_FILES + COUNT))
        TOTAL_USERS=$((TOTAL_USERS + 1))
    done < <(get_user_homes)

    # Cerca anche server RDP nel registro
    echo ""
    info "$(L "Ricerca server RDP recenti nel registro (NTUSER.DAT)..." "Searching recent RDP servers in registry (NTUSER.DAT)...")"
    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local NTUSER
        NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        if check_regipy; then
            local RDP_SERVERS
            RDP_SERVERS=$("$PY3" - "$NTUSER" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    try:
        key = hive.get_key('Software\\Microsoft\\Terminal Server Client\\Servers')
        for sk in key.iter_subkeys():
            hostname = sk.name
            uname = ''
            for v in sk.get_values():
                if v.name.lower() == 'usernamehinint' or v.name.lower() == 'usernamehint':
                    uname = str(v.value)
            print(f"{hostname}\t{uname}")
    except: pass
    try:
        key2 = hive.get_key('Software\\Microsoft\\Terminal Server Client\\Default')
        for v in key2.get_values():
            if v.name.lower().startswith('mru'):
                print(f"MRU: {v.value}\t")
    except: pass
except: pass
PYEOF
)
            if [[ -n "$RDP_SERVERS" ]]; then
                echo -e "  ${GREEN}${BOLD}$USERNAME - $(L "Server RDP trovati:" "RDP servers found:")${RESET}"
                while IFS=$'\t' read -r HOST UNAME; do
                    printf "      ${CYAN}%-40s${RESET}  ${DIM}utente: %s${RESET}\n" "$HOST" "${UNAME:--}"
                done <<< "$RDP_SERVERS"
            fi
        fi
    done < <(get_user_homes)

    separator
    info "Utenti con cache RDP: ${BOLD}$TOTAL_USERS${RESET}  |  $(L "File totali:" "Total files:") ${BOLD}$TOTAL_FILES"

    if [[ $TOTAL_FILES -gt 0 ]]; then
        echo ""
        info "${BOLD}$(L "Per analizzare le tile bitmap della cache usa bmc-tools:" "To analyze cache bitmap tiles use bmc-tools:")${RESET}"
        echo -e "    ${DIM}git clone https://github.com/ANSSI-FR/bmc-tools${RESET}"
        echo -e "    ${DIM}${PY3} bmc-tools.py -s <dir_cache> -d ./output/ -b${RESET}"
    fi

    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun file cache RDP trovato." "No RDP cache files found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "rdp_cache")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""

    for i in "${!RES_USERS[@]}"; do
        local USER="${RES_USERS[$i]}"
        local DIR="${RES_DIRS[$i]}"
        local ROWS=""
        IFS='|' read -ra FENTRIES <<< "${RES_FILES_LIST[$i]}"
        # Ordina per FMTIME (k3, resto della riga) decrescente — FMTIME può contenere ':'
        mapfile -t FENTRIES < <(printf '%s\n' "${FENTRIES[@]}" | sort -t':' -k3r)
        for FE in "${FENTRIES[@]}"; do
            IFS=':' read -r FNAME FSIZE FMTIME <<< "$FE"
            ROWS+="<tr>
              <td class='mono'>$(html_esc "$FNAME")</td>
              <td class='mono mid' style='white-space:nowrap'>${FSIZE} B</td>
              <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
            </tr>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>RD</div>
            <div><div class='uname'>$(html_esc "$USER")</div><div class='upath'>$(html_esc "$DIR")</div></div>
            <div class='badge'>${#FENTRIES[@]} file</div>
          </div>
          <table><thead><tr><th>File</th><th>$(L "Dimensione" "Size")</th><th>$(L "Ultima modifica" "Last modified")</th></tr></thead>
          <tbody>${ROWS}</tbody></table>
        </div>"
    done

    {
        html_header "RDP Cache"
        html_page_header "RD" "RDP Cache <span>Forensics</span>" \
            "%LOCALAPPDATA%\\Microsoft\\Terminal Server Client\\Cache" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File cache</div><div class='value'>${TOTAL_FILES}</div></div>
          <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>
        </div>
        <main>
        <div class='stitle'>File cache per utente</div>
        <div class='cards'>${CARDS_HTML}</div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Analisi tile bitmap</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            git clone https://github.com/ANSSI-FR/bmc-tools<br>
            ${PY3} bmc-tools.py -s &lt;dir_cache&gt; -d ./output/ -b
          </div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
