
# ================================================================
#  MODULO 32 — Office MRU & Recent Docs
# ================================================================
module_office_mru() {
    section_header "$(L "Office MRU - File Recenti Office" "Office MRU - Recent Office Files")" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"; return 1
    fi


    declare -a ALL_MRU=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        info "Parsing Office MRU: $USERNAME"

        local MRU_DATA
        MRU_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    APPS = ['Word', 'Excel', 'PowerPoint', 'Access', 'Publisher', 'Visio', 'Project', 'OneNote']
    # Office versioni comuni: 14.0 (2010), 15.0 (2013), 16.0 (2016/2019/365)
    VERSIONS = ['14.0', '15.0', '16.0']
    for ver in VERSIONS:
        for app in APPS:
            for key_path in [
                f'Software\\Microsoft\\Office\\{ver}\\{app}\\User MRU',
                f'Software\\Microsoft\\Office\\{ver}\\{app}\\File MRU',
            ]:
                try:
                    root = hive.get_key(key_path)
                    # Cerca sub-chiavi AD_* (per User MRU) o valori diretti
                    try:
                        for sk in root.iter_subkeys():
                            if not sk.name.startswith('AD_'): continue
                            for mru_key_path in [
                                f'{key_path}\\{sk.name}\\File MRU',
                                f'{key_path}\\{sk.name}',
                            ]:
                                try:
                                    mru_key = hive.get_key(mru_key_path)
                                    for v in mru_key.get_values():
                                        if v.name.lower() in ('mruliste','mrulist','max display'): continue
                                        val = str(v.value)
                                        if '\\' in val or '/' in val:
                                            print(f"{user}\t{app}\t{ver}\t{val[:300]}")
                                except: pass
                    except:
                        for v in root.get_values():
                            if v.name.lower() in ('mrulist','mrulistex','max display'): continue
                            val = str(v.value)
                            if '\\' in val or '/' in val:
                                print(f"{user}\t{app}\t{ver}\t{val[:300]}")
                except: continue
except Exception as e:
    pass
PYEOF
)
        while IFS=$'\t' read -r U APP VER FPATH; do
            [[ -z "$FPATH" ]] && continue
            local IS_SUSP="false"
            echo "$FPATH" | grep -qi "temp\|appdata\|public\|programdata\|\\\\temp\\\\" && IS_SUSP="true"
            printf "  ${DIM}[%-15s]${RESET}  ${GREEN}%-12s${RESET}  [%s]  %s\n" "$U" "$APP" "$VER" "${FPATH:0:80}"
            ALL_MRU+=("${U}§${APP}§${VER}§${FPATH}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
        done <<< "$MRU_DATA"
    done < <(get_user_homes)

    separator
    info "File Office recenti: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun file Office recente trovato." "No recent Office files found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "office_mru")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_MRU[@]}"; do
        IFS='§' read -r U APP VER FPATH IS_SUSP <<< "$E"
        local RS="" NC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && NC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$APP")</td>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$VER")</td>
          <td class='${NC}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$FPATH")</td>
        </tr>"
    done
    {
        html_header "Office MRU"
        html_page_header "OF" "Office MRU - <span>File Recenti</span>" \
            "NTUSER.DAT → Software\\Microsoft\\Office" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File recenti</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>File aperti con Word · Excel · PowerPoint · Access · OneNote</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:12%'>App</th><th style='width:7%'>$(L "Versione" "Version")</th><th>Path file</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
