
# ================================================================
#  MODULO 29 — AppX / UWP Packages
# ================================================================
module_appx() {
    section_header "$(L "AppX / UWP - Pacchetti Store Installati" "AppX / UWP - Installed Store Packages")" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"; return 1
    fi


    declare -a ALL_PKGS=()
    local TOTAL=0 SUSP_COUNT=0

    # SOFTWARE hive (sistema)
    local SW_HIVE; SW_HIVE=$(get_hive "SOFTWARE")
    if [[ -n "$SW_HIVE" ]]; then
        info "$(L "Parsing AppX packages da SOFTWARE hive..." "Parsing AppX packages from SOFTWARE hive...")"
        local PKG_DATA
        PKG_DATA=$("$PY3" - "$SW_HIVE" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for key_path in [
        'Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages',
        'Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages',
    ]:
        try:
            root = hive.get_key(key_path)
            for sk in root.iter_subkeys():
                vals = {v.name.lower(): str(v.value) for v in sk.get_values()}
                name = sk.name
                path = vals.get('packagerootfolder', vals.get('installlocation',''))
                pub  = vals.get('packageid.publisher', vals.get('publisher',''))
                ver  = vals.get('packageid.version','')
                arch = vals.get('packageid.processorarchitecture','')
                print(f"SYS\t{name}\t{path}\t{pub}\t{ver}\t{arch}")
        except: continue
except: pass
PYEOF
)
        while IFS=$'\t' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER PKG_ARCH; do
            [[ -z "$PKG_NAME" ]] && continue
            local IS_SUSP="false"
            echo "$PKG_PATH" | grep -qiv "program files\|windows apps\|system32" && \
                [[ -n "$PKG_PATH" ]] && IS_SUSP="true"
            [[ "${PKG_PUB,,}" == *"cn="* ]] || echo "$PKG_NAME" | grep -qi "sideload\|test\|debug" && IS_SUSP="true"
            ALL_PKGS+=("${SCOPE}§${PKG_NAME}§${PKG_PATH}§${PKG_PUB}§${PKG_VER}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
            if [[ "$IS_SUSP" == "true" ]]; then
                printf "  ${RED}[!] %-50s${RESET}  %s\n" "$PKG_NAME" "$PKG_PATH"
            else
                printf "  ${DIM}%-50s${RESET}\n" "${PKG_NAME:0:60}"
            fi
        done <<< "$PKG_DATA"
    fi

    # Per utente: NTUSER.DAT
    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        local UPKG_DATA
        UPKG_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    for key_path in ['Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages']:
        try:
            root = hive.get_key(key_path)
            for sk in root.iter_subkeys():
                vals = {v.name.lower(): str(v.value) for v in sk.get_values()}
                path = vals.get('packagerootfolder','')
                pub  = vals.get('publisher','')
                print(f"USR:{user}\t{sk.name}\t{path}\t{pub}\t\t")
        except: continue
except: pass
PYEOF
)
        while IFS=$'\t' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER PKG_ARCH; do
            [[ -z "$PKG_NAME" ]] && continue
            local IS_SUSP="false"
            echo "$PKG_PATH" | grep -qiv "program files\|windows apps" && [[ -n "$PKG_PATH" ]] && IS_SUSP="true"
            ALL_PKGS+=("${SCOPE}§${PKG_NAME}§${PKG_PATH}§${PKG_PUB}§${PKG_VER}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
        done <<< "$UPKG_DATA"
    done < <(get_user_homes)

    separator
    info "Pacchetti AppX: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun pacchetto AppX trovato." "No AppX packages found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "appx")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_PKGS[@]}"; do
        IFS='§' read -r SCOPE PKG_NAME PKG_PATH PKG_PUB PKG_VER IS_SUSP <<< "$E"
        local RS=""; [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        local NC="mono"; [[ "$IS_SUSP" == "true" ]] && NC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap;font-size:.68rem'>$(html_esc "$SCOPE")</td>
          <td class='${NC}' style='word-break:break-all;font-size:.72rem'>$(html_esc "${PKG_NAME:0:80}")</td>
          <td class='mono dim' style='word-break:break-all;font-size:.68rem'>$(html_esc "$PKG_PUB")</td>
          <td class='mono mid' style='word-break:break-all;font-size:.68rem'>$(html_esc "$PKG_PATH")</td>
        </tr>"
    done
    {
        html_header "AppX"
        html_page_header "AX" "AppX / UWP - <span>Pacchetti Store</span>" \
            "AppModel\\Repository\\Packages" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Pacchetti</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Pacchetti UWP installati - path fuori Windows Apps = potenziale sideload</div>
        <div class='card'><table>
          <thead><tr><th style='width:10%'>Scope</th><th>Package Name</th><th style='width:20%'>Publisher</th><th>Install Path</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
