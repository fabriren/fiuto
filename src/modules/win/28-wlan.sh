
# ================================================================
#  MODULO 28 — WLAN & VPN Profiles
# ================================================================
module_wlan() {
    section_header "$(L "WLAN & VPN - Profili di Rete" "WLAN & VPN - Network Profiles")" "$CYAN"
    check_win_root || return 1


    declare -a ALL_WLAN=()
    local TOTAL_WLAN=0

    # WLAN Profiles: ProgramData\Microsoft\Wlansvc\Profiles\Interfaces\*\*.xml
    local WLANSVC_DIR
    WLANSVC_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Wlansvc/Profiles/Interfaces")
    if [[ -n "$WLANSVC_DIR" && -d "$WLANSVC_DIR" ]]; then
        info "WLAN Profiles: $WLANSVC_DIR"
        while IFS= read -r WXML; do
            [[ -f "$WXML" ]] || continue
            local WOUT
            WOUT=$("$PY3" - "$WXML" << 'PYEOF' 2>/dev/null || true
import sys, xml.etree.ElementTree as ET
try:
    tree = ET.parse(sys.argv[1])
    root = tree.getroot()
    ns = {'w': 'http://www.microsoft.com/networking/WLAN/profile/v1'}
    def gt(tag):
        el = root.find(f'.//{{http://www.microsoft.com/networking/WLAN/profile/v1}}{tag}')
        return el.text.strip() if el is not None and el.text else ''
    ssid   = gt('name') or gt('SSID') or gt('SSIDConfig')
    auth   = gt('authentication')
    enc    = gt('encryption')
    key    = gt('keyMaterial')  # WEP/WPA PSK in chiaro
    print(f"{ssid}\t{auth}\t{enc}\t{key}")
except Exception as e:
    print(f"ERROR\t{e}\t\t")
PYEOF
)
            IFS=$'\t' read -r SSID AUTH ENC KEYMAT <<< "$WOUT"
            [[ "$SSID" == "ERROR" || -z "$SSID" ]] && continue
            local HAS_KEY=""; [[ -n "$KEYMAT" ]] && HAS_KEY="${RED}[KEY IN CHIARO]${RESET}"
            printf "  ${CYAN}%-35s${RESET}  %-15s  %-10s  %b\n" "$SSID" "$AUTH" "$ENC" "$HAS_KEY"
            ALL_WLAN+=("WLAN§${SSID}§${AUTH}§${ENC}§${KEYMAT}§${WXML}")
            TOTAL_WLAN=$((TOTAL_WLAN + 1))
        done < <(find "$WLANSVC_DIR" -name "*.xml" -type f 2>/dev/null)
    else
        warn "$(L "WLAN Profiles non trovati (ProgramData/Microsoft/Wlansvc/Profiles)" "WLAN Profiles not found (ProgramData/Microsoft/Wlansvc/Profiles)")"
    fi

    # VPN / Network profiles dal registro SOFTWARE
    echo ""
    info "$(L "Profili di rete (NetworkList) dal registro..." "Network profiles (NetworkList) from registry...")"
    if check_regipy; then
        local SW_HIVE; SW_HIVE=$(get_hive "SOFTWARE")
        if [[ -n "$SW_HIVE" ]]; then
            local NET_DATA
            NET_DATA=$("$PY3" - "$SW_HIVE" << 'PYEOF' 2>/dev/null || true
import sys, datetime
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for key_path in ['Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles',
                     'Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Signatures\\Unmanaged',
                     'Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Signatures\\Managed']:
        try:
            key = hive.get_key(key_path)
            for sk in key.iter_subkeys():
                vals = {v.name.lower(): v.value for v in sk.get_values()}
                name = str(vals.get('profilename', vals.get('firstname', sk.name)))
                desc = str(vals.get('description',''))
                cat  = str(vals.get('category',''))
                gw   = str(vals.get('defaultgatewayipv4', vals.get('defaultgatewayipv6','')))
                dns  = str(vals.get('dnsserver',''))
                print(f"NET\t{name}\t{desc}\t{cat}\t{gw}\t{dns}")
        except: continue
except Exception as e:
    pass
PYEOF
)
            while IFS=$'\t' read -r TYPE NAME DESC CAT GW DNS; do
                [[ -z "$NAME" ]] && continue
                printf "  ${GREEN}%-40s${RESET}  ${DIM}cat:%s  gw:%s${RESET}\n" "$NAME" "$CAT" "$GW"
                ALL_WLAN+=("VPN§${NAME}§${CAT}§${GW}§${DNS}§-")
                TOTAL_WLAN=$((TOTAL_WLAN + 1))
            done <<< "$NET_DATA"
        fi
    fi

    separator
    info "$(L "Profili di rete totali:" "Total network profiles:") ${BOLD}$TOTAL_WLAN"
    [[ $TOTAL_WLAN -eq 0 ]] && { warn "$(L "Nessun profilo trovato." "No profiles found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wlan")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_WLAN[@]}"; do
        IFS='§' read -r TYPE NAME F2 F3 F4 F5 <<< "$E"
        local CSS=""; [[ "$TYPE" == "WLAN" && -n "$F4" ]] && CSS="style='background:rgba(255,123,114,.07)'"
        local KEYHTML=""; [[ "$TYPE" == "WLAN" && -n "$F4" ]] && KEYHTML="<span class='bad' style='font-size:.68rem'>KEY: $(html_esc "${F4:0:40}")</span>"
        ROWS+="<tr ${CSS}>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$TYPE")</td>
          <td class='mono'>$(html_esc "$NAME") ${KEYHTML}</td>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$F2")</td>
          <td class='mono mid' style='white-space:nowrap'>$(html_esc "$F3")</td>
        </tr>"
    done
    {
        html_header "WLAN VPN"
        html_page_header "WL" "WLAN / VPN - <span>Profili di Rete</span>" \
            "ProgramData\\Microsoft\\Wlansvc + NetworkList" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Profili totali</div><div class='value'>${TOTAL_WLAN}</div></div>
        </div><main>
        <div class='stitle'>Reti WiFi e profili VPN/network memorizzati</div>
        <div class='card'><table>
          <thead><tr><th style='width:8%'>$(L "Tipo" "Type")</th><th>Nome / SSID</th><th style='width:12%'>Auth/Cat</th><th style='width:15%'>Gateway/Enc</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
