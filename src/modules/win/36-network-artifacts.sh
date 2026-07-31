
# ================================================================
#  MODULO 36 — Network Artifacts (DNS Cache, Interfacce, NetworkList)
# ================================================================
module_network_artifacts() {
    section_header "$(L "Network Artifacts - DNS · Interfacce · NetworkList" "Network Artifacts - DNS · Interfaces · NetworkList")" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "regipy non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"; return 1
    fi

    local SYS_HIVE; SYS_HIVE=$(get_hive "SOFTWARE")
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "$(L "Hive SOFTWARE non trovato" "SOFTWARE hive not found")"; return 1
    fi

    info "$(L "Parsing NetworkList da hive SOFTWARE:" "Parsing NetworkList from SOFTWARE hive:") $SYS_HIVE"

    local NET_DATA
    NET_DATA=$("$PY3" - "$SYS_HIVE" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    try:
        nl = hive.get_key('Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles')
        for sk in nl.iter_subkeys():
            vals = {v.name: str(v.value) for v in sk.get_values()}
            name       = vals.get('ProfileName', '')
            desc       = vals.get('Description', '')
            category   = vals.get('Category', '')
            date_first = vals.get('DateFirstConnected', '')
            date_last  = vals.get('DateLastConnected', '')
            cat_str = {'0':'Public','1':'Private','2':'Domain'}.get(str(category), str(category))
            print(f"PROFILE\t{name}\t{desc}\t{cat_str}\t{date_first}\t{date_last}")
    except: pass
    try:
        sigs = hive.get_key('Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Signatures\\Unmanaged')
        for sk in sigs.iter_subkeys():
            vals = {v.name: str(v.value) for v in sk.get_values()}
            desc     = vals.get('Description', '')
            dns      = vals.get('DnsSuffix', '')
            mac      = vals.get('DefaultGatewayMac', '')
            ssid     = vals.get('FirstNetwork', '')
            print(f"SIGNATURE\t{desc}\t{dns}\t{mac}\t{ssid}\t")
    except: pass
except Exception as e:
    print(f"ERROR\t{e}", file=sys.stderr)
PYEOF
)

    local SYS_HIVE2; SYS_HIVE2=$(get_hive "SYSTEM")
    local IFACE_DATA=""
    if [[ -n "$SYS_HIVE2" && -f "$SYS_HIVE2" ]]; then
        info "$(L "Parsing interfacce TCP/IP da SYSTEM hive:" "Parsing TCP/IP interfaces from SYSTEM hive:") $SYS_HIVE2"
        IFACE_DATA=$("$PY3" - "$SYS_HIVE2" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for cs in ['ControlSet001','ControlSet002','CurrentControlSet']:
        try:
            ifaces = hive.get_key(f'{cs}\\Services\\Tcpip\\Parameters\\Interfaces')
            for sk in ifaces.iter_subkeys():
                vals = {v.name: str(v.value) for v in sk.get_values()}
                ip    = vals.get('IPAddress','') or vals.get('DhcpIPAddress','')
                gw    = vals.get('DefaultGateway','') or vals.get('DhcpDefaultGateway','')
                dns   = vals.get('NameServer','') or vals.get('DhcpNameServer','')
                dhcp  = vals.get('EnableDHCP','')
                if ip and ip not in ('0.0.0.0',''):
                    print(f"IFACE\t{cs}\t{sk.name}\t{ip}\t{gw}\t{dns}\t{dhcp}")
            break
        except: continue
except Exception as e:
    pass
PYEOF
)
    fi

    declare -a PROFILES=() SIGNATURES=() IFACES=()
    local TOTAL_P=0 TOTAL_S=0 TOTAL_I=0

    while IFS=$'\t' read -r TYPE F1 F2 F3 F4 F5; do
        case "$TYPE" in
            PROFILE)
                printf "  ${GREEN}[P]${RESET}  ${BOLD}%-30s${RESET}  ${DIM}%s${RESET}  [%s]\n" "$F1" "$F3" "$F4"
                PROFILES+=("${F1}§${F2}§${F3}§${F4}§${F5}")
                TOTAL_P=$((TOTAL_P+1))
                ;;
            SIGNATURE)
                printf "  ${BLUE}[S]${RESET}  %-30s  DNS:${DIM}%s${RESET}  MAC:${DIM}%s${RESET}\n" "$F1" "$F2" "$F3"
                SIGNATURES+=("${F1}§${F2}§${F3}§${F4}")
                TOTAL_S=$((TOTAL_S+1))
                ;;
        esac
    done <<< "$NET_DATA"

    while IFS=$'\t' read -r TYPE CS IFACE IP GW DNS DHCP; do
        [[ "$TYPE" != "IFACE" ]] && continue
        printf "  ${CYAN}[I]${RESET}  ${DIM}%s${RESET}  IP:${BOLD}%s${RESET}  GW:%s  DNS:%s\n" "$IFACE" "$IP" "$GW" "$DNS"
        IFACES+=("${CS}§${IFACE}§${IP}§${GW}§${DNS}§${DHCP}")
        TOTAL_I=$((TOTAL_I+1))
    done <<< "$IFACE_DATA"

    separator
    info "$(L "Profili rete:" "Network profiles:") ${BOLD}$TOTAL_P${RESET}  |  Firme: ${BOLD}$TOTAL_S${RESET}  |  Interfacce TCP/IP: ${BOLD}$TOTAL_I"
    [[ $((TOTAL_P + TOTAL_S + TOTAL_I)) -eq 0 ]] && { warn "$(L "Nessun dato di rete trovato." "No network data found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "network_artifacts")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local HIVE_HASH; HIVE_HASH=$(sha256_file "$SYS_HIVE")

    # Ordina profili rete per ultima connessione decrescente (campo 5: DLAST)
    mapfile -t PROFILES < <(printf '%s\n' "${PROFILES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+5*[""])[4],reverse=True); print("\n".join(L))')
    local ROWS_P="" ROWS_S="" ROWS_I=""
    for E in "${PROFILES[@]}"; do
        IFS='§' read -r NAME DESC CAT DFIRST DLAST <<< "$E"
        local CAT_COLOR="ok"
        [[ "$CAT" == "Public" ]] && CAT_COLOR="warn"
        ROWS_P+="<tr><td class='mono'>$(html_esc "$NAME")</td><td class='mono dim'>$(html_esc "$DESC")</td>
          <td class='mono ${CAT_COLOR}'>$(html_esc "$CAT")</td>
          <td class='mono dim' style='font-size:.68rem'>${DFIRST}</td>
          <td class='mono ok'  style='font-size:.68rem'>${DLAST}</td></tr>"
    done
    for E in "${SIGNATURES[@]}"; do
        IFS='§' read -r NAME DNS_S MAC SSID <<< "$E"
        ROWS_S+="<tr><td class='mono'>$(html_esc "$NAME")</td><td class='mono dim'>$(html_esc "$DNS_S")</td>
          <td class='mono warn'>$(html_esc "$MAC")</td><td class='mono'>$(html_esc "$SSID")</td></tr>"
    done
    for E in "${IFACES[@]}"; do
        IFS='§' read -r CS IFACE IP GW DNS_I DHCP <<< "$E"
        ROWS_I+="<tr><td class='mono dim'>$(html_esc "$CS")</td><td class='mono dim' style='font-size:.65rem'>$(html_esc "$IFACE")</td>
          <td class='mono ok'>$(html_esc "$IP")</td><td class='mono'>$(html_esc "$GW")</td>
          <td class='mono dim'>$(html_esc "$DNS_I")</td><td class='mono'>$(html_esc "$DHCP")</td></tr>"
    done
    {
        html_header "Network Artifacts"
        html_page_header "NET" "Network Artifacts - <span>Profili · Firme · Interfacce</span>" \
            "SOFTWARE · SYSTEM hive" "$SCAN" "$WIN_ROOT" "$HIVE_HASH"
        echo "<div class='statsbar'>
          <div class='stat info'><div class='label'>Profili rete</div><div class='value'>${TOTAL_P}</div></div>
          <div class='stat'><div class='label'>Firme</div><div class='value'>${TOTAL_S}</div></div>
          <div class='stat ok'><div class='label'>Interfacce TCP/IP</div><div class='value'>${TOTAL_I}</div></div>
        </div><main>"
        if [[ -n "$ROWS_P" ]]; then
            echo "<div class='stitle'>Profili di rete noti (NetworkList\\Profiles)</div>
            <div class='card'><table>
              <thead><tr><th>$(L "Nome" "Name")</th><th>$(L "Descrizione" "Description")</th><th>$(L "Categoria" "Category")</th><th>$(L "Prima connessione" "First connection")</th><th>$(L "Ultima connessione" "Last connection")</th></tr></thead>
              <tbody>${ROWS_P}</tbody></table></div>"
        fi
        if [[ -n "$ROWS_S" ]]; then
            echo "<div class='stitle' style='margin-top:2rem'>Firme di rete (Signatures\\Unmanaged)</div>
            <div class='card'><table>
              <thead><tr><th>$(L "Descrizione" "Description")</th><th>DNS Suffix</th><th>MAC Gateway</th><th>SSID</th></tr></thead>
              <tbody>${ROWS_S}</tbody></table></div>"
        fi
        if [[ -n "$ROWS_I" ]]; then
            echo "<div class='stitle' style='margin-top:2rem'>Interfacce TCP/IP (SYSTEM\\Tcpip\\Parameters)</div>
            <div class='card'><table>
              <thead><tr><th>ControlSet</th><th>$(L "GUID interfaccia" "Interface GUID")</th><th>IP</th><th>Gateway</th><th>DNS</th><th>DHCP</th></tr></thead>
              <tbody>${ROWS_I}</tbody></table></div>"
        fi
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
