
# ================================================================
#  MODULO 9 — USB / Dispositivi Rimovibili
# ================================================================
module_usb() {
    section_header "$(L "USB — Dispositivi Rimovibili" "USB — Removable Devices")" "$BLUE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SYS
    HIVE_SYS=$(get_hive "SYSTEM")
    [[ -z "$HIVE_SYS" ]] && { err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found")"; return 1; }

    info "$(L "Parsing dispositivi USB da hive SYSTEM..." "Parsing USB devices from SYSTEM hive...")"

    local USB_DATA
    USB_DATA=$("$PY3" - "$HIVE_SYS" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    for cs in ['ControlSet001', 'ControlSet002']:
        try:
            usbstor = hive.get_key(f'{cs}\\Enum\\USBSTOR')
            for dtype in usbstor.iter_subkeys():
                for instance in dtype.iter_subkeys():
                    friendly = ''
                    serial = instance.name
                    for v in instance.get_values():
                        if v.name.lower() == 'friendlyname': friendly = str(v.value)
                    print(f"{dtype.name}\t{serial}\t{friendly}")
            break
        except: continue
except Exception as e:
    print(f"ERROR\t{e}\t")
PYEOF
)

    local TOTAL=0
    declare -a USB_ROWS=()
    echo ""
    printf "  ${BOLD}%-50s %-30s %s${RESET}\n" "Tipo dispositivo" "Seriale" "Nome"
    separator
    while IFS=$'\t' read -r DTYPE SERIAL FNAME; do
        [[ -z "$DTYPE" ]] && continue
        if [[ "$DTYPE" == "ERROR" ]]; then
            err "$SERIAL"
            continue
        fi
        printf "  ${CYAN}%-50s${RESET} ${DIM}%-30s${RESET} %s\n" "$DTYPE" "$SERIAL" "${FNAME:--}"
        USB_ROWS+=("${DTYPE}§${SERIAL}§${FNAME}")
        TOTAL=$((TOTAL + 1))
    done <<< "$USB_DATA"

    separator
    info "$(L "Dispositivi USB storici:" "Historical USB devices:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun dispositivo USB trovato nell'hive." "No USB device found in hive.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "usb")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for USB_E in "${USB_ROWS[@]}"; do
        IFS='§' read -r DTYPE SERIAL FNAME <<< "$USB_E"
        ROWS+="<tr>
          <td class='mono'>$(html_esc "$DTYPE")</td>
          <td class='mono dim'>$(html_esc "$SERIAL")</td>
          <td class='mono'>$([[ -n "$FNAME" ]] && html_esc "$FNAME" || echo "-")</td>
        </tr>"
    done

    {
        html_header "USB Devices"
        html_page_header "USB" "USB <span>Device History</span>" \
            "SYSTEM\\ControlSet001\\Enum\\USBSTOR" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Dispositivi</div><div class='value'>${TOTAL}</div></div>
        </div>
        <main><div class='stitle'>Dispositivi USB storici</div>
        <div class='card'><table>
          <thead><tr><th>$(L "Tipo" "Type")</th><th>$(L "Seriale" "Serial")</th><th>$(L "Nome" "Name")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
