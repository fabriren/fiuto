
# ================================================================
#  MODULO 11 — Services (Servizi Windows)
# ================================================================
module_services() {
    section_header "$(L "Services - Servizi Windows" "Services - Windows Services")" "$RED"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SYS
    HIVE_SYS=$(get_hive "SYSTEM")
    [[ -z "$HIVE_SYS" || ! -f "$HIVE_SYS" ]] && { err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found")"; return 1; }

    info "$(L "Parsing servizi da hive SYSTEM:" "Parsing services from SYSTEM hive:") $HIVE_SYS"
    echo ""

    local SVCDATA
    SVCDATA=$("$PY3" - "$HIVE_SYS" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])

    START_MAP = {0:'Boot',1:'System',2:'Automatic',3:'Manual',4:'Disabled'}
    TYPE_MAP  = {1:'Kernel Driver',2:'FS Driver',4:'Adapter',8:'Recognizer',
                 16:'Own Process',32:'Share Process',256:'Interactive'}

    found_cs = None
    for cs in ['ControlSet001','ControlSet002','CurrentControlSet']:
        try:
            hive.get_key(f'{cs}\\Services'); found_cs = cs; break
        except: continue

    if not found_cs:
        print("ERROR\tChiave Services non trovata", file=sys.stderr); sys.exit(1)

    for svc in hive.get_key(f'{found_cs}\\Services').iter_subkeys():
        name = svc.name
        vals = {v.name.lower(): v.value for v in svc.get_values()}
        start    = START_MAP.get(vals.get('start'),    str(vals.get('start','')))
        svc_type = TYPE_MAP.get(vals.get('type'),      str(vals.get('type','')))
        image    = str(vals.get('imagepath',''))
        display  = str(vals.get('displayname',''))
        desc     = str(vals.get('description',''))[:100]
        obj_name = str(vals.get('objectname',''))

        # Flag sospetto: ImagePath fuori dalle dir di sistema standard
        img_l = image.lower()
        susp = 0
        if image and not any(p in img_l for p in [
            '\\system32\\','\\syswow64\\','\\drivers\\',
            '%systemroot%','%windir%','system32','syswow64',
            'svchost','ntosknrl'
        ]):
            susp = 1
        # ObjectName non standard
        if obj_name and obj_name.lower() not in (
            'localsystem','localservice','networkservice',''):
            susp = 1

        print(f"{name}\t{display}\t{start}\t{svc_type}\t{image}\t{obj_name}\t{desc}\t{susp}")

except Exception as e:
    print(f"ERROR\t{e}", file=sys.stderr)
PYEOF
)

    local TOTAL=0 SUSP_COUNT=0
    declare -a ALL_SVCS=()

    printf "  ${BOLD}%-30s %-12s %-18s %s${RESET}\n" "Nome" "Start" "Tipo" "ImagePath"
    separator

    while IFS=$'\t' read -r SVC_NAME SVC_DISP SVC_START SVC_TYPE SVC_IMG SVC_OBJ SVC_DESC SVC_SUSP; do
        [[ -z "$SVC_NAME" ]] && continue
        ALL_SVCS+=("${SVC_NAME}§${SVC_DISP}§${SVC_START}§${SVC_TYPE}§${SVC_IMG}§${SVC_OBJ}§${SVC_DESC}§${SVC_SUSP}")
        TOTAL=$((TOTAL + 1))
        if [[ "$SVC_SUSP" == "1" ]]; then
            SUSP_COUNT=$((SUSP_COUNT + 1))
            printf "  ${RED}%-30s${RESET} ${YELLOW}%-12s${RESET} %-18s ${RED}%s${RESET}\n" \
                "$SVC_NAME" "$SVC_START" "$SVC_TYPE" "$SVC_IMG"
            [[ -n "$SVC_OBJ" && "${SVC_OBJ,,}" != "localsystem" ]] && \
                echo -e "      ${YELLOW}ObjectName: $SVC_OBJ${RESET}"
        else
            printf "  ${DIM}%-30s${RESET} %-12s %-18s %s\n" \
                "$SVC_NAME" "$SVC_START" "$SVC_TYPE" "$SVC_IMG"
        fi
    done <<< "$SVCDATA"

    separator
    info "Servizi totali: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun servizio trovato." "No services found.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "services")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Escape HTML — usata anche per Services
    _esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

    local ROWS=""
    for E in "${ALL_SVCS[@]}"; do
        IFS='§' read -r SVC_NAME SVC_DISP SVC_START SVC_TYPE SVC_IMG SVC_OBJ SVC_DESC SVC_SUSP <<< "$E"

        local ROW_STYLE="" START_CSS="mono mid"
        [[ "$SVC_SUSP" == "1" ]] && ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        case "$SVC_START" in
            Boot|System)   START_CSS="mono bad"  ;;
            Automatic)     START_CSS="mono warn" ;;
            Disabled)      START_CSS="mono dim"  ;;
        esac

        local IMG_CSS="mono fld"
        [[ "$SVC_SUSP" == "1" ]] && IMG_CSS="mono fld bad"

        local DETAILS=""
        [[ -n "$SVC_IMG"  ]] && DETAILS+="<div class='drow'><span class='lbl'>IMG</span><span class='${IMG_CSS}'>$(_esc "$SVC_IMG")</span></div>"
        if [[ -n "$SVC_OBJ" && "${SVC_OBJ,,}" != "localsystem" ]]; then
            DETAILS+="<div class='drow'><span class='lbl'>OBJ</span><span class='mono fld warn'>$(_esc "$SVC_OBJ")</span></div>"
        fi
        [[ -n "$SVC_DESC" ]] && DETAILS+="<div class='drow'><span class='lbl'>DSC</span><span class='fld' style='font-size:.72rem;color:var(--text-mid)'>$(_esc "$SVC_DESC")</span></div>"
        [[ -z "$DETAILS"  ]] && DETAILS="<span class='dim' style='font-size:.72rem'>-</span>"

        local DISP_HTML=""
        [[ -n "$SVC_DISP" && "$SVC_DISP" != "$SVC_NAME" ]] && \
            DISP_HTML="<div class='dim' style='font-size:.65rem;margin-top:.1rem'>$(_esc "$SVC_DISP")</div>"

        ROWS+="<tr ${ROW_STYLE}>
          <td class='tname' style='vertical-align:top;padding-top:.5rem'>
            <span class='mono'>$(_esc "$SVC_NAME")</span>${DISP_HTML}
          </td>
          <td style='vertical-align:top;padding:.5rem .8rem;white-space:nowrap'>
            <span class='${START_CSS}'>${SVC_START}</span>
          </td>
          <td class='mono dim' style='vertical-align:top;padding:.5rem .8rem;white-space:nowrap;font-size:.72rem'>${SVC_TYPE}</td>
          <td style='vertical-align:top;padding:.35rem .9rem'>${DETAILS}</td>
        </tr>"
    done

    {
        html_header "Services"
        html_page_header "SV" "Windows <span>Services</span> Forensics" \
            "SYSTEM\\ControlSet001\\Services" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Servizi trovati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div>
        <main>
        <style>
          .lbl{display:inline-block;font-family:var(--mono);font-size:.58rem;color:var(--text-dim);
               width:2.6rem;margin-right:.5rem;text-transform:uppercase;vertical-align:top;padding-top:.1rem;flex-shrink:0}
          .drow{display:flex;align-items:flex-start;margin-bottom:.22rem}
          .drow:last-child{margin-bottom:0}
          .fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap;flex:1}
          .tname{font-family:var(--mono);font-size:.72rem;color:var(--text-mid);word-break:break-all;
                 min-width:140px;max-width:240px;padding:.5rem .8rem .5rem 1rem}
        </style>
        <div class='stitle'>Servizi Windows - IMG · OBJ · DSC</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:18%'>Nome / Display</th>
            <th style='width:9%'>Start</th>
            <th style='width:13%'>$(L "Tipo" "Type")</th>
            <th>$(L "Dettagli" "Details")</th>
          </tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>ℹ Criteri di flagging sospetto</div>
          <div style='color:var(--text-mid);line-height:1.8'>
            • ImagePath fuori da System32 / SysWOW64 / drivers<br>
            • ObjectName diverso da LocalSystem / LocalService / NetworkService
          </div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
