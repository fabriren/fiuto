
# ================================================================
#  MODULO 34 — PowerShell Script Block Logging (EID 4104)
# ================================================================
module_ps_scriptblock() {
    section_header "PowerShell Script Block Logging — EID 4104" "$MAGENTA"
    check_win_root || return 1

    local EVTX_DIR
    EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    if [[ -z "$EVTX_DIR" ]]; then
        warn "$(L "Directory EVTX non trovata" "EVTX directory not found")"
        return 1
    fi

    local PS_EVTX
    PS_EVTX=$(find "$EVTX_DIR" -maxdepth 1 -iname "Microsoft-Windows-PowerShell%4Operational.evtx" -type f 2>/dev/null | head -1)
    if [[ -z "$PS_EVTX" || ! -f "$PS_EVTX" ]]; then
        warn "$(L "Log PowerShell Operational non trovato" "PowerShell Operational log not found")"
        info "$(L "Atteso: Windows/System32/winevt/Logs/Microsoft-Windows-PowerShell%4Operational.evtx" "Expected: Windows/System32/winevt/Logs/Microsoft-Windows-PowerShell%4Operational.evtx")"
        return 0
    fi

    info "Parsing: $PS_EVTX"

    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        warn "$(L "python-evtx non trovato." "python-evtx not found.") Installalo con: ${PY3} -m pip install python-evtx"
        return 1
    fi

    local BLOCK_DATA
    BLOCK_DATA=$("$PY3" - "$PS_EVTX" << 'PYEOF' 2>/dev/null || true
import sys, xml.etree.ElementTree as ET
try:
    from Evtx.Evtx import Evtx
    SENS = ['invoke-expression','iex','downloadstring','encodedcommand','-enc','bypass',
            'frombase64string','webclient','net user','net localgroup','mimikatz',
            'password','token','secret','hidden','noprofile','windowstyle']

    # Accumula frammenti per ScriptBlockId (EID 4104 suddivide script grandi
    # in N eventi consecutivi con MessageNumber/MessageTotal)
    blocks = {}  # block_id -> {'total':N, 'parts':{num:text}, 'ts':'', 'path':''}

    with Evtx(sys.argv[1]) as log:
        for record in log.records():
            try:
                xml_str = record.xml()
                root = ET.fromstring(xml_str)
                ns = {'e': 'http://schemas.microsoft.com/win/2004/08/events/event'}
                eid_el = root.find('.//e:EventID', ns)
                if eid_el is None or eid_el.text != '4104': continue

                vals = {}
                for ed in root.findall('.//e:EventData/e:Data', ns):
                    vals[ed.get('Name', '')] = ed.text or ''

                sbt = vals.get('ScriptBlockText', '')
                if not sbt: continue

                block_id = vals.get('ScriptBlockId', '') or id(vals)
                msg_num  = int(vals.get('MessageNumber', '1') or '1')
                msg_tot  = int(vals.get('MessageTotal',  '1') or '1')
                path     = vals.get('Path', '')

                ts_el = root.find('.//e:TimeCreated', ns)
                ts = ts_el.get('SystemTime', '') if ts_el is not None else ''

                if block_id not in blocks:
                    blocks[block_id] = {'total': msg_tot, 'parts': {}, 'ts': ts, 'path': path}
                blocks[block_id]['parts'][msg_num] = sbt
                if ts and not blocks[block_id]['ts']:
                    blocks[block_id]['ts'] = ts
            except: continue

    # Riassembla ogni script e stampa come singola riga:
    # TYPE \t TIMESTAMP \t PATH \t SCRIPT_FLAT
    # I newline nello script sono codificati come \x1f (unit separator)
    # per non rompere il parsing bash riga-per-riga.
    for info in blocks.values():
        parts = info['parts']
        total = info['total']
        full  = ''.join(parts.get(i, '') for i in range(1, total + 1))
        if not full:
            full = ''.join(parts[k] for k in sorted(parts))
        ts    = info['ts']
        path  = info['path']
        is_sens = any(k in full.lower() for k in SENS)
        flat  = full.replace('\r\n', '\x1f').replace('\n', '\x1f').replace('\r', '\x1f')
        sys.stdout.buffer.write(
            f"{'SENS' if is_sens else 'NORM'}\x1e{ts}\x1e{path}\x1e{flat[:1000000]}\n".encode('utf-8', errors='replace')
        )

except Exception as e:
    sys.stderr.write(f"ERROR\x1e\x1e\x1e{e}\n")
PYEOF
)

    local TOTAL=0 SUSP=0
    declare -a BLOCKS=()
    while IFS=$'\x1e' read -r TYPE TS FPATH SCRIPT; do
        [[ "$TYPE" == "ERROR" ]] && { err "$(L "Errore parser:" "Parser error:") $SCRIPT"; break; }
        # Salta righe orfane (non devono esistere dopo il fix, ma per sicurezza)
        [[ -z "$TS" && -z "$FPATH" ]] && continue
        TOTAL=$((TOTAL+1))
        [[ "$TYPE" == "SENS" ]] && SUSP=$((SUSP+1))
        # Ripristina i newline (\x1f → \n) per il display terminale
        local DISP; DISP="${SCRIPT:0:120}"
        DISP="${DISP//$'\x1f'/$'\n'      }"
        if [[ "$TYPE" == "SENS" ]]; then
            printf "  ${RED}[!]${RESET} ${DIM}%s${RESET}  ${DIM}%s${RESET}\n  ${RED}%s${RESET}\n\n" "$TS" "$FPATH" "$DISP"
        else
            printf "  ${DIM}[i] %s  %s${RESET}\n" "$TS" "$DISP"
        fi
        BLOCKS+=("${TYPE}§${TS}§${FPATH}§${SCRIPT}")
    done <<< "$BLOCK_DATA"

    separator
    info "Script block unici: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun EID 4104 trovato (Script Block Logging potrebbe non essere abilitato)." "No EID 4104 found (Script Block Logging may not be enabled).")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ps_scriptblock")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local HIVE_HASH; HIVE_HASH=$(sha256_file "$PS_EVTX")
    # Ordina per timestamp decrescente (campo 2: TS)
    mapfile -t BLOCKS < <(printf '%s\n' "${BLOCKS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+2*[""])[1],reverse=True); print("\n".join(L))')
    local ROWS=""
    for E in "${BLOCKS[@]}"; do
        IFS='§' read -r TYPE TS FPATH SCRIPT <<< "$E"
        local RS="" TC="mono"
        [[ "$TYPE" == "SENS" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && TC="mono bad"
        # Ripristina i newline (\x1f → LF) e applica HTML escape
        local SCRIPT_ESC
        SCRIPT_ESC=$(printf '%s' "${SCRIPT:0:1000000}" \
            | sed $'s/\x1f/\\n/g' \
            | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g')
        local PATH_ESC; PATH_ESC=$(html_esc "$FPATH")
        local PATH_ATTR; PATH_ATTR=$(html_attr "$FPATH")
        ROWS+="<tr ${RS}>
          <td class='mono ok' style='white-space:nowrap;font-size:.68rem'>${TS}</td>
          <td class='mono dim' style='white-space:nowrap;font-size:.65rem;max-width:220px;overflow:hidden;text-overflow:ellipsis' title='${PATH_ATTR}'>${PATH_ESC}</td>
          <td class='${TC}' style='word-break:break-all;font-size:.72rem;white-space:pre-wrap'>${SCRIPT_ESC}</td>
        </tr>"
    done
    {
        html_header "PS ScriptBlock Log"
        html_page_header "PS4" "PS Script Block Logging — <span>EID 4104</span>" \
            "Microsoft-Windows-PowerShell%4Operational.evtx" "$SCAN" "$WIN_ROOT" "$HIVE_HASH"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Script Block unici</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP}</div></div>
        </div><main>
        <div class='stitle'>Event ID 4104 — Frammenti riassemblati per ScriptBlockId — ScriptBlockText (max 1000000 char)</div>
        <div class='card'><table>
          <thead><tr><th style='width:16%'>Timestamp</th><th style='width:18%'>Path script</th><th>ScriptBlockText</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
