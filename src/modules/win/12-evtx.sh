
# ================================================================
#  MODULO 0 — Esegui tutti i moduli
# ================================================================
# ================================================================
#  MODULO 12 — Event Log (Security / System / PowerShell / RDP)
# ================================================================
module_evtx() {
    section_header "$(L "Event Log - Analisi .evtx" "Event Log - .evtx Analysis")" "$RED"
    check_win_root || return 1

    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        warn "$(L "python-evtx non trovato." "python-evtx not found.")"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install python-evtx"
        return 1
    fi

    local EVTX_DIR
    EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    if [[ -z "$EVTX_DIR" || ! -d "$EVTX_DIR" ]]; then
        err "$(L "Directory Event Log non trovata" "Event Log directory not found")"; return 1
    fi
    info "$(L "Directory log:" "Log directory:") $EVTX_DIR"
    echo ""

    declare -A EVTX_MAP=(
        ["Security.evtx"]="4624 4625 4648 4688 4698 4702 4720 4726 4732 4733 4756 4757"
        ["System.evtx"]="7045 7034 7036"
        ["Microsoft-Windows-PowerShell%4Operational.evtx"]="4104"
        ["Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx"]="21 23 24 25"
        ["Microsoft-Windows-TaskScheduler%4Operational.evtx"]="106 140 141 200 201"
        ["Microsoft-Windows-Windows Defender%4Operational.evtx"]="1116 1117 1118 1119"
    )

    # Parser: legge .evtx, filtra per EID, scrive JSON su file (non su argv/stdout grande)
    local PARSER_EVTX=""
    PARSER_EVTX=$(mktemp /tmp/dfir_evtxp_XXXXXX.py)
    cat > "$PARSER_EVTX" << 'PYEOF'
import sys, json, re
try:
    import Evtx.Evtx as evtx
except ImportError:
    json.dump([], open(sys.argv[3],'w')); sys.exit(1)

path, eids_csv, outf = sys.argv[1], sys.argv[2], sys.argv[3]
target = set(eids_csv.split(','))
NS = 'http://schemas.microsoft.com/win/2004/08/events/event'

def sn(tag): return re.sub(r'\{[^}]+\}','',tag)

results = []
try:
    with evtx.Evtx(path) as log:
        for rec in log.records():
            try:
                root = rec.lxml()
                sys_el = root.find(f'{{{NS}}}System')
                if sys_el is None: continue
                eid_el = sys_el.find(f'{{{NS}}}EventID')
                if eid_el is None or eid_el.text not in target: continue
                ts_el  = sys_el.find(f'{{{NS}}}TimeCreated')
                cmp_el = sys_el.find(f'{{{NS}}}Computer')
                ts   = (ts_el.attrib.get('SystemTime','') if ts_el is not None else '')[:19].replace('T',' ')
                comp = cmp_el.text if cmp_el is not None else ''
                de   = root.find(f'.//{{{NS}}}EventData')
                data = {}
                if de is not None:
                    for it in de:
                        data[it.attrib.get('Name', sn(it.tag))] = (it.text or '').strip()
                results.append({'eid':eid_el.text,'ts':ts,'computer':comp,'data':data})
                if len(results) >= 500: break
            except: continue
except Exception as e:
    results = [{'eid':'ERROR','ts':'','computer':'','data':{'msg':str(e)}}]
json.dump(results, open(outf,'w'))
PYEOF

    # File JSONL globale — accumula tutti gli eventi senza limiti argv
    local TMP_ALL=""
    TMP_ALL=$(mktemp /tmp/dfir_allevt_XXXXXX.jsonl)
    register_tmp "$PARSER_EVTX"; register_tmp "$TMP_ALL"

    local TOTAL_FOUND=0 SUSP_COUNT=0

    for EVTX_NAME in "${!EVTX_MAP[@]}"; do
        local EVTX_FILE
        EVTX_FILE=$(ci_find_file "$EVTX_DIR" "$EVTX_NAME")
        [[ -z "$EVTX_FILE" ]] && { dim_msg "$EVTX_NAME - $(L "non trovato" "not found")"; continue; }

        local EIDS_CSV="${EVTX_MAP[$EVTX_NAME]// /,}"
        local LABEL="${EVTX_NAME%.evtx}"; LABEL="${LABEL//%4/\/}"
        info "Parsing: $LABEL"

        local TMP_OUT; TMP_OUT=$(mktemp /tmp/dfir_evtxo_XXXXXX.json)
        "$PY3" "$PARSER_EVTX" "$EVTX_FILE" "$EIDS_CSV" "$TMP_OUT" 2>/dev/null \
            || echo "[]" > "$TMP_OUT"

        local COUNT
        COUNT=$("$PY3" -c "import json,sys; d=json.load(open(sys.argv[1])); print(len(d))" \
            "$TMP_OUT" 2>/dev/null || echo 0)
        [[ "$COUNT" -eq 0 ]] && { dim_msg "  $(L "Nessun evento corrispondente" "No matching events")"; rm -f "$TMP_OUT"; continue; }

        TOTAL_FOUND=$((TOTAL_FOUND + COUNT))
        ok "  $COUNT $(L "eventi trovati" "events found")"

        # Stampa terminale
        "$PY3" - "$TMP_OUT" << 'PYEOF' 2>/dev/null || true
import sys,json
EID_DESC={"4624":"Logon OK","4625":"Logon FAIL","4648":"Logon espl.","4688":"Proc. creato",
           "4698":"Task creato","4702":"Task modif.","4720":"Account creato","4726":"Account elim.",
           "4732":"Aggiunto gruppo","7045":"Servizio inst.","7034":"Svc crash",
           "4104":"PS ScriptBlock","21":"RDP Logon","23":"RDP Logoff",
           "1116":"Malware rilevato","1117":"Defender azione"}
SUSP={"4648","4698","4702","4720","7045","4104","1116","1117","4688"}
events=json.load(open(sys.argv[1]))
for ev in events[:15]:
    eid=ev['eid']; ts=ev['ts']
    mark="\033[0;31m⚠\033[0m " if eid in SUSP else "  "
    d=ev['data']
    detail=(d.get('SubjectUserName','') or d.get('TargetUserName','') or
            d.get('ImagePath','') or d.get('ProcessCommandLine','') or
            d.get('ScriptBlockText','')[:60] or d.get('TaskName','') or
            d.get('ServiceName','') or d.get('AccountName','') or '')
    print(f"    {mark}[{eid}] {ts}  {EID_DESC.get(eid,''):<22}  {detail[:60]}")
PYEOF

        # Accoda al file JSONL globale (con label sorgente)
        local LBL="$LABEL"
        "$PY3" -c "
import json, sys
events=json.load(open(sys.argv[1]))
for e in events: e['src']=sys.argv[2]
with open(sys.argv[3],'a') as f:
    for e in events: f.write(json.dumps(e)+'\n')
" "$TMP_OUT" "$LBL" "$TMP_ALL" 2>/dev/null || true

        # Conta sospetti
        local SC
        SC=$("$PY3" -c "
import json,sys
SUSP={'4648','4698','4702','4720','7045','4104','1116','1117','4688'}
print(sum(1 for e in json.load(open(sys.argv[1])) if e.get('eid') in SUSP))
" "$TMP_OUT" 2>/dev/null || echo 0)
        SUSP_COUNT=$((SUSP_COUNT + SC))
        rm -f "$TMP_OUT"
    done

    separator
    info "Totale eventi: ${BOLD}$TOTAL_FOUND${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL_FOUND -eq 0 ]] && { warn "$(L "Nessun evento trovato." "No events found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "evtx")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Genera righe HTML leggendo dal JSONL — mai da argv
    local TMP_ROWS; TMP_ROWS=$(mktemp /tmp/dfir_evtxr_XXXXXX.html)
    "$PY3" - "$TMP_ALL" "$TMP_ROWS" << 'PYEOF' 2>/dev/null || true
import sys, json, html as H
EID_DESC={"4624":"Logon riuscito","4625":"Logon fallito","4648":"Logon credenziali espl.",
           "4688":"Processo creato","4698":"Task creato","4702":"Task modificato",
           "4720":"Account creato","4726":"Account eliminato",
           "4732":"Aggiunto a gruppo loc.","4733":"Rimosso da gruppo loc.",
           "4756":"Aggiunto a gruppo glob.","4757":"Rimosso da gruppo glob.",
           "7045":"Nuovo servizio installato","7034":"Servizio crash","7036":"Svc avv/fermato",
           "4104":"PS Script Block Logging","21":"RDP Logon","23":"RDP Logoff",
           "24":"RDP Disco.","25":"RDP Ricollega","106":"Task registrato",
           "140":"Task aggiornato","141":"Task rimosso","200":"Task eseguito","201":"Task completato",
           "1116":"Malware rilevato","1117":"Defender azione",
           "1118":"Defender scan avviato","1119":"Defender scan completo"}
SUSP={"4648","4698","4702","4720","7045","4104","1116","1117","4688"}
SKIP={'-','-1','0','%%1796','S-1-0-0','%%1833','%%1842','%%1843',''}

events=[]
with open(sys.argv[1]) as f:
    for line in f:
        line=line.strip()
        if line:
            try: events.append(json.loads(line))
            except: pass
events.sort(key=lambda e: e.get('ts',''), reverse=True)

rows=[]
for ev in events:
    eid  = H.escape(str(ev.get('eid','')))
    ts   = H.escape(str(ev.get('ts','')))
    src  = H.escape(str(ev.get('src','')))
    desc = H.escape(EID_DESC.get(ev.get('eid',''),''))
    d    = ev.get('data',{})
    parts=[]
    for k,v in d.items():
        sv=str(v).strip()
        if sv and sv not in SKIP:
            parts.append(f"<div class='drow'><span class='lbl'>{H.escape(k)}</span>"
                         f"<span class='fld mono'>{H.escape(sv[:200])}</span></div>")
    detail=''.join(parts[:8]) or "<span class='dim'>-</span>"
    susp=ev.get('eid','') in SUSP
    rs="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if susp else ""
    ec="bad" if susp else "ok"
    rows.append(f"<tr {rs}><td style='white-space:nowrap;vertical-align:top;padding:.4rem .7rem'>"
                f"<span class='mono {ec}' style='font-size:.8rem;font-weight:700'>{eid}</span><br>"
                f"<span class='dim' style='font-size:.62rem'>{desc}</span></td>"
                f"<td class='mono' style='white-space:nowrap;vertical-align:top;padding:.4rem .7rem;"
                f"font-size:.72rem;color:var(--accent4)'>{ts}</td>"
                f"<td style='vertical-align:top;padding:.35rem .8rem;font-size:.7rem;"
                f"color:var(--text-dim)'>{src}</td>"
                f"<td style='vertical-align:top;padding:.35rem .9rem'>{detail}</td></tr>")

with open(sys.argv[2],'w') as f:
    f.write('\n'.join(rows))
PYEOF

    local TABLE_ROWS
    TABLE_ROWS=$(cat "$TMP_ROWS" 2>/dev/null || echo "")
    rm -f "$TMP_ROWS"

    {
        html_header "Event Log"
        html_page_header "EV" "Event Log <span>Forensics</span>" \
            "Windows\\System32\\winevt\\Logs" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Eventi estratti</div><div class='value'>${TOTAL_FOUND}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <style>
          .lbl{display:inline-block;font-family:var(--mono);font-size:.58rem;color:var(--text-dim);
               min-width:9rem;margin-right:.4rem;flex-shrink:0;vertical-align:top;padding-top:.1rem}
          .drow{display:flex;align-items:flex-start;margin-bottom:.18rem}
          .fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap;flex:1}
        </style>
        <div class='stitle'>$(L "Eventi per timestamp - EID · Timestamp · Sorgente · Dettagli" "Events by timestamp - EID · Timestamp · Source · Details")</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:11%'>EID</th>
            <th style='width:13%'>Timestamp</th>
            <th style='width:18%'>Log</th>
            <th>$(L "Dettagli" "Details")</th>
          </tr></thead>
          <tbody>${TABLE_ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
