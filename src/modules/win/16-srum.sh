
# ================================================================
#  MODULO 16 — SRUM (System Resource Usage Monitor)
# ================================================================
module_srum() {
    section_header "SRUM - System Resource Usage Monitor" "$BLUE"
    check_win_root || return 1

    local SRUM_PATH
    SRUM_PATH=$(ci_find_file \
        "$(ci_find_dir "$WIN_ROOT" "Windows/System32/sru")" \
        "SRUDB.dat")
    [[ -z "$SRUM_PATH" ]] && \
        SRUM_PATH=$(find "$WIN_ROOT" -maxdepth 8 -iname "SRUDB.dat" 2>/dev/null | head -1)

    if [[ -z "$SRUM_PATH" ]]; then
        warn "$(L "SRUDB.dat non trovato" "SRUDB.dat not found")"
        return 0
    fi

    info "SRUM DB: $SRUM_PATH"

    # Verifica pyesedb o libesedb-python
    local HAS_ESE=false
    "$PY3" -c "import pyesedb" 2>/dev/null && HAS_ESE=true

    if ! $HAS_ESE; then
        warn "pyesedb non disponibile - parsing tramite strings (risultati parziali)"
        warn "Per parsing completo: pip install pyesedb  (richiede libesedb)"
        echo ""
        # Fallback strings: cerca nomi applicazione e byte trasferiti
        info "$(L "Estrazione tramite strings..." "Extracting via strings...")"
        local STRINGS_OUT
        STRINGS_OUT=$(strings "$SRUM_PATH" 2>/dev/null | \
            grep -iE '\.(exe|dll|bat|ps1|vbs|py)|\\Device\\|\\\\[A-Za-z]' | \
            grep -v '^.\{1\}$' | sort -u | head -200 || true)

        echo "$STRINGS_OUT" | while IFS= read -r LINE; do
            printf "  ${DIM}%s${RESET}\n" "$LINE"
        done | head -50

        separator
        info "Strings estratte (prime 200 uniche). Per analisi completa installa pyesedb."
        ask_yn "Generare report HTML (strings)?" || return 0

        local REPORT_HTML
        local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "srum")
        local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
        local ROWS=""
        echo "$STRINGS_OUT" | while IFS= read -r LINE; do
            local ESC_LINE; ESC_LINE=$(printf '%s' "$LINE" | sed 's/&/\&amp;/g;s/</\&lt;/g;s/>/\&gt;/g')
            ROWS+="<tr><td class='mono' style='font-size:.72rem;word-break:break-all'>${ESC_LINE}</td></tr>"
        done

        {
            html_header "SRUM"
            html_page_header "SR" "SRUM <span>Database</span> Forensics" \
                "Windows\\System32\\sru\\SRUDB.dat" "$SCAN" "$WIN_ROOT"
            echo "<div class='statsbar'>
              <div class='stat'><div class='label'>Modalità</div><div class='value' style='font-size:1rem;padding-top:.4rem;color:var(--accent4)'>strings</div></div>
            </div><main>
            <div class='stitle'>Stringhe estratte da SRUDB.dat (pyesedb non disponibile)</div>
            <div class='card'><table><thead><tr><th>$(L "Stringa" "String")</th></tr></thead><tbody>${ROWS}</tbody></table></div>
            <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px'>
              <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Per analisi completa con traffico di rete e CPU per applicazione:</div>
              <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>pip install pyesedb<br>oppure usa: srum-dump (Mark Baggett) / ese2csv</div>
            </div></main>"
            html_footer "$SCAN" "$WIN_ROOT"
        } > "$REPORT_HTML"
        register_report "$REPORT_HTML"
        ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
        open_report_prompt "$REPORT_HTML"
        return 0
    fi

    # Parser completo con pyesedb
    info "Parsing ESE con pyesedb..."
    local SRUM_JSON
    SRUM_JSON=$("$PY3" - "$SRUM_PATH" << 'PYEOF' 2>/dev/null || echo "[]"
import sys, json, datetime
try:
    import pyesedb
except ImportError:
    print("[]"); sys.exit()

try:
    db = pyesedb.open(sys.argv[1])
except Exception as e:
    print(json.dumps([{"error": str(e)}])); sys.exit()

# GUID tabelle principali
TABLE_NAMES = {
    '{5C8CF1C7-7257-4F13-B223-970EF5939312}': 'App Resource Usage',
    '{973F5D5C-1D90-4944-BE8E-24B94231A174}': 'Network Data Usage',
    '{D10CA2FE-6FCF-4F6D-848E-B2E99266FA89}': 'Push Notifications',
    '{DD6636C4-8929-4683-974E-22C046A43763}': 'Network Connectivity',
    '{FEE4E14F-02A9-4550-B5CE-5FA2DA202E37}': 'Energy Usage',
}

results = []
for i in range(db.number_of_tables):
    tbl = db.get_table(i)
    tname = tbl.name
    label = TABLE_NAMES.get(tname, tname)
    if 'SruDbIdMapTable' in tname or 'MSysObjects' in tname:
        continue
    cols = [tbl.get_column(j).name for j in range(tbl.number_of_columns)]
    rows_data = []
    rec = tbl.first_record
    count = 0
    while rec and count < 300:
        row = {}
        for j, col in enumerate(cols):
            try:
                val = rec.get_value_data_as_string(j)
                if val: row[col] = val[:200]
            except: pass
        if row:
            rows_data.append(row)
        rec = rec.next_record
        count += 1
    if rows_data:
        results.append({'table': label, 'guid': tname, 'rows': rows_data})

print(json.dumps(results))
PYEOF
)

    local TABLE_COUNT
    TABLE_COUNT=$("$PY3" -c "import sys,json; d=json.loads(sys.stdin.read()); print(len(d))" <<< "$SRUM_JSON" 2>/dev/null || echo 0)
    local ROW_COUNT
    ROW_COUNT=$("$PY3" -c "import sys,json; d=json.loads(sys.stdin.read()); print(sum(len(t.get('rows',[])) for t in d))" <<< "$SRUM_JSON" 2>/dev/null || echo 0)

    info "Tabelle SRUM: $TABLE_COUNT  |  Record: $ROW_COUNT"
    echo ""

    "$PY3" - "$SRUM_JSON" << 'PYEOF' 2>/dev/null || true
import sys, json
tables = json.loads(sys.argv[1])
for tbl in tables[:4]:
    print(f"  [{tbl['table']}]  {len(tbl.get('rows',[]))} record")
    for row in tbl.get('rows',[])[:5]:
        interesting = {k:v for k,v in row.items() if any(
            x in k.lower() for x in ['app','exe','byte','sent','recv','connect','user','time']
        )}
        if interesting:
            parts = [f"{k}={v[:40]}" for k,v in list(interesting.items())[:4]]
            print(f"    {', '.join(parts)}")
    print()
PYEOF

    separator
    info "Record totali: ${BOLD}$ROW_COUNT"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "srum")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local SECTIONS_HTML
    SECTIONS_HTML=$("$PY3" - "$SRUM_JSON" << 'PYEOF' 2>/dev/null || echo ""
import sys, json, html as html_mod
tables = json.loads(sys.argv[1])
out = []
for tbl in tables:
    label = html_mod.escape(tbl.get('table',''))
    rows  = tbl.get('rows', [])
    if not rows: continue
    all_cols = list(dict.fromkeys(k for r in rows for k in r.keys()))
    show_cols = all_cols[:10]
    th = ''.join(f"<th>{html_mod.escape(c)}</th>" for c in show_cols)
    trs = []
    for row in rows[:200]:
        tds = ''.join(f"<td class='mono' style='font-size:.68rem;word-break:break-all;max-width:200px'>{html_mod.escape(str(row.get(c,''))[:100])}</td>" for c in show_cols)
        trs.append(f"<tr>{tds}</tr>")
    out.append(f"""
<div class='stitle' style='margin-top:2rem'>{label} ({len(rows)} record)</div>
<div class='card' style='overflow-x:auto'><table style='min-width:600px'>
  <thead><tr>{th}</tr></thead>
  <tbody>{''.join(trs)}</tbody>
</table></div>""")
print('\n'.join(out))
PYEOF
)

    {
        html_header "SRUM"
        html_page_header "SR" "SRUM <span>Database</span> Forensics" \
            "Windows\\System32\\sru\\SRUDB.dat" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Tabelle</div><div class='value'>${TABLE_COUNT}</div></div>
          <div class='stat info'><div class='label'>Record</div><div class='value'>${ROW_COUNT}</div></div>
        </div><main>${SECTIONS_HTML}</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
