
# ================================================================
#  MASTER TIMELINE CROSS-MODULO (Linux/macOS)
#  Aggrega tutte le evidenze con timestamp dai report generati in sessione.
#  Pensato per girare per ULTIMO (è l'ultima voce dei registri Linux/macOS):
#  legge GENERATED_REPORTS, che a quel punto contiene gli altri moduli.
# ================================================================
module_xplat_master_timeline() {
    section_header "$(L "Master Timeline - Aggregazione Cross-Moduli" "Master Timeline - Cross-Module Aggregation")" "$YELLOW"
    check_target_root || return 1

    local -a SRC=()
    local r
    for r in "${GENERATED_REPORTS[@]}"; do
        [[ "$r" == *master_timeline* ]] && continue
        [[ "$r" == */index.html ]] && continue
        [[ -f "$r" ]] && SRC+=("$r")
    done
    if [[ ${#SRC[@]} -eq 0 ]]; then
        warn "$(L "Nessun report generato in questa sessione." "No reports generated in this session.")"
        info "$(L "Esegui prima gli altri moduli (o usa 'Esegui TUTTI'), poi la Master Timeline." "Run the other modules first (or 'Run ALL'), then the Master Timeline.")"
        return 0
    fi
    info "$(L "Report da aggregare:" "Reports to aggregate:") ${BOLD}${#SRC[@]}"

    # Estrae gli eventi con timestamp da tabelle (<tr>) e righe di testo (<span class="line">).
    local YEAR; YEAR=$(date +%Y)
    local TL_RAW
    TL_RAW=$("$PY3" - "$YEAR" "${SRC[@]}" << 'PYEOF' 2>/dev/null || true
import sys, re, html as H
year = sys.argv[1]
MONTHS = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
          'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
TS_ISO = re.compile(r'\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}')
TS_SYS = re.compile(r'\b(' + '|'.join(MONTHS) + r')\s+(\d{1,2})\s+(\d{2}:\d{2}:\d{2})')
TD     = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL | re.I)
PRE    = re.compile(r'<pre[^>]*>(.*?)</pre>', re.DOTALL | re.I)
TAG    = re.compile(r'<[^>]+>')
def strip(s): return TAG.sub('', H.unescape(s)).strip()
def iso(m):   return m.group(0).replace('T', ' ')[:19]
def sysfmt(m):return f"{year}-{MONTHS[m.group(1)]}-{int(m.group(2)):02d} {m.group(3)}"

seen = set(); out = []
def add(ts, mod, desc):
    desc = ' '.join(desc.split())[:200]
    k = (ts, mod, desc[:50])
    if k in seen: return
    seen.add(k); out.append(f"{ts}\t{mod}\t{desc}")

for rp in sys.argv[2:]:
    parts = rp.rstrip('/').split('/')
    mod = re.sub(r'_\d{8}_\d{6}$', '', parts[-2]) if len(parts) >= 2 else parts[-1]
    try:
        content = open(rp, encoding='utf-8', errors='replace').read()
    except Exception:
        continue
    # 1) Righe di tabella con un timestamp in una cella
    for tr in re.finditer(r'<tr[^>]*>(.*?)</tr>', content, re.DOTALL | re.I):
        cells = [strip(m.group(1)) for m in TD.finditer(tr.group(1))]
        if not cells: continue
        ts = ''
        for c in cells:
            m = TS_ISO.search(c)
            if m: ts = iso(m); break
        if not ts: continue
        desc = ' | '.join(c[:90] for c in cells
                          if c and not TS_ISO.fullmatch(c.strip()) and not re.fullmatch(r'\d+', c.strip()))
        add(ts, mod, desc)
    # 2) Blocchi <pre> (log, history, config): tolgo i tag mantenendo i newline,
    #    poi valuto ogni riga (con timestamp ISO, [ISO] o syslog "Mon DD HH:MM:SS").
    for pm in PRE.finditer(content):
        block = H.unescape(TAG.sub('', pm.group(1)))
        for raw in block.split('\n'):
            txt = re.sub(r'^\s*\d+\s+', '', raw).strip()   # toglie il numero di riga
            if not txt: continue
            m = TS_ISO.search(txt)
            if m:
                ts = iso(m)
            else:
                m = TS_SYS.search(txt)
                if not m: continue
                ts = sysfmt(m)
            add(ts, mod, txt)

print('\n'.join(out))
PYEOF
)

    local TOTAL; TOTAL=$(printf '%s' "$TL_RAW" | grep -c $'\t' 2>/dev/null || echo 0)
    local _TMP; _TMP=$(mktemp)
    printf '%s\n' "$TL_RAW" | awk 'NF' | sort -t$'\t' -k1 -r | head -100000 > "$_TMP"
    local SHOW; SHOW=$(wc -l < "$_TMP" 2>/dev/null || echo 0)

    separator
    info "$(L "Eventi con timestamp:" "Events with timestamp:") ${BOLD}$TOTAL${RESET} ($(L "mostrati" "showing") $SHOW)"
    if [[ "$TOTAL" -eq 0 ]]; then
        rm -f "$_TMP"
        warn "$(L "Nessun timestamp trovato nei report aggregati." "No timestamp found in the aggregated reports.")"
        return 0
    fi
    ask_yn "$(L "Generare report HTML Master Timeline?" "Generate Master Timeline HTML report?")" || { rm -f "$_TMP"; return 0; }

    local TABLE; TABLE=$("$PY3" - "$_TMP" << 'PYEOF'
import sys, html
print("<table><tr><th style='width:15%'>Timestamp</th><th style='width:16%'>Modulo</th><th>Descrizione</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    c = line.rstrip('\n').split('\t')
    if len(c) < 3: continue
    print(f"<tr><td class='mono ok' style='white-space:nowrap'>{html.escape(c[0])}</td>"
          f"<td class='mono'>{html.escape(c[1])}</td>"
          f"<td class='mono' style='color:var(--text)'>{html.escape(c[2])}</td></tr>")
print("</table>")
PYEOF
)
    rm -f "$_TMP"
    local BODY; BODY=$(generic_card_html "$(L "Timeline cronologica (decrescente)" "Chronological timeline (descending)")" "$WIN_ROOT" "$SHOW $(L "eventi" "events")" "$TABLE" "◷")
    local STATS; STATS="$(stat_box "$(L "Report aggregati" "Aggregated reports")" "${#SRC[@]}")$(stat_box "$(L "Eventi" "Events")" "$TOTAL" "info")$(stat_box "$(L "Mostrati" "Shown")" "$SHOW" "ok")"
    finish_report "master_timeline" "Master Timeline" "TL" "$(L "Aggregazione cross-moduli" "Cross-module aggregation")" "$STATS" "<div class='cards'>$BODY</div>"
}
