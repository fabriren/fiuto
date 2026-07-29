
# ----------------------------------------------------------------
#  Export JSONL (schema Timesketch / plaso)
#
#  I report HTML sono ottimi per l'analista e inutilizzabili per una
#  pipeline: non si correlano con altre sorgenti e non si caricano in un
#  SIEM. Qui gli stessi eventi vengono riemessi in JSON Lines con i campi
#  attesi da Timesketch (datetime, timestamp_desc, message), cosi' il
#  risultato di FIUTO entra direttamente in una super-timeline.
#
#  L'estrazione lavora sull'HTML gia' prodotto invece che sui dati grezzi
#  dei singoli moduli: e' l'unico punto in cui il formato e' omogeneo per
#  tutti e 64+ i moduli, e non richiede di toccarli uno per uno.
# ----------------------------------------------------------------
export_report_jsonl() {
    local HTML="$1"
    local DIR; DIR=$(dirname "$HTML")
    local SLUG; SLUG=$(basename "$DIR" | sed -E 's/_[0-9]{8}_[0-9]{6}$//')
    local OUT="${DIR}/report.jsonl"

    "$PY3" - "$HTML" "$SLUG" "${WIN_ROOT:-}" "${HOST_NAME:-}" "${OS_TYPE:-}" > "$OUT" << 'PYEOF' 2>/dev/null
import sys, re, json, html as H, datetime

html_path, slug = sys.argv[1], sys.argv[2]
volume  = sys.argv[3] if len(sys.argv) > 3 else ''
host    = sys.argv[4] if len(sys.argv) > 4 else ''
os_type = sys.argv[5] if len(sys.argv) > 5 else ''

MONTHS = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
          'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
TS_ISO = re.compile(r'\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}')
TS_SYS = re.compile(r'\b(' + '|'.join(MONTHS) + r')\s+(\d{1,2})\s+(\d{2}:\d{2}:\d{2})')
TD     = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL | re.I)
PRE    = re.compile(r'<pre[^>]*>(.*?)</pre>', re.DOTALL | re.I)
TAG    = re.compile(r'<[^>]+>')

def strip(s):
    return TAG.sub('', H.unescape(s)).strip()

# I log syslog non portano l'anno: si usa quello del file HTML (l'analisi e'
# contestuale all'acquisizione). Approssimazione esplicita, non silenziosa.
year = str(datetime.date.today().year)
year_assumed = False

def norm_iso(raw):
    return raw.replace('T', ' ')[:19].replace(' ', 'T')

def from_sys(m):
    global year_assumed
    year_assumed = True
    return f"{year}-{MONTHS[m.group(1)]}-{int(m.group(2)):02d}T{m.group(3)}"

try:
    content = open(html_path, encoding='utf-8', errors='replace').read()
except Exception:
    sys.exit(0)

seen = set()
out  = []

def emit(dt, message, assumed):
    message = ' '.join(message.split())[:2000]
    if not message:
        return
    key = (dt, message[:120])
    if key in seen:
        return
    seen.add(key)
    rec = {
        # Campi richiesti da Timesketch
        "datetime": dt,
        "timestamp_desc": f"FIUTO {slug}",
        "message": message,
        # Contesto aggiuntivo
        "data_type": f"fiuto:{slug}",
        "module": slug,
        "source_volume": volume,
        "hostname": host,
        "os": os_type,
    }
    if assumed:
        # L'anno non era nel dato di origine: va dichiarato, non nascosto.
        rec["year_inferred"] = True
    out.append(rec)

# 1) Righe di tabella in cui una cella contiene un timestamp
for tr in re.finditer(r'<tr[^>]*>(.*?)</tr>', content, re.DOTALL | re.I):
    cells = [strip(m.group(1)) for m in TD.finditer(tr.group(1))]
    if not cells:
        continue
    dt = None
    for c in cells:
        m = TS_ISO.search(c)
        if m:
            dt = norm_iso(m.group(0))
            break
    if not dt:
        continue
    msg = ' | '.join(c for c in cells
                     if c and not TS_ISO.fullmatch(c.strip()) and not re.fullmatch(r'\d+', c.strip()))
    emit(dt, msg, False)

# 2) Blocchi <pre> (log, history, config): una riga per evento
for pm in PRE.finditer(content):
    block = H.unescape(TAG.sub('', pm.group(1)))
    for raw in block.split('\n'):
        txt = re.sub(r'^\s*\d+\s+', '', raw).strip()   # via il numero di riga
        if not txt:
            continue
        assumed = False
        m = TS_ISO.search(txt)
        if m:
            dt = norm_iso(m.group(0))
        else:
            m = TS_SYS.search(txt)
            if not m:
                continue
            dt = from_sys(m)
            assumed = True
        emit(dt, txt, assumed)

for rec in sorted(out, key=lambda r: r["datetime"]):
    print(json.dumps(rec, ensure_ascii=False))
PYEOF

    local N=0
    [[ -s "$OUT" ]] && N=$(wc -l < "$OUT")
    if [[ "$N" -eq 0 ]]; then
        rm -f "$OUT"
        return 0
    fi

    # Timeline unica di sessione: e' il file da caricare in Timesketch.
    local COMBINED="${REPORT_BASE_DIR}/fiuto_timeline.jsonl"
    cat "$OUT" >> "$COMBINED" 2>/dev/null || true
    ok "$(L "Export JSONL:" "JSONL export:") ${BOLD}${N}$(L " eventi" " events")${RESET} → $(basename "$OUT")"
    log_msg "[JSONL] $OUT — $N eventi"
}
