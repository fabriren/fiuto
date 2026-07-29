
# ================================================================
#  MODULO 17 — Browser History (Chrome / Edge / Firefox)
# ================================================================
module_browser() {
    section_header "Browser History — Chrome / Edge / Firefox" "$CYAN"
    check_win_root || return 1


    # Verifica sqlite3 disponibile
    if ! "$PY3" -c "import sqlite3" 2>/dev/null; then
        err "$(L "sqlite3 non disponibile" "sqlite3 not available") nel Python rilevato"
        return 1
    fi

    declare -a PROFILES=(
        # [label]  [path relativo nella home utente]  [tipo]
        "Chrome|AppData/Local/Google/Chrome/User Data/Default/History|chromium"
        "Edge|AppData/Local/Microsoft/Edge/User Data/Default/History|chromium"
        "Chromium|AppData/Local/Chromium/User Data/Default/History|chromium"
        "Brave|AppData/Local/BraveSoftware/Brave-Browser/User Data/Default/History|chromium"
        "Firefox|AppData/Roaming/Mozilla/Firefox/Profiles|firefox"
    )

    ALL_VISITS=()
    TOTAL=0
    SUSP_COUNT=0

    local TMP_ROWS
    TMP_ROWS=$(mktemp /tmp/dfir_br_rows_XXXXXX.html 2>/dev/null) || {
        err "Impossibile creare file temporaneo per le righe HTML"
        return 1
    }

    local HIST_PARSER=""
    HIST_PARSER=$(mktemp /tmp/dfir_brhist_XXXXXX.py 2>/dev/null) || {
        err "Impossibile creare file temporaneo per il parser browser"
        rm -f "$TMP_ROWS" 2>/dev/null
        return 1
    }

    register_tmp "$HIST_PARSER"; register_tmp "$TMP_ROWS"
    cat > "$HIST_PARSER" << 'PYEOF'
import sys, sqlite3, shutil, os, tempfile, datetime, json

mode     = sys.argv[1]   # chromium | firefox
db_path  = sys.argv[2]
browser  = sys.argv[3]
user     = sys.argv[4]

SUSP_DOMAINS = ['pastebin','transfer.sh','mega.nz','anonfiles','ufile.io',
                'ngrok','serveo','temp-mail','guerrillamail',
                '10minutemail','hastebin','ghostbin']
SUSP_PATTERNS = ['.exe','.ps1','.bat','.vbs','.dll','.hta','.msi',
                 'powershell','base64','invoke','download','rat','trojan',
                 'metasploit','cobalt','mimikatz','payload']

def is_susp(url):
    ul = url.lower()
    return any(p in ul for p in SUSP_DOMAINS + SUSP_PATTERNS)

def chrome_time(t):
    if not t: return ''
    try:
        return (datetime.datetime(1601,1,1) +
                datetime.timedelta(microseconds=t)).strftime('%Y-%m-%d %H:%M:%S')
    except: return ''

# Copia DB in tmp per evitare lock
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    results = []
    if mode == 'chromium':
        cur.execute("""
            SELECT v.visit_time, u.url, u.title, u.visit_count
            FROM visits v JOIN urls u ON v.url=u.id
            ORDER BY v.visit_time DESC LIMIT 8000
        """)
        for row in cur.fetchall():
            ts  = chrome_time(row[0])
            url = row[1] or ''
            title = row[2] or ''
            cnt   = row[3] or 0
            susp  = is_susp(url)
            results.append({'ts':ts,'url':url,'title':title[:100],
                            'visits':cnt,'browser':browser,'user':user,'susp':susp})
    elif mode == 'firefox':
        cur.execute("""
            SELECT h.last_visit_date, p.url, p.title, p.visit_count
            FROM moz_historyvisits h JOIN moz_places p ON h.place_id=p.id
            ORDER BY h.last_visit_date DESC LIMIT 8000
        """)
        for row in cur.fetchall():
            ts_raw = row[0]
            ts = ''
            if ts_raw:
                try:
                    ts = (datetime.datetime(1970,1,1) +
                          datetime.timedelta(microseconds=ts_raw)).strftime('%Y-%m-%d %H:%M:%S')
                except: pass
            url   = row[1] or ''
            title = row[2] or ''
            cnt   = row[3] or 0
            susp  = is_susp(url)
            results.append({'ts':ts,'url':url,'title':title[:100],
                            'visits':cnt,'browser':browser,'user':user,'susp':susp})
    conn.close()
    print(json.dumps(results))
except Exception as e:
    print(json.dumps([{"error": str(e), "browser": browser, "user": user}]))
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        for PROFILE_SPEC in "${PROFILES[@]}"; do
            IFS='|' read -r BLABEL BPATH BTYPE <<< "$PROFILE_SPEC"

            if [[ "$BTYPE" == "firefox" ]]; then
                # Firefox: cerca in tutti i profili
                local FF_BASE
                FF_BASE=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Mozilla/Firefox/Profiles")
                [[ -z "$FF_BASE" ]] && continue
                for FF_PROF in "$FF_BASE"/*/; do
                    local FF_DB="${FF_PROF}places.sqlite"
                    [[ -f "$FF_DB" ]] || continue
                    local JSON_OUT
                    JSON_OUT=$("$PY3" "$HIST_PARSER" "firefox" "$FF_DB" "Firefox" "$USERNAME" 2>/dev/null || echo "[]")
                    _process_browser_json "$JSON_OUT" "$USERNAME" "$BLABEL"
                done
            else
                # Chromium-based
                local HIST_FILE
                HIST_FILE=$(ci_find_dir "$USER_DIR" "$(dirname "$BPATH")")
                [[ -z "$HIST_FILE" ]] && continue
                HIST_FILE="${HIST_FILE}/$(basename "$BPATH")"
                [[ -f "$HIST_FILE" ]] || continue
                local JSON_OUT
                JSON_OUT=$("$PY3" "$HIST_PARSER" "chromium" "$HIST_FILE" "$BLABEL" "$USERNAME" 2>/dev/null || echo "[]")
                _process_browser_json "$JSON_OUT" "$USERNAME" "$BLABEL"
            fi
        done
    done < <(get_user_homes)

    separator
    info "Visite totali: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history browser trovata." "No browser history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "browser")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Chiamata del parser su tutte le home (questo aggiungerà l'HTML a TMP_ROWS e modificherà TOTAL/SUSP_COUNT)


    {
        html_header "Browser History"
        html_page_header "BR" "Browser <span>History</span> Forensics" \
            "Chrome · Edge · Firefox · Brave" "$SCAN" "$WIN_ROOT"
        cat << 'FILTERHTML'
<style>
  .fbtn{background:var(--bg3);border:1px solid var(--border);color:var(--text-mid);
    font-family:var(--mono);font-size:.72rem;padding:.35rem .9rem;border-radius:4px;
    cursor:pointer;transition:all .15s}
  .fbtn:hover{border-color:var(--accent);color:var(--accent)}
  .fbtn-on{border-color:var(--accent2)!important;color:var(--accent2)!important;
    background:rgba(255,123,114,.12)!important}
  .user-sep td{padding:.4rem 1rem;background:var(--bg4);border-top:2px solid var(--border);
    border-bottom:1px solid var(--border)}
  .user-sep-first td{border-top:none}
  .sep-label{font-family:var(--mono);font-size:.68rem;text-transform:uppercase;
    letter-spacing:.15em;color:var(--accent);font-weight:700}
</style>
<div class='statsbar' style='align-items:center;flex-wrap:wrap'>
  <div class='stat'><div class='label'>Visite totali</div>
    <div class='value' id='cnt-all' style='cursor:pointer' onclick="setFilter('all')" title='Mostra tutto'>0</div></div>
  <div class='stat'><div class='label'>⚠ Sospette — clicca per filtrare</div>
    <div class='value' id='cnt-susp' style='color:var(--accent2);cursor:pointer' onclick="setFilter('susp')" title='Filtra solo sospette'>0</div></div>
  <div style='margin-left:auto;display:flex;gap:.6rem;align-items:center'>
    <button id='btn-all'  class='fbtn fbtn-on' onclick="setFilter('all')">Tutte</button>
    <button id='btn-susp' class='fbtn'          onclick="setFilter('susp')">⚠ Solo sospette</button>
    <span id='finfo' style='font-family:var(--mono);font-size:.68rem;color:var(--text-dim);margin-left:.6rem'></span>
  </div>
</div>
FILTERHTML
        echo "<main>
        <style>.fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}</style>
        <div class='stitle'>$(L "Cronologia navigazione — Timestamp · Utente/Browser · URL · Titolo · Visite" "Browsing history — Timestamp · User/Browser · URL · Title · Visits")</div>
        <div class='card'><table id='brtable'>
          <thead><tr>
            <th style='width:12%'>Timestamp</th>
            <th style='width:12%'>$(L "Utente/Browser" "User/Browser")</th>
            <th>URL</th>
            <th style='width:18%'>Titolo</th>
            <th style='width:5%'>N.</th>
          </tr></thead>
          <tbody id='brtbody'>"
        cat "$TMP_ROWS"
        echo "</tbody>
        </table></div></main>"
        cat << 'FILTERJS'
<script>
(function(){
  var allRows = Array.from(document.querySelectorAll('#brtbody tr'));
  var dataRows = allRows.filter(function(r){ return r.dataset.susp !== 'sep'; });
  var sepRows  = allRows.filter(function(r){ return r.dataset.susp === 'sep'; });
  var total = dataRows.length;
  var susp  = dataRows.filter(function(r){ return r.dataset.susp === 'true'; }).length;
  document.getElementById('cnt-all').textContent  = total;
  document.getElementById('cnt-susp').textContent = susp;
  function setFilter(mode){
    window._brFilter = mode;
    var vis = 0;
    // Prima nascondi/mostra le righe dati
    dataRows.forEach(function(r){
      var show = mode === 'all' || r.dataset.susp === 'true';
      r.style.display = show ? '' : 'none';
      if(show) vis++;
    });
    // Separatori: visibili solo se almeno una riga del loro blocco è visibile
    sepRows.forEach(function(sep){
      var next = sep.nextElementSibling;
      var hasVisible = false;
      while(next && next.dataset.susp !== 'sep'){
        if(next.style.display !== 'none'){ hasVisible = true; break; }
        next = next.nextElementSibling;
      }
      sep.style.display = hasVisible ? '' : 'none';
    });
    document.getElementById('btn-all').className  = 'fbtn' + (mode === 'all'  ? ' fbtn-on' : '');
    document.getElementById('btn-susp').className = 'fbtn' + (mode === 'susp' ? ' fbtn-on' : '');
    document.getElementById('finfo').textContent  =
      mode === 'susp' ? 'Mostrando ' + vis + ' sospette su ' + total + ' totali'
                      : 'Mostrando tutte le ' + total + ' voci';
  }
  window.setFilter = setFilter;
  setFilter('all');
})();
</script>
FILTERJS
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# Funzione helper usata da module_browser
# Funzione helper usata da module_browser
# Aggiorna variabili totali e accoda l'HTML per il report in TMP_ROWS
_process_browser_json() {
    local JSON_OUT="$1"
    local USERNAME="$2"
    local BLABEL="$3"

    # Facciamo fare un job completo a Python: restituisce COUNT su stdout,
    # e scrive l'HTML convertito su file. Ritorna pure il count sospetti formattato come COUNT|SUSP
    local COUNTS
    COUNTS=$(printf '%s' "$JSON_OUT" | "$PY3" -c "
import sys, json, html
try:
    data = json.loads(sys.stdin.read())
except:
    sys.exit(0)

d = [x for x in data if 'error' not in x][:8000]
if not d:
    print('0|0')
    sys.exit(0)

susp_count = sum(1 for e in d if e.get('susp', False))
print(f'{len(d)}|{susp_count}')

tmp_file = sys.argv[1]
with open(tmp_file, 'a', encoding='utf-8') as f:
    f.write(\"<tr data-susp='sep' class='user-sep'><td colspan='5'><span class='sep-label'>▶ \" + html.escape(sys.argv[2]) + \" / \" + html.escape(sys.argv[3]) + \"</span></td></tr>\\n\")
    for v in d:
        is_susp = v.get('susp', False)
        susp_str = 'true' if is_susp else 'false'
        row_style = \"style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'\" if is_susp else ''
        url_css = \"mono fld bad\" if is_susp else \"mono fld\"
        u = html.escape(v.get('url','')[:300])
        t = html.escape(v.get('title','')[:100])
        ts = html.escape(v.get('ts',''))
        cnt = str(v.get('visits',0))
        f.write(f\"<tr data-susp='{susp_str}' {row_style}><td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{ts}</td><td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{html.escape(sys.argv[2])} / {html.escape(sys.argv[3])}</td><td style='word-break:break-all;font-size:.72rem'><span class='{url_css}'>{u}</span></td><td style='font-size:.7rem;color:var(--text-mid);max-width:200px;word-break:break-word'>{t}</td><td class='mono mid' style='white-space:nowrap;font-size:.7rem;text-align:center'>{cnt}</td></tr>\\n\")
" "$TMP_ROWS" "$USERNAME" "$BLABEL" 2>/dev/null)

    [[ -z "$COUNTS" || "$COUNTS" == "0|0" ]] && return

    local P_COUNT="${COUNTS%|*}"
    local P_SUSP="${COUNTS#*|}"

    ok "$USERNAME / $BLABEL — $P_COUNT visite"
    TOTAL=$((TOTAL + P_COUNT))
    SUSP_COUNT=$((SUSP_COUNT + P_SUSP))

    # Stampa in console un riepilogo rapido delle sospette
    printf '%s' "$JSON_OUT" | "$PY3" -c "
import sys, json
data = json.loads(sys.stdin.read())
susp = [x for x in data if x.get('susp') and 'error' not in x][:5]
normal = [x for x in data if not x.get('susp') and 'error' not in x][:5]
for v in susp:
    print(f\"  \033[0;31m[!]\033[0m {v.get('ts','')}  {v.get('url','')[:80]}\")
for v in normal:
    print(f\"      {v.get('ts','')}  {v.get('url','')[:80]}\")
" 2>/dev/null || true
}
