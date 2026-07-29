
# Prepara la directory report e restituisce il path del file HTML.
# La directory base viene creata SOLO qui, al primo report effettivo,
# per evitare directory vuote quando l'utente non genera alcun report.
prepare_report_dir() {
    local MODULE_NAME="$1"
    [[ -n "$REPORT_BASE_DIR" && ! -d "$REPORT_BASE_DIR" ]] && mkdir -p "$REPORT_BASE_DIR"
    local DIR="${REPORT_BASE_DIR}/${MODULE_NAME}_$(date +%Y%m%d_%H%M%S)"
    mkdir -p "$DIR"
    echo "${DIR}/report.html"
}

# Registra il report nel riepilogo di sessione.
# Chiamata DOPO la scrittura effettiva del file, fuori da qualsiasi subshell.
#
# E' anche il punto di aggancio dell'export JSONL: ogni modulo passa di qui
# dopo aver scritto il proprio HTML, quindi l'export copre automaticamente
# tutti i moduli — inclusi quelli Windows, che generano l'HTML per conto
# proprio senza passare da finish_report.
register_report() {
    [[ -n "${1:-}" && -f "$1" ]] || return 0
    GENERATED_REPORTS+=("$1")
    [[ "$EXPORT_JSONL" == "true" ]] && export_report_jsonl "$1"
    return 0
}

# HTML boilerplate header comune a tutti i report
html_header() {
    local TITLE="$1"
    local ACCENT="${2:-#58a6ff}"
    cat << HTMLEOF
<!DOCTYPE html>
<html lang="$(L "it" "en")">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${TITLE} — DFIR Report</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@300;400;500;700&family=DM+Sans:wght@400;500;700&display=swap');
  :root {
    --bg:#080b0f; --bg2:#0d1117; --bg3:#131920; --bg4:#1a2332;
    --border:#1e2d3d; --accent:${ACCENT}; --accent2:#ff7b72;
    --accent3:#3fb950; --accent4:#f0883e;
    --text:#c9d1d9; --text-dim:#3d5166; --text-mid:#6e8898;
    --mono:'Fira Code',monospace; --sans:'DM Sans',sans-serif;
  }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { background:var(--bg); color:var(--text); font-family:var(--sans); min-height:100vh; }
  body::before {
    content:''; position:fixed; inset:0; pointer-events:none; z-index:0;
    background-image:linear-gradient(rgba(88,166,255,.03) 1px,transparent 1px),
      linear-gradient(90deg,rgba(88,166,255,.03) 1px,transparent 1px);
    background-size:40px 40px;
  }
  header { position:relative; z-index:1; background:var(--bg2); border-bottom:1px solid var(--border);
    padding:1.8rem 3rem; display:flex; align-items:center; gap:2rem; }
  .hicon { width:3rem; height:3rem; background:linear-gradient(135deg,var(--accent),#1f6feb);
    display:flex; align-items:center; justify-content:center; font-family:var(--mono);
    font-size:1.1rem; color:#fff; font-weight:700; flex-shrink:0;
    clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%); }
  .htxt h1 { font-size:1.35rem; font-weight:700; color:#fff; }
  .htxt h1 span { color:var(--accent); }
  .htxt .sub { font-size:.72rem; color:var(--text-dim); font-family:var(--mono); margin-top:.25rem; }
  .hmeta { margin-left:auto; font-family:var(--mono); font-size:.7rem; color:var(--text-dim); text-align:right; line-height:2; }
  .hmeta .val { color:var(--accent); }
  .statsbar { position:relative; z-index:1; background:var(--bg3); border-bottom:1px solid var(--border);
    padding:1rem 3rem; display:flex; gap:3rem; flex-wrap:wrap; }
  .stat .label { font-size:.6rem; text-transform:uppercase; letter-spacing:.15em; color:var(--text-dim); font-family:var(--mono); }
  .stat .value { font-size:1.8rem; font-weight:800; font-family:var(--mono); color:var(--accent2); }
  .stat.info .value { color:var(--accent); }
  .stat.ok .value { color:var(--accent3); }
  main { position:relative; z-index:1; padding:2rem 3rem; max-width:1400px; margin:0 auto; }
  .stitle { font-size:.65rem; text-transform:uppercase; letter-spacing:.2em; color:var(--text-dim);
    font-family:var(--mono); margin-bottom:1.5rem; padding-bottom:.5rem; border-bottom:1px solid var(--border); }
  .stitle::before { content:'> '; color:var(--accent3); }
  .cards { display:flex; flex-direction:column; gap:1.5rem; }
  .card { background:var(--bg2); border:1px solid var(--border); border-radius:6px; overflow:hidden; }
  .card:hover { border-color:rgba(88,166,255,.25); }
  .card-header { display:flex; align-items:center; gap:1rem; padding:1rem 1.5rem;
    background:var(--bg3); border-bottom:1px solid var(--border); }
  .uicon { width:2.2rem; height:2.2rem; background:linear-gradient(135deg,var(--accent),#1f6feb);
    color:#fff; display:flex; align-items:center; justify-content:center; font-size:.9rem; flex-shrink:0;
    clip-path:polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%); }
  .uname { font-weight:700; font-size:1rem; color:#fff; }
  .upath { font-family:var(--mono); font-size:.62rem; color:var(--text-dim); margin-top:.1rem; word-break:break-all; }
  .badge { margin-left:auto; background:rgba(88,166,255,.1); border:1px solid rgba(88,166,255,.3);
    color:var(--accent); font-family:var(--mono); font-size:.72rem; padding:.2rem .7rem;
    border-radius:2px; white-space:nowrap; }
  .badge.warn { background:rgba(255,123,114,.1); border-color:rgba(255,123,114,.3); color:var(--accent2); }
  table { width:100%; border-collapse:collapse; font-family:var(--mono); font-size:.78rem; }
  th { padding:.5rem 1rem; text-align:left; font-size:.6rem; text-transform:uppercase;
    letter-spacing:.15em; color:var(--text-dim); border-bottom:1px solid var(--border);
    background:rgba(0,0,0,.3); }
  td { padding:.5rem 1rem; border-bottom:1px solid rgba(30,45,61,.5); vertical-align:top; }
  tr:last-child td { border-bottom:none; }
  tr:hover { background:rgba(88,166,255,.03); }
  .mono { font-family:var(--mono); }
  .dim  { color:var(--text-dim); }
  .mid  { color:var(--text-mid); }
  .ok   { color:var(--accent3); }
  .warn { color:var(--accent4); }
  .bad  { color:var(--accent2); background:rgba(255,123,114,.07);
    border-left:2px solid rgba(255,123,114,.5); padding-left:calc(.5rem - 2px); }
  pre { white-space:pre-wrap; word-break:break-word; }
  footer { position:relative; z-index:1; border-top:1px solid var(--border); padding:1rem 3rem;
    font-family:var(--mono); font-size:.62rem; color:var(--text-dim); margin-top:3rem;
    display:flex; justify-content:space-between; }
  footer span { color:var(--accent); }
  /* ── Barra di ricerca / Search bar ── */
  #fz-toolbar { position:fixed; top:.7rem; right:.9rem; z-index:50; display:flex; align-items:center;
    gap:.4rem; font-family:var(--mono); }
  #fz-csv { display:flex; align-items:center; gap:.35rem; background:var(--bg3);
    border:1px solid var(--border); border-radius:6px; color:var(--text); cursor:pointer;
    font-family:var(--mono); font-size:.74rem; padding:.4rem .6rem; box-shadow:0 4px 18px rgba(0,0,0,.45); }
  #fz-csv:hover { border-color:var(--accent3); color:var(--accent3); }
  #fz-csv svg { width:.95rem; height:.95rem; }
  #fz-search { display:flex; align-items:center; gap:.4rem; background:var(--bg3);
    border:1px solid var(--border); border-radius:6px; padding:.35rem .5rem;
    box-shadow:0 4px 18px rgba(0,0,0,.45); }
  #fz-search input { background:var(--bg); border:1px solid var(--border); color:var(--text);
    font-family:var(--mono); font-size:.78rem; padding:.3rem .55rem; border-radius:4px; width:15rem;
    outline:none; }
  #fz-search input:focus { border-color:var(--accent); }
  #fz-search button { background:var(--bg4); border:1px solid var(--border); color:var(--text);
    font-family:var(--mono); font-size:.8rem; line-height:1; padding:.28rem .5rem; border-radius:4px;
    cursor:pointer; }
  #fz-search button:hover { border-color:var(--accent); color:var(--accent); }
  #fz-close:hover { border-color:var(--accent2); color:var(--accent2); }
  #fz-count { font-size:.7rem; color:var(--text-dim); min-width:4.2rem; text-align:center; }
  #fz-search.fz-hidden { display:none; }
  #fz-toggle { display:none; align-items:center; justify-content:center; width:2.1rem; height:2.1rem;
    background:var(--bg3); border:1px solid var(--border); border-radius:6px; color:var(--text-mid);
    cursor:pointer; box-shadow:0 4px 18px rgba(0,0,0,.45); }
  #fz-toggle.fz-show { display:flex; }
  #fz-toggle:hover { border-color:var(--accent); color:var(--accent); }
  #fz-toggle svg { width:1.05rem; height:1.05rem; }
  mark.fz-hit { background:rgba(240,136,62,.32); color:inherit; border-radius:2px; padding:0 .05em; }
  mark.fz-hit.fz-active { background:var(--accent); color:#08121f; }
</style>
</head>
<body>
<div id="fz-toolbar">
  <button id="fz-csv" title="$(L "Scarica tutti i dati in CSV" "Download all data as CSV")"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"></path><path d="M7 11l5 5 5-5"></path><path d="M5 21h14"></path></svg>CSV</button>
  <div id="fz-search">
    <input id="fz-q" type="search" placeholder="$(L "Cerca nel report…" "Search report…")" autocomplete="off" spellcheck="false">
    <span id="fz-count"></span>
    <button id="fz-prev" title="$(L "Precedente" "Previous")">▲</button>
    <button id="fz-next" title="$(L "Successivo" "Next")">▼</button>
    <button id="fz-close" title="$(L "Chiudi ricerca" "Close search")">✕</button>
  </div>
  <button id="fz-toggle" title="$(L "Cerca nel report" "Search report")"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"></circle><line x1="21" y1="21" x2="16.5" y2="16.5"></line></svg></button>
</div>
HTMLEOF
}

html_page_header() {
    local ICON="$1"
    local TITLE="$2"
    local SUBTITLE="$3"
    local SCAN="$4"
    local SRC="$5"
    local HASH="${6:-}"
    cat << HTMLEOF
<header>
  <div class="hicon">${ICON}</div>
  <div class="htxt">
    <h1>${TITLE}</h1>
    <div class="sub">${SUBTITLE}</div>
  </div>
  <div class="hmeta">
    <div>$(L "Scansione:" "Scan:") <span class="val">${SCAN}</span></div>
    <div>$(L "Sorgente:" "Source:") <span class="val">${SRC}</span></div>$([ -n "$HASH" ] && echo "
    <div style='font-size:.6rem'>SHA256: <span class=\"val\" style='word-break:break-all'>${HASH}</span></div>")
  </div>
</header>
HTMLEOF
}

html_footer() {
    local SCAN="$1"
    local SRC="$2"
    cat << HTMLEOF
<footer>
  <div>fiuto.sh — $(L "generato il" "generated on") ${SCAN}</div>
  <div>$(L "Sorgente:" "Source:") <span>${SRC}</span></div>
</footer>
HTMLEOF
    # Script di ricerca/evidenziazione (heredoc quotato: nessuna espansione shell)
    cat << 'JSEOF'
<script>
(function(){
  var q=document.getElementById('fz-q'), cnt=document.getElementById('fz-count'),
      prev=document.getElementById('fz-prev'), next=document.getElementById('fz-next'),
      bar=document.getElementById('fz-search'), closeBtn=document.getElementById('fz-close'),
      toggle=document.getElementById('fz-toggle'), csvBtn=document.getElementById('fz-csv');
  // ── Export CSV: serializza TUTTE le tabelle e i blocchi <pre> del report ──
  function csvCell(v){ v=(v==null?'':String(v)); return /[",\r\n]/.test(v) ? '"'+v.replace(/"/g,'""')+'"' : v; }
  function exportCSV(){
    var root=document.querySelector('main')||document.body;
    var nodes=root.querySelectorAll('table, pre');
    var lines=[], lastLabel='';
    nodes.forEach(function(node){
      var card=node.closest ? node.closest('.card') : null, label='';
      if(card){
        var u=card.querySelector('.uname'), p=card.querySelector('.upath');
        label=[u&&u.textContent.trim(), p&&p.textContent.trim()].filter(Boolean).join(' — ');
      }
      if(label && label!==lastLabel){ lines.push('# '+csvCell(label)); lastLabel=label; }
      if(node.tagName==='TABLE'){
        node.querySelectorAll('tr').forEach(function(tr){
          var row=[];
          tr.querySelectorAll('th,td').forEach(function(c){ row.push(csvCell(c.textContent.replace(/\s+/g,' ').trim())); });
          if(row.length) lines.push(row.join(','));
        });
      } else {
        node.textContent.split('\n').forEach(function(l){
          l=l.replace(/^\s*\d+\s/,'').replace(/\s+$/,'');   // toglie il numero di riga
          if(l.trim()!=='') lines.push(csvCell(l));
        });
      }
      lines.push('');
    });
    if(!lines.length){ return; }
    var blob=new Blob(['\ufeff'+lines.join('\r\n')], {type:'text/csv;charset=utf-8'});
    var name=(document.title||'fiuto_report').replace(/\s*—.*$/,'').replace(/[^\w.-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,80)||'fiuto_report';
    var a=document.createElement('a');
    a.href=URL.createObjectURL(blob); a.download=name+'.csv';
    document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 200);
  }
  if(csvBtn) csvBtn.addEventListener('click', exportCSV);
  if(!q) return;
  var hits=[], idx=-1, timer=null;
  function openBar(){ bar.classList.remove('fz-hidden'); if(toggle) toggle.classList.remove('fz-show'); q.focus(); q.select(); }
  function closeBar(){ q.value=''; clearMarks(); cnt.textContent=''; bar.classList.add('fz-hidden'); if(toggle) toggle.classList.add('fz-show'); }
  function esc(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
  function clearMarks(){
    document.querySelectorAll('mark.fz-hit').forEach(function(m){
      var p=m.parentNode; if(!p) return;
      p.replaceChild(document.createTextNode(m.textContent), m); p.normalize();
    });
    hits=[]; idx=-1;
  }
  function activate(){
    hits.forEach(function(m){ m.classList.remove('fz-active'); });
    if(idx>=0 && idx<hits.length){
      var m=hits[idx]; m.classList.add('fz-active');
      m.scrollIntoView({block:'center', behavior:'smooth'});
      cnt.textContent=(idx+1)+'/'+hits.length;
    }
  }
  function go(d){ if(!hits.length) return; idx=(idx+d+hits.length)%hits.length; activate(); }
  function search(term){
    clearMarks();
    if(!term || term.length<2){ cnt.textContent=''; return; }
    var reTest=new RegExp(esc(term),'i'), reG=new RegExp(esc(term),'gi');
    var walker=document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
      acceptNode:function(node){
        if(!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        var p=node.parentNode;
        while(p){
          if(p.id==='fz-search') return NodeFilter.FILTER_REJECT;
          var t=p.nodeName;
          if(t==='SCRIPT'||t==='STYLE'||t==='MARK') return NodeFilter.FILTER_REJECT;
          p=p.parentNode;
        }
        return reTest.test(node.nodeValue)?NodeFilter.FILTER_ACCEPT:NodeFilter.FILTER_REJECT;
      }
    });
    var nodes=[], n;
    while(n=walker.nextNode()) nodes.push(n);
    nodes.forEach(function(node){
      var text=node.nodeValue, frag=document.createDocumentFragment(), last=0, m;
      reG.lastIndex=0;
      while((m=reG.exec(text))){
        if(m.index>last) frag.appendChild(document.createTextNode(text.slice(last,m.index)));
        var mk=document.createElement('mark'); mk.className='fz-hit'; mk.textContent=m[0];
        frag.appendChild(mk); hits.push(mk);
        last=m.index+m[0].length;
        if(m[0].length===0) reG.lastIndex++;
      }
      if(last<text.length) frag.appendChild(document.createTextNode(text.slice(last)));
      if(node.parentNode) node.parentNode.replaceChild(frag, node);
    });
    if(hits.length){ idx=0; activate(); } else { cnt.textContent='0'; }
  }
  q.addEventListener('input', function(){ clearTimeout(timer); timer=setTimeout(function(){ search(q.value); }, 200); });
  q.addEventListener('keydown', function(e){
    if(e.key==='Enter'){ e.preventDefault(); go(e.shiftKey?-1:1); }
    else if(e.key==='Escape'){ e.preventDefault(); closeBar(); }
  });
  prev.addEventListener('click', function(){ go(-1); });
  next.addEventListener('click', function(){ go(1); });
  if(closeBtn) closeBtn.addEventListener('click', closeBar);
  if(toggle) toggle.addEventListener('click', openBar);
  document.addEventListener('keydown', function(e){
    if((e.ctrlKey||e.metaKey) && (e.key==='f'||e.key==='F')){ e.preventDefault(); openBar(); }
  });
})();
</script>
</body></html>
JSEOF
}

# ================================================================
#  MENU PRINCIPALE
# ================================================================
# ================================================================
#  HELPER CONDIVISI PER I MODULI LINUX / macOS
# ================================================================

# Blocco <style> per i <pre> con numeri di riga ed evidenziazione (riuso dal modulo PS)
pre_style_block() {
    cat << 'EOF'
<style>
  .hist-pre{font-family:var(--mono);font-size:.75rem;line-height:1.7;padding:.8rem 1rem;
    overflow-x:auto;max-height:520px;overflow-y:auto;}
  .hist-pre::-webkit-scrollbar{width:5px;height:5px}
  .hist-pre::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
  .line{display:block;color:var(--text);white-space:pre;padding:.05rem .5rem;border-radius:2px}
  .line:hover{background:rgba(88,166,255,.05)}
  .line.sensitive{color:var(--accent2);background:rgba(255,123,114,.07);
    border-left:2px solid rgba(255,123,114,.5);padding-left:calc(.5rem - 2px)}
  .lnum{color:var(--text-dim);user-select:none;margin-right:1rem;font-size:.7rem}
  .grp{margin-bottom:1.5rem}
</style>
EOF
}

# Genera il contenuto di un <pre> (numeri di riga + escape HTML + evidenziazione IoC).
# $1 = file, $2 = keyword separate da '|' (case-insensitive) per marcare le righe sensibili.
render_pre_block() {
    local FILE="$1" KW="$2" MODE="${3:-}"
    "$PY3" - "$FILE" "$KW" "$MODE" << 'PYEOF'
import sys, html, re, datetime
path, kw = sys.argv[1], sys.argv[2].lower()
mode = sys.argv[3] if len(sys.argv) > 3 else ''
keys = [k for k in kw.split('|') if k]

# Decodifica i timestamp UNIX nelle history di shell in formato leggibile.
# zsh extended_history:  ": <epoch>:<elapsed>;<comando>"
# bash con HISTTIMEFORMAT: una riga "#<epoch>" prima del comando
_ZSH = re.compile(r'^: (\d{9,12}):(\d+);(.*)$', re.S)
_BASH = re.compile(r'^#(\d{9,12})$')
def fmt(ep):
    try:
        return datetime.datetime.utcfromtimestamp(int(ep)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ep
def decode_histts(line):
    m = _ZSH.match(line)
    if m:
        return f"[{fmt(m.group(1))}]  {m.group(3)}"
    m = _BASH.match(line)
    if m:
        return f"[{fmt(m.group(1))}]"
    return line

try:
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('utf-8', 'replace').replace('\r\n', '\n').replace('\r', '\n')
    out = []
    for i, line in enumerate(text.split('\n'), 1):
        if mode == 'histts':
            line = decode_histts(line)
        esc = html.escape(line)
        css = 'line sensitive' if any(k in line.lower() for k in keys) else 'line'
        out.append(f'<span class="{css}"><span class="lnum">{i:5d}</span> {esc}</span>')
    print('\n'.join(out))
except Exception as e:
    print(f'<span class="line bad">{html.escape(str(e))}</span>')
PYEOF
}

# Stampa a console le righe di un file con evidenziazione IoC (rosso sulle corrispondenze).
# $1 = file, $2 = regex grep (-iE), $3 = max righe (default 200)
print_file_lines() {
    local FILE="$1" KW="$2" MAX="${3:-200}"
    [[ -f "$FILE" ]] || return
    local LN=0
    while IFS= read -r LINE || [[ -n "$LINE" ]]; do
        LN=$((LN + 1))
        if [[ $LN -gt $MAX ]]; then
            echo -e "      ${DIM}... ($(L "troncato a" "truncated at") $MAX $(L "righe" "lines"))${RESET}"
            break
        fi
        if [[ -n "$KW" ]] && printf '%s' "$LINE" | grep -qiE "$KW"; then
            printf "      ${RED}%5d  %s${RESET}\n" "$LN" "$LINE"
        else
            printf "      ${DIM}%5d${RESET}  %s\n" "$LN" "$LINE"
        fi
    done < "$FILE"
}

# Card HTML per un singolo file di testo (header con metadati + <pre> evidenziato).
# $1 = file, $2 = keyword IoC, $3 = icona (default ≣)
file_card_html() {
    local F="$1" KW="$2" ICON="${3:-≣}" MODE="${4:-}"
    local SZ MT BODY
    SZ=$(stat -c %s "$F" 2>/dev/null || echo "?")
    MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1 || echo "?")
    BODY=$(render_pre_block "$F" "$KW" "$MODE")
    printf "<div class='card' style='margin-bottom:.8rem'><div class='card-header'><div class='uicon' style='font-size:.7rem'>%s</div><div><div class='uname' style='font-size:.85rem'>%s</div><div class='upath'>%s</div></div><div style='margin-left:auto;text-align:right;font-family:var(--mono);font-size:.65rem;color:var(--text-dim)'><div class='mid'>%s</div><div>%s B</div></div></div><div class='hist-content'><pre class='hist-pre'>%s</pre></div></div>" \
        "$ICON" "$(html_esc "$(basename "$F")")" "$(html_esc "$F")" "$MT" "$SZ" "$BODY"
}

# Card HTML generica con corpo arbitrario (tabella/pre già formattati).
# $1 = titolo, $2 = sottopath, $3 = badge, $4 = corpo HTML, $5 = icona
generic_card_html() {
    printf "<div class='card'><div class='card-header'><div class='uicon'>%s</div><div class='user-info'><div class='uname'>%s</div><div class='upath'>%s</div></div><div class='badge'>%s</div></div><div style='padding:1rem 1.5rem'>%s</div></div>" \
        "${5:-▣}" "$(html_esc "$1")" "$(html_esc "$2")" "$3" "$4"
}

# Scrive il report HTML finale e lo registra.
# $1 slug · $2 titolo · $3 icona · $4 sottotitolo · $5 stats_html · $6 body_html
finish_report() {
    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "$1")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    {
        html_header "$2"
        html_page_header "$3" "$2" "$4" "$SCAN" "$WIN_ROOT"
        [[ -n "$5" ]] && printf "<div class='statsbar'>%s</div>\n" "$5"
        echo "<main>"
        pre_style_block
        printf '%s\n' "$6"
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}

# Helper per una stat della statsbar
stat_box() { printf "<div class='stat %s'><div class='label'>%s</div><div class='value'>%s</div></div>" "${3:-}" "$1" "$2"; }

# ================================================================
#  MODULI macOS
# ================================================================

# Renderizza una tabella HTML da righe tab-separated. $1=righe, $2.. = intestazioni
_rows_to_table() {
    local ROWS="$1"; shift
    local _RTMP; _RTMP=$(mktemp); printf '%s\n' "$ROWS" > "$_RTMP"
    printf '%s\n' "$@" > "${_RTMP}.h"
    "$PY3" - "$_RTMP" "${_RTMP}.h" << 'PYEOF'
import sys, html
heads=[h.rstrip('\n') for h in open(sys.argv[2])]
print("<table><tr>"+''.join(f'<th>{html.escape(h)}</th>' for h in heads)+"</tr>")
for line in open(sys.argv[1], errors='replace'):
    if not line.strip(): continue
    cells=line.rstrip('\n').split('\t')
    tds=''.join(f"<td class='mono'>{html.escape(c)}</td>" for c in cells)
    print(f"<tr>{tds}</tr>")
print("</table>")
PYEOF
    rm -f "$_RTMP" "${_RTMP}.h"
}
