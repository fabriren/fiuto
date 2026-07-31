
# ================================================================
#  MODULO 39 — AI Chat / Query History
# ================================================================
# Recupera la cronologia delle chat con assistenti AI lasciata sul
# disco dagli applicativi desktop/CLI piu' diffusi. Riferimento:
#   https://r3nzsec.github.io/irflow-timeline/dfir-tips/ai-query-history
#
# Tutti gli artefatti sono cercati sotto le home utente Windows
# (Users\<utente>\...). I formati JSON/JSONL vengono parsati e i
# messaggi estratti; i DB SQLite (state.vscdb, store.db, state.sqlite)
# vengono interrogati per le chiavi di chat; i bundle cifrati / LevelDB /
# protobuf Cascade sono solo inventariati (path, dimensione, SHA256).
module_ai_chat() {
    section_header "AI Chat / Query History" "$CYAN"
    check_win_root || return 1

    local MANIFEST; MANIFEST=$(mktemp); register_tmp "$MANIFEST"
    local IOCTMP;   IOCTMP=$(mktemp);   register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"

    TOTAL_FILES=0
    local TOTAL_USERS=0

    # Aggiunge al manifest le righe lette da stdin (una per path trovato).
    # Gira nella shell corrente (chiamata con < <(...)) cosi' i contatori
    # globali vengono aggiornati. Formato manifest: tool<TAB>kind<TAB>user<TAB>path
    _ai_collect() {
        local _tool="$1" _kind="$2" _user="$3" _f
        while IFS= read -r _f; do
            [[ -n "$_f" ]] || continue
            printf '%s\t%s\t%s\t%s\n' "$_tool" "$_kind" "$_user" "$_f" >> "$MANIFEST"
            TOTAL_FILES=$((TOTAL_FILES+1))
            _USER_FOUND=$((_USER_FOUND+1))
        done
    }

    info "$(L "Ricerca artefatti chat AI nelle home utente..." "Searching AI chat artifacts in user homes...")"

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local _USER_FOUND=0

        # ── Claude Code (CLI) ──────────────────────────────────────
        _ai_collect "Claude Code (CLI)" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.claude/history.jsonl' -type f 2>/dev/null)
        _ai_collect "Claude Code (CLI)" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 6 -ipath '*/.claude/projects/*' -iname '*.jsonl' -type f 2>/dev/null)
        # ── Claude Desktop ─────────────────────────────────────────
        _ai_collect "Claude Desktop" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 7 -ipath '*/AppData/Roaming/Claude/claude-code-sessions/*' -iname '*.json' -type f 2>/dev/null)
        _ai_collect "Claude Desktop" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 7 -ipath '*/AppData/Roaming/Claude/local-agent-mode-sessions/*' -iname '*.json' -type f 2>/dev/null)
        # ── ChatGPT Desktop (standalone) ───────────────────────────
        _ai_collect "ChatGPT Desktop" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/*' \( -iname '*.ldb' -o -iname '*.log' \) -type f 2>/dev/null)
        _ai_collect "ChatGPT Desktop" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 12 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/IndexedDB/*.indexeddb.blob/*' -type f 2>/dev/null)
        _ai_collect "ChatGPT Desktop" inventory "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/*' -iname 'conversations-v2-*' -type f 2>/dev/null)
        _ai_collect "ChatGPT Desktop" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/OpenAI/ChatGPT/*' \( -iname '*.sqlite' -o -iname '*.db' \) -type f 2>/dev/null)
        # ── ChatGPT (Microsoft Store) ──────────────────────────────
        _ai_collect "ChatGPT (Store)" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/*' \( -iname '*.ldb' -o -iname '*.log' \) -type f 2>/dev/null)
        _ai_collect "ChatGPT (Store)" leveldb "$USERNAME" < <(find "$USER_DIR" -maxdepth 14 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/IndexedDB/*.indexeddb.blob/*' -type f 2>/dev/null)
        _ai_collect "ChatGPT (Store)" inventory "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/*' -iname 'conversations-v2-*' -type f 2>/dev/null)
        _ai_collect "ChatGPT (Store)" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/AppData/Local/Packages/OpenAI.ChatGPT-Desktop_*/LocalCache/Roaming/ChatGPT/*' \( -iname '*.sqlite' -o -iname '*.db' \) -type f 2>/dev/null)
        # ── OpenAI Codex ───────────────────────────────────────────
        _ai_collect "OpenAI Codex" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.codex/history.jsonl' -type f 2>/dev/null)
        _ai_collect "OpenAI Codex" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.codex/session_index.jsonl' -type f 2>/dev/null)
        _ai_collect "OpenAI Codex" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 8 \( -ipath '*/.codex/sessions/*' -o -ipath '*/.codex/archived_sessions/*' \) -iname 'rollout-*.jsonl' -type f 2>/dev/null)
        _ai_collect "OpenAI Codex" sqlite_codex "$USERNAME" < <(find "$USER_DIR" -maxdepth 2 -ipath '*/.codex/state.sqlite' -type f 2>/dev/null)
        # ── Gemini CLI ─────────────────────────────────────────────
        _ai_collect "Gemini CLI" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 6 -ipath '*/.gemini/tmp/*/chats/*' -iname 'session-*.json' -type f 2>/dev/null)
        _ai_collect "Gemini CLI" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 5 -ipath '*/.gemini/tmp/*' \( -iname 'logs.json' -o -iname 'checkpoint-*.json' \) -type f 2>/dev/null)
        # ── Cursor ─────────────────────────────────────────────────
        _ai_collect "Cursor" jsonl "$USERNAME" < <(find "$USER_DIR" -maxdepth 7 -ipath '*/.cursor/projects/*/agent-transcripts/*' -iname '*.jsonl' -type f 2>/dev/null)
        _ai_collect "Cursor" sqlite_store "$USERNAME" < <(find "$USER_DIR" -maxdepth 6 -ipath '*/.cursor/chats/*' -iname 'store.db' -type f 2>/dev/null)
        _ai_collect "Cursor" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/Cursor/User/*Storage/*' -iname 'state.vscdb' -type f 2>/dev/null)
        # ── GitHub Copilot (VS Code / VSCodium) ────────────────────
        local _APP
        for _APP in "Code" "Code - Insiders" "VSCodium"; do
            _ai_collect "GitHub Copilot ($_APP)" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath "*/AppData/Roaming/$_APP/User/workspaceStorage/*/chatSessions/*" -iname '*.json' -type f 2>/dev/null)
            _ai_collect "GitHub Copilot ($_APP)" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath "*/AppData/Roaming/$_APP/User/globalStorage/emptyWindowChatSessions/*" -iname '*.json' -type f 2>/dev/null)
            _ai_collect "GitHub Copilot ($_APP)" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath "*/AppData/Roaming/$_APP/User/*Storage/*" -iname 'state.vscdb' -type f 2>/dev/null)
        done
        # ── Windsurf ───────────────────────────────────────────────
        _ai_collect "Windsurf" sqlite_vscdb "$USERNAME" < <(find "$USER_DIR" -maxdepth 9 -ipath '*/AppData/Roaming/Windsurf/User/*Storage/*' -iname 'state.vscdb' -type f 2>/dev/null)
        _ai_collect "Windsurf" inventory "$USERNAME" < <(find "$USER_DIR" -maxdepth 11 -ipath '*/windsurf.cascade/*' -iname '*.pb' -type f 2>/dev/null)
        # ── Continue.dev ───────────────────────────────────────────
        _ai_collect "Continue.dev" json "$USERNAME" < <(find "$USER_DIR" -maxdepth 4 -ipath '*/.continue/sessions/*' -iname '*.json' -type f 2>/dev/null)

        if [[ $_USER_FOUND -gt 0 ]]; then
            ok "$USERNAME - ${BOLD}${_USER_FOUND}${RESET} $(L "artefatti AI" "AI artifacts")"
            TOTAL_USERS=$((TOTAL_USERS+1))
        else
            dim_msg "$USERNAME - $(L "nessun artefatto AI" "no AI artifacts")"
        fi
    done < <(get_user_homes)
    unset -f _ai_collect

    separator
    info "$(L "Utenti con chat AI:" "Users with AI chats:") ${BOLD}$TOTAL_USERS${RESET}  |  $(L "File totali:" "Total files:") ${BOLD}$TOTAL_FILES"

    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun artefatto di chat AI trovato." "No AI chat artifact found.")"; return 0; }
    ask_yn "$(L "Generare report HTML?" "Generate HTML report?")" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ai_chat")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Parser Python: legge il manifest + il file IoC e produce
    #   riga 1: STATS:<file>|<utenti>|<tool>|<sensibili>|<ioc>
    #   resto : HTML delle card (contenuto di <div class='cards'>)
    local PARSER_OUT
    PARSER_OUT=$(run_py_with_lib pylib_leveldb "$MANIFEST" "$IOCTMP" << 'PYEOF'
import sys, os, re, json, html, sqlite3, shutil, tempfile, hashlib, datetime
from collections import OrderedDict

ISO_TS = re.compile(r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}')

def norm_ts(ts):
    """Normalizza un timestamp eterogeneo (ISO o epoch s/ms) in
    'YYYY-MM-DD HH:MM:SS' (UTC). Ritorna la stringa grezza se non parsabile."""
    if not ts:
        return ''
    ts = str(ts).strip()
    m = ISO_TS.search(ts)
    if m:
        return m.group(0).replace('T', ' ')[:19]
    try:
        v = float(ts)
        if v > 1e12:      # millisecondi
            v /= 1000.0
        if 1e8 < v < 1e11:
            return datetime.datetime.fromtimestamp(
                v, datetime.timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        pass
    return ts

manifest_path, ioc_path = sys.argv[1], sys.argv[2]

IOCS = []
try:
    with open(ioc_path, encoding='utf-8', errors='replace') as f:
        for ln in f:
            ln = ln.strip().lower()
            if ln and not ln.startswith('#'):
                IOCS.append(ln)
except Exception:
    pass

SENS = ['api_key','apikey','api-key','secret','password','passwd','token',
        'bearer ','authorization','sk-','ghp_','gho_','xoxb','xoxp','aws_secret',
        'private_key','-----begin','client_secret','credential','passphrase',
        'connectionstring']

MAX_MSGS = 300
MAX_TXT  = 6000

TEXT_KEYS  = ('text','content','display','message','prompt','body','parts',
              'summary','thinking','reasoning','response','rawText','value')
ROLE_KEYS  = ('role','type','sender','author','speaker','kind')
TS_KEYS    = ('timestamp','time','created_at','createdAt','ts','date',
              'updated_at','updatedAt','startTime','responseTime')
ARRAY_KEYS = ('messages','requests','history','chats','entries','items',
              'turns','conversation','events','log','records','responses')

def gstr(d, keys):
    for k in keys:
        if k in d and isinstance(d[k], (str, int, float)):
            return str(d[k])
    return ''

def collect_text(v, depth=0, acc=None):
    if acc is None: acc = []
    if depth > 6 or len(acc) > 40: return acc
    if isinstance(v, str):
        s = v.strip()
        if s: acc.append(s)
    elif isinstance(v, dict):
        for k in TEXT_KEYS:
            if k in v:
                collect_text(v[k], depth+1, acc)
    elif isinstance(v, list):
        for it in v:
            collect_text(it, depth+1, acc)
    return acc

def mk_msg(m):
    if isinstance(m, str):
        s = m.strip()
        return ('', '', s) if s else None
    if not isinstance(m, dict):
        return None
    txt = '\n'.join(collect_text(m)).strip()
    if not txt:
        return None
    role = gstr(m, ROLE_KEYS)
    # Voci della prompt-history (es. Claude history.jsonl: {"display": "..."}) →
    # sono prompt dell'utente anche se prive di campo ruolo.
    if not role and ('display' in m or 'prompt' in m):
        role = 'user'
    return (gstr(m, TS_KEYS), role, txt)

def extract(obj, out, depth=0):
    if len(out) >= MAX_MSGS or depth > 5: return
    if isinstance(obj, dict):
        handled = False
        for ak in ARRAY_KEYS:
            v = obj.get(ak)
            if isinstance(v, list) and v and isinstance(v[0], dict):
                for m in v:
                    mm = mk_msg(m)
                    if mm: out.append(mm)
                    if len(out) >= MAX_MSGS: return
                handled = True
        if not handled:
            if any(k in obj for k in TEXT_KEYS):
                mm = mk_msg(obj)
                if mm: out.append(mm)
            else:
                for vv in obj.values():
                    extract(vv, out, depth+1)
    elif isinstance(obj, list):
        if obj and isinstance(obj[0], dict) and any(
                isinstance(it, dict) and any(k in it for k in TEXT_KEYS) for it in obj[:5]):
            for m in obj:
                mm = mk_msg(m)
                if mm: out.append(mm)
                if len(out) >= MAX_MSGS: return
        else:
            for it in obj:
                extract(it, out, depth+1)

# Tipi di voce NON conversazionali nei session log (Claude Code, ecc.) da saltare.
JSONL_SKIP_TYPES = {'queue-operation', 'file-history-snapshot', 'ai-title',
                    'last-prompt', 'attachment', 'summary', 'progress'}

def parse_jsonl(path):
    out = []
    with open(path, encoding='utf-8', errors='replace') as f:
        for line in f:
            line = line.strip()
            if not line: continue
            try:
                obj = json.loads(line)
            except Exception:
                out.append(('', '', line[:MAX_TXT])); continue
            if isinstance(obj, dict) and obj.get('type') in JSONL_SKIP_TYPES:
                continue
            if isinstance(obj, dict) and any(k in obj for k in TEXT_KEYS):
                mm = mk_msg(obj)
                if mm: out.append(mm)
            else:
                extract(obj, out)
            if len(out) >= MAX_MSGS: break
    return out

def parse_json(path):
    out = []
    with open(path, encoding='utf-8', errors='replace') as f:
        obj = json.load(f)
    extract(obj, out)
    return out

def sqlite_rows(path, like_keys, tables):
    _fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd); out = []
    try:
        shutil.copy2(path, tmp)
        con = sqlite3.connect(tmp)
        con.text_factory = lambda b: b.decode('utf-8', 'replace') if isinstance(b, bytes) else b
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        present = {r[0] for r in cur.fetchall()}
        for t in tables:
            if t not in present: continue
            cur.execute("PRAGMA table_info(%s)" % t)
            cols = [c[1] for c in cur.fetchall()]
            if 'key' not in cols or 'value' not in cols: continue
            where = ' OR '.join("lower(key) LIKE ?" for _ in like_keys)
            cur.execute("SELECT key,value FROM %s WHERE %s" % (t, where),
                        tuple('%%%s%%' % k for k in like_keys))
            for k, v in cur.fetchall():
                if v is None: continue
                vs = v if isinstance(v, str) else str(v)
                sub = []
                try:
                    extract(json.loads(vs), sub)
                except Exception:
                    pass
                if sub:
                    for ts, role, txt in sub[:60]:
                        out.append((ts, role or k, txt))
                else:
                    out.append(('', k, vs[:MAX_TXT]))
                if len(out) >= MAX_MSGS: break
        con.close()
    except Exception as e:
        out.append(('', 'error', 'SQLite: %s' % e))
    finally:
        try: os.remove(tmp)
        except Exception: pass
    return out

def sqlite_meta(path):
    _fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd); out = []
    try:
        shutil.copy2(path, tmp)
        con = sqlite3.connect(tmp); cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        info = []
        for (t,) in cur.fetchall():
            try:
                cur.execute("SELECT count(*) FROM %s" % t)
                info.append('%s: %d' % (t, cur.fetchone()[0]))
            except Exception:
                info.append(t)
        con.close()
        out.append(('', 'metadata', 'Tabelle/rows: ' + ', '.join(info)))
    except Exception as e:
        out.append(('', 'error', str(e)))
    finally:
        try: os.remove(tmp)
        except Exception: pass
    return out

def sha256_short(p):
    try:
        h = hashlib.sha256()
        with open(p, 'rb') as f:
            for ch in iter(lambda: f.read(65536), b''):
                h.update(ch)
        return h.hexdigest()[:16]
    except Exception:
        return '?'


VSCDB_KEYS = ['chat','copilot','aichat','interactive.session','cascade',
              'composer','aiservice','windsurf','continue','prompt']
STORE_KEYS = ['composer','message','chat','bubble']

def parse_file(kind, path):
    if kind == 'jsonl':         return parse_jsonl(path)
    if kind == 'json':          return parse_json(path)
    if kind == 'sqlite_vscdb':  return sqlite_rows(path, VSCDB_KEYS, ['ItemTable', 'cursorDiskKV'])
    if kind == 'sqlite_store':  return sqlite_rows(path, STORE_KEYS, ['cursorDiskKV', 'ItemTable'])
    if kind == 'sqlite_codex':  return sqlite_meta(path)
    if kind == 'leveldb':       return carve_leveldb(path)
    return []  # inventory

def role_class(role):
    r = (role or '').lower()
    if 'user' in r or 'human' in r or 'prompt' in r: return 'r-user'
    if 'assist' in r or 'model' in r or 'gemini' in r or 'ai' in r or 'response' in r: return 'r-ai'
    if 'tool' in r or 'function' in r or 'shell' in r: return 'r-tool'
    if 'error' in r: return 'r-err'
    return 'r-meta'

def ai_product(tool):
    """Nome leggibile dell'assistente AL dal tool del record."""
    t = (tool or '').lower()
    if 'chatgpt' in t or ('openai' in t and 'codex' not in t): return 'ChatGPT'
    if 'copilot' in t:  return 'Copilot'
    if 'claude' in t:   return 'Claude'
    if 'gemini' in t:   return 'Gemini'
    if 'codex' in t:    return 'Codex'
    if 'cursor' in t:   return 'Cursor'
    if 'windsurf' in t: return 'Windsurf'
    if 'continue' in t: return 'Continue'
    return re.sub(r'\s*\(.*\)$', '', tool).strip() or 'AI'

def render(msgs, ai='AI'):
    rows, sens_n, ioc_n = [], 0, 0
    for ts, role, txt in msgs:
        t = txt[:MAX_TXT]
        low = t.lower()
        is_sens = any(s in low for s in SENS)
        is_ioc  = bool(IOCS) and any(i in low for i in IOCS)
        if is_sens: sens_n += 1
        if is_ioc:  ioc_n += 1
        rc = role_class(role)
        cls = 'msg ' + rc
        if is_sens: cls += ' sensitive'
        if is_ioc:  cls += ' ioc'
        if rc == 'r-user':
            label = 'UTENTE'
        elif rc == 'r-ai':
            label = 'AI - ' + html.escape(ai)
        else:
            label = html.escape(role) if role else '-'
        nts   = norm_ts(ts)
        meta  = html.escape(nts) if nts else ''
        tag   = ''
        if is_ioc:  tag += "<span class='tag tag-ioc'>IOC</span>"
        if is_sens: tag += "<span class='tag tag-sec'>SENS</span>"
        rows.append(
            "<div class='%s'><div class='msg-h'><span class='role'>%s</span>%s"
            "<span class='mts'>%s</span></div><pre class='msg-t'>%s</pre></div>"
            % (cls, label, tag, meta, html.escape(t)))
    return rows, sens_n, ioc_n

data = OrderedDict()
with open(manifest_path, encoding='utf-8', errors='replace') as f:
    for line in f:
        line = line.rstrip('\n')
        if not line: continue
        parts = line.split('\t')
        if len(parts) < 4: continue
        tool, kind, user = parts[0], parts[1], parts[2]
        path = '\t'.join(parts[3:])
        data.setdefault(user, OrderedDict()).setdefault(tool, []).append((kind, path))

tot_files = tot_sens = tot_ioc = 0
tools_seen = set()
cards = []

for user, tools in data.items():
    tool_blocks = []
    user_files = 0
    for tool, files in tools.items():
        tools_seen.add(tool)
        file_blocks = []
        for kind, path in files:
            tot_files += 1; user_files += 1
            try:
                st = os.stat(path)
                size = st.st_size
                mt = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                size, mt = 0, '?'
            try:
                msgs = parse_file(kind, path)
            except Exception as e:
                msgs = [('', 'error', '%s' % e)]
            rows, sn, ic = render(msgs, ai_product(tool))
            tot_sens += sn; tot_ioc += ic
            sha = sha256_short(path)
            base = html.escape(os.path.basename(path))
            meta = ("<span>%s</span> &middot; <span>%d B</span> &middot; "
                    "<span>mtime %s</span> &middot; <span>sha256 %s</span>"
                    % (html.escape(kind), size, mt, sha))
            badges = "<span class='fb'>%d msg</span>" % len(rows)
            if sn: badges += "<span class='fb fb-sec'>%d sens</span>" % sn
            if ic: badges += "<span class='fb fb-ioc'>%d ioc</span>" % ic
            if rows:
                note = ''
                if kind == 'leveldb':
                    note = ("<div class='inv'>%s</div>" % html.escape(
                        "Contenuto ricavato per string-carving dal LevelDB "
                        "(best-effort, non strutturato - verificare manualmente)."))
                body = note + "<div class='thread'>%s</div>" % ''.join(rows)
            elif kind in ('inventory', 'pb_inventory'):
                body = ("<div class='inv'>%s</div>" %
                        html.escape("Solo inventario - contenuto cifrato/proprietario non decodificato "
                                    "(LevelDB / bundle conversations-v2 / protobuf Cascade)."))
            else:
                body = "<div class='inv'>%s</div>" % html.escape("Nessun messaggio estratto.")
            file_blocks.append(
                "<div class='fblk'><div class='fhdr'><div class='fname'>%s</div>"
                "<div class='fbadges'>%s</div></div><div class='fpath'>%s</div>"
                "<div class='fmeta'>%s</div>%s</div>"
                % (base, badges, html.escape(path), meta, body))
        tool_blocks.append(
            "<div class='toolsec'><div class='toolname'>%s <span class='dim'>(%d)</span></div>%s</div>"
            % (html.escape(tool), len(files), ''.join(file_blocks)))
    cards.append(
        "<div class='card'><div class='card-header'><div class='uicon'>AI</div>"
        "<div class='user-info'><div class='uname'>%s</div>"
        "<div class='upath'>%d file</div></div>"
        "<div class='badge'>%d tool</div></div>"
        "<div style='padding:1rem 1.5rem'>%s</div></div>"
        % (html.escape(user), user_files, len(tools), ''.join(tool_blocks)))

print("STATS:%d|%d|%d|%d|%d" % (tot_files, len(data), len(tools_seen), tot_sens, tot_ioc))
sys.stdout.write(''.join(cards))
PYEOF
)

    local STATS_LINE; STATS_LINE=$(printf '%s\n' "$PARSER_OUT" | head -1)
    local CARDS_HTML; CARDS_HTML=$(printf '%s\n' "$PARSER_OUT" | tail -n +2)
    local _S="${STATS_LINE#STATS:}"
    local ST_FILES ST_USERS ST_TOOLS ST_SENS ST_IOC
    IFS='|' read -r ST_FILES ST_USERS ST_TOOLS ST_SENS ST_IOC <<< "$_S"

    {
        html_header "AI Chat History" "#a371f7"
        html_page_header "AI" "AI Chat / <span>Query History</span> Forensics" \
            "Claude · ChatGPT · Copilot · Cursor · Gemini · Codex · Windsurf · Continue" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>"
        echo "  <div class='stat'><div class='label'>$(L "File" "Files")</div><div class='value'>${ST_FILES:-0}</div></div>"
        echo "  <div class='stat info'><div class='label'>$(L "Utenti" "Users")</div><div class='value'>${ST_USERS:-0}</div></div>"
        echo "  <div class='stat info'><div class='label'>Tool</div><div class='value'>${ST_TOOLS:-0}</div></div>"
        echo "  <div class='stat'><div class='label'>$(L "Sensibili" "Sensitive")</div><div class='value'>${ST_SENS:-0}</div></div>"
        echo "  <div class='stat ok'><div class='label'>IoC</div><div class='value'>${ST_IOC:-0}</div></div>"
        echo "</div>"
        echo "<main><div class='stitle'>$(L "Risultati per utente" "Results per user")</div>"
        cat << 'CSSEOF'
<style>
  .toolsec{margin-bottom:1rem;border:1px solid var(--border);border-radius:6px;overflow:hidden}
  .toolname{font-family:var(--mono);font-size:.8rem;font-weight:700;color:var(--accent);
    background:var(--bg3);padding:.5rem .9rem;border-bottom:1px solid var(--border)}
  .toolname .dim{color:var(--text-dim);font-weight:400}
  .fblk{padding:.7rem .9rem;border-bottom:1px solid rgba(30,45,61,.5)}
  .fblk:last-child{border-bottom:none}
  .fhdr{display:flex;align-items:center;gap:1rem}
  .fname{font-family:var(--mono);font-size:.78rem;color:#fff;font-weight:500}
  .fbadges{margin-left:auto;display:flex;gap:.35rem}
  .fb{font-family:var(--mono);font-size:.6rem;padding:.1rem .5rem;border-radius:2px;
    background:rgba(88,166,255,.1);border:1px solid rgba(88,166,255,.3);color:var(--accent)}
  .fb-sec{background:rgba(240,136,62,.1);border-color:rgba(240,136,62,.4);color:var(--accent4)}
  .fb-ioc{background:rgba(255,123,114,.12);border-color:rgba(255,123,114,.4);color:var(--accent2)}
  .fpath{font-family:var(--mono);font-size:.6rem;color:var(--text-dim);word-break:break-all;margin:.2rem 0}
  .fmeta{font-family:var(--mono);font-size:.6rem;color:var(--text-mid);margin-bottom:.5rem}
  .thread{display:flex;flex-direction:column;gap:.4rem;max-height:520px;overflow-y:auto;
    padding:.3rem;background:var(--bg);border-radius:4px}
  .thread::-webkit-scrollbar{width:5px}
  .thread::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
  .msg{border-left:3px solid var(--border);padding:.35rem .6rem;border-radius:0 3px 3px 0;background:var(--bg2)}
  .msg-h{display:flex;align-items:center;gap:.5rem;margin-bottom:.2rem}
  .role{font-family:var(--mono);font-size:.6rem;text-transform:uppercase;letter-spacing:.08em;color:var(--text-mid)}
  .mts{margin-left:auto;font-family:var(--mono);font-size:.58rem;color:var(--text-dim)}
  .msg-t{font-family:var(--mono);font-size:.72rem;line-height:1.5;white-space:pre-wrap;word-break:break-word;color:var(--text)}
  .r-user{border-left:3px solid var(--accent);background:rgba(88,166,255,.08);margin-left:2.5rem}
  .r-user .role{color:var(--accent)}
  .r-user .role::before{content:'\1F464  '}
  .r-ai{border-left:3px solid var(--accent3);background:rgba(63,185,80,.07);margin-right:2.5rem}
  .r-ai .role{color:var(--accent3)}
  .r-ai .role::before{content:'\1F916  '}
  .r-tool{border-left-color:var(--accent4)}
  .r-tool .role{color:var(--accent4)}
  .r-err{border-left-color:var(--accent2)}
  .msg.sensitive{background:rgba(240,136,62,.07);border-left-color:var(--accent4)}
  .msg.ioc{background:rgba(255,123,114,.1);border-left-color:var(--accent2)}
  .tag{font-family:var(--mono);font-size:.55rem;padding:0 .35rem;border-radius:2px;font-weight:700}
  .tag-sec{background:rgba(240,136,62,.2);color:var(--accent4)}
  .tag-ioc{background:rgba(255,123,114,.2);color:var(--accent2)}
  .inv{font-family:var(--mono);font-size:.68rem;color:var(--text-dim);font-style:italic;padding:.3rem 0}
</style>
CSSEOF
        echo "<div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
