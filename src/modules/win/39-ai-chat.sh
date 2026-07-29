
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
            ok "$USERNAME — ${BOLD}${_USER_FOUND}${RESET} $(L "artefatti AI" "AI artifacts")"
            TOTAL_USERS=$((TOTAL_USERS+1))
        else
            dim_msg "$USERNAME — $(L "nessun artefatto AI" "no AI artifacts")"
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
    PARSER_OUT=$("$PY3" - "$MANIFEST" "$IOCTMP" << 'PYEOF'
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

# Carving di stringhe da file LevelDB (.ldb/.log) usati da ChatGPT Desktop
# (IndexedDB / Local Storage). I file SSTable non sono parsabili senza la
# libreria leveldb, ma i messaggi delle conversazioni sono memorizzati come
# stringhe in chiaro UTF-8/UTF-16. Estraiamo le run leggibili e teniamo SOLO
# quelle che sembrano linguaggio naturale (prosa), scartando i frammenti
# JSON/serializzati di configurazione (feature flag, chiavi JWK, metadati) che
# popolano soprattutto il Local Storage e non hanno valore investigativo.
MIN_LDB = 6

# Marcatori di dato serializzato: "key":  ,"  :{  :[  }] ,  ecc.
_JSON_FRAG = re.compile(r'["\}\]]\s*[:,]|[:,]\s*["\{\[]')
_WORD      = re.compile(r'[A-Za-zÀ-ÿ]{2,}')
# Prefisso V8/IndexedDB: chiave nota + " + 1 byte-tag (rumore di serializzazione)
_VPFX = re.compile(
    r'^(?:text|parts|content|content_type|message|title|name|author|value|role)".')
# Byte di controllo C0/C1 (tranne tab/CR/LF): tag e varint di lunghezza V8 che
# capitano stampabili e sporcano il testo (es. \x02 davanti a un messaggio).
_CTRL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')
# Segnali di risposta AI nel contenuto (markdown/struttura): grassetto, heading,
# liste, code fence, tabelle, link/citazioni. Più affidabile dell'encoding per
# distinguere assistente da utente (un prompt utente con € o emoji è comunque UTF-16).
_AI_SIG = re.compile(
    r'\*\*.+?\*\*|^#{1,6}\s|^\s*[-*]\s|^\s*\d+\.\s|```|\|.+\||\]\(https?://',
    re.MULTILINE)
# Marcatori dei record di metadati IndexedDB (id, utenti, flag) → mai nella prosa
_META = re.compile(
    r'accountUserId|authUserId|isArchived|client-created-root|conversationId|'
    r'currentNode|asyncStatus|updateTime[A-Z]|messagesa[A-Z]|gizmoId|'
    r'id"\$?[0-9a-f]{8}-[0-9a-f]{4}-')

def _ldb_interesting(s):
    if len(s) < 12:
        return False
    # Una frase reale contiene spazi → scarta blob base64/hex/cert/token
    if ' ' not in s:
        return False
    # Scarta i record di metadati IndexedDB (id conversazione, userId, flag…)
    if _META.search(s):
        return False
    # Scarta frammenti JSON / config / JWK / feature-flag
    if _JSON_FRAG.search(s):
        return False
    if s.count('"') >= 2 and (':' in s or ',' in s):
        return False
    # Richiede almeno 3 parole "vere" → frase, non token isolato
    if len(_WORD.findall(s)) < 3:
        return False
    # Prevalenza di caratteri alfabetici → testo, non blob/codice
    alpha = sum(c.isalpha() for c in s)
    if alpha / max(len(s), 1) < 0.55:
        return False
    return True

def _carve_bytes(data):
    """Estrae le run di testo stampabile da un buffer, restituendole in ordine
    di posizione. Nei valori V8/IndexedDB i prompt utente sono spesso one-byte
    (Latin-1) e le risposte AI two-byte (UTF-16, per markdown/caratteri speciali):
    mescolandole per offset la conversazione si legge in ordine cronologico."""
    # Caratteri base di una run: stampabili ASCII + supplemento Latin-1
    # (accenti à,è,é,ù…) + tab/CR/LF (i messaggi multi-riga restano interi).
    # One-byte = Latin-1, two-byte = UTF-16LE (high byte 0x00).
    def _dec(s, e, enc):
        return data[s:e].decode('latin-1' if enc == 'l' else 'utf-16-le', errors='replace')
    runs = []  # [start, end, enc, text, is_meta]
    for m in re.finditer(rb'[\x09\x0a\x0d\x20-\x7e\xa0-\xff]{%d,}' % MIN_LDB, data):
        t = _dec(m.start(), m.end(), 'l')
        runs.append([m.start(), m.end(), 'l', t, bool(_META.search(t))])
    for m in re.finditer(rb'(?:[\x09\x0a\x0d\x20-\x7e\xa0-\xff]\x00){%d,}' % MIN_LDB, data):
        t = _dec(m.start(), m.end(), 'u')
        runs.append([m.start(), m.end(), 'u', t, bool(_META.search(t))])
    runs.sort(key=lambda x: x[0])
    # Fusione di run di CONTENUTO adiacenti, stesso encoding, gap ≤12 byte: è
    # quasi sempre un simbolo Unicode (→ • ✓ —, emoji) o un a-capo che spezza il
    # testo; ri-decodificando lo span viene assorbito. NON si fonde mai attraverso
    # un record di metadati (id/userId/flag), per non trascinare via il testo utente.
    merged = []
    for r in runs:
        if merged:
            last = merged[-1]
            gap = r[0] - last[1]
            if (not r[4] and not last[4] and r[2] == last[2] and 0 <= gap <= 12
                    and (r[2] == 'l' or (gap % 2 == 0 and last[0] % 2 == r[0] % 2))):
                last[1] = r[1]
                last[3] = _dec(last[0], last[1], last[2])
                continue
        merged.append(r)
    # Ritorna (testo, encoding): one-byte 'l' ≈ prompt utente (testo semplice),
    # two-byte 'u' ≈ risposta AI (markdown/simboli/emoji → UTF-16). Euristica.
    return [(r[3], r[2]) for r in merged]

# ── Decompressione Snappy (opzionale) ──────────────────────────────
# I blocchi dati delle SSTable LevelDB sono spesso compressi con Snappy
# (raw block format). Senza decomprimerli, gran parte del testo delle
# conversazioni resta non carvabile. Proviamo python-snappy, poi cramjam.
_SNAPPY = None
try:
    import snappy as _sn
    _SNAPPY = ('snappy', _sn)
except Exception:
    try:
        import cramjam as _cj
        _SNAPPY = ('cramjam', _cj)
    except Exception:
        _SNAPPY = None

# Decompressore Snappy "raw" in puro Python (nessuna dipendenza): preambolo
# varint con la lunghezza non compressa, poi tag literal/copy LZ77. Usato come
# fallback quando python-snappy/cramjam non sono installati nel Python in uso.
def _snappy_py(buf):
    try:
        n = len(buf); pos = 0
        shift = 0
        while pos < n:                       # preambolo: lunghezza (scartata)
            b = buf[pos]; pos += 1
            if not (b & 0x80):
                break
            shift += 7
        out = bytearray()
        while pos < n:
            tag = buf[pos]; pos += 1
            t = tag & 0x03
            if t == 0:                       # literal
                length = (tag >> 2) + 1
                if length > 60:
                    nb = length - 60
                    length = 1
                    for i in range(nb):
                        length += buf[pos + i] << (8 * i)
                    pos += nb
                out += buf[pos:pos + length]
                pos += length
            else:
                if t == 1:                   # copy, offset 1 byte
                    length = ((tag >> 2) & 0x07) + 4
                    offset = ((tag >> 5) << 8) | buf[pos]; pos += 1
                elif t == 2:                 # copy, offset 2 byte
                    length = (tag >> 2) + 1
                    offset = buf[pos] | (buf[pos + 1] << 8); pos += 2
                else:                        # copy, offset 4 byte
                    length = (tag >> 2) + 1
                    offset = (buf[pos] | (buf[pos + 1] << 8) |
                              (buf[pos + 2] << 16) | (buf[pos + 3] << 24)); pos += 4
                if offset == 0 or offset > len(out):
                    break
                start = len(out) - offset
                for i in range(length):      # copia byte-a-byte (gestisce overlap)
                    out.append(out[start + i])
        return bytes(out)
    except Exception:
        return b''

def _snappy_raw(buf):
    if _SNAPPY is not None:
        kind, mod = _SNAPPY
        try:
            if kind == 'snappy':
                return mod.uncompress(buf)
            return bytes(mod.snappy.decompress_raw(buf))
        except Exception:
            pass
    return _snappy_py(buf)

def _uvarint(buf, pos):
    result = shift = 0
    while pos < len(buf):
        b = buf[pos]; pos += 1
        result |= (b & 0x7f) << shift
        if not (b & 0x80):
            return result, pos
        shift += 7
    return result, pos

# Legge un blocco SSTable a (offset,size); il trailer è 1 byte tipo + 4 byte CRC.
def _read_block(data, off, size):
    if off < 0 or off + size > len(data):
        return b''
    raw = data[off:off + size]
    ctype = data[off + size] if off + size < len(data) else 0
    if ctype == 0:
        return raw
    if ctype == 1:                       # Snappy
        return _snappy_raw(raw)
    return b''                            # zstd/altro non gestito

# Estrae i valori (record) da un blocco SSTable (formato entry + restart array).
def _block_values(block):
    if len(block) < 4:
        return []
    num_restarts = int.from_bytes(block[-4:], 'little')
    restart_start = len(block) - 4 - num_restarts * 4
    if restart_start < 0:
        return []
    pos = 0; last_key = b''; vals = []
    while pos < restart_start and len(vals) < 100000:
        shared, pos = _uvarint(block, pos)
        nonshared, pos = _uvarint(block, pos)
        vlen, pos = _uvarint(block, pos)
        key = last_key[:shared] + block[pos:pos + nonshared]; pos += nonshared
        value = block[pos:pos + vlen]; pos += vlen
        last_key = key
        vals.append(value)
    return vals

_LDB_MAGIC = 0xdb4775248b80fb57

def _printable_ratio(b, n=4096):
    s = b[:n]
    if not s:
        return 0.0
    return sum(0x20 <= c < 0x7f or c in (9, 10, 13) for c in s) / len(s)

# Chromium/Blink comprime i singoli valori IndexedDB con Snappy, dietro un
# breve header wrapper (lunghezza variabile). Proviamo a decomprimere saltando
# 0..15 byte iniziali; se otteniamo testo plausibile e più lungo, lo usiamo.
def _decompress_value(v):
    if len(v) < 32:
        return v
    for k in range(0, 16):
        d = _snappy_raw(v[k:])
        # Soglia bassa: i valori IndexedDB V8 contengono stringhe UTF-16
        # (byte \x00 alternati) che abbassano la quota di stampabili.
        if d and len(d) > len(v) * 1.2 and _printable_ratio(d) > 0.45:
            return d
    return v

# Estrae il contenuto applicativo da una SSTable (.ldb): decomprime i data-block
# (livello LevelDB) e poi decomprime i singoli valori dei record (livello Blink).
# Best-effort: b'' se il file non è una SSTable valida.
def _sstable_data(data):
    if len(data) < 48:
        return b''
    footer = data[-48:]
    if int.from_bytes(footer[40:48], 'little') != _LDB_MAGIC:
        return b''
    try:
        pos = 0
        _mi_off, pos = _uvarint(footer, pos); _mi_sz, pos = _uvarint(footer, pos)
        ix_off, pos = _uvarint(footer, pos); ix_sz, pos = _uvarint(footer, pos)
        index_block = _read_block(data, ix_off, ix_sz)
        out = bytearray()
        for h in _block_values(index_block):
            p = 0
            off, p = _uvarint(h, p); sz, p = _uvarint(h, p)
            blk = _read_block(data, off, sz)        # livello LevelDB (Snappy/none)
            if not blk:
                continue
            for v in _block_values(blk):            # valori dei record
                out += _decompress_value(v) + b'\x00'   # livello Blink (Snappy)
                if len(out) > 96 * 1024 * 1024:
                    return bytes(out)
        return bytes(out)
    except Exception:
        return b''

# Estrae i valori scritti da un write-ahead log LevelDB (.log): blocchi fisici
# da 32 KB con record header crc(4)+len(2)+type(1); i record logici (FULL o
# FIRST/MIDDLE/LAST) sono WriteBatch (seq 8 + count 4 + voci put/delete). Per
# ogni PUT decomprimiamo il valore (livello Blink). Best-effort.
_WAL_BLOCK = 32768

def _wal_data(data):
    try:
        logical, cur, pos, n = [], bytearray(), 0, len(data)
        while pos + 7 <= n:
            off = pos % _WAL_BLOCK
            if _WAL_BLOCK - off < 7:          # trailer di blocco → blocco succ.
                pos += _WAL_BLOCK - off
                continue
            length = data[pos + 4] | (data[pos + 5] << 8)
            rtype = data[pos + 6]
            payload = data[pos + 7:pos + 7 + length]
            pos += 7 + length
            if rtype == 1:                     # FULL
                logical.append(bytes(payload))
            elif rtype == 2:                   # FIRST
                cur = bytearray(payload)
            elif rtype == 3:                   # MIDDLE
                cur += payload
            elif rtype == 4:                   # LAST
                cur += payload; logical.append(bytes(cur)); cur = bytearray()
            elif length == 0:                  # padding → blocco successivo
                pos = (pos // _WAL_BLOCK + 1) * _WAL_BLOCK
        out = bytearray()
        for rec in logical:
            p, L = 12, len(rec)                # salta seq(8)+count(4)
            while p < L:
                t = rec[p]; p += 1
                if t not in (0, 1):
                    break                      # formato inatteso
                klen, p = _uvarint(rec, p); p += klen      # salta key
                if t == 1:                     # kTypeValue
                    vlen, p = _uvarint(rec, p)
                    out += _decompress_value(rec[p:p + vlen]) + b'\x00'
                    p += vlen
                if len(out) > 96 * 1024 * 1024:
                    return bytes(out)
        return bytes(out)
    except Exception:
        return b''

def carve_leveldb(path, cap=2500):
    try:
        with open(path, 'rb') as f:
            data = f.read(64 * 1024 * 1024)
    except Exception as e:
        return [('', 'error', str(e))]
    # Decomprimiamo il contenuto e carviamo SOLO quello (pulito): per le SSTable
    # (.ldb) i data-block, per i write-ahead log (.log) i record WriteBatch. In
    # entrambi i casi i valori IndexedDB vengono poi de-comprimati (livello Blink).
    # Il carving del file grezzo darebbe frammenti spezzati → solo come fallback.
    low = path.lower()
    if low.endswith('.ldb'):
        blocks = _sstable_data(data)
    elif low.endswith('.log'):
        blocks = _wal_data(data)
    else:
        # File blob esterno IndexedDB (valore grande wrappato+Snappy da Blink):
        # un singolo valore → proviamo a decomprimerlo direttamente.
        dv = _decompress_value(data)
        blocks = dv if len(dv) > len(data) else b''
    cand = _carve_bytes(blocks) if blocks else _carve_bytes(data)
    out, seen = [], set()
    for s, enc in cand:
        s = s.strip()
        # Strip conservativo del prefisso V8: chiave IndexedDB nota + " + 1
        # byte-tag (es. text"T… → non…). Solo se resta del testo.
        s2 = _VPFX.sub('', s)
        if s2:
            s = s2.strip()
        # Rimuove i byte di controllo residui (tag/varint V8) ovunque nel testo.
        s = _CTRL.sub('', s).strip()
        # Rimuove il tag V8 di fine-oggetto '{' (kEndJSObject) in coda al messaggio.
        s = re.sub(r'\s*\{+\s*$', '', s)
        if not _ldb_interesting(s):
            continue
        key = s.lower()[:80]
        if key in seen:
            continue
        seen.add(key)
        # Ruolo dal contenuto: la presenza di markdown/struttura indica una
        # risposta AI; altrimenti è un prompt utente (più robusto dell'encoding).
        role = 'assistant' if _AI_SIG.search(s) else 'user'
        out.append(('', role, s[:MAX_TXT]))
        if len(out) >= cap:
            break
    if not out:
        out.append(('', 'info',
                    'Nessuna stringa leggibile estratta (file vuoto, binario o solo metadati).'))
    return out

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
            label = 'AI — ' + html.escape(ai)
        else:
            label = html.escape(role) if role else '—'
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
                        "(best-effort, non strutturato — verificare manualmente)."))
                body = note + "<div class='thread'>%s</div>" % ''.join(rows)
            elif kind in ('inventory', 'pb_inventory'):
                body = ("<div class='inv'>%s</div>" %
                        html.escape("Solo inventario — contenuto cifrato/proprietario non decodificato "
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
