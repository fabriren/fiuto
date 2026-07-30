
# ================================================================
#  LIBRERIA PYTHON CONDIVISA — LevelDB / Snappy
#
#  Le app Electron (ChatGPT Desktop, Slack, Discord, Teams) memorizzano i
#  messaggi in LevelDB: SSTable .ldb e write-ahead log .log, con i data-block
#  compressi in Snappy e i singoli valori IndexedDB compressi una seconda
#  volta da Blink.
#
#  Questo codice nasceva dentro il modulo 39 (AI Chat). Serve identico al
#  modulo 48 (chat desktop), quindi vive qui: duplicarlo avrebbe significato
#  correggere ogni bug in due punti.
#
#  Uso da un modulo:
#      OUT=$(run_py_with_lib pylib_leveldb "$ARG" << 'PYEOF'
#      import sys
#      for _, role, text in carve_leveldb(sys.argv[1]):
#          ...
#      PYEOF
#      )
# ================================================================

# Emette il sorgente della libreria sullo stdout.
pylib_leveldb() {
    cat << 'FIUTO_PYLIB_EOF'
import re

# Lunghezza massima del testo estratto per messaggio. Il modulo chiamante puo'
# ridefinirla: la sua assegnazione viene dopo questa e quindi prevale.
MAX_TXT = 6000

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
FIUTO_PYLIB_EOF
}

# Esegue uno script Python letto da stdin, anteponendogli una libreria condivisa.
#
# Concatenare invece di importare tiene il programma finale identico a com'era
# quando la libreria stava dentro il modulo: stesso spazio dei nomi, nessuna
# differenza di comportamento da dimostrare.
#
#   run_py_with_lib <funzione_libreria> [argomenti...] << 'PYEOF' ... PYEOF
run_py_with_lib() {
    local LIBF="$1"; shift
    local TMP; TMP=$(mktemp) || return 1
    { "$LIBF"; printf '\n'; cat; } > "$TMP"
    "$PY3" "$TMP" "$@"
    local RC=$?
    rm -f "$TMP"
    return $RC
}
