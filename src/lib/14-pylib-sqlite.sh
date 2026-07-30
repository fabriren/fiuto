
# ================================================================
#  LIBRERIA PYTHON CONDIVISA — recupero da spazio libero SQLite
#
#  query_sqlite() legge i record vivi (e copia il -wal per vedere le
#  transazioni non consolidate), ma un record CANCELLATO non sparisce dal
#  file: la sua pagina finisce nella freelist, o lo spazio che occupava
#  diventa un freeblock dentro la pagina. Il contenuto resta li' finche' non
#  viene sovrascritto.
#
#  Su un artefatto come la cronologia del browser, TCC.db o chat.db questo e'
#  spesso l'unico modo di vedere cio' che l'utente ha cancellato.
#
#  Livello di supporto: si estraggono le STRINGHE dallo spazio non allocato,
#  non i record ricostruiti. Ricostruire un record cancellato richiede lo
#  schema e l'interpretazione dei serial type: e' possibile ma fragile, e un
#  record ricomposto male in un report forense e' peggio di nessun record.
#
#  Uso:
#      run_py_with_lib pylib_sqlite_recover "$DB" << 'PYEOF'
#      import sys
#      for kind, text in sqlite_carve(sys.argv[1]):
#          ...
#      PYEOF
# ================================================================

pylib_sqlite_recover() {
    cat << 'FIUTO_SQLPY_EOF'
import re
import struct

# Lunghezza minima di una stringa perche' valga la pena riportarla.
SQLITE_MIN_STR = 5

_PRINTABLE = re.compile(rb'[\x20-\x7e\xc0-\xf4][\x20-\x7e\x80-\xbf]{%d,}' % (SQLITE_MIN_STR - 1))


def _pages(data):
    """Genera (numero_pagina, byte) per ogni pagina del database."""
    if len(data) < 100 or data[:15] != b'SQLite format 3':
        return
    page_size = struct.unpack_from('>H', data, 16)[0]
    # 1 significa 65536 (il campo e' a 16 bit).
    if page_size == 1:
        page_size = 65536
    if page_size < 512 or page_size & (page_size - 1):
        return
    total = len(data) // page_size
    for i in range(total):
        yield i + 1, data[i * page_size:(i + 1) * page_size]


def _freelist_pages(data):
    """Numeri delle pagine nella freelist, seguendo la catena dei trunk."""
    out = set()
    if len(data) < 100:
        return out
    page_size = struct.unpack_from('>H', data, 16)[0]
    if page_size == 1:
        page_size = 65536
    try:
        trunk = struct.unpack_from('>I', data, 32)[0]
        count = struct.unpack_from('>I', data, 36)[0]
    except Exception:
        return out
    seen = set()
    guard = 0
    while trunk and trunk not in seen and guard < 100000:
        guard += 1
        seen.add(trunk)
        off = (trunk - 1) * page_size
        if off < 0 or off + 8 > len(data):
            break
        out.add(trunk)
        try:
            nxt, n = struct.unpack_from('>II', data, off)
        except Exception:
            break
        n = min(n, (page_size - 8) // 4)
        for i in range(n):
            try:
                leaf = struct.unpack_from('>I', data, off + 8 + 4 * i)[0]
            except Exception:
                break
            if leaf:
                out.add(leaf)
        trunk = nxt
    return out


def _unallocated(page, is_first):
    """Spazio non allocato e freeblock di una pagina b-tree.

    Layout: header, array dei puntatori alle celle, spazio libero, contenuto
    delle celle. Cio' che sta fra la fine dell'array e l'inizio del contenuto
    non e' in uso — ed e' li' che restano i record cancellati.
    """
    base = 100 if is_first else 0
    if len(page) < base + 8:
        return b''
    ptype = page[base]
    if ptype not in (0x02, 0x05, 0x0a, 0x0d):
        return b''
    hdr = 12 if ptype in (0x02, 0x05) else 8
    try:
        ncells = struct.unpack_from('>H', page, base + 3)[0]
        content = struct.unpack_from('>H', page, base + 5)[0]
        freeblk = struct.unpack_from('>H', page, base + 1)[0]
    except Exception:
        return b''
    if content == 0:
        content = 65536
    start = base + hdr + 2 * ncells
    out = bytearray()
    if 0 < start < content <= len(page):
        out += page[start:content]
    # Catena dei freeblock: spazio liberato dentro l'area del contenuto.
    guard = 0
    while freeblk and freeblk + 4 <= len(page) and guard < 10000:
        guard += 1
        try:
            nxt, size = struct.unpack_from('>HH', page, freeblk)
        except Exception:
            break
        if size and freeblk + size <= len(page):
            out += page[freeblk:freeblk + size]
        if nxt <= freeblk:
            break
        freeblk = nxt
    return bytes(out)


def sqlite_carve(path, cap=20000):
    """Ritorna [(origine, stringa)] recuperate dallo spazio non allocato.

    origine: 'freelist' per le pagine liberate per intero, 'unallocated' per
    lo spazio non usato dentro pagine ancora in uso.
    """
    try:
        with open(path, 'rb') as fh:
            data = fh.read(512 * 1024 * 1024)
    except Exception:
        return []
    if data[:15] != b'SQLite format 3':
        return []

    free = _freelist_pages(data)
    out = []
    seen = set()

    def harvest(blob, origin):
        for m in _PRINTABLE.finditer(blob):
            # errors='ignore' e non 'strict': il pattern ammette i byte di
            # continuazione UTF-8 (\x80-\xbf) per non spezzare i caratteri
            # accentati, ma in una pagina SQLite l'header del record segue il
            # payload, quindi il match finisce quasi sempre con un byte di
            # continuazione isolato. Con 'strict' l'intera stringa veniva
            # scartata in silenzio — e sono proprio quelle interessanti.
            s = m.group(0).decode('utf-8', 'ignore')
            # Via i residui di controllo dell'header del record.
            s = ''.join(ch for ch in s if ch >= ' ' and ch != '\x7f').strip()
            if len(s) < SQLITE_MIN_STR:
                continue
            # Scarta le sequenze senza alcuna lettera: sono quasi sempre
            # residui binari che capitano nell'intervallo stampabile.
            if not re.search(r'[A-Za-z]{3}', s):
                continue
            k = s.lower()
            if k in seen:
                continue
            seen.add(k)
            out.append((origin, s[:500]))
            if len(out) >= cap:
                return True
        return False

    for num, page in _pages(data):
        if len(out) >= cap:
            break
        if num in free:
            if harvest(page, 'freelist'):
                break
        else:
            blob = _unallocated(page, num == 1)
            if blob and harvest(blob, 'unallocated'):
                break
    return out
FIUTO_SQLPY_EOF
}
