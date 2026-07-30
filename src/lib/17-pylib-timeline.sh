
# ================================================================
#  LIBRERIA PYTHON CONDIVISA — estrazione eventi dai report HTML
#
#  L'HTML gia' prodotto dai moduli e' l'unico punto in cui il formato e'
#  omogeneo per tutti e 91: ogni modulo emette tabelle e blocchi <pre>, non
#  esiste una struttura dati comune a monte. Estrarre gli eventi da li' e'
#  quello che permette all'export JSONL e all'executive summary di coprire
#  l'intero toolkit senza toccare i moduli uno per uno.
#
#  Il codice nasceva dentro export_report_jsonl. Serve identico all'executive
#  summary, che deve poter correlare gli eventi anche quando --jsonl non e'
#  stato chiesto: vive qui perche' due copie divergerebbero al primo bugfix.
#
#  Uso da un modulo:
#      run_py_with_lib pylib_timeline "$HTML" << 'PYEOF'
#      import sys
#      for ev in extract_events(sys.argv[1]):
#          print(ev['datetime'], ev['message'])
#      PYEOF
# ================================================================

# Emette il sorgente della libreria sullo stdout.
pylib_timeline() {
    cat << 'FIUTO_PYLIB_EOF'
import re as _re
import html as _H
import datetime as _dt

_MONTHS = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
           'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
_TS_ISO = _re.compile(r'\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}')
_TS_SYS = _re.compile(r'\b(' + '|'.join(_MONTHS) + r')\s+(\d{1,2})\s+(\d{2}:\d{2}:\d{2})')
_TD = _re.compile(r'<td[^>]*>(.*?)</td>', _re.DOTALL | _re.I)
_TR = _re.compile(r'<tr[^>]*>(.*?)</tr>', _re.DOTALL | _re.I)
_PRE = _re.compile(r'<pre[^>]*>(.*?)</pre>', _re.DOTALL | _re.I)
_TAG = _re.compile(r'<[^>]+>')


def strip_tags(s):
    return _TAG.sub('', _H.unescape(s)).strip()


_SENS = _re.compile(r'<span class="line sensitive">(.*?)</span>', _re.DOTALL)


def read_report_data(path):
    """Solo il DATO del report: righe di tabella e righe dei blocchi <pre>.

    Non il testo completo della pagina. I report contengono cartigli
    esplicativi scritti da FIUTO ("...tipico del malware che...") e un motore
    di detection che cercasse parole nell'intera pagina scatterebbe sulla
    propria prosa. Qui la prosa non entra proprio.

    L'unita' e' la RIGA, non la cella: un conteggio di celle non significa
    niente per chi legge ("3 riscontri" su una riga sola), e una regola che
    deve vedere data e percorso insieme non funzionerebbe.
    """
    try:
        content = open(path, encoding='utf-8', errors='replace').read()
    except Exception:
        return []
    items = []
    for tr in _TR.finditer(content):
        cells = [strip_tags(m.group(1)) for m in _TD.finditer(tr.group(1))]
        row = ' | '.join(c for c in cells if c)
        if row:
            items.append(row)
    for pm in _PRE.finditer(content):
        block = _H.unescape(_TAG.sub('', pm.group(1)))
        for line in block.split('\n'):
            line = _re.sub(r'^\s*\d+\s+', '', line).strip()
            if line:
                items.append(line)
    return items


def read_sensitive_lines(path):
    """Righe che il modulo stesso ha marcato come degne di attenzione.

    render_pre_block applica la classe 'sensitive' alle righe che contengono
    le parole chiave passate dal modulo: e' un giudizio gia' espresso da chi
    conosce l'artefatto, molto piu' affidabile di una regola generica scritta
    a valle.
    """
    try:
        content = open(path, encoding='utf-8', errors='replace').read()
    except Exception:
        return []
    out = []
    for m in _SENS.finditer(content):
        s = _re.sub(r'^\s*\d+\s+', '', strip_tags(m.group(1))).strip()
        if s:
            out.append(s)
    return out


def extract_events(path):
    """Eventi datati di un report HTML, deduplicati e ordinati.

    Ritorna [{'datetime','message','year_inferred'}]. Le date NON sono
    riportate a un fuso comune: vedi src/lib/16-time.sh per il perche'.
    """
    try:
        content = open(path, encoding='utf-8', errors='replace').read()
    except Exception:
        return []

    # I log syslog non portano l'anno: si usa quello corrente (l'analisi e'
    # contestuale all'acquisizione). Approssimazione esplicita, non silenziosa.
    year = str(_dt.date.today().year)
    out = []
    seen = set()

    def emit(dt, message, assumed):
        message = ' '.join(message.split())[:2000]
        if not message:
            return
        key = (dt, message[:120])
        if key in seen:
            return
        seen.add(key)
        out.append({'datetime': dt, 'message': message, 'year_inferred': assumed})

    # 1) Righe di tabella in cui una cella contiene un timestamp
    for tr in _TR.finditer(content):
        cells = [strip_tags(m.group(1)) for m in _TD.finditer(tr.group(1))]
        if not cells:
            continue
        dt = None
        for c in cells:
            m = _TS_ISO.search(c)
            if m:
                dt = m.group(0).replace('T', ' ')[:19].replace(' ', 'T')
                break
        if not dt:
            continue
        msg = ' | '.join(c for c in cells
                         if c and not _TS_ISO.fullmatch(c.strip())
                         and not _re.fullmatch(r'\d+', c.strip()))
        emit(dt, msg, False)

    # 2) Blocchi <pre> (log, history, config): una riga per evento
    for pm in _PRE.finditer(content):
        block = _H.unescape(_TAG.sub('', pm.group(1)))
        for raw in block.split('\n'):
            txt = _re.sub(r'^\s*\d+\s+', '', raw).strip()   # via il numero di riga
            if not txt:
                continue
            m = _TS_ISO.search(txt)
            if m:
                emit(m.group(0).replace('T', ' ')[:19].replace(' ', 'T'), txt, False)
                continue
            m = _TS_SYS.search(txt)
            if m:
                emit(f"{year}-{_MONTHS[m.group(1)]}-{int(m.group(2)):02d}T{m.group(3)}",
                     txt, True)

    out.sort(key=lambda r: r['datetime'])
    return out
FIUTO_PYLIB_EOF
}
