
# ================================================================
#  LIBRERIA PYTHON CONDIVISA — compilatore di regole Sigma
#
#  Sigma e' un linguaggio ampio e questo e' un motore parziale. La parte
#  rischiosa non e' leggere gli EVTX, e' decidere cosa una regola significa:
#  un modificatore interpretato male, o una condizione valutata al contrario,
#  produce un falso negativo che nessuno vede. Per questo il compilatore vive
#  qui e non dentro il modulo: cosi' e' esercitabile dai test con eventi
#  sintetici, senza bisogno di un .evtx.
#
#  SOTTOINSIEME SUPPORTATO — dichiarato, non implicito:
#    selezioni     mappa campo/valore, liste di valori (OR), liste di mappe (OR)
#    modificatori  contains, startswith, endswith, re, all, cased
#    condizioni    "sel", "a and b", "a or b", "a and not b", "not a",
#                  "1 of x*", "all of x*", "1 of them", "all of them"
#    valori        null = campo assente o vuoto
#
#  Tutto il resto solleva Unsupported con il motivo, e il chiamante e' tenuto
#  a ELENCARE le regole scartate: una regola non valutata non e' una regola
#  non scattata.
#
#  Uso:
#      run_py_with_lib pylib_sigma "$RULES" << 'PYEOF'
#      rules, rejected = load_sigma_rules(sys.argv[1], {'security.evtx': '/x/Security.evtx'})
#      PYEOF
# ================================================================

pylib_sigma() {
    cat << 'FIUTO_PYLIB_EOF'
import os as _os
import re as _re
import glob as _glob

import yaml as _yaml


class Unsupported(Exception):
    """La regola usa una costruzione che questo motore non implementa."""


SUPPORTED_MODS = {'contains', 'startswith', 'endswith', 're', 'all', 'cased'}

# Sigma descrive la sorgente in astratto; sul disco ci sono file con nomi
# precisi. Una logsource che non sappiamo mappare NON viene fatta girare a
# tappeto su tutti i log: la regola verrebbe valutata su campi che in quel
# canale non esistono, e il "non scattata" sarebbe privo di significato.
SERVICE_MAP = {
    'security': ['Security.evtx'],
    'system': ['System.evtx'],
    'application': ['Application.evtx'],
    'sysmon': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'powershell': ['Microsoft-Windows-PowerShell%4Operational.evtx'],
    'powershell-classic': ['Windows PowerShell.evtx'],
    'taskscheduler': ['Microsoft-Windows-TaskScheduler%4Operational.evtx'],
    'windefend': ['Microsoft-Windows-Windows Defender%4Operational.evtx'],
    'terminalservices-localsessionmanager':
        ['Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx'],
    'wmi': ['Microsoft-Windows-WMI-Activity%4Operational.evtx'],
    'bits-client': ['Microsoft-Windows-Bits-Client%4Operational.evtx'],
    'smbclient-security': ['Microsoft-Windows-SmbClient%4Security.evtx'],
    'ntlm': ['Microsoft-Windows-NTLM%4Operational.evtx'],
}
# process_creation sta sia in Security (4688) sia in Sysmon (EID 1): una regola
# scritta per l'uno vale spesso per l'altro, quindi si applica a entrambi.
CATEGORY_MAP = {
    'process_creation': ['Security.evtx', 'Microsoft-Windows-Sysmon%4Operational.evtx'],
    'network_connection': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'image_load': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'file_event': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'registry_event': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'registry_set': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'registry_add': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'process_access': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'pipe_created': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'dns_query': ['Microsoft-Windows-Sysmon%4Operational.evtx'],
    'ps_script': ['Microsoft-Windows-PowerShell%4Operational.evtx'],
    'ps_module': ['Microsoft-Windows-PowerShell%4Operational.evtx'],
    'ps_classic_start': ['Windows PowerShell.evtx'],
}


def compile_field(key, raw):
    """(nome_campo, predicato) da una chiave 'Campo|modificatore'."""
    parts = key.split('|')
    field = parts[0]
    mods = [m.lower() for m in parts[1:]]
    for m in mods:
        if m not in SUPPORTED_MODS:
            raise Unsupported('modificatore |%s' % m)
    cased = 'cased' in mods
    values = raw if isinstance(raw, list) else [raw]

    def norm(v):
        return None if v is None else (str(v) if cased else str(v).lower())

    vals = [norm(v) for v in values]

    def actual_of(a):
        a = '' if a is None else str(a)
        return a if cased else a.lower()

    if 're' in mods:
        rxs = [_re.compile(str(v), 0 if cased else _re.I) for v in values]

        def pred(a):
            return any(r.search('' if a is None else str(a)) for r in rxs)
    elif 'contains' in mods:
        if 'all' in mods:
            def pred(a):
                s = actual_of(a)
                return all(v is not None and v in s for v in vals)
        else:
            def pred(a):
                s = actual_of(a)
                return any(v is not None and v in s for v in vals)
    elif 'startswith' in mods:
        def pred(a):
            s = actual_of(a)
            return any(v is not None and s.startswith(v) for v in vals)
    elif 'endswith' in mods:
        def pred(a):
            s = actual_of(a)
            return any(v is not None and s.endswith(v) for v in vals)
    else:
        def pred(a):
            s = actual_of(a)
            # null in Sigma significa "campo assente o vuoto".
            return any((v is None and not s) or (v is not None and s == v) for v in vals)

    return field, pred


def compile_selection(sel):
    """funzione(evento)->bool. Una lista di mappe e' un OR fra le mappe."""
    if isinstance(sel, list):
        subs = [compile_selection(s) for s in sel]
        return lambda ev: any(s(ev) for s in subs)
    if not isinstance(sel, dict):
        raise Unsupported('selezione non e\' una mappa')
    checks = [compile_field(k, v) for k, v in sel.items()]

    def run(ev):
        return all(pred(ev.get(field, '')) for field, pred in checks)
    return run


_COND_RX = _re.compile(
    r'^(?:(1|all)\s+of\s+(\S+)|not\s+(\S+)|(\S+))'
    r'(?:\s+(and|or)\s+(?:(not)\s+)?(?:(1|all)\s+of\s+)?(\S+))?$', _re.I)


def compile_condition(cond, sels):
    """Sottoinsieme documentato della grammatica delle condizioni Sigma."""
    if not isinstance(cond, str):
        raise Unsupported('condizione non testuale')
    c = ' '.join(cond.split())
    if '(' in c or '|' in c:
        raise Unsupported('condizione con parentesi o aggregazione')

    def resolve(name):
        n = name.strip()
        if n.lower() == 'them':
            got = list(sels.values())
        elif n.endswith('*'):
            got = [v for k, v in sels.items() if k.startswith(n[:-1])]
        else:
            got = [sels[n]] if n in sels else []
        if not got:
            raise Unsupported('selezione %s inesistente' % n)
        return got

    m = _COND_RX.match(c)
    if not m:
        raise Unsupported('condizione non riconosciuta: %s' % c[:60])
    quant, quant_name, not_name, plain, op, op_not, rquant, rhs = m.groups()

    def side(quantifier, name, negated):
        group = resolve(name)
        if (quantifier or '').lower() == 'all':
            base = lambda ev: all(f(ev) for f in group)   # noqa: E731
        else:
            base = lambda ev: any(f(ev) for f in group)   # noqa: E731
        return (lambda ev: not base(ev)) if negated else base

    if quant:
        left = side(quant, quant_name, False)
    elif not_name:
        left = side(None, not_name, True)
    else:
        left = side(None, plain, False)

    if not op:
        return left

    right = side(rquant, rhs, bool(op_not))
    if op.lower() == 'and':
        return lambda ev: left(ev) and right(ev)
    return lambda ev: left(ev) or right(ev)


def _rule_files(path):
    if _os.path.isdir(path):
        out = []
        for ext in ('yml', 'yaml'):
            out += _glob.glob(_os.path.join(path, '**', '*.' + ext), recursive=True)
        return sorted(out)
    return [path]


def load_sigma_rules(rules_path, available):
    """(rules, rejected).

    `available` mappa nome-file-evtx-minuscolo -> percorso reale sul volume.
    Ogni regola scartata finisce in `rejected` con il motivo: e' quella lista
    che impedisce di leggere "nessun match" come "nessuna minaccia".
    """
    rules, rejected = [], []
    for rf in _rule_files(rules_path):
        try:
            with open(rf, encoding='utf-8', errors='replace') as fh:
                docs = [d for d in _yaml.safe_load_all(fh) if isinstance(d, dict)]
        except Exception as exc:
            rejected.append({'rule': _os.path.basename(rf),
                             'reason': 'YAML illeggibile: %s' % str(exc)[:120]})
            continue
        for doc in docs:
            name = doc.get('title') or _os.path.basename(rf)
            det = doc.get('detection')
            if not isinstance(det, dict) or 'condition' not in det:
                rejected.append({'rule': name, 'reason': 'detection assente o senza condition'})
                continue
            ls = doc.get('logsource') or {}
            svc = str(ls.get('service', '')).lower()
            cat = str(ls.get('category', '')).lower()
            targets = SERVICE_MAP.get(svc) or CATEGORY_MAP.get(cat) or []
            if not targets:
                rejected.append({'rule': name, 'reason': 'logsource non mappata (%s/%s)'
                                 % (ls.get('product', '-'), svc or cat or '-')})
                continue
            present = [available[t.lower()] for t in targets if t.lower() in available]
            if not present:
                rejected.append({'rule': name, 'reason': 'canale non presente sul volume (%s)'
                                 % ', '.join(targets)})
                continue
            try:
                sels = {k: compile_selection(v) for k, v in det.items() if k != 'condition'}
                pred = compile_condition(det['condition'], sels)
            except Unsupported as exc:
                rejected.append({'rule': name, 'reason': str(exc)})
                continue
            except Exception as exc:
                rejected.append({'rule': name,
                                 'reason': 'compilazione fallita: %s' % str(exc)[:120]})
                continue
            rules.append({
                'title': name,
                'id': doc.get('id', ''),
                'level': str(doc.get('level', 'medium')),
                'tags': [t for t in (doc.get('tags') or []) if isinstance(t, str)],
                'files': present,
                'pred': pred,
            })
    return rules, rejected
FIUTO_PYLIB_EOF
}
