
# ================================================================
#  MODULO 51 — Sigma sugli Event Log (EVTX)
#
#  Sigma e' il formato in cui la comunita' pubblica le detection: SigmaHQ,
#  i CERT e i vendor distribuiscono migliaia di regole YAML. Applicarle agli
#  EVTX di un disco acquisito e' quello che fanno Chainsaw e Hayabusa, ed e' il
#  passo che trasforma una raccolta di log in un triage.
#
#  IL SOTTOINSIEME SUPPORTATO E' DICHIARATO, NON IMPLICITO. Sigma e' un
#  linguaggio ampio: modificatori base64, CIDR, condizioni con parentesi,
#  aggregazioni temporali. Implementarne una parte e far finta di supportarlo
#  tutto significherebbe che una regola non valutata risulta "non scattata" —
#  cioe' un falso negativo silenzioso, il difetto peggiore che una detection
#  possa avere. Qui le regole che il motore non sa valutare vengono CONTATE ED
#  ELENCATE nel report con il motivo.
#
#  Supportato:
#    - selezioni: mappa campo/valore, liste di valori (OR), liste di mappe (OR)
#    - modificatori: contains, startswith, endswith, re, all, cased
#    - condizioni: "sel", "a and b", "a or b", "a and not b", "not a",
#                  "1 of x*", "all of x*", "1 of them", "all of them"
#    - null come valore (campo assente o vuoto)
#
#  Non supportato (regola scartata e dichiarata):
#    - condizioni con parentesi o aggregazioni (| count, near, timeframe)
#    - modificatori base64/base64offset, utf16, wide, cidr, gt/lt
#    - logsource non mappabile a un canale EVTX presente sul volume
#
#  Il compilatore vive in src/lib/19-pylib-sigma.sh: e' la parte rischiosa —
#  un modificatore interpretato male produce un falso negativo invisibile — e
#  li' e' esercitabile dai test con eventi sintetici, senza un .evtx. Qui resta
#  la lettura degli EVTX e la presentazione.
# ================================================================

_guard_sigma() {
    if [[ -z "${SIGMA_RULES:-}" ]]; then
        L "nessuna regola (--sigma)" "no rules (--sigma)"
        return 1
    fi
    return 0
}

module_sigma() {
    section_header "Sigma - Event Log" "$RED"
    check_win_root || return 1

    if [[ -z "${SIGMA_RULES:-}" ]]; then
        warn "$(L "Nessuna regola indicata." "No rules given.")"
        info "$(L "Uso: --sigma /percorso/regole/  (file .yml o directory, anche annidata)" \
                 "Usage: --sigma /path/rules/  (a .yml file or a directory, nested is fine)")"
        return 0
    fi
    [[ -e "$SIGMA_RULES" ]] || { err "$(L "Percorso regole inesistente:" "Rules path does not exist:") $SIGMA_RULES"; return 1; }

    if ! "$PY3" -c "import yaml" 2>/dev/null; then
        err "$(L "PyYAML non disponibile: le regole Sigma sono file YAML." \
                 "PyYAML unavailable: Sigma rules are YAML files.")"
        info "$(L "Installa con:" "Install with:") ${PY3} -m pip install pyyaml"
        return 1
    fi
    if ! "$PY3" -c "import Evtx" 2>/dev/null; then
        err "$(L "python-evtx non disponibile: senza non si leggono gli EVTX." \
                 "python-evtx unavailable: without it EVTX cannot be read.")"
        info "$(L "Installa con:" "Install with:") ${PY3} -m pip install python-evtx"
        return 1
    fi

    local EVTX_DIR; EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/winevt/Logs")
    [[ -z "$EVTX_DIR" ]] && EVTX_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
    if [[ -z "$EVTX_DIR" ]]; then
        warn "$(L "Directory dei log eventi non trovata." "Event log directory not found.")"
        return 0
    fi
    info "$(L "Log eventi:" "Event logs:") ${BOLD}${EVTX_DIR}"
    info "$(L "Regole:" "Rules:") ${BOLD}${SIGMA_RULES}"
    info "$(L "Valutazione in corso (dipende dal numero di regole e dalla dimensione dei log)..." \
             "Evaluating (depends on rule count and log size)...")"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local STATS; STATS=$(mktemp); register_tmp "$STATS"

    run_py_with_lib pylib_sigma "$SIGMA_RULES" "$EVTX_DIR" "$OUT" "$STATS" "$SIGMA_MAX_RECORDS" << 'PYEOF'
import sys, os, re, json, glob

rules_path, evtx_dir, out_path, stats_path = sys.argv[1:5]
MAX_RECORDS = int(sys.argv[5])


import Evtx.Evtx as evtx

NS = 'http://schemas.microsoft.com/win/2004/08/events/event'

available = {os.path.basename(p).lower(): p
             for p in glob.glob(os.path.join(evtx_dir, '*.evtx'))}
rules, rejected = load_sigma_rules(rules_path, available)

stats = {
    'rules_loaded': len(rules) + len(rejected),
    'rules_active': len(rules),
    'rules_rejected': rejected[:400],
    'rules_rejected_total': len(rejected),
    'channels': [],
    'records_read': 0,
    'stopped_at_cap': False,
}

if not rules:
    json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
    sys.exit(0)

# --- valutazione -----------------------------------------------------------
by_file = {}
for r in rules:
    for f in r['files']:
        by_file.setdefault(f, []).append(r)

_TAG = re.compile(r'\{[^}]+\}')
rows = []
seen = set()

for path, rlist in sorted(by_file.items()):
    n = 0
    try:
        with evtx.Evtx(path) as log:
            for rec in log.records():
                if stats['records_read'] >= MAX_RECORDS:
                    stats['stopped_at_cap'] = True
                    break
                n += 1
                stats['records_read'] += 1
                try:
                    root = rec.lxml()
                except Exception:
                    continue
                sysel = root.find('{%s}System' % NS)
                if sysel is None:
                    continue
                ev = {}
                eid = ''
                ts = ''
                for child in sysel:
                    tag = _TAG.sub('', child.tag)
                    if tag == 'EventID':
                        eid = (child.text or '').strip()
                        ev['EventID'] = eid
                    elif tag == 'TimeCreated':
                        ts = child.attrib.get('SystemTime', '')[:19]
                    elif tag == 'Provider':
                        ev['Provider_Name'] = child.attrib.get('Name', '')
                    elif tag == 'Channel':
                        ev['Channel'] = (child.text or '').strip()
                    elif tag == 'Computer':
                        ev['Computer'] = (child.text or '').strip()
                de = root.find('.//{%s}EventData' % NS)
                if de is not None:
                    for it in de:
                        key = it.attrib.get('Name') or _TAG.sub('', it.tag)
                        ev[key] = (it.text or '').strip()
                for r in rlist:
                    try:
                        if not r['pred'](ev):
                            continue
                    except Exception:
                        continue
                    key = (r['title'], ts, eid)
                    if key in seen:
                        continue
                    seen.add(key)
                    # Il contesto e' cio' che rende il match verificabile: i
                    # campi piu' parlanti, non il record intero.
                    ctx = ' | '.join(
                        '%s=%s' % (k, v[:120])
                        for k, v in ev.items()
                        if k in ('Image', 'CommandLine', 'ParentImage', 'ParentCommandLine',
                                 'TargetUserName', 'SubjectUserName', 'ServiceName',
                                 'ScriptBlockText', 'TargetFilename', 'DestinationIp',
                                 'QueryName', 'TargetObject', 'Details') and v)[:600]
                    rows.append((ts.replace('T', ' '), r['level'], r['title'],
                                 os.path.basename(path), eid, ctx,
                                 ','.join(t for t in r['tags'] if t.startswith('attack.'))[:100]))
    except Exception as exc:
        stats['channels'].append({'file': os.path.basename(path), 'records': n,
                                  'error': str(exc)[:150]})
        continue
    stats['channels'].append({'file': os.path.basename(path), 'records': n,
                              'rules': len(rlist)})
    if stats['stopped_at_cap']:
        break

LEVEL_ORDER = {'critical': 0, 'high': 1, 'medium': 2, 'low': 3, 'informational': 4}
rows.sort(key=lambda r: (LEVEL_ORDER.get(r[1], 9), r[0]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
stats['matches'] = len(rows)
json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
PYEOF

    [[ -s "$STATS" ]] || { err "$(L "Valutazione non riuscita." "Evaluation failed.")"; return 1; }
    local NLOAD NACT NREJ NREC CAPPED
    read -r NLOAD NACT NREJ NREC CAPPED < <("$PY3" -c '
import json, sys
s = json.load(open(sys.argv[1]))
print(s["rules_loaded"], s["rules_active"], s["rules_rejected_total"],
      s["records_read"], str(s["stopped_at_cap"]).lower())' "$STATS" 2>/dev/null)

    if [[ "${NACT:-0}" -eq 0 ]]; then
        warn "$(L "Nessuna regola valutabile su questo volume." "No rule evaluable on this volume.")"
        info "$(L "Regole caricate:" "Rules loaded:") ${NLOAD:-0}  ·  $(L "scartate:" "rejected:") ${NREJ:-0}"
        "$PY3" -c '
import json, sys
for r in json.load(open(sys.argv[1]))["rules_rejected"][:10]:
    print("      %s - %s" % (r["rule"][:60], r["reason"]))' "$STATS" 2>/dev/null
        return 0
    fi

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    ok "$(L "Regole attive:" "Active rules:") ${BOLD}${NACT}${RESET}/${NLOAD}  ·  $(L "record letti:" "records read:") ${BOLD}${NREC}"
    [[ "${NREJ:-0}" -gt 0 ]] && warn "$(L "Regole non valutate (elencate nel report):" "Rules not evaluated (listed in the report):") ${BOLD}${NREJ}"
    [[ "$CAPPED" == "true" ]] && warn "$(L "Raggiunto il tetto di record: la valutazione e' PARZIALE." \
                                          "Record cap reached: the evaluation is PARTIAL.")"

    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessun match." "No match.")"
    else
        warn "$(L "Match:" "Matches:") ${BOLD}${TOTAL}"
        awk -F'\t' '{printf "      [%s] %s  %s\n", $2, $1, substr($3,1,70)}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local TABLE
    if [[ "$TOTAL" -gt 0 ]]; then
        TABLE=$(_rows_to_table "$(head -20000 "$OUT")" \
            "$(L "Data" "Date")" "$(L "Livello" "Level")" "$(L "Regola" "Rule")" \
            "$(L "Canale" "Channel")" "EventID" "$(L "Contesto" "Context")" "ATT&CK")
    else
        TABLE="<div style='padding:.6rem 0;font-size:.85rem'>$(L "Nessun match." "No match.")</div>"
    fi

    local COV; COV=$("$PY3" - "$STATS" "${LANG:-en}" << 'PYEOF' 2>/dev/null
import json, sys, html

s = json.load(open(sys.argv[1]))
it = sys.argv[2] == 'it'


def L(i, e):
    return i if it else e


o = ['<div class="card" style="margin-bottom:1rem">'
     '<div style="padding:1rem 1.5rem;font-size:.8rem;line-height:1.7">']
o.append('<b>' + L('Copertura effettiva', 'Actual coverage') + '</b><br>')
o.append(L("Sigma e' un linguaggio ampio e questo motore ne implementa un sottoinsieme. "
           "Una regola non valutata non e' una regola non scattata: le regole scartate "
           "sono contate ed elencate qui sotto con il motivo, perche' un falso negativo "
           "silenzioso e' il difetto peggiore che una detection possa avere.",
           'Sigma is a broad language and this engine implements a subset of it. '
           'A rule that was not evaluated is not a rule that did not fire: rejected rules '
           'are counted and listed below with the reason, because a silent false negative '
           'is the worst defect a detection can have.') + '<br><br>')
o.append('<b>%d</b> %s · <b>%d</b> %s · <b>%d</b> %s<br><br>' % (
    s['rules_loaded'], L('regole caricate', 'rules loaded'),
    s['rules_active'], L('valutate', 'evaluated'),
    s['rules_rejected_total'], L('scartate', 'rejected')))

o.append('<table><tr><th>' + L('Canale', 'Channel') + '</th><th>'
         + L('Record letti', 'Records read') + '</th><th>'
         + L('Regole applicate', 'Rules applied') + '</th></tr>')
for c in s['channels']:
    o.append('<tr><td class="mono">%s</td><td class="mono">%d</td><td class="mono">%s</td></tr>'
             % (html.escape(c['file']), c.get('records', 0),
                html.escape(str(c.get('rules', c.get('error', '-'))))))
o.append('</table>')
if s['stopped_at_cap']:
    o.append('<br><b>' + L("Raggiunto il tetto di record: la valutazione e' PARZIALE.",
                           'Record cap reached: the evaluation is PARTIAL.') + '</b>')

if s['rules_rejected']:
    reasons = {}
    for r in s['rules_rejected']:
        key = r['reason'].split(':')[0].split('(')[0].strip()
        reasons.setdefault(key, []).append(r['rule'])
    o.append('<br><br><b>' + L('Regole scartate, per motivo', 'Rejected rules, by reason') + '</b>')
    o.append('<table><tr><th>' + L('Motivo', 'Reason') + '</th><th>'
             + L('Regole', 'Rules') + '</th><th>' + L('Esempi', 'Examples') + '</th></tr>')
    for reason, names in sorted(reasons.items(), key=lambda kv: -len(kv[1])):
        o.append('<tr><td class="mono">%s</td><td class="mono">%d</td><td class="mono">%s</td></tr>'
                 % (html.escape(reason), len(names),
                    html.escape(', '.join(n[:50] for n in names[:3]))))
    o.append('</table>')
    if s['rules_rejected_total'] > len(s['rules_rejected']):
        o.append('<br>' + L('Elenco troncato: scartate in totale %d.'
                            % s['rules_rejected_total'],
                            'List truncated: %d rejected in total.'
                            % s['rules_rejected_total']))
o.append('</div></div>')
print(''.join(o))
PYEOF
    )

    local SB
    SB="$(stat_box "Match" "$TOTAL" "$([[ "$TOTAL" -gt 0 ]] && echo warn || echo ok)")"
    SB+="$(stat_box "$(L "Regole valutate" "Rules evaluated")" "${NACT:-0}" "info")"
    SB+="$(stat_box "$(L "Regole scartate" "Rules rejected")" "${NREJ:-0}" "$([[ "${NREJ:-0}" -gt 0 ]] && echo warn || echo info)")"
    SB+="$(stat_box "$(L "Record letti" "Records read")" "${NREC:-0}" "info")"
    finish_report "sigma" "Sigma" "SIG" "$(basename "$SIGMA_RULES")" "$SB" \
        "${COV}<div class='cards'>$(generic_card_html "$(L "Riscontri" "Matches")" "$SIGMA_RULES" "$TOTAL" "$TABLE" "⚐")</div>"
}
