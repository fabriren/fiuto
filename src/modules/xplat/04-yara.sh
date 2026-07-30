
# ================================================================
#  CROSS-OS — Scansione YARA
#
#  YARA e' il formato con cui l'industria distribuisce le firme: un feed di
#  threat intelligence, l'IR di un vendor o il CERT nazionale mandano regole
#  .yar, e finora FIUTO non aveva modo di applicarle. Con --yara le applica.
#
#  IL PUNTO DELICATO E' L'AMBITO. Scansionare un volume da un terabyte file per
#  file non e' praticabile su una workstation forense, e un report che non
#  dichiarasse cosa ha guardato sarebbe peggio di nessun report: "nessun match"
#  verrebbe letto come "il disco e' pulito". Il modulo scansiona quindi un
#  insieme limitato di posizioni — quelle scrivibili senza privilegi, dove il
#  codice non installato dai pacchetti finisce quasi sempre — e ELENCA NEL
#  REPORT esattamente cosa ha scansionato, cosa ha saltato e perche'.
#
#  Con --yara-scan si indica un percorso diverso e l'ambito diventa quello.
# ================================================================

# Guardia: senza regole non c'e' niente da fare, e in batch il modulo va
# saltato con un motivo invece di produrre un report vuoto.
_guard_yara() {
    if [[ -z "${YARA_RULES:-}" ]]; then
        L "nessuna regola (--yara)" "no rules (--yara)"
        return 1
    fi
    return 0
}

module_xplat_yara() {
    section_header "YARA" "$RED"
    check_target_root || return 1

    if [[ -z "${YARA_RULES:-}" ]]; then
        warn "$(L "Nessuna regola indicata." "No rules given.")"
        info "$(L "Uso: --yara /percorso/regole.yar oppure --yara /percorso/directory/" \
                 "Usage: --yara /path/rules.yar or --yara /path/directory/")"
        return 0
    fi
    if [[ ! -e "$YARA_RULES" ]]; then
        err "$(L "Percorso regole inesistente:" "Rules path does not exist:") $YARA_RULES"
        return 1
    fi
    if ! "$PY3" -c "import yara" 2>/dev/null; then
        err "$(L "yara-python non disponibile." "yara-python unavailable.")"
        info "$(L "Installa con: pip install yara-python" "Install with: pip install yara-python")"
        info "$(L "Senza il motore non esiste alcun ripiego: una scansione YARA senza YARA non e' una scansione." \
                 "Without the engine there is no fallback: a YARA scan without YARA is not a scan.")"
        return 1
    fi

    # --- ambito ------------------------------------------------------------
    # Posizioni scrivibili senza privilegi, piu' i punti in cui il sistema
    # deposita eseguibili scaricati. Non e' l'intero volume: e' dichiarato.
    local -a TARGETS=()
    local D
    if [[ -n "${YARA_SCAN_PATH:-}" ]]; then
        [[ -d "$YARA_SCAN_PATH" ]] || { err "$(L "Percorso da scansionare inesistente:" "Scan path does not exist:") $YARA_SCAN_PATH"; return 1; }
        TARGETS=("$YARA_SCAN_PATH")
    else
        local -a CAND=()
        case "$OS_TYPE" in
            windows)
                CAND=("Windows/Temp" "ProgramData" "Users" "Windows/Tasks"
                      "Windows/System32/Tasks" "PerfLogs" "EFI"
                      "ProgramData/Microsoft/Windows Defender/Quarantine")
                ;;
            linux)
                CAND=("tmp" "var/tmp" "dev/shm" "home" "root" "opt"
                      "usr/local" "var/www" "etc/cron.d" "etc/systemd/system" "EFI")
                ;;
            macos)
                CAND=("tmp" "private/tmp" "private/var/tmp" "Users" "Library/LaunchAgents"
                      "Library/LaunchDaemons" "Library/Application Support" "EFI")
                ;;
        esac
        for D in "${CAND[@]}"; do
            local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
            [[ -n "$R" && -d "$R" ]] && TARGETS+=("$R")
        done
    fi

    if [[ ${#TARGETS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna delle posizioni previste esiste su questo volume." \
                 "None of the expected locations exists on this volume.")"
        return 0
    fi

    info "$(L "Regole:" "Rules:") ${BOLD}${YARA_RULES}"
    info "$(L "Posizioni da scansionare:" "Locations to scan:") ${BOLD}${#TARGETS[@]}${RESET}  ·  $(L "tetto per file:" "per-file cap:") ${BOLD}${YARA_MAX_MB} MB"
    info "$(L "Scansione in corso (puo' richiedere molto tempo)..." "Scanning (this can take a long time)...")"

    local TLIST; TLIST=$(mktemp); register_tmp "$TLIST"
    printf '%s\n' "${TARGETS[@]}" > "$TLIST"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local STATS; STATS=$(mktemp); register_tmp "$STATS"

    "$PY3" - "$YARA_RULES" "$TLIST" "$OUT" "$STATS" "$YARA_MAX_MB" "$YARA_MAX_FILES" << 'PYEOF'
import sys, os, json, hashlib

rules_path, tlist, out_path, stats_path = sys.argv[1:5]
max_bytes = int(sys.argv[5]) * 1024 * 1024
max_files = int(sys.argv[6])

import yara

# --- compilazione ----------------------------------------------------------
# Un file di regole con un errore di sintassi non deve far fallire tutto il
# resto: si compila file per file e si dichiara quali sono stati scartati.
sources = {}
bad_rules = []
if os.path.isdir(rules_path):
    for root, _dirs, files in os.walk(rules_path):
        for fn in sorted(files):
            if fn.lower().endswith(('.yar', '.yara')):
                sources[os.path.relpath(os.path.join(root, fn), rules_path)] = \
                    os.path.join(root, fn)
else:
    sources[os.path.basename(rules_path)] = rules_path

compiled = {}
for ns, path in sources.items():
    try:
        compiled[ns] = yara.compile(filepath=path)
    except Exception as exc:
        bad_rules.append({'file': ns, 'error': str(exc)[:300]})

stats = {
    'rule_files': len(sources),
    'rule_files_compiled': len(compiled),
    'rule_files_rejected': bad_rules,
    'scanned': 0,
    'skipped_too_big': 0,
    'skipped_unreadable': 0,
    'stopped_at_cap': False,
    'targets': [],
}

if not compiled:
    json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
    sys.exit(0)

targets = [t.strip() for t in open(tlist, encoding='utf-8') if t.strip()]
rows = []
seen_files = 0


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


for target in targets:
    per_target = 0
    for root, dirs, files in os.walk(target, followlinks=False):
        for fn in files:
            if seen_files >= max_files:
                stats['stopped_at_cap'] = True
                break
            p = os.path.join(root, fn)
            try:
                if os.path.islink(p):
                    continue
                sz = os.path.getsize(p)
            except OSError:
                stats['skipped_unreadable'] += 1
                continue
            if sz == 0:
                continue
            if sz > max_bytes:
                # Dichiarato, non omesso: un payload dentro un file da 2 GB
                # non verrebbe visto, e chi legge deve saperlo.
                stats['skipped_too_big'] += 1
                continue
            seen_files += 1
            per_target += 1
            stats['scanned'] += 1
            for ns, rl in compiled.items():
                try:
                    matches = rl.match(p, timeout=30)
                except Exception:
                    stats['skipped_unreadable'] += 1
                    continue
                for m in matches:
                    # Le stringhe che hanno fatto match sono la prova del
                    # riscontro: senza, resta solo un nome di regola.
                    frag = []
                    try:
                        for s in m.strings[:4]:
                            for inst in s.instances[:2]:
                                frag.append('%s@%d:%s' % (
                                    s.identifier, inst.offset,
                                    inst.matched_data[:60].decode('utf-8', 'replace')))
                    except Exception:
                        pass
                    try:
                        digest = sha256(p)
                    except Exception:
                        digest = ''
                    rows.append((m.rule, ns, p, str(sz), digest,
                                 ' | '.join(frag)[:400],
                                 ','.join(m.tags)[:80]))
        if stats['stopped_at_cap']:
            break
    stats['targets'].append({'path': target, 'files': per_target})
    if stats['stopped_at_cap']:
        break

rows.sort()
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
json.dump(stats, open(stats_path, 'w'), ensure_ascii=False)
PYEOF

    local NSCAN=0 NBIG=0 NBAD=0 NRULES=0 NREJ=0 CAPPED=false
    if [[ -s "$STATS" ]]; then
        read -r NSCAN NBIG NBAD NRULES NREJ CAPPED < <("$PY3" -c '
import json, sys
s = json.load(open(sys.argv[1]))
print(s["scanned"], s["skipped_too_big"], s["skipped_unreadable"],
      s["rule_files_compiled"], len(s["rule_files_rejected"]),
      str(s["stopped_at_cap"]).lower())' "$STATS" 2>/dev/null)
    fi

    if [[ "${NRULES:-0}" -eq 0 ]]; then
        err "$(L "Nessun file di regole compilato." "No rule file compiled.")"
        [[ "${NREJ:-0}" -gt 0 ]] && "$PY3" -c '
import json, sys
for r in json.load(open(sys.argv[1]))["rule_files_rejected"]:
    print("      %s: %s" % (r["file"], r["error"]))' "$STATS" 2>/dev/null
        return 1
    fi
    [[ "${NREJ:-0}" -gt 0 ]] && warn "$(L "File di regole scartati (errore di sintassi):" "Rule files rejected (syntax error):") ${BOLD}${NREJ}"

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    ok "$(L "File scansionati:" "Files scanned:") ${BOLD}${NSCAN}"
    [[ "${NBIG:-0}" -gt 0 ]] && info "$(L "Saltati perche' oltre il tetto:" "Skipped as over the cap:") ${BOLD}${NBIG}"
    [[ "$CAPPED" == "true" ]] && warn "$(L "Raggiunto il tetto di file: la scansione e' PARZIALE." \
                                          "File cap reached: the scan is PARTIAL.")"

    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessun match." "No match.")"
        info "$(L "Nessun match sulle posizioni scansionate: non equivale a un volume pulito." \
                 "No match in the scanned locations: this is not equivalent to a clean volume.")"
    else
        warn "$(L "Match:" "Matches:") ${BOLD}${TOTAL}"
        awk -F'\t' '{printf "      %s  %s\n", $1, $3}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local TABLE
    if [[ "$TOTAL" -gt 0 ]]; then
        TABLE=$(_rows_to_table "$(head -5000 "$OUT")" \
            "$(L "Regola" "Rule")" "$(L "File regole" "Rule file")" "$(L "Percorso" "Path")" \
            "$(L "Byte" "Bytes")" "SHA-256" "$(L "Stringhe" "Strings")" "Tag")
    else
        TABLE="<div style='padding:.6rem 0;font-size:.85rem'>$(L "Nessun match." "No match.")</div>"
    fi

    # Il cartiglio dell'ambito non e' un dettaglio: e' cio' che rende il
    # risultato interpretabile.
    local SCOPE; SCOPE=$("$PY3" - "$STATS" "${LANG:-en}" << 'PYEOF' 2>/dev/null
import json, sys, html

s = json.load(open(sys.argv[1]))
it = sys.argv[2] == 'it'


def L(i, e):
    return i if it else e


out = ['<div class="card" style="margin-bottom:1rem">'
       '<div style="padding:1rem 1.5rem;font-size:.8rem;line-height:1.7">']
out.append('<b>' + L('Ambito effettivo della scansione', 'Actual scope of the scan') + '</b><br>')
out.append(L("Non e' stato scansionato l'intero volume. Un \"nessun match\" vale solo per le "
             "posizioni elencate qui sotto.",
             'The whole volume was not scanned. "No match" only holds for the locations '
             'listed below.') + '<br><br>')
out.append('<table><tr><th>' + L('Posizione', 'Location') + '</th><th>'
           + L('File scansionati', 'Files scanned') + '</th></tr>')
for t in s['targets']:
    out.append('<tr><td class="mono">%s</td><td class="mono">%d</td></tr>'
               % (html.escape(t['path']), t['files']))
out.append('</table><br>')
nums = (s['skipped_too_big'], s['skipped_unreadable'], s['rule_files_compiled'])
out.append(L("File saltati perche' oltre il tetto per file: <b>%d</b>. Illeggibili o in "
             "errore: <b>%d</b>. File di regole compilati: <b>%d</b>." % nums,
             'Files skipped as over the per-file cap: <b>%d</b>. Unreadable or errored: '
             '<b>%d</b>. Rule files compiled: <b>%d</b>.' % nums))
if s['stopped_at_cap']:
    out.append('<br><b>' + L("La scansione si e' fermata al tetto massimo di file: e' PARZIALE.",
                             'The scan stopped at the maximum file cap: it is PARTIAL.') + '</b>')
for r in s['rule_files_rejected']:
    out.append('<br>' + L('Regole scartate: ', 'Rules rejected: ')
               + html.escape(r['file']) + ' — ' + html.escape(r['error']))
out.append('</div></div>')
print(''.join(out))
PYEOF
    )

    local SB
    SB="$(stat_box "Match" "$TOTAL" "$([[ "$TOTAL" -gt 0 ]] && echo warn || echo ok)")"
    SB+="$(stat_box "$(L "File scansionati" "Files scanned")" "${NSCAN:-0}" "info")"
    SB+="$(stat_box "$(L "Saltati" "Skipped")" "$(( ${NBIG:-0} + ${NBAD:-0} ))" "info")"
    SB+="$(stat_box "$(L "File regole" "Rule files")" "${NRULES:-0}" "info")"
    finish_report "yara" "YARA" "YAR" "$(basename "$YARA_RULES")" "$SB" \
        "${SCOPE}<div class='cards'>$(generic_card_html "$(L "Riscontri" "Matches")" "$YARA_RULES" "$TOTAL" "$TABLE" "⚑")</div>"
}
