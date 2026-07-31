
# --- LINUX 15 — auditd (/var/log/audit) ---
#
# Su RHEL/CentOS/Fedora e su ogni sistema hardenizzato auditd e' la fonte
# primaria di intrusion detection: registra syscall, esecuzioni, autenticazioni
# e violazioni di policy con un dettaglio che syslog non ha. Il modulo "System
# Logs" copre /var/log testuali ma non tocca audit.log, che ha un formato
# proprio (campi chiave=valore, stringhe in esadecimale, timestamp epoch).
module_linux_auditd() {
    section_header "Linux - auditd" "$RED"
    check_target_root || return 1

    local AUDITDIR; AUDITDIR=$(ci_find_dir "$WIN_ROOT" "var/log/audit")
    if [[ -z "$AUDITDIR" ]]; then
        warn "$(L "Directory /var/log/audit non trovata (auditd non installato o log altrove)." \
                 "/var/log/audit directory not found (auditd not installed or logs elsewhere).")"
        return 0
    fi
    mapfile -t LOGS < <(find "$AUDITDIR" -maxdepth 1 -type f -name 'audit.log*' -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
    if [[ ${#LOGS[@]} -eq 0 ]]; then
        warn "$(L "Nessun audit.log presente." "No audit.log present.")"
        return 0
    fi
    info "$(L "File di audit trovati:" "Audit files found:") ${BOLD}${#LOGS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local SUM; SUM=$(mktemp); register_tmp "$SUM"
    "$PY3" - "$OUT" "$SUM" "${LOGS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, datetime, collections, binascii

out_path, sum_path = sys.argv[1], sys.argv[2]
files = sys.argv[3:]

HDR = re.compile(r'type=(?P<type>\S+)\s+msg=audit\((?P<epoch>\d+)\.(?P<ms>\d+):(?P<serial>\d+)\):\s*(?P<rest>.*)')
KV  = re.compile(r'(\w+)=("([^"]*)"|\S+)')

# Campi che auditd puo' emettere in esadecimale quando contengono spazi o
# caratteri speciali: senza decodifica il comando dell'attaccante resta
# illeggibile proprio nei casi piu' interessanti.
HEXFIELDS = {"proctitle", "exe", "cmd", "comm", "name", "cwd", "a0", "a1", "a2", "a3"}

# Tipi che pesano in un'indagine.
NOTABLE = {
    "USER_AUTH", "USER_LOGIN", "USER_ACCT", "USER_START", "USER_CMD",
    "ADD_USER", "DEL_USER", "ADD_GROUP", "DEL_GROUP", "USER_CHAUTHTOK",
    "USER_ROLE_CHANGE", "ROLE_ASSIGN", "ANOM_ABEND", "AVC", "SECCOMP",
    "EXECVE", "ANOM_PROMISCUOUS", "CONFIG_CHANGE", "MAC_POLICY_LOAD",
}

def unhex(val):
    """auditd codifica alcuni valori come esadecimale puro (senza virgolette)."""
    if len(val) >= 4 and len(val) % 2 == 0 and re.fullmatch(r'[0-9A-Fa-f]+', val):
        try:
            return binascii.unhexlify(val).decode('utf-8', 'replace').replace('\x00', ' ').strip()
        except Exception:
            return val
    return val

counts = collections.Counter()
rows = []

for fp in files:
    try:
        fh = open(fp, encoding='utf-8', errors='replace')
    except Exception:
        continue
    with fh:
        for line in fh:
            m = HDR.match(line.strip())
            if not m:
                continue
            typ = m.group('type')
            counts[typ] += 1
            try:
                ts = datetime.datetime.utcfromtimestamp(int(m.group('epoch'))).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                ts = ''
            fields = {}
            for km in KV.finditer(m.group('rest')):
                key = km.group(1)
                if km.group(3) is not None:
                    val = km.group(3)          # valore fra virgolette
                else:
                    # Valore non quotato. auditd annida i sottocampi dentro
                    # msg='...': l'ultimo di questi si porta dietro l'apice di
                    # chiusura (res=failed'), che va tolto o il confronto salta.
                    val = km.group(2).strip().rstrip("',")
                    if key in HEXFIELDS:
                        val = unhex(val)
                fields[key] = val

            # Ricompone il comando dagli argomenti a0..aN di EXECVE.
            if typ == 'EXECVE':
                args = []
                i = 0
                while f'a{i}' in fields:
                    args.append(fields[f'a{i}'])
                    i += 1
                detail = ' '.join(args) if args else m.group('rest')[:300]
            else:
                parts = []
                for key in ('res', 'acct', 'uid', 'auid', 'exe', 'comm', 'terminal',
                            'hostname', 'addr', 'op', 'key', 'proctitle', 'cmd', 'name'):
                    if key in fields and fields[key] not in ('', '?', '(none)'):
                        parts.append(f"{key}={fields[key]}")
                detail = ' '.join(parts) or m.group('rest')[:300]

            user = fields.get('acct') or fields.get('auid') or fields.get('uid') or ''
            failed = '1' if fields.get('res') in ('failed', 'fail', '0') else '0'
            notable = '1' if (typ in NOTABLE or failed == '1') else '0'
            rows.append((ts, typ, str(user), detail[:500], notable, failed))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write("\t".join(x.replace("\t", " ").replace("\n", " ") for x in r) + "\n")

with open(sum_path, 'w', encoding='utf-8') as fh:
    for typ, n in counts.most_common():
        fh.write(f"{typ}\t{n}\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun record auditd interpretabile." "No parsable auditd record.")"
        return 0
    fi
    local NFAIL NNOTE
    NFAIL=$(awk -F'\t' '$6=="1"' "$OUT" | wc -l)
    NNOTE=$(awk -F'\t' '$5=="1"' "$OUT" | wc -l)
    local NEXEC; NEXEC=$(awk -F'\t' '$2=="EXECVE"' "$OUT" | wc -l)

    ok "$(L "Record auditd:" "auditd records:") ${BOLD}$TOTAL"
    info "EXECVE: ${BOLD}${NEXEC}${RESET}  ·  $(L "esiti negativi:" "failed outcomes:") ${BOLD}${NFAIL}"
    echo ""
    info "$(L "Distribuzione per tipo:" "Breakdown by type:")"
    head -12 "$SUM" | while IFS=$'\t' read -r T N; do
        printf "      ${DIM}%-24s %6s${RESET}\n" "$T" "$N"
    done

    ask_yn "Generare report HTML?" || return 0

    # Prima gli eventi rilevanti, poi il resto.
    local ROWS; ROWS=$( { awk -F'\t' '$5=="1"{print $1"\t"$2"\t"$3"\t"$4}' "$OUT";
                          awk -F'\t' '$5=="0"{print $1"\t"$2"\t"$3"\t"$4}' "$OUT"; } | head -20000 )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Data (UTC)" "Date (UTC)")" "$(L "Tipo" "Type")" "$(L "Utente" "User")" "$(L "Dettaglio" "Detail")")
    local SUMTABLE; SUMTABLE=$(_rows_to_table "$(cat "$SUM")" "$(L "Tipo" "Type")" "$(L "Occorrenze" "Occurrences")")

    local BODY="<div class='cards'>"
    BODY+=$(generic_card_html "$(L "Distribuzione per tipo" "Breakdown by type")" "$AUDITDIR" "$TOTAL" "$SUMTABLE" "∑")
    BODY+=$(generic_card_html "$(L "Eventi" "Events")" "$(L "rilevanti in testa" "notable first")" "$TOTAL" "$TABLE" "⚑")
    BODY+="</div>"

    local STATS
    STATS="$(stat_box "$(L "Record" "Records")" "$TOTAL")"
    STATS+="$(stat_box "EXECVE" "$NEXEC" "info")"
    STATS+="$(stat_box "$(L "Esiti negativi" "Failed")" "$NFAIL" "warn")"
    STATS+="$(stat_box "$(L "Rilevanti" "Notable")" "$NNOTE" "warn")"
    finish_report "linux_auditd" "Linux auditd" "AUD" "/var/log/audit/audit.log" "$STATS" "$BODY"
}
