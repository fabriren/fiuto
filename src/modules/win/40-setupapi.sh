
# ================================================================
#  MODULO 40 — SetupAPI Device Log
#
#  setupapi.dev.log registra l'installazione dei driver di ogni dispositivo
#  collegato alla macchina. E' l'unica fonte che data con precisione la PRIMA
#  connessione di un dispositivo USB: il registro (modulo 8, USBSTOR) conserva
#  l'ultima connessione e i metadati, ma non il momento in cui il dispositivo
#  e' comparso per la prima volta.
#
#  In un caso di esfiltrazione la differenza conta: "questa chiavetta e' stata
#  collegata per la prima volta il giorno X" e' un'affermazione che si sostiene
#  solo con questo log.
# ================================================================
module_setupapi() {
    section_header "SetupAPI — $(L "Prima installazione dispositivi" "Device First Install")" "$BLUE"
    check_win_root || return 1

    local -a LOGS=()
    local D F
    D=$(ci_find_dir "$WIN_ROOT" "Windows/INF")
    if [[ -n "$D" ]]; then
        while IFS= read -r F; do
            [[ -n "$F" ]] && LOGS+=("$F")
        done < <(find "$D" -maxdepth 1 -type f -iname 'setupapi.dev*.log' 2>/dev/null)
    fi
    # Percorso legacy (XP/2003)
    D=$(ci_find_dir "$WIN_ROOT" "Windows")
    if [[ -n "$D" ]]; then
        F=$(ci_find_file "$D" "setupapi.log")
        [[ -n "$F" ]] && LOGS+=("$F")
    fi

    if [[ ${#LOGS[@]} -eq 0 ]]; then
        warn "$(L "Nessun setupapi.dev.log trovato." "No setupapi.dev.log found.")"
        return 0
    fi
    info "$(L "Log trovati:" "Logs found:") ${BOLD}${#LOGS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${LOGS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os

out_path = sys.argv[1]
logs = sys.argv[2:]

# Struttura del log:
#   >>>  [Device Install (Hardware initiated) - SWD\WPDBUSENUM\{GUID}#...]
#   >>>  Section start 2024/01/15 10:23:45.123
DEV = re.compile(r'^>>>\s+\[(?P<what>[^\]]+)\]')
SEC = re.compile(r'^>>>\s+Section start\s+(?P<ts>\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})')

# Bus che identificano supporti rimovibili: sono quelli che interessano in
# un'indagine su esfiltrazione.
REMOVABLE = ('USBSTOR', 'USB\\', 'WPDBUSENUM', 'SWD\\WPDBUSENUM', 'SCSI\\DISK')

rows = []
for path in logs:
    src = os.path.basename(path)
    try:
        fh = open(path, encoding='utf-8', errors='replace')
    except Exception:
        continue
    pending = None
    with fh:
        for line in fh:
            line = line.rstrip('\n')
            m = DEV.match(line)
            if m:
                pending = m.group('what').strip()
                continue
            m = SEC.match(line)
            if m and pending:
                ts = m.group('ts').replace('/', '-')
                what = pending
                pending = None
                # "Device Install (Hardware initiated) - <istanza>"
                inst = what.split(' - ', 1)[1] if ' - ' in what else what
                kind = what.split(' - ', 1)[0] if ' - ' in what else ''
                up = inst.upper()
                removable = '1' if any(k in up for k in REMOVABLE) else '0'
                rows.append((ts, kind, inst, removable, src))

rows.sort(key=lambda r: r[0])
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ') for x in r) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna installazione di dispositivo interpretabile." "No parsable device installation.")"
        return 0
    fi
    local NREM; NREM=$(awk -F'\t' '$4=="1"' "$OUT" | wc -l)

    ok "$(L "Installazioni registrate:" "Recorded installations:") ${BOLD}$TOTAL"
    info "$(L "Di cui supporti rimovibili:" "Of which removable media:") ${BOLD}$NREM"
    echo ""
    awk -F'\t' '$4=="1"{printf "      %s  %s\n", $1, substr($3,1,90)}' "$OUT" | head -25 | while IFS= read -r LN; do
        echo -e "      ${CYAN}${LN}${RESET}"
    done

    # Marca gli IoC eventualmente caricati.
    if [[ ${#IOC_LIST[@]} -gt 0 ]]; then
        local NIOC=0 LINE
        while IFS= read -r LINE; do
            check_ioc "$LINE" && NIOC=$((NIOC + 1))
        done < "$OUT"
        [[ $NIOC -gt 0 ]] && warn "$(L "Righe con match IoC:" "Lines matching IoC:") ${BOLD}$NIOC"
    fi

    ask_yn "Generare report HTML?" || return 0

    # Prima i rimovibili: sono il motivo per cui si guarda questo log.
    local ROWS; ROWS=$( { awk -F'\t' '$4=="1"{print $1"\t"$2"\t"$3"\t"$5}' "$OUT";
                          awk -F'\t' '$4=="0"{print $1"\t"$2"\t"$3"\t"$5}' "$OUT"; } )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Prima installazione" "First install")" "$(L "Tipo" "Kind")" \
        "$(L "Istanza dispositivo" "Device instance")" "Log")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' conta" "Why it matters")</b><br>"
    NOTE+="$(L "Il registro USBSTOR (modulo 8) conserva l'ULTIMA connessione di un dispositivo. Questo log conserva la PRIMA: e' l'unica fonte che permette di affermare quando un supporto e' comparso per la prima volta sulla macchina. I due vanno letti insieme." \
        "The USBSTOR registry (module 8) keeps a device's LAST connection. This log keeps the FIRST: it is the only source supporting a claim about when a device first appeared on the machine. Read the two together.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Installazioni" "Installations")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Rimovibili" "Removable")" "$NREM" "warn")"
    STATS+="$(stat_box "Log" "${#LOGS[@]}" "info")"
    finish_report "setupapi" "SetupAPI Device Log" "DEV" "Windows/INF/setupapi.dev.log" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Installazioni dispositivi" "Device installations")" "${LOGS[0]}" "$TOTAL" "$TABLE" "⇄")</div>"
}
