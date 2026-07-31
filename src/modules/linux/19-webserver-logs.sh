
# ================================================================
#  LINUX 19 — Log dei web server
#
#  Su un server esposto la compromissione iniziale passa quasi sempre da qui.
#  Il modulo System Logs legge /var/log testuali generici, ma non tocca gli
#  access log di nginx e Apache, che hanno un formato proprio e un volume tale
#  da richiedere un'analisi mirata.
#
#  Cosa cerca: richieste verso webshell note, path traversal, tentativi di
#  SQL injection, upload, User-Agent di strumenti offensivi, e i codici di
#  risposta che distinguono un tentativo fallito da uno riuscito.
# ================================================================
module_linux_webserver_logs() {
    section_header "Linux - $(L "Log web server" "Web server logs")" "$ORANGE"
    check_target_root || return 1

    local -a LOGS=()
    local D F
    for D in "var/log/nginx" "var/log/apache2" "var/log/httpd" "var/log/lighttpd" "var/log/caddy"; do
        local DIR; DIR=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$DIR" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && LOGS+=("$F")
        done < <(find "$DIR" -maxdepth 2 -type f \( -name '*access*' -o -name '*error*' \) ! -name '*.gz' 2>/dev/null | sort)
    done
    if [[ ${#LOGS[@]} -eq 0 ]]; then
        warn "$(L "Nessun log di web server trovato." "No web server log found.")"
        return 0
    fi
    info "$(L "Log trovati:" "Logs found:") ${BOLD}${#LOGS[@]}"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local SUM; SUM=$(mktemp); register_tmp "$SUM"

    "$PY3" - "$OUT" "$SUM" "$IOCTMP" "${LOGS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os, collections

out_path, sum_path, ioc_path = sys.argv[1], sys.argv[2], sys.argv[3]
logs = sys.argv[4:]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

# Combined log format: IP - user [data] "METODO path proto" status size "ref" "ua"
CLF = re.compile(
    r'^(?P<ip>\S+)\s+\S+\s+(?P<user>\S+)\s+\[(?P<ts>[^\]]+)\]\s+'
    r'"(?P<method>[A-Z]+)\s+(?P<path>[^"\s]*)[^"]*"\s+(?P<status>\d{3})\s+(?P<size>\S+)'
    r'(?:\s+"(?P<ref>[^"]*)"\s+"(?P<ua>[^"]*)")?')

MONTHS = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
          'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}

def iso(ts):
    # 10/Oct/2024:13:55:36 +0000
    try:
        d, t = ts.split(':', 1)
        day, mon, year = d.split('/')
        return f"{year}-{MONTHS.get(mon,'01')}-{int(day):02d} {t.split()[0]}"
    except Exception:
        return ''

# Pattern di attacco: il nome del gruppo diventa l'etichetta nel report.
ATTACKS = [
    ('WEBSHELL',   re.compile(r'(?i)(c99|r57|b374k|wso\.php|shell\.php|cmd\.php|adminer\.php|alfa\.php|tiny\.php|/\.well-known/[^ ]*\.php)')),
    ('TRAVERSAL',  re.compile(r'(?i)(\.\./|%2e%2e[/%]|/etc/passwd|/proc/self/environ|\\\.\\\.)')),
    ('SQLI',       re.compile(r"(?i)(union[\s+]+select|' or '1'='1|sleep\(\d|benchmark\(|information_schema|xp_cmdshell)")),
    ('RCE',        re.compile(r'(?i)(\bwget\b|\bcurl\b.+http|/bin/(ba)?sh|nc\s+-e|python\s+-c|base64\s+-d|\$\(.*\)|%24%28)')),
    ('UPLOAD',     re.compile(r'(?i)(multipart/form-data|\.php[3457]?(\?|$)|\.jsp(\?|$)|\.aspx?(\?|$)).*(POST|PUT)')),
    ('LOG4J',      re.compile(r'(?i)\$\{jndi:')),
    ('SCANNER_UA', re.compile(r'(?i)(sqlmap|nikto|nmap|masscan|dirbuster|gobuster|wpscan|acunetix|nuclei|feroxbuster|zgrab)')),
]

rows = []
per_ip = collections.Counter()
per_attack = collections.Counter()
status_of_attack = collections.Counter()

for path in logs:
    src = os.path.basename(path)
    try:
        fh = open(path, encoding='utf-8', errors='replace')
    except Exception:
        continue
    with fh:
        for line in fh:
            line = line.rstrip('\n')
            if not line:
                continue
            m = CLF.match(line)
            ip = m.group('ip') if m else ''
            status = m.group('status') if m else ''
            when = iso(m.group('ts')) if m else ''
            req = (m.group('path') if m else line)
            ua = (m.group('ua') or '') if m else ''
            hay = line
            hits = [name for name, rx in ATTACKS if rx.search(hay)]
            if iocs and any(i in hay.lower() for i in iocs):
                hits.append('IOC')
            if not hits:
                continue
            for h in hits:
                per_attack[h] += 1
            if ip:
                per_ip[ip] += 1
            # Uno status 200/500 su una richiesta di attacco vale molto piu' di
            # un 404: distingue il tentativo dall'esito.
            if status.startswith(('2', '5')) or status == '301':
                status_of_attack['riuscita_o_errore_server'] += 1
            rows.append((when, ip, status, ';'.join(hits), req[:300], ua[:150], src))

rows.sort(key=lambda r: r[0])
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')

with open(sum_path, 'w', encoding='utf-8') as fh:
    for name, n in per_attack.most_common():
        fh.write(f"attacco\t{name}\t{n}\n")
    for ip, n in per_ip.most_common(40):
        fh.write(f"ip\t{ip}\t{n}\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessuna richiesta sospetta rilevata nei log analizzati." "No suspicious request found in the analysed logs.")"
        return 0
    fi
    local NHIT2XX; NHIT2XX=$(awk -F'\t' '$3 ~ /^[25]/' "$OUT" | wc -l)

    warn "$(L "Richieste sospette:" "Suspicious requests:") ${BOLD}$TOTAL"
    warn "$(L "Di cui con risposta 2xx/5xx (possibile successo):" "Of which answered 2xx/5xx (possible success):") ${BOLD}$NHIT2XX"
    echo ""
    info "$(L "Per tipo:" "By type:")"
    awk -F'\t' '$1=="attacco"{printf "      %-14s %6s\n", $2, $3}' "$SUM" | head -10
    echo ""
    info "$(L "IP piu' attivi:" "Most active IPs:")"
    awk -F'\t' '$1=="ip"{printf "      %-40s %6s\n", $2, $3}' "$SUM" | head -8

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$3 ~ /^[25]/' "$OUT"; awk -F'\t' '$3 !~ /^[25]/' "$OUT"; } | head -20000 )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data" "Date")" "IP" "$(L "Stato" "Status")" "$(L "Tipo" "Type")" \
        "$(L "Richiesta" "Request")" "User-Agent" "Log")
    local ATABLE; ATABLE=$(_rows_to_table "$(awk -F'\t' '$1=="attacco"{print $2"\t"$3}' "$SUM")" "$(L "Tipo" "Type")" "$(L "Occorrenze" "Occurrences")")
    local ITABLE; ITABLE=$(_rows_to_table "$(awk -F'\t' '$1=="ip"{print $2"\t"$3}' "$SUM")" "IP" "$(L "Richieste sospette" "Suspicious requests")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Lo stato HTTP e' la parte che conta" "The HTTP status is the part that matters")</b><br>"
    NOTE+="$(L "Un log esposto su Internet contiene sempre migliaia di tentativi automatici: la loro presenza non e' un incidente. Cio' che distingue un tentativo da una compromissione e' la risposta: un 404 su una webshell e' rumore di fondo, un 200 sulla stessa richiesta significa che la webshell esiste. Le righe con risposta 2xx e 5xx sono messe in cima per questo motivo." \
        "A log exposed to the Internet always contains thousands of automated attempts: their presence is not an incident. What separates an attempt from a compromise is the response: a 404 on a webshell is background noise, a 200 on the same request means the webshell is there. Rows answered 2xx and 5xx are listed first for this reason.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Richieste sospette" "Suspicious requests")" "$TOTAL" "warn")"
    STATS+="$(stat_box "2xx/5xx" "$NHIT2XX" "$([[ "$NHIT2XX" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Log" "${#LOGS[@]}" "info")"
    finish_report "linux_webserver" "Linux Web Server Logs" "WWW" "nginx · apache · lighttpd" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Per tipo di attacco" "By attack type")" "$(L "riepilogo" "summary")" "$TOTAL" "$ATABLE" "∑")$(generic_card_html "$(L "IP piu' attivi" "Most active IPs")" "$(L "riepilogo" "summary")" "$TOTAL" "$ITABLE" "◉")</div><div class='cards'>$(generic_card_html "$(L "Richieste" "Requests")" "$(L "risposte 2xx/5xx in testa" "2xx/5xx responses first")" "$TOTAL" "$TABLE" "⚑")</div>"
}
