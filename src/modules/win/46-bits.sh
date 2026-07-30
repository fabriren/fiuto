
# ================================================================
#  MODULO 46 — BITS Jobs (Background Intelligent Transfer Service)
#
#  BITS e' il servizio che Windows usa per scaricare aggiornamenti in
#  background. Essendo un componente firmato e legittimo, viene usato dagli
#  attaccanti per scaricare payload e per mantenere persistenza: un job BITS
#  con /SetNotifyCmdLine rieseguita un comando a ogni completamento, e il
#  traffico appare come una normale attivita' di sistema (T1197).
#
#  La coda dei job sta in qmgr.db (Windows 10+, formato ESE) o nei vecchi
#  qmgr0.dat/qmgr1.dat. Qui si estraggono URL, percorsi locali e nomi dei job
#  dalle stringhe UTF-16LE: e' l'approccio che funziona su tutti i formati e
#  senza dipendenze, a costo di non ricostruire la struttura dei record.
# ================================================================
module_bits() {
    section_header "BITS Jobs" "$ORANGE"
    check_win_root || return 1

    local -a DBS=()
    local D F
    D=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Network/Downloader")
    [[ -z "$D" ]] && D=$(ci_find_dir "$WIN_ROOT" "Documents and Settings/All Users/Application Data/Microsoft/Network/Downloader")
    if [[ -n "$D" ]]; then
        while IFS= read -r F; do
            [[ -n "$F" ]] && DBS+=("$F")
        done < <(find "$D" -maxdepth 1 -type f \( -iname 'qmgr.db' -o -iname 'qmgr[01].dat' \) 2>/dev/null)
    fi

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna coda BITS trovata (qmgr.db / qmgr0.dat)." "No BITS queue found (qmgr.db / qmgr0.dat).")"
        return 0
    fi
    info "$(L "Code BITS trovate:" "BITS queues found:") ${BOLD}${#DBS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os

out_path = sys.argv[1]
dbs = sys.argv[2:]

# Le stringhe in qmgr sono UTF-16LE. Si estraggono URL, percorsi locali e
# nomi di job, poi si classifica cosa merita attenzione.
URL  = re.compile(r'(?:https?|ftp)://[!-~]{4,400}')
WPATH = re.compile(r'[A-Za-z]:\\[^\x00<>|?*"]{3,250}')

# Domini di aggiornamento legittimi: servono a separare il rumore dal resto,
# non a dichiarare "sicuro" cio' che vi corrisponde.
MS_HOSTS = ('microsoft.com', 'windowsupdate.com', 'msftncsi.com', 'msedge.net',
            'windows.com', 'msn.com', 'live.com', 'office.net', 'office.com',
            'azureedge.net', 'akamaized.net', 'delivery.mp.microsoft.com')

# Estensioni che, scaricate via BITS, sono di per se' un segnale.
RISKY_EXT = ('.exe', '.dll', '.ps1', '.bat', '.cmd', '.scr', '.vbs', '.js',
             '.hta', '.jar', '.zip', '.7z', '.tmp', '.dat')

rows = []
seen = set()

def add(kind, value, src, flags):
    key = (kind, value)
    if key in seen:
        return
    seen.add(key)
    rows.append((kind, value, ';'.join(flags), src))

for db in dbs:
    src = os.path.basename(db)
    try:
        with open(db, 'rb') as fh:
            raw = fh.read(512 * 1024 * 1024)
    except Exception:
        continue
    # UTF-16LE e, per sicurezza, anche ASCII
    for enc in ('utf-16-le', 'latin-1'):
        try:
            text = raw.decode(enc, 'ignore')
        except Exception:
            continue
        for m in URL.finditer(text):
            u = m.group(0).rstrip('\x00').strip()
            low = u.lower()
            flags = []
            if not any(h in low for h in MS_HOSTS):
                flags.append('HOST_NON_MICROSOFT')
            if any(low.split('?')[0].endswith(e) for e in RISKY_EXT):
                flags.append('ESTENSIONE_A_RISCHIO')
            if low.startswith('http://'):
                flags.append('IN_CHIARO')
            add('URL', u[:400], src, flags)
        for m in WPATH.finditer(text):
            p = m.group(0).rstrip('\x00').strip()
            low = p.lower()
            if not any(low.endswith(e) for e in RISKY_EXT) and '\\temp\\' not in low and '\\users\\public\\' not in low:
                continue
            flags = []
            if '\\temp\\' in low or '\\users\\public\\' in low or '\\programdata\\' in low:
                flags.append('PERCORSO_SOSPETTO')
            if any(low.endswith(e) for e in RISKY_EXT):
                flags.append('ESTENSIONE_A_RISCHIO')
            add('PATH', p[:250], src, flags)

# Prima le voci segnalate.
rows.sort(key=lambda r: (r[2] == '', r[0], r[1]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ') for x in r) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun URL o percorso estratto dalla coda BITS." "No URL or path extracted from the BITS queue.")"
        return 0
    fi
    local NFLAG NURL
    NFLAG=$(awk -F'\t' '$3!=""' "$OUT" | wc -l)
    NURL=$(awk -F'\t' '$1=="URL"' "$OUT" | wc -l)

    ok "$(L "Voci estratte:" "Entries extracted:") ${BOLD}$TOTAL"
    info "URL: ${BOLD}${NURL}"
    if [[ "$NFLAG" -gt 0 ]]; then
        warn "$(L "Voci segnalate:" "Flagged entries:") ${BOLD}$NFLAG"
        awk -F'\t' '$3!=""{printf "      [%s] %s\n", $3, substr($2,1,100)}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${ORANGE}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4}' "$OUT" | head -20000)
    local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Tipo" "Type")" "$(L "Valore" "Value")" "$(L "Segnalazioni" "Flags")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Come leggere le segnalazioni" "How to read the flags")</b><br>"
    NOTE+="$(L "HOST_NON_MICROSOFT non significa 'malevolo': molti software legittimi usano BITS. Significa che quella voce non e' spiegabile come Windows Update e va verificata. Un job BITS verso un host non Microsoft che scarica un eseguibile in %TEMP% e' invece un indicatore forte." \
        "HOST_NON_MICROSOFT does not mean 'malicious': plenty of legitimate software uses BITS. It means the entry is not explainable as Windows Update and needs checking. A BITS job to a non-Microsoft host downloading an executable into %TEMP% is a strong indicator instead.")<br><br>"
    NOTE+="$(L "Da correlare con gli Event ID 3, 59, 60 del log Microsoft-Windows-Bits-Client/Operational (modulo Event Log)." \
        "Correlate with Event IDs 3, 59, 60 in Microsoft-Windows-Bits-Client/Operational (Event Log module).")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Voci" "Entries")" "$TOTAL")"
    STATS+="$(stat_box "URL" "$NURL" "info")"
    STATS+="$(stat_box "$(L "Segnalate" "Flagged")" "$NFLAG" "$([[ "$NFLAG" -gt 0 ]] && echo warn || echo info)")"
    finish_report "bits" "BITS Jobs" "BIT" "ProgramData/Microsoft/Network/Downloader" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Coda BITS" "BITS queue")" "${DBS[0]}" "$TOTAL" "$TABLE" "⇩")</div>"
}
