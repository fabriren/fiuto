
# ================================================================
#  CROSS-OS — Recupero da spazio libero SQLite
#
#  Praticamente ogni artefatto moderno e' un database SQLite: cronologia dei
#  browser, TCC, KnowledgeC, quarantena, Messages, cookie di Chrome, chat.
#  Tutti i moduli che li leggono vedono pero' solo i record VIVI.
#
#  Un record cancellato non sparisce dal file: la pagina finisce nella
#  freelist, o lo spazio che occupava diventa un freeblock. Il contenuto resta
#  leggibile finche' non viene sovrascritto. Quando l'utente "ha cancellato la
#  cronologia", questo e' spesso l'unico posto dove quella cronologia esiste
#  ancora.
#
#  Il modulo attraversa i database del volume e ne estrae le stringhe dallo
#  spazio non allocato. Vale per tutti e tre i sistemi operativi.
# ================================================================
module_xplat_sqlite_recovery() {
    section_header "$(L "Recupero record cancellati — SQLite" "Deleted record recovery — SQLite")" "$MAGENTA"
    check_target_root || return 1

    # Database che valgono la pena in un'indagine, per OS.
    local -a PATTERNS=(
        "History" "places.sqlite" "cookies.sqlite" "Cookies" "Web Data" "Login Data"
        "chat.db" "TCC.db" "knowledgeC.db" "QuarantineEventsV2*" "Downloads.sqlite"
        "ActivitiesCache.db" "SRUDB.dat" "index.sqlite" "Extension Cookies"
        "History.db" "*.sqlite" "*.db"
    )

    info "$(L "Ricerca dei database SQLite..." "Searching for SQLite databases...")"
    local MANIFEST; MANIFEST=$(mktemp); register_tmp "$MANIFEST"

    # Si limita alle aree utente e alle directory applicative: una scansione
    # dell'intero volume su un disco reale richiederebbe ore e restituirebbe
    # soprattutto database di sistema senza interesse.
    local -a ROOTS=()
    local HOME_DIR
    while IFS= read -r HOME_DIR; do
        [[ -n "$HOME_DIR" ]] && ROOTS+=("$HOME_DIR")
    done < <(get_target_user_homes)
    local D
    for D in "ProgramData" "private/var/db" "var/db" "var/lib"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ROOTS+=("$R")
    done
    [[ ${#ROOTS[@]} -eq 0 ]] && ROOTS=("$WIN_ROOT")

    local R F
    for R in "${ROOTS[@]}"; do
        while IFS= read -r F; do
            # Il magic e' l'unico criterio affidabile: molti database SQLite
            # non hanno estensione .db o .sqlite (Chrome li chiama "History").
            [[ -s "$F" ]] || continue
            if [[ "$(head -c 15 "$F" 2>/dev/null)" == "SQLite format 3" ]]; then
                printf '%s\n' "$F" >> "$MANIFEST"
            fi
        done < <(find "$R" -maxdepth 8 -type f -size +1k -size -512M 2>/dev/null | head -4000)
    done

    if [[ ! -s "$MANIFEST" ]]; then
        warn "$(L "Nessun database SQLite trovato." "No SQLite database found.")"
        return 0
    fi
    sort -u "$MANIFEST" -o "$MANIFEST"
    local NDB; NDB=$(wc -l < "$MANIFEST")
    info "$(L "Database SQLite individuati:" "SQLite databases found:") ${BOLD}$NDB"
    info "$(L "Carving dello spazio non allocato..." "Carving unallocated space...")"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local SUM; SUM=$(mktemp); register_tmp "$SUM"

    run_py_with_lib pylib_sqlite_recover "$MANIFEST" "$IOCTMP" "$OUT" "$SUM" << 'PYEOF' 2>/dev/null
import sys, os, re

manifest, ioc_path, out_path, sum_path = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

# Un flag binario "rilevante" non aiuta: su una cronologia cancellata ogni
# stringa contiene un URL e finirebbe marcata, azzerando il valore di triage.
# Si classifica invece per tipo, e si ordina mettendo per prime le categorie
# rare (credenziali, eseguibili) rispetto a quelle abbondanti (URL).
CAT_CRED = re.compile(r'(?i)(password|passwd|token|secret|api[_-]?key|bearer |authorization|BEGIN [A-Z ]*PRIVATE KEY)')
CAT_EXE  = re.compile(r'(?i)\.(exe|dll|ps1|vbs|js|bat|cmd|scr|sh|jar|apk|dmg|pkg)\b')
CAT_DOC  = re.compile(r'(?i)\.(docx?|xlsx?|pptx?|pdf|zip|7z|rar|csv|sql|bak)\b')
CAT_PATH = re.compile(r'(?i)(/Users/|/home/|[A-Z]:\\\\)')
CAT_URL  = re.compile(r'(?i)(https?://|ftp://)')

def classify(t):
    if CAT_CRED.search(t):
        return 'credenziale'
    if CAT_EXE.search(t):
        return 'eseguibile'
    if CAT_DOC.search(t):
        return 'documento'
    if CAT_PATH.search(t):
        return 'percorso'
    if CAT_URL.search(t):
        return 'url'
    return ''

# Ordine di priorita' per il triage: le categorie rare prima.
PRIORITY = {'credenziale': 0, 'eseguibile': 1, 'documento': 2, 'percorso': 3, 'url': 4, '': 5}

rows = []
per_db = []
for line in open(manifest, encoding='utf-8', errors='replace'):
    path = line.rstrip('\n')
    if not path:
        continue
    try:
        found = sqlite_carve(path, cap=4000)
    except Exception:
        continue
    if not found:
        continue
    name = os.path.basename(path)
    nnot = 0
    for origin, text in found:
        cat = classify(text)
        if iocs and any(i in text.lower() for i in iocs):
            cat = (cat + ';IOC') if cat else 'IOC'
        if cat.split(';')[0] in ('credenziale', 'eseguibile', 'documento') or 'IOC' in cat:
            nnot += 1
        rows.append((name, origin, cat, text, path))
    per_db.append((name, str(len(found)), str(nnot), path))

rows.sort(key=lambda r: (0 if 'IOC' in r[2] else 1,
                         PRIORITY.get(r[2].split(';')[0], 5), r[0]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
with open(sum_path, 'w', encoding='utf-8') as fh:
    for r in sorted(per_db, key=lambda x: -int(x[1])):
        fh.write('\t'.join(r) + '\n')
PYEOF

    local TOTAL=0 NDBHIT=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    [[ -s "$SUM" ]] && NDBHIT=$(wc -l < "$SUM")
    if [[ "$TOTAL" -eq 0 ]]; then
        ok "$(L "Nessun contenuto recuperabile dallo spazio libero dei database." "No recoverable content in the databases' free space.")"
        return 0
    fi
    local NNOT NIOC
    # "Prioritarie" = credenziali, eseguibili, documenti. Gli URL sono esclusi
    # di proposito: su una cronologia cancellata sono la norma, non il segnale.
    NNOT=$(awk -F'\t' '$3 ~ /credenziale|eseguibile|documento/' "$OUT" | wc -l)
    NIOC=$(awk -F'\t' '$3 ~ /IOC/' "$OUT" | wc -l)

    ok "$(L "Stringhe recuperate:" "Strings recovered:") ${BOLD}$TOTAL"
    info "$(L "Da" "From") ${BOLD}${NDBHIT}${RESET} $(L "database su" "databases out of") ${NDB}"
    [[ "$NNOT" -gt 0 ]] && warn "$(L "Prioritarie (credenziali, eseguibili, documenti):" "Priority (credentials, executables, documents):") ${BOLD}$NNOT"
    [[ "$NIOC" -gt 0 ]] && warn "$(L "Con match IoC:" "With IoC match:") ${BOLD}$NIOC"
    echo ""
    awk -F'\t' '$3 ~ /credenziale|eseguibile|documento|IOC/{printf "      %-14s %-12s %s\n", $1, $3, substr($4,1,70)}' "$OUT" | head -15 | while IFS= read -r LN; do
        echo -e "      ${MAGENTA}${LN}${RESET}"
    done

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(head -25000 "$OUT" | awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4}')
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "Database" "$(L "Origine" "Origin")" "$(L "Categoria" "Category")" "$(L "Contenuto recuperato" "Recovered content")")
    local STABLE; STABLE=$(_rows_to_table "$(cat "$SUM")" \
        "Database" "$(L "Stringhe" "Strings")" "$(L "Prioritarie" "Priority")" "$(L "Percorso" "Path")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Che cosa sono queste stringhe" "What these strings are")</b><br>"
    NOTE+="$(L "Provengono dallo spazio che SQLite considera libero: pagine finite nella freelist (colonna 'freelist') e spazio non allocato dentro pagine ancora in uso (colonna 'unallocated'). In pratica sono i resti di record CANCELLATI, ancora presenti perche' non sovrascritti." \
        "They come from space SQLite considers free: pages moved to the freelist (origin 'freelist') and unallocated space inside pages still in use (origin 'unallocated'). In practice they are the remains of DELETED records, still present because not yet overwritten.")<br><br>"
    NOTE+="<b>$(L "Limite: sono frammenti, non record" "Limitation: fragments, not records")</b><br>"
    NOTE+="$(L "Il modulo estrae stringhe, non ricostruisce righe di tabella: campi adiacenti possono comparire concatenati e all'inizio puo' esserci qualche byte di intestazione del record. Ricostruire un record cancellato richiede lo schema e l'interpretazione dei serial type, ed e' un'operazione fragile: un record ricomposto male in una perizia e' peggio di nessun record. Va quindi trattato come indizio da corroborare, non come contenuto autoritativo." \
        "The module extracts strings, it does not rebuild table rows: adjacent fields may appear concatenated and a few record-header bytes may prefix the text. Rebuilding a deleted record requires the schema and serial-type interpretation, and is fragile: a badly reassembled record in an expert report is worse than none. Treat this as a lead to corroborate, not as authoritative content.")<br><br>"
    NOTE+="<b>$(L "Le categorie" "The categories")</b><br>"
    NOTE+="$(L "Le stringhe sono classificate per tipo e ordinate mettendo per prime le categorie rare. Gli URL sono deliberatamente in fondo: recuperare una cronologia cancellata produce migliaia di URL, quindi marcarli tutti come rilevanti non aiuterebbe a decidere da dove iniziare." \
        "Strings are classified by type and ordered with the rare categories first. URLs are deliberately last: recovering a deleted history yields thousands of URLs, so flagging them all as notable would not help decide where to start.")<br><br>"
    NOTE+="$(L "Non trovare nulla non significa che l'utente non abbia cancellato: significa che lo spazio e' stato riutilizzato, o che il database e' stato compattato con VACUUM — operazione che azzera proprio questo tipo di recupero ed e' essa stessa degna di nota." \
        "Finding nothing does not mean the user deleted nothing: it means the space was reused, or the database was compacted with VACUUM — an operation that wipes exactly this kind of recovery and is itself worth noting.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Stringhe" "Strings")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Prioritarie" "Priority")" "$NNOT" "$([[ "$NNOT" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Database" "${NDBHIT}/${NDB}" "info")"
    STATS+="$(stat_box "IoC" "$NIOC" "$([[ "$NIOC" -gt 0 ]] && echo warn || echo info)")"
    finish_report "xplat_sqlite_recovery" "SQLite — $(L "record cancellati" "deleted records")" "SQL" \
        "$(L "freelist e spazio non allocato" "freelist and unallocated space")" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Resa per database" "Yield per database")" "$(L "riepilogo" "summary")" "$NDBHIT" "$STABLE" "∑")</div><div class='cards'>$(generic_card_html "$(L "Contenuto recuperato" "Recovered content")" "$(L "segnalati in testa" "flagged first")" "$TOTAL" "$TABLE" "♺")</div>"
}
