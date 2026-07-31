
# --- macOS 12 — Spotlight (provenienza download e metadati) ---
#
# Lo store Spotlight indicizza attributi che sopravvivono al file indicizzato,
# fra cui kMDItemWhereFroms — l'URL da cui un file e' stato scaricato — e i nomi
# di file poi cancellati.
#
# Il formato di store.db e' proprietario, compresso a blocchi, e un parser
# completo esula da un tool senza dipendenze: qui si fa un'estrazione EURISTICA
# delle stringhe leggibili (URL e percorsi). Il report lo dichiara apertamente e
# rimanda a spotlight_parser per l'analisi strutturata.
module_macos_spotlight() {
    section_header "macOS - Spotlight" "$CYAN"
    check_target_root || return 1

    local -a STORES=()
    local D
    while IFS= read -r D; do
        [[ -n "$D" ]] && STORES+=("$D")
    done < <(find "$WIN_ROOT" -maxdepth 6 -type d -name "Store-V2" 2>/dev/null
             find "$WIN_ROOT" -maxdepth 6 -type d -name ".Spotlight-V100" 2>/dev/null)

    mapfile -t DBS < <(find "$WIN_ROOT" -maxdepth 8 -type f \( -name "store.db" -o -name ".store.db" \) 2>/dev/null)
    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessuno store Spotlight trovato." "No Spotlight store found.")"
        return 0
    fi
    info "$(L "Store Spotlight trovati:" "Spotlight stores found:") ${BOLD}${#DBS[@]}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, re, os

out_path = sys.argv[1]
dbs = sys.argv[2:]

# Estrazione euristica: si cercano URL e percorsi nelle stringhe leggibili.
URL = re.compile(rb'(?:https?|ftp)://[!-~]{4,300}')
# Percorsi utente: evitano il rumore dei path di sistema.
PATH = re.compile(rb'/Users/[A-Za-z0-9._-]{1,40}/[!-~]{3,200}')

rows = []
seen = set()
for db in dbs:
    try:
        size = os.path.getsize(db)
        with open(db, 'rb') as fh:
            data = fh.read(256 * 1024 * 1024)   # tetto di sicurezza: 256 MB
    except Exception:
        continue
    src = db
    for kind, rx in (("URL", URL), ("path", PATH)):
        for m in rx.finditer(data):
            try:
                val = m.group(0).decode('utf-8', 'strict')
            except Exception:
                continue
            # Scarta stringhe con caratteri di controllo residui
            if any(ord(c) < 32 for c in val):
                continue
            key = (kind, val)
            if key in seen:
                continue
            seen.add(key)
            rows.append((kind, val, src))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write("\t".join(x.replace("\t", " ") for x in r) + "\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    local NURL=0 NPATH=0
    if [[ "$TOTAL" -gt 0 ]]; then
        NURL=$(awk -F'\t' '$1=="URL"' "$OUT" | wc -l)
        NPATH=$(awk -F'\t' '$1=="path"' "$OUT" | wc -l)
    fi

    # Inventario degli store, utile anche quando l'estrazione non produce nulla.
    local INV="" 
    local DB
    for DB in "${DBS[@]}"; do
        local SZ MT
        SZ=$(stat -c %s "$DB" 2>/dev/null || echo "?")
        MT=$(stat -c %y "$DB" 2>/dev/null | cut -d. -f1 || echo "?")
        INV+="$(html_esc "$DB")	${SZ}	${MT}
"
    done
    local INVTABLE; INVTABLE=$(_rows_to_table "$INV" "store.db" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")")

    separator
    ok "$(L "Stringhe estratte:" "Strings extracted:") ${BOLD}$TOTAL"
    info "URL: ${BOLD}${NURL}${RESET}  ·  path: ${BOLD}${NPATH}"

    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna stringa utile estratta dagli store Spotlight." "No useful string extracted from the Spotlight stores.")"
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Metodo" "Method")</b><br>"
    BODY+="$(L "store.db ha un formato proprietario compresso a blocchi. Questo modulo NON lo decodifica: esegue un'estrazione euristica delle stringhe leggibili (URL e percorsi utente), utile per recuperare provenienza dei download e nomi di file poi cancellati. I risultati vanno considerati indizi da confermare, non un dump strutturato dell'indice." \
        "store.db uses a proprietary block-compressed format. This module does NOT decode it: it performs a heuristic extraction of readable strings (URLs and user paths), useful to recover download provenance and names of later-deleted files. Treat results as leads to confirm, not as a structured index dump.")<br><br>"
    BODY+="$(L "Per un'analisi strutturata (kMDItemWhereFroms, date, attributi completi) usa" "For structured analysis (kMDItemWhereFroms, dates, full attributes) use") "
    BODY+="<code>spotlight_parser</code> (Yogesh Khatri)."
    BODY+="</div></div>"
    BODY+="<div class='cards'>$(generic_card_html "$(L "Store trovati" "Stores found")" "$WIN_ROOT" "${#DBS[@]}" "$INVTABLE" "▤")</div>"

    if [[ "$TOTAL" -gt 0 ]]; then
        local ROWS; ROWS=$( { awk -F'\t' '$1=="URL"' "$OUT"; awk -F'\t' '$1=="path"' "$OUT"; } | head -20000 )
        local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Tipo" "Type")" "$(L "Valore" "Value")" "store.db")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Stringhe estratte" "Extracted strings")" "$(L "estrazione euristica" "heuristic extraction")" "$TOTAL" "$TABLE" "⌕")</div>"
    fi

    local STATS
    STATS="$(stat_box "URL" "$NURL" "warn")"
    STATS+="$(stat_box "$(L "Percorsi" "Paths")" "$NPATH")"
    STATS+="$(stat_box "Store" "${#DBS[@]}" "info")"
    finish_report "macos_spotlight" "macOS Spotlight" "SPT" ".Spotlight-V100 / store.db" "$STATS" "$BODY"
}
