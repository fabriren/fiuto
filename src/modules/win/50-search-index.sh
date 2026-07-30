
# ================================================================
#  MODULO 50 — Windows Search Index (Windows.edb)
#
#  L'indice di ricerca di Windows conserva nome, percorso, autore e — per molti
#  formati — un estratto del CONTENUTO dei file indicizzati. L'indice non viene
#  ripulito quando un file viene cancellato: la voce sopravvive fino al
#  successivo passaggio dell'indicizzatore.
#
#  Di conseguenza Windows.edb e' spesso l'unica fonte che conserva testo di
#  documenti eliminati, e i percorsi di file che non esistono piu' sul volume.
#
#  Il file e' un database ESE che puo' superare il gigabyte. Con libesedb si
#  legge la tabella SystemIndex_Gthr; senza, si ripiega sull'estrazione dei
#  percorsi dalle stringhe, che resta utile per sapere COSA c'era.
# ================================================================
module_search_index() {
    section_header "Windows Search Index" "$YELLOW"
    check_win_root || return 1

    local -a DBS=()
    local D F
    for D in "ProgramData/Microsoft/Search/Data/Applications/Windows" \
             "Documents and Settings/All Users/Application Data/Microsoft/Search/Data/Applications/Windows"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$R" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && DBS+=("$F")
        done < <(find "$R" -maxdepth 1 -type f -iname 'Windows.edb' 2>/dev/null)
    done

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Windows.edb non trovato (indicizzazione disattivata o percorso non standard)." \
                 "Windows.edb not found (indexing disabled or non-standard path).")"
        return 0
    fi
    local SZ; SZ=$(stat -c %s "${DBS[0]}" 2>/dev/null || echo 0)
    info "Windows.edb — ${BOLD}$(numfmt --to=iec "$SZ" 2>/dev/null || echo "$SZ")"

    local HAVE_ESE=false
    "$PY3" -c "import pyesedb" 2>/dev/null && HAVE_ESE=true
    $HAVE_ESE || info "$(L "libesedb (pyesedb) non disponibile: si usera' l'estrazione dalle stringhe." \
                          "libesedb (pyesedb) unavailable: string extraction will be used.")"

    info "$(L "Analisi in corso (il database puo' essere molto grande)..." "Analysing (the database can be very large)...")"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local METHOD; METHOD=$("$PY3" - "$OUT" "$($HAVE_ESE && echo 1 || echo 0)" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, os, re

out_path = sys.argv[1]
have_ese = sys.argv[2] == '1'
dbs = sys.argv[3:]

MAX_ROWS = 60000
rows = []
seen = set()

def add(path, extra, src):
    key = path.lower()[:200]
    if key in seen:
        return
    seen.add(key)
    rows.append((path[:400], extra[:200], src))

def parse_ese():
    import pyesedb
    for db in dbs:
        src = os.path.basename(db)
        try:
            f = pyesedb.file()
            f.open(db)
        except Exception:
            continue
        for i in range(f.get_number_of_tables()):
            if len(rows) >= MAX_ROWS:
                break
            try:
                t = f.get_table(i)
                tname = t.get_name()
            except Exception:
                continue
            # SystemIndex_Gthr contiene i percorsi indicizzati e i tempi di
            # ultima indicizzazione; SystemIndex_PropertyStore le proprieta'.
            if 'Gthr' not in tname and 'PropertyStore' not in tname:
                continue
            try:
                cols = [t.get_column(c).get_name() for c in range(t.get_number_of_columns())]
                ncols = len(cols)
            except Exception:
                continue
            for r in range(min(t.get_number_of_records(), MAX_ROWS)):
                try:
                    rec = t.get_record(r)
                except Exception:
                    continue
                path = ''
                extra = ''
                for n in range(ncols):
                    try:
                        v = rec.get_value_data_as_string(n)
                    except Exception:
                        continue
                    if not v:
                        continue
                    if not path and ('://' in v or re.match(r'^[A-Za-z]:\\', v)):
                        path = v
                    elif len(v) > 8 and not extra:
                        extra = v
                if path:
                    add(path, f"{tname} {extra}".strip(), src)
                if len(rows) >= MAX_ROWS:
                    break
        try:
            f.close()
        except Exception:
            pass


def parse_strings():
    """I percorsi indicizzati restano leggibili anche se l'ESE non si apre."""
    PATH = re.compile(r'(?:file:///)?[A-Za-z]:\\[^\x00<>|?*"\r\n]{4,250}')
    for db in dbs:
        src = os.path.basename(db)
        try:
            with open(db, 'rb') as fh:
                # Tetto di lettura: Windows.edb puo' superare il gigabyte.
                raw = fh.read(2 * 1024 * 1024 * 1024)
        except Exception:
            continue
        for enc in ('utf-16-le', 'latin-1'):
            try:
                text = raw.decode(enc, 'ignore')
            except Exception:
                continue
            for m in PATH.finditer(text):
                add(m.group(0).strip(), '', src)
                if len(rows) >= MAX_ROWS:
                    break

# Windows.edb e' quasi sempre in uso al momento dell'acquisizione, quindi
# spesso dirty e non apribile: ripiegare sulle stringhe evita un report vuoto.
method = 'strings'
if have_ese:
    parse_ese()
    if rows:
        method = 'ESE'
if not rows:
    parse_strings()
print(method if rows else 'nessuno')

rows.sort()
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')
PYEOF
    )

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna voce estratta dall'indice." "No entry extracted from the index.")"
        return 0
    fi
    ok "$(L "Voci indicizzate estratte:" "Indexed entries extracted:") ${BOLD}$TOTAL${RESET} ($(L "parser" "parser"): ${METHOD:-strings})"
    if [[ "$METHOD" != "ESE" ]] && $HAVE_ESE; then
        warn "$(L "Database non apribile con libesedb (in uso al momento dell'acquisizione): estrazione dalle stringhe." \
                 "Database not openable with libesedb (in use when acquired): string extraction.")"
    fi

    # Il valore forense sta nelle voci che puntano a file non piu' presenti, ma
    # il confronto non e' automatizzabile in modo affidabile: i percorsi
    # indicizzati usano la lettera di unita' vista dal sistema (C:\...), che non
    # corrisponde al punto di mount in analisi. Meglio dirlo che dedurre a caso.
    info "$(L "Il confronto con i file ancora presenti richiede la mappatura delle lettere di unita': va fatto a mano." \
             "Comparing against files still present requires drive-letter mapping: do it manually.")"

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(head -30000 "$OUT")
    local TABLE; TABLE=$(_rows_to_table "$ROWS" "$(L "Percorso indicizzato" "Indexed path")" "$(L "Dettaglio" "Detail")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Come usarlo" "How to use it")</b><br>"
    NOTE+="$(L "L'indice non viene ripulito alla cancellazione di un file: una voce che punta a un percorso non piu' esistente sul volume e' la traccia di un file eliminato. Il confronto va fatto tenendo conto della lettera di unita': il volume montato in analisi non corrisponde necessariamente a C: come lo vedeva il sistema." \
        "The index is not purged when a file is deleted: an entry pointing to a path no longer present on the volume is the trace of a removed file. Compare with the drive letter in mind: the mounted volume does not necessarily correspond to C: as the system saw it.")"
    if [[ "$METHOD" != "ESE" ]]; then
        NOTE+="<br><br>$(L "Sono stati estratti i soli percorsi: gli estratti di contenuto e le date di indicizzazione richiedono il parsing ESE, non riuscito su questo database." \
            "Only paths were extracted: content excerpts and indexing dates require ESE parsing, which did not succeed on this database.")"
    fi
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Voci" "Entries")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Dimensione DB" "DB size")" "$(numfmt --to=iec "$SZ" 2>/dev/null || echo "$SZ")" "info")"
    STATS+="$(stat_box "$(L "Parser" "Parser")" "${METHOD:-strings}" "$([[ "$METHOD" == "ESE" ]] && echo info || echo warn)")"
    finish_report "search_index" "Windows Search Index" "IDX" "Windows.edb" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Percorsi indicizzati" "Indexed paths")" "${DBS[0]}" "$TOTAL" "$TABLE" "⌕")</div>"
}
