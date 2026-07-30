
# ================================================================
#  MODULO 47 — Thumbcache / IconCache
#
#  Windows conserva le miniature dei file visualizzati in Esplora risorse in
#  database centralizzati. La miniatura SOPRAVVIVE alla cancellazione del file
#  originale: e' spesso l'unica prova visiva residua di un documento o di
#  un'immagine poi eliminata, e non e' toccata dalla pulizia del cestino.
#
#  I formati dei record cambiano fra Vista, 7, 8 e 10/11. Invece di parsare la
#  struttura — fragile fra versioni — qui si esegue il CARVING delle immagini
#  incorporate cercandone le firme. E' robusto su tutte le versioni; in cambio
#  non si ottiene la corrispondenza miniatura-nome file, che nel formato non
#  e' comunque presente in chiaro (la chiave e' un hash del percorso).
#
#  Le immagini estratte vengono scritte accanto al report e mostrate in
#  galleria, cosi' l'analista le sfoglia invece di aprirle una per una.
# ================================================================
module_thumbcache() {
    section_header "Thumbcache / IconCache" "$GREEN"
    check_win_root || return 1

    local -a DBS=()
    local HOME_DIR F
    while IFS= read -r HOME_DIR; do
        local EXPL; EXPL=$(ci_find_dir "$HOME_DIR" "AppData/Local/Microsoft/Windows/Explorer")
        [[ -z "$EXPL" ]] && continue
        while IFS= read -r F; do
            [[ -n "$F" ]] && DBS+=("$F")
        done < <(find "$EXPL" -maxdepth 1 -type f \( -iname 'thumbcache_*.db' -o -iname 'iconcache_*.db' \) -size +1k 2>/dev/null)
    done < <(get_user_homes)

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessun database thumbcache/iconcache trovato." "No thumbcache/iconcache database found.")"
        return 0
    fi
    info "$(L "Database trovati:" "Databases found:") ${BOLD}${#DBS[@]}"

    # Le immagini estratte vanno accanto al report: si prepara la cartella qui,
    # perche' prepare_report_dir ne crea una nuova a ogni chiamata.
    [[ -n "$REPORT_BASE_DIR" && ! -d "$REPORT_BASE_DIR" ]] && mkdir -p "$REPORT_BASE_DIR"
    local RDIR="${REPORT_BASE_DIR}/thumbcache_$(date +%Y%m%d_%H%M%S)"
    mkdir -p "$RDIR/images" 2>/dev/null || {
        err "$(L "Impossibile creare la cartella del report." "Cannot create the report directory.")"
        return 1
    }

    info "$(L "Carving delle miniature in corso..." "Carving thumbnails...")"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "$RDIR/images" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, os, struct

out_path, img_dir = sys.argv[1], sys.argv[2]
dbs = sys.argv[3:]

MAX_IMAGES = 4000          # tetto: una thumbcache puo' contenerne decine di migliaia
MIN_SIZE   = 512           # sotto questa soglia sono icone di sistema, rumore

def carve(data):
    """Genera (offset, estensione, byte) per ogni immagine incorporata."""
    n = len(data)
    i = 0
    while i < n - 8:
        # PNG
        if data[i:i+8] == b'\x89PNG\r\n\x1a\n':
            end = data.find(b'IEND', i)
            if end > 0:
                end += 8
                yield i, 'png', data[i:end]
                i = end
                continue
        # JPEG
        if data[i:i+3] == b'\xff\xd8\xff':
            end = data.find(b'\xff\xd9', i + 3)
            if end > 0:
                end += 2
                yield i, 'jpg', data[i:end]
                i = end
                continue
        # BMP: 'BM' + dimensione dichiarata nell'header
        if data[i:i+2] == b'BM' and i + 6 <= n:
            try:
                size = struct.unpack_from('<I', data, i + 2)[0]
            except Exception:
                size = 0
            if 100 < size < 8 * 1024 * 1024 and i + size <= n:
                yield i, 'bmp', data[i:i+size]
                i += size
                continue
        i += 1

rows = []
count = 0
for db in dbs:
    base = os.path.basename(db)
    try:
        with open(db, 'rb') as fh:
            data = fh.read()
    except Exception:
        continue

    # L'intestazione dichiara versione e tipo di cache (la dimensione delle
    # miniature): utile per sapere quale cache si sta guardando.
    version = cache_type = ''
    if data[:4] == b'CMMM':
        try:
            version, cache_type = struct.unpack_from('<II', data, 4)
        except Exception:
            pass

    for off, ext, blob in carve(data):
        if len(blob) < MIN_SIZE:
            continue
        if count >= MAX_IMAGES:
            break
        name = f"{base.replace('.db','')}_{off:08x}.{ext}"
        try:
            with open(os.path.join(img_dir, name), 'wb') as out:
                out.write(blob)
        except Exception:
            continue
        count += 1
        rows.append((base, str(version), str(cache_type), ext, str(len(blob)), name))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(r) + '\n')
PYEOF

    local N=0
    [[ -s "$OUT" ]] && N=$(wc -l < "$OUT")
    if [[ "$N" -eq 0 ]]; then
        warn "$(L "Nessuna miniatura estratta." "No thumbnail extracted.")"
        rmdir "$RDIR/images" "$RDIR" 2>/dev/null
        return 0
    fi
    ok "$(L "Miniature estratte:" "Thumbnails extracted:") ${BOLD}$N"
    info "$(L "Salvate in:" "Saved to:") ${DIM}${RDIR}/images${RESET}"

    ask_yn "Generare report HTML?" || return 0

    # Galleria: le miniature sono l'informazione, non la tabella.
    local GAL="" ROW
    local SHOWN=0
    while IFS=$'\t' read -r SRC VER CT EXT SZ NAME; do
        [[ -z "$NAME" ]] && continue
        SHOWN=$((SHOWN + 1))
        [[ $SHOWN -gt 1500 ]] && break
        GAL+="<figure style='margin:0;text-align:center'><img src='images/$(html_attr "$NAME")' loading='lazy' style='max-width:120px;max-height:120px;border:1px solid var(--border);border-radius:4px;background:#0d1117'><figcaption style='font-family:var(--mono);font-size:.55rem;color:var(--text-dim);word-break:break-all'>$(html_esc "${SRC}")</figcaption></figure>"
    done < "$OUT"

    local SUM; SUM=$(awk -F'\t' '{c[$1"\t"$2"\t"$3]++} END{for(k in c) print k"\t"c[k]}' "$OUT" | sort)
    local SUMTABLE; SUMTABLE=$(_rows_to_table "$SUM" "Database" "$(L "Versione" "Version")" "$(L "Tipo cache" "Cache type")" "$(L "Miniature" "Thumbnails")")

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Cosa sono e cosa non sono" "What these are and are not")</b><br>"
    BODY+="$(L "Ogni miniatura e' la prova che un file e' stato visualizzato in Esplora risorse: sopravvive alla cancellazione dell'originale e allo svuotamento del cestino. NON e' pero' possibile risalire al nome del file: nel formato la chiave e' un hash del percorso, non il percorso stesso. La miniatura prova l'esistenza e il contenuto, non l'ubicazione." \
        "Each thumbnail proves a file was viewed in Explorer: it survives deletion of the original and emptying the Recycle Bin. It is NOT possible to recover the file name: the format keys entries by a hash of the path, not the path itself. A thumbnail proves existence and content, not location.")<br><br>"
    BODY+="$(L "Estrazione per carving delle firme immagine, non per parsing dei record: robusta su tutte le versioni di Windows." \
        "Extraction is by image-signature carving, not record parsing: robust across all Windows versions.")"
    BODY+="</div></div>"
    BODY+="<div class='cards'>$(generic_card_html "$(L "Database analizzati" "Databases analysed")" "AppData/Local/Microsoft/Windows/Explorer" "${#DBS[@]}" "$SUMTABLE" "▦")</div>"
    BODY+="<div class='stitle'>$(L "Galleria miniature" "Thumbnail gallery")</div>"
    BODY+="<div class='card'><div style='padding:1rem;display:flex;flex-wrap:wrap;gap:.6rem'>${GAL}</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Miniature" "Thumbnails")" "$N")"
    STATS+="$(stat_box "Database" "${#DBS[@]}" "info")"
    [[ "$N" -gt 1500 ]] && STATS+="$(stat_box "$(L "In galleria" "In gallery")" "1500" "info")"

    # Il report va scritto nella cartella gia' creata, accanto alle immagini.
    local REPORT_HTML="${RDIR}/report.html"
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    {
        html_header "Thumbcache"
        html_page_header "THU" "Thumbcache / IconCache" "thumbcache_*.db" "$SCAN" "$WIN_ROOT"
        printf "<div class='statsbar'>%s</div>\n" "$STATS"
        echo "<main>"
        pre_style_block
        printf '%s\n' "$BODY"
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
