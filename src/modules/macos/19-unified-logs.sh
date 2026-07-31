
# ================================================================
#  macOS 19 — Unified Logs (.tracev3)
#
#  E' il registro centrale di macOS: esecuzioni, autenticazioni, rete,
#  installazioni, XProtect. Nelle versioni precedenti FIUTO lo dichiarava
#  fuori scope, ed era il gap piu' evidente dell'analisi macOS.
#
#  ONESTA' SUL LIVELLO DI SUPPORTO. Un parser completo di .tracev3 deve
#  ricostruire il catalogo, risolvere i riferimenti alle stringhe nel
#  dyld_shared_cache e nei file .uuidtext, e reidratare i formati: e' un
#  progetto a se' (cfr. mandiant/macos-UnifiedLogs). Questo modulo NON lo fa.
#
#  Fa una cosa piu' modesta e verificabile: decomprime i chunk LZ4 (bv41) di
#  cui il file e' composto ed estrae le stringhe leggibili che ne emergono —
#  percorsi di processo, nomi di bundle, frammenti di messaggi gia' in chiaro.
#  Senza la decompressione un `strings` sul file grezzo non restituirebbe
#  quasi nulla, perche' il contenuto e' compresso.
#
#  Risultato: indizi datati approssimativamente e cercabili, non una timeline
#  di log ricostruita. Per quella serve `log show --archive` su un Mac.
# ================================================================
module_macos_unified_logs() {
    section_header "macOS - Unified Logs (.tracev3)" "$MAGENTA"
    check_target_root || return 1

    local -a FILES=()
    local D F
    for D in "private/var/db/diagnostics" "var/db/diagnostics"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$R" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && FILES+=("$F")
        done < <(find "$R" -type f -name '*.tracev3' 2>/dev/null | sort)
    done

    if [[ ${#FILES[@]} -eq 0 ]]; then
        warn "$(L "Nessun file .tracev3 trovato." "No .tracev3 file found.")"
        return 0
    fi
    local TOTB=0
    for F in "${FILES[@]}"; do
        TOTB=$(( TOTB + $(stat -c %s "$F" 2>/dev/null || echo 0) ))
    done
    info "$(L "File .tracev3:" ".tracev3 files:") ${BOLD}${#FILES[@]}${RESET} - $(numfmt --to=iec "$TOTB" 2>/dev/null || echo "$TOTB B")"
    info "$(L "Decompressione dei chunk in corso..." "Decompressing chunks...")"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local STATSF; STATSF=$(mktemp); register_tmp "$STATSF"

    "$PY3" - "$OUT" "$STATSF" "$IOCTMP" "${FILES[@]}" << 'PYEOF' 2>/dev/null
import sys, os, re, struct, datetime

out_path, stats_path, ioc_path = sys.argv[1], sys.argv[2], sys.argv[3]
files = sys.argv[4:]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []


def lz4_block(src, expected):
    """Decompressore LZ4 block format in puro Python (nessuna dipendenza).

    Apple incapsula i chunk in 'bv41' + dimensioni + blocco LZ4. Senza questo
    passaggio il contenuto del file resta compresso e illeggibile.
    """
    out = bytearray()
    i, n = 0, len(src)
    while i < n:
        token = src[i]; i += 1
        lit = token >> 4
        if lit == 15:
            while i < n:
                b = src[i]; i += 1
                lit += b
                if b != 255:
                    break
        out += src[i:i + lit]
        i += lit
        if i + 2 > n:
            break
        off = src[i] | (src[i + 1] << 8); i += 2
        if off == 0:
            break
        mlen = token & 0x0F
        if mlen == 15:
            while i < n:
                b = src[i]; i += 1
                mlen += b
                if b != 255:
                    break
        mlen += 4
        start = len(out) - off
        if start < 0:
            break
        for k in range(mlen):
            out.append(out[start + k])
        if expected and len(out) >= expected:
            break
    return bytes(out)


def decompress(data):
    """Concatena il contenuto di tutti i chunk bv41/bv4- del file."""
    out = bytearray()
    pos = 0
    n = len(data)
    while True:
        k = data.find(b'bv4', pos)
        if k < 0 or k + 12 > n:
            break
        tag = data[k:k + 4]
        if tag == b'bv41':
            try:
                usize, csize = struct.unpack_from('<II', data, k + 4)
            except Exception:
                pos = k + 4
                continue
            blob = data[k + 12:k + 12 + csize]
            if 0 < csize <= n and usize < 64 * 1024 * 1024:
                out += lz4_block(blob, usize)
            pos = k + 12 + max(csize, 1)
        elif tag == b'bv4-':
            try:
                usize = struct.unpack_from('<I', data, k + 4)[0]
            except Exception:
                pos = k + 4
                continue
            out += data[k + 8:k + 8 + usize]
            pos = k + 8 + max(usize, 1)
        else:
            pos = k + 3
        if len(out) > 256 * 1024 * 1024:
            break
    return bytes(out)


# Stringhe che vale la pena tenere: percorsi, bundle id, comandi, URL.
KEEP = re.compile(
    rb'(?:/(?:usr|bin|sbin|Applications|Library|System|Users|private|tmp|var|opt)/[!-~]{3,180}'
    rb'|[a-z][a-z0-9]+(?:\.[a-z0-9-]+){2,6}'          # bundle id / reverse-dns
    rb'|https?://[!-~]{4,200}'
    rb'|[A-Za-z0-9_.-]+\.(?:app|kext|plist|dylib|sh|py|pkg|dmg)\b)')

# Rumore ricorrente dei log di sistema: escluderlo rende il resto leggibile.
NOISE = re.compile(rb'(?i)(com\.apple\.(?:coreanimation|uikit|coretext|metal|springboard|coreui))')

rows = []
seen = set()
stats = []

for path in files:
    src = os.path.basename(path)
    try:
        with open(path, 'rb') as fh:
            raw = fh.read(512 * 1024 * 1024)
    except Exception:
        continue
    # Unico ancoraggio temporale disponibile senza ricostruire il catalogo.
    try:
        approx = datetime.datetime.utcfromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        approx = ''
    plain = decompress(raw)
    stats.append((src, str(len(raw)), str(len(plain))))
    if not plain:
        continue
    for m in KEEP.finditer(plain):
        s = m.group(0)
        if NOISE.search(s):
            continue
        try:
            t = s.decode('utf-8', 'strict')
        except Exception:
            continue
        if len(t) < 6:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        flag = 'IOC' if (iocs and any(i in key for i in iocs)) else ''
        rows.append((approx, t[:300], flag, src))

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ') for x in r) + '\n')
with open(stats_path, 'w', encoding='utf-8') as fh:
    for s in stats:
        fh.write('\t'.join(s) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    local DECOMP=0
    [[ -s "$STATSF" ]] && DECOMP=$(awk -F'\t' '{s+=$3} END{print s+0}' "$STATSF")

    separator
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna stringa estratta: i chunk non sono stati decompressi (formato non riconosciuto)." \
                 "No string extracted: chunks were not decompressed (unrecognised format).")"
        info "$(L "Per l'analisi completa serve un Mac: log show --archive <cartella>.logarchive" \
                 "Full analysis requires a Mac: log show --archive <folder>.logarchive")"
        return 0
    fi
    ok "$(L "Byte decompressi:" "Bytes decompressed:") ${BOLD}$(numfmt --to=iec "$DECOMP" 2>/dev/null || echo "$DECOMP")"
    ok "$(L "Stringhe rilevanti estratte:" "Relevant strings extracted:") ${BOLD}$TOTAL"
    local NIOC; NIOC=$(awk -F'\t' '$3=="IOC"' "$OUT" | wc -l)
    [[ "$NIOC" -gt 0 ]] && warn "$(L "Con match IoC:" "With IoC match:") ${BOLD}$NIOC"

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$3=="IOC"' "$OUT"; awk -F'\t' '$3!="IOC"' "$OUT"; } | head -30000 )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data file (approx.)" "File date (approx.)")" "$(L "Stringa" "String")" "IoC" "$(L "Origine" "Source")")
    local STABLE; STABLE=$(_rows_to_table "$(cat "$STATSF")" "$(L "File" "File")" "$(L "Byte su disco" "Bytes on disk")" "$(L "Byte decompressi" "Bytes decompressed")")

    local NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,166,87,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Livello di supporto: parziale, e dichiarato" "Support level: partial, and stated")</b><br>"
    NOTE+="$(L "Questo modulo NON ricostruisce i messaggi di log. Un parser completo di .tracev3 deve interpretare il catalogo e risolvere i riferimenti alle stringhe nei file .uuidtext e nel dyld_shared_cache: e' un progetto a se'. Qui i chunk LZ4 vengono decompressi e se ne estraggono le stringhe gia' leggibili - percorsi, bundle id, URL, nomi di file." \
        "This module does NOT reconstruct log messages. A complete .tracev3 parser must interpret the catalogue and resolve string references in .uuidtext files and the dyld_shared_cache: that is a project of its own. Here the LZ4 chunks are decompressed and the already-readable strings are extracted - paths, bundle ids, URLs, file names.")<br><br>"
    NOTE+="$(L "Le date sono quelle di modifica del file che contiene la stringa: un limite superiore approssimato, non l'istante dell'evento." \
        "Dates are the modification times of the file containing the string: an approximate upper bound, not the moment of the event.")<br><br>"
    NOTE+="<b>$(L "Per l'analisi completa" "For full analysis")</b><br>"
    NOTE+="<code>log show --archive /percorso/diagnostics.logarchive --info --debug</code> ($(L "richiede un Mac" "requires a Mac")) - "
    NOTE+="$(L "oppure" "or") <code>mandiant/macos-UnifiedLogs</code>."
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Stringhe" "Strings")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Decompressi" "Decompressed")" "$(numfmt --to=iec "$DECOMP" 2>/dev/null || echo "$DECOMP")" "info")"
    STATS+="$(stat_box "File" "${#FILES[@]}" "info")"
    STATS+="$(stat_box "IoC" "$NIOC" "$([[ "$NIOC" -gt 0 ]] && echo warn || echo info)")"
    finish_report "macos_unified_logs" "macOS Unified Logs" "ULG" "private/var/db/diagnostics/*.tracev3" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Resa della decompressione" "Decompression yield")" "diagnostics" "${#FILES[@]}" "$STABLE" "∑")</div><div class='cards'>$(generic_card_html "$(L "Stringhe estratte" "Extracted strings")" "$(L "estrazione parziale" "partial extraction")" "$TOTAL" "$TABLE" "⌕")</div>"
}
