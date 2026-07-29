
# --- macOS 11 — FSEvents (/.fseventsd) ---
#
# FSEvents e' il registro delle modifiche al filesystem tenuto da macOS: l'analogo
# dello USN Journal di NTFS, e la fonte piu' ricca per ricostruire creazione,
# rinomina e cancellazione di file — comprese quelle di file non piu' presenti.
#
# Formato: file gzip in /.fseventsd, ciascuno con una o piu' pagine DLS1/DLS2.
# Ogni record e' path NUL-terminato + event id (u64 LE) + flag (u32 LE), piu'
# un node id (u64 LE) nelle pagine DLS2/DLS3.
#
# ATTENZIONE sui tempi: i record NON contengono un timestamp. L'event id e' un
# contatore monotono. L'unico riferimento temporale e' l'intervallo coperto dal
# file che li contiene, quindi le date qui sono un LIMITE SUPERIORE approssimato
# (mtime del file di log), non l'istante dell'evento. Il report lo dichiara.
module_macos_fsevents() {
    section_header "macOS — FSEvents" "$MAGENTA"
    check_target_root || return 1

    local FSEDIR; FSEDIR=$(ci_find_dir "$WIN_ROOT" ".fseventsd")
    [[ -z "$FSEDIR" ]] && { warn "$(L "Directory /.fseventsd non trovata." "/.fseventsd directory not found.")"; return 0; }

    mapfile -t FSEFILES < <(find "$FSEDIR" -maxdepth 1 -type f ! -name 'fseventsd-uuid' -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
    if [[ ${#FSEFILES[@]} -eq 0 ]]; then
        warn "$(L "Nessun log FSEvents presente." "No FSEvents log present.")"
        return 0
    fi
    info "$(L "File FSEvents trovati:" "FSEvents logs found:") ${BOLD}${#FSEFILES[@]}"

    local KW="${IOC_LIST[*]:-}"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${FSEFILES[@]}" << 'PYEOF' 2>/dev/null
import sys, gzip, struct, os, datetime

out_path = sys.argv[1]
files = sys.argv[2:]

# Bitmask degli eventi FSEvents.
FLAGS = [
    (0x00000001, "FolderEvent"),      (0x00000002, "Mount"),
    (0x00000004, "Unmount"),          (0x00000020, "EndOfTransaction"),
    (0x00000800, "LastHardLinkRemoved"), (0x00001000, "HardLink"),
    (0x00004000, "SymbolicLink"),     (0x00008000, "FileEvent"),
    (0x00010000, "PermissionChange"), (0x00020000, "XattrModified"),
    (0x00040000, "XattrRemoved"),     (0x00100000, "DocumentRevision"),
    (0x00400000, "ItemCloned"),       (0x01000000, "Created"),
    (0x02000000, "Removed"),          (0x04000000, "InodeMetaMod"),
    (0x08000000, "Renamed"),          (0x10000000, "Modified"),
    (0x20000000, "Exchange"),         (0x40000000, "FinderInfoMod"),
    (0x80000000, "FolderCreated"),
]

# Eventi che meritano attenzione in un'indagine: cancellazioni e rinomine
# sono il segnale tipico di anti-forensics e di stage/esfiltrazione.
NOTABLE = {"Removed", "Renamed", "LastHardLinkRemoved", "PermissionChange", "XattrRemoved"}

def decode_flags(v):
    names = [n for bit, n in FLAGS if v & bit]
    return ";".join(names) if names else f"0x{v:08x}"

def parse(data):
    """Genera (path, event_id, flags) per ogni record delle pagine del file."""
    off = 0
    total = len(data)
    while off + 12 <= total:
        magic = data[off:off + 4]
        if magic not in (b"1SLD", b"2SLD", b"3SLD"):
            break
        # header: magic(4) + unknown(4) + page_size(4)
        page_size = struct.unpack_from("<I", data, off + 8)[0]
        if page_size < 12 or off + page_size > total:
            break
        # DLS2 e DLS3 aggiungono un node id da 8 byte per record
        extra = 8 if magic in (b"2SLD", b"3SLD") else 0
        p = off + 12
        end = off + page_size
        while p < end:
            nul = data.find(b"\x00", p, end)
            if nul < 0:
                break
            path = data[p:nul].decode("utf-8", "replace")
            p = nul + 1
            if p + 12 + extra > end:
                break
            event_id, flags = struct.unpack_from("<QI", data, p)
            p += 12 + extra
            if path:
                yield path, event_id, flags
        off = end

rows = []
for fp in files:
    try:
        with gzip.open(fp, "rb") as fh:
            data = fh.read()
    except Exception:
        continue
    # Unico ancoraggio temporale disponibile: il mtime del file di log.
    try:
        approx = datetime.datetime.utcfromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        approx = ""
    src = os.path.basename(fp)
    for path, event_id, flags in parse(data):
        names = decode_flags(flags)
        notable = "1" if any(n in NOTABLE for n in names.split(";")) else "0"
        rows.append((approx, path, names, str(event_id), src, notable))

with open(out_path, "w", encoding="utf-8") as fh:
    for r in rows:
        fh.write("\t".join(x.replace("\t", " ") for x in r) + "\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun record FSEvents decodificato (log vuoti o formato non riconosciuto)." \
                 "No FSEvents record decoded (empty logs or unrecognised format).")"
        return 0
    fi

    local NOTABLE; NOTABLE=$(awk -F'\t' '$6=="1"' "$OUT" | wc -l)
    ok "$(L "Record FSEvents decodificati:" "FSEvents records decoded:") ${BOLD}$TOTAL"
    info "$(L "Eventi di cancellazione/rinomina:" "Delete/rename events:") ${BOLD}$NOTABLE"

    # Anteprima a console: solo gli eventi rilevanti, altrimenti e' illeggibile.
    awk -F'\t' '$6=="1"{print $1"  "$3"  "$2}' "$OUT" | head -40 | while IFS= read -r LN; do
        echo -e "      ${DIM}${LN}${RESET}"
    done

    ask_yn "Generare report HTML?" || return 0

    # Il report completo puo' contenere milioni di record: si limita a un
    # volume gestibile dal browser, privilegiando gli eventi rilevanti.
    local LIMIT=20000
    local ROWS; ROWS=$( { awk -F'\t' '$6=="1"{print $1"\t"$3"\t"$2"\t"$4"\t"$5}' "$OUT";
                          awk -F'\t' '$6=="0"{print $1"\t"$3"\t"$2"\t"$4"\t"$5}' "$OUT"; } | head -$LIMIT )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data (approx.)" "Date (approx.)")" "$(L "Evento" "Event")" "Path" "Event ID" "$(L "File di log" "Log file")")

    local NOTE
    NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Nota sui tempi" "Note on timestamps")</b><br>"
    NOTE+="$(L "I record FSEvents non contengono un timestamp: l'event ID e' un contatore monotono. La colonna data riporta il mtime del file di log che contiene il record, cioe' un LIMITE SUPERIORE approssimato dell'istante dell'evento — non la sua ora esatta. Per una datazione precisa va correlato con altri artefatti." \
        "FSEvents records carry no timestamp: the event ID is a monotonic counter. The date column shows the mtime of the log file containing the record, i.e. an approximate UPPER BOUND of when the event happened — not its exact time. Precise dating requires correlation with other artefacts.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Record" "Records")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Cancellazioni/rinomine" "Deletes/renames")" "$NOTABLE" "warn")"
    STATS+="$(stat_box "$(L "File di log" "Log files")" "${#FSEFILES[@]}" "info")"
    [[ "$TOTAL" -gt "$LIMIT" ]] && STATS+="$(stat_box "$(L "Mostrati" "Shown")" "$LIMIT" "info")"

    finish_report "macos_fsevents" "macOS FSEvents" "FSE" "/.fseventsd" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "FSEvents" "$FSEDIR" "$TOTAL" "$TABLE" "⟳")</div>"
}
