
# ================================================================
#  MODULO 44 — Outlook PST / OST
#
#  FIUTO non aveva alcuna copertura della posta locale, che in un caso di
#  phishing, BEC o esfiltrazione e' spesso il punto di partenza: il messaggio
#  di ingresso, l'allegato eseguito, le regole di inoltro create
#  dall'attaccante e la corrispondenza esfiltrata stanno tutti qui.
#
#  Un OST conserva anche messaggi cancellati lato server ma non ancora
#  sincronizzati: a volte e' l'unica copia rimasta.
#
#  Richiede pypff (libpff). Senza, il modulo si limita all'inventario dei file
#  con dimensioni e date, che e' comunque utile per l'acquisizione.
# ================================================================
module_pst_ost() {
    section_header "Outlook PST / OST" "$YELLOW"
    check_win_root || return 1

    local -a STORES=()
    local F HOME_DIR
    while IFS= read -r HOME_DIR; do
        while IFS= read -r F; do
            [[ -n "$F" ]] && STORES+=("$F")
        done < <(find "$HOME_DIR" -maxdepth 6 -type f \( -iname '*.pst' -o -iname '*.ost' \) 2>/dev/null)
    done < <(get_user_homes)

    if [[ ${#STORES[@]} -eq 0 ]]; then
        warn "$(L "Nessun archivio PST/OST trovato." "No PST/OST store found.")"
        return 0
    fi
    ok "$(L "Archivi trovati:" "Stores found:") ${BOLD}${#STORES[@]}"

    # Inventario: utile anche senza pypff (dimensioni, date, hash per la catena
    # di custodia, stima dei tempi di elaborazione).
    local INV="" TOTBYTES=0
    for F in "${STORES[@]}"; do
        local SZ MT
        SZ=$(stat -c %s "$F" 2>/dev/null || echo 0)
        MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1 || echo "?")
        TOTBYTES=$((TOTBYTES + SZ))
        info "$(basename "$F") — ${BOLD}$(numfmt --to=iec "$SZ" 2>/dev/null || echo "$SZ")${RESET} ${DIM}${MT}${RESET}"
        INV+="${MT}	$(basename "$F")	${SZ}	${F}
"
    done

    local HAVE_PYPFF=false
    "$PY3" -c "import pypff" 2>/dev/null && HAVE_PYPFF=true

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local NMSG=0 NSUSP=0
    if $HAVE_PYPFF; then
        info "$(L "Estrazione messaggi con pypff (puo' richiedere tempo)..." "Extracting messages with pypff (may take a while)...")"
        local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
        printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"
        "$PY3" - "$OUT" "$IOCTMP" "${STORES[@]}" << 'PYEOF' 2>/dev/null
import sys, os, datetime

out_path, ioc_path = sys.argv[1], sys.argv[2]
stores = sys.argv[3:]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

import pypff

# Indicatori tipici di phishing/BEC negli header e negli allegati.
SUSP_EXT = ('.exe', '.scr', '.js', '.vbs', '.jse', '.wsf', '.hta', '.lnk',
            '.iso', '.img', '.7z', '.ace', '.docm', '.xlsm', '.pptm', '.jar', '.ps1')

rows = []

def walk(folder, path, src):
    try:
        subs = folder.number_of_sub_folders
    except Exception:
        subs = 0
    for i in range(subs):
        try:
            sub = folder.get_sub_folder(i)
        except Exception:
            continue
        name = ''
        try:
            name = sub.name or ''
        except Exception:
            pass
        walk(sub, f"{path}/{name}", src)
    try:
        n = folder.number_of_sub_messages
    except Exception:
        return
    for i in range(n):
        try:
            msg = folder.get_sub_message(i)
        except Exception:
            continue
        def g(attr):
            try:
                v = getattr(msg, attr)
                return v if isinstance(v, str) else (v or '')
            except Exception:
                return ''
        subject = g('subject')
        sender = g('sender_name')
        try:
            when = msg.get_delivery_time()
            when = when.strftime('%Y-%m-%d %H:%M:%S') if when else ''
        except Exception:
            when = ''
        # Allegati
        atts = []
        try:
            for a in range(msg.number_of_attachments):
                try:
                    att = msg.get_attachment(a)
                    an = ''
                    try:
                        an = att.get_name() or ''
                    except Exception:
                        pass
                    atts.append(an)
                except Exception:
                    continue
        except Exception:
            pass
        att_s = '; '.join(x for x in atts if x)
        flags = []
        low_att = att_s.lower()
        if any(low_att.endswith(e) or (e + ';') in low_att or (e + ' ') in low_att for e in SUSP_EXT):
            flags.append('ALLEGATO_RISCHIOSO')
        blob = f"{subject} {sender} {att_s}".lower()
        if iocs and any(i in blob for i in iocs):
            flags.append('IOC')
        rows.append((when, path, sender[:120], subject[:200], att_s[:200], ';'.join(flags), src))

for sp in stores:
    src = os.path.basename(sp)
    try:
        pff = pypff.file()
        pff.open(sp)
        root = pff.get_root_folder()
    except Exception:
        continue
    try:
        walk(root, '', src)
    except Exception:
        pass
    try:
        pff.close()
    except Exception:
        pass

rows.sort(key=lambda r: r[0])
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
PYEOF
        [[ -s "$OUT" ]] && NMSG=$(wc -l < "$OUT")
        [[ "$NMSG" -gt 0 ]] && NSUSP=$(awk -F'\t' '$6!=""' "$OUT" | wc -l)
        ok "$(L "Messaggi estratti:" "Messages extracted:") ${BOLD}$NMSG"
        [[ "$NSUSP" -gt 0 ]] && warn "$(L "Messaggi con allegato rischioso o match IoC:" "Messages with risky attachment or IoC match:") ${BOLD}$NSUSP"
    else
        warn "$(L "pypff non disponibile: estratto solo l'inventario degli archivi." \
                 "pypff unavailable: only the store inventory was extracted.")"
        info "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install libpff-python"
    fi

    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    local INVTABLE; INVTABLE=$(_rows_to_table "$INV" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Archivio" "Store")" "$(L "Byte" "Bytes")" "$(L "Percorso" "Path")")
    BODY+="<div class='cards'>$(generic_card_html "$(L "Archivi di posta" "Mail stores")" "$(L "inventario" "inventory")" "${#STORES[@]}" "$INVTABLE" "✉")</div>"

    if [[ "$NMSG" -gt 0 ]]; then
        local ROWS; ROWS=$( { awk -F'\t' '$6!=""{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6}' "$OUT";
                              awk -F'\t' '$6==""{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6}' "$OUT"; } | head -20000 )
        local TABLE; TABLE=$(_rows_to_table "$ROWS" \
            "$(L "Data" "Date")" "$(L "Cartella" "Folder")" "$(L "Mittente" "Sender")" \
            "$(L "Oggetto" "Subject")" "$(L "Allegati" "Attachments")" "$(L "Segnalazioni" "Flags")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Messaggi" "Messages")" "$(L "sospetti in testa" "suspicious first")" "$NMSG" "$TABLE" "✉")</div>"
    else
        BODY+="<div class='card'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
        BODY+="$(L "Contenuto dei messaggi non estratto: manca pypff (libpff). Gli archivi sono comunque inventariati sopra e possono essere analizzati a parte con readpst o pffexport." \
            "Message content not extracted: pypff (libpff) is missing. The stores are inventoried above and can be analysed separately with readpst or pffexport.")<br><br>"
        BODY+="<code>${PY3} -m pip install libpff-python</code><br>"
        BODY+="<code>pffexport -q archivio.pst</code>"
        BODY+="</div></div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Archivi" "Stores")" "${#STORES[@]}")"
    STATS+="$(stat_box "$(L "Dimensione" "Size")" "$(numfmt --to=iec "$TOTBYTES" 2>/dev/null || echo "$TOTBYTES")" "info")"
    if [[ "$NMSG" -gt 0 ]]; then
        STATS+="$(stat_box "$(L "Messaggi" "Messages")" "$NMSG" "info")"
        STATS+="$(stat_box "$(L "Sospetti" "Suspicious")" "$NSUSP" "$([[ "$NSUSP" -gt 0 ]] && echo warn || echo info)")"
    fi
    finish_report "pst_ost" "Outlook PST / OST" "PST" "*.pst · *.ost" "$STATS" "$BODY"
}
