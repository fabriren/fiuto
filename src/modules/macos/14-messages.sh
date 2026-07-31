
# ================================================================
#  macOS 14 — Messages (chat.db)
#
#  iMessage e SMS inoltrati dall'iPhone finiscono in chat.db sul Mac. In
#  un'indagine contano per lo smishing (il link arrivato via SMS che ha
#  iniziato la compromissione), per i codici 2FA intercettati e per la
#  corrispondenza scambiata fuori dai canali aziendali.
#
#  Il database conserva anche i messaggi con `is_deleted`, e gli allegati
#  restano su disco in ~/Library/Messages/Attachments anche dopo la
#  cancellazione della conversazione.
# ================================================================
module_macos_messages() {
    section_header "macOS - Messages" "$CYAN"
    check_target_root || return 1

    local BODY="" TOTAL=0 NATT=0 NSUSP=0
    local HOME_DIR
    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local DB; DB=$(ci_find_file "$HOME_DIR" "Library/Messages/chat.db")
        [[ -s "$DB" ]] || continue

        # La colonna date e' in nanosecondi dal 2001 su macOS recenti e in
        # secondi su quelli vecchi: si normalizza in SQL guardando l'ordine di
        # grandezza, altrimenti le date risultano nel futuro remoto.
        local ROWS
        ROWS=$(query_sqlite "$DB" "
            SELECT datetime(CASE WHEN m.date > 100000000000
                                 THEN m.date/1000000000 + 978307200
                                 ELSE m.date + 978307200 END, 'unixepoch'),
                   COALESCE(h.id,'?'),
                   CASE m.is_from_me WHEN 1 THEN 'inviato' ELSE 'ricevuto' END,
                   COALESCE(m.text,''),
                   COALESCE(m.service,'')
            FROM message m LEFT JOIN handle h ON m.handle_id = h.ROWID
            WHERE m.text IS NOT NULL AND m.text <> ''
            ORDER BY m.date DESC LIMIT 50000")
        [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue

        local N; N=$(printf '%s\n' "$ROWS" | grep -c . || true)
        TOTAL=$((TOTAL + N))
        ok "$U - ${BOLD}${N}${RESET} $(L "messaggi" "messages")"

        # Messaggi con link o riferimenti a credenziali: sono quelli che
        # spiegano un accesso iniziale o una compromissione di account.
        local SUSP
        SUSP=$(printf '%s\n' "$ROWS" | grep -iE 'https?://|bit\.ly|tinyurl|codice|code|verifica|verif|otp|password|accedi|login|urgente|urgent|bloccat|blocked' || true)
        local NS; NS=$(printf '%s\n' "$SUSP" | grep -c . || true)
        NSUSP=$((NSUSP + NS))

        local TABLE; TABLE=$(_rows_to_table "$ROWS" \
            "$(L "Data" "Date")" "$(L "Interlocutore" "Handle")" "$(L "Direzione" "Direction")" \
            "$(L "Testo" "Text")" "$(L "Servizio" "Service")")
        BODY+=$(generic_card_html "$U" "$DB" "$N" "$TABLE" "✉")

        if [[ "$NS" -gt 0 ]]; then
            local STABLE; STABLE=$(_rows_to_table "$SUSP" \
                "$(L "Data" "Date")" "$(L "Interlocutore" "Handle")" "$(L "Direzione" "Direction")" \
                "$(L "Testo" "Text")" "$(L "Servizio" "Service")")
            BODY="<div class='cards'>$(generic_card_html "$U - $(L "messaggi con link o credenziali" "messages with links or credentials")" "$DB" "$NS" "$STABLE" "⚑")</div>${BODY}"
        fi

        # Allegati: restano su disco anche dopo la cancellazione della chat.
        local ATTDIR; ATTDIR=$(ci_find_dir "$HOME_DIR" "Library/Messages/Attachments")
        if [[ -n "$ATTDIR" ]]; then
            local NA; NA=$(find "$ATTDIR" -type f 2>/dev/null | wc -l)
            NATT=$((NATT + NA))
            [[ "$NA" -gt 0 ]] && info "  $(L "allegati su disco:" "attachments on disk:") ${BOLD}${NA}"
        fi
    done < <(get_macos_user_homes)

    separator
    info "$(L "Messaggi totali:" "Total messages:") ${BOLD}$TOTAL"
    [[ "$NSUSP" -gt 0 ]] && warn "$(L "Con link o riferimenti a credenziali:" "With links or credential references:") ${BOLD}$NSUSP"
    info "$(L "Allegati su disco:" "Attachments on disk:") ${BOLD}$NATT"
    [[ "$TOTAL" -eq 0 ]] && { warn "$(L "Nessun database Messages leggibile." "No readable Messages database.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="$(L "Gli allegati restano in ~/Library/Messages/Attachments anche dopo la cancellazione della conversazione: vanno esaminati a parte, non compaiono in questa tabella. La colonna data e' normalizzata dal formato Apple (riferimento 2001), che su macOS recenti e' in nanosecondi." \
        "Attachments remain in ~/Library/Messages/Attachments even after the conversation is deleted: examine them separately, they are not in this table. The date column is normalised from the Apple epoch (2001 reference), which on recent macOS is in nanoseconds.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Messaggi" "Messages")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Segnalati" "Flagged")" "$NSUSP" "$([[ "$NSUSP" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Allegati" "Attachments")" "$NATT" "info")"
    finish_report "macos_messages" "macOS Messages" "MSG" "Library/Messages/chat.db" "$STATS" "${NOTE}<div class='cards'>$BODY</div>"
}
