
# ================================================================
#  MODULO 41 — PowerShell Transcript
#
#  Il modulo 1 legge PSReadLine, che conserva solo le righe DIGITATE nella
#  console. I transcript (Start-Transcript, o la policy "Turn on PowerShell
#  Transcription") registrano invece la sessione completa: comandi, output,
#  utente, host e processo.
#
#  Differenza pratica: PSReadLine non vede nulla di quanto eseguito da script,
#  da -EncodedCommand o da una sessione remota; i transcript sì, e includono
#  anche l'output — spesso l'unica traccia rimasta di cosa un comando abbia
#  effettivamente restituito.
# ================================================================
module_ps_transcripts() {
    section_header "PowerShell Transcript" "$MAGENTA"
    check_win_root || return 1

    local -a FILES=()
    local F

    # I transcript finiscono in Documenti per default, ma la policy consente
    # una OutputDirectory arbitraria: si cercano anche nelle posizioni comuni.
    local -a ROOTS=()
    while IFS= read -r F; do
        [[ -n "$F" ]] && ROOTS+=("$F")
    done < <(get_user_homes)
    local D
    for D in "Windows/Temp" "Temp" "ProgramData/Microsoft/Windows/PowerShell" "Transcripts" "PSTranscripts"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ROOTS+=("$R")
    done

    local R
    for R in "${ROOTS[@]}"; do
        while IFS= read -r F; do
            [[ -n "$F" ]] && FILES+=("$F")
        done < <(find "$R" -maxdepth 5 -type f -iname 'PowerShell_transcript*.txt' 2>/dev/null)
    done

    if [[ ${#FILES[@]} -eq 0 ]]; then
        warn "$(L "Nessun transcript PowerShell trovato." "No PowerShell transcript found.")"
        info "$(L "I transcript esistono solo se attivati (Start-Transcript o policy di trascrizione)." \
                 "Transcripts exist only if enabled (Start-Transcript or transcription policy).")"
        return 0
    fi
    info "$(L "Transcript trovati:" "Transcripts found:") ${BOLD}${#FILES[@]}"

    # Parole che segnalano attivita' offensiva dentro una sessione PowerShell.
    local KW="downloadstring|downloadfile|iex|invoke-expression|invoke-webrequest|frombase64string|encodedcommand|-enc |bypass|hidden|net user|net localgroup|mimikatz|invoke-mimikatz|add-mppreference|set-mppreference|disable-windowsoptionalfeature|schtasks|reg add|vssadmin|bcdedit|wevtutil|certutil|bitsadmin|password|secret|token"
    local -a IOCS=("${IOC_LIST[@]:-}")
    local I
    for I in "${IOCS[@]}"; do
        [[ -n "$I" ]] && KW+="|${I,,}"
    done

    local BODY="" TOTAL=0 SUSP=0
    local META; META=$(mktemp); register_tmp "$META"
    for F in "${FILES[@]}"; do
        [[ -s "$F" ]] || continue
        TOTAL=$((TOTAL + 1))
        local N; N=$(grep -icE "$KW" "$F" 2>/dev/null || echo 0)
        [[ "$N" -gt 0 ]] && SUSP=$((SUSP + 1))

        # L'intestazione del transcript contiene utente, host e PID.
        local U H P ST
        U=$(grep -m1 -iE '^Username:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        H=$(grep -m1 -iE '^Host Application:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        P=$(grep -m1 -iE '^Process ID:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        ST=$(grep -m1 -iE '^Start time:' "$F" 2>/dev/null | cut -d: -f2- | xargs || true)
        printf '%s\t%s\t%s\t%s\t%s\t%s\n' \
            "${ST:-?}" "${U:-?}" "${P:-?}" "$N" "${H:0:120}" "$F" >> "$META"

        if [[ "$N" -gt 0 ]]; then
            ok "$(basename "$F") — ${BOLD}${N}$(L " righe sospette" " suspicious lines")"
        else
            dim_msg "$(basename "$F")"
        fi
        BODY+=$(file_card_html "$F" "$KW" "PS")
    done

    separator
    info "$(L "Transcript analizzati:" "Transcripts analysed:") ${BOLD}$TOTAL"
    [[ $SUSP -gt 0 ]] && warn "$(L "Con contenuto sospetto:" "With suspicious content:") ${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && return 0

    ask_yn "Generare report HTML?" || return 0

    local SUMTABLE; SUMTABLE=$(_rows_to_table "$(sort "$META")" \
        "$(L "Inizio" "Start")" "$(L "Utente" "User")" "PID" \
        "$(L "Righe sospette" "Suspicious lines")" "Host Application" "$(L "Percorso" "Path")")

    local STATS
    STATS="$(stat_box "Transcript" "$TOTAL")"
    STATS+="$(stat_box "$(L "Sospetti" "Suspicious")" "$SUSP" "$([[ $SUSP -gt 0 ]] && echo warn || echo info)")"
    finish_report "ps_transcripts" "PowerShell Transcript" "PS" "PowerShell_transcript*.txt" "$STATS" \
        "<div class='cards'>$(generic_card_html "$(L "Sessioni" "Sessions")" "$(L "intestazioni dei transcript" "transcript headers")" "$TOTAL" "$SUMTABLE" "▤")</div><div class='cards'>${BODY}</div>"
}
