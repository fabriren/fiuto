
# ================================================================
#  MODULO 43 — Volume Shadow Copies
#
#  Le shadow copy sono fotografie precedenti del volume. Valgono molto in DFIR
#  per due ragioni:
#
#  - contengono file che l'attaccante ha poi cancellato o modificato, e hive di
#    registro anteriori alla compromissione;
#  - permettono l'analisi DIFFERENZIALE: cosa e' cambiato fra due istanti.
#
#  Sono anche un bersaglio: la loro cancellazione (vssadmin delete shadows) e'
#  un passo standard del ransomware, quindi l'ASSENZA di shadow copy su una
#  macchina che dovrebbe averne e' essa stessa un indicatore.
#
#  Nota sui limiti: montare una shadow copy richiede libvshadow (vshadowmount),
#  privilegi e un mount point. Questo modulo inventaria e prepara i comandi;
#  non monta nulla da solo, perche' un mount silenzioso su una postazione
#  forense non e' un effetto collaterale accettabile.
# ================================================================
module_vss() {
    section_header "Volume Shadow Copies" "$CYAN"
    check_win_root || return 1

    local SVI; SVI=$(ci_find_dir "$WIN_ROOT" "System Volume Information")
    local -a STORES=()
    local F

    if [[ -n "$SVI" ]]; then
        # I contenitori delle shadow copy hanno il GUID del provider VSS nel nome.
        while IFS= read -r F; do
            [[ -n "$F" ]] && STORES+=("$F")
        done < <(find "$SVI" -maxdepth 1 -type f -iname '*3808876b-c176-4e48-b7ae-04046e6cc752*' 2>/dev/null)
    fi

    local NSTORE=${#STORES[@]}
    local ROWS="" TOTBYTES=0
    for F in "${STORES[@]}"; do
        local SZ MT
        SZ=$(stat -c %s "$F" 2>/dev/null || echo 0)
        MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1 || echo "?")
        TOTBYTES=$((TOTBYTES + SZ))
        ROWS+="${MT}	$(basename "$F")	${SZ}
"
    done

    # vshadowinfo legge il catalogo VSS, ma opera sul DEVICE o sull'immagine
    # raw, non su un filesystem gia' montato: da qui si puo' solo segnalarne la
    # disponibilita' e preparare il comando, non eseguirlo.
    local HAVE_VSHADOW=false
    command -v vshadowinfo > /dev/null 2>&1 && HAVE_VSHADOW=true

    separator
    if [[ "$NSTORE" -eq 0 ]]; then
        warn "$(L "Nessun contenitore di shadow copy trovato in System Volume Information." \
                 "No shadow copy container found in System Volume Information.")"
        echo ""
        warn "$(L "L'assenza non e' neutra: la cancellazione delle shadow copy (vssadmin delete shadows) e' un passo tipico del ransomware. Verifica gli Event ID 524/8224 nel modulo Event Log e la presenza di vssadmin/wmic nelle history." \
                 "Absence is not neutral: deleting shadow copies (vssadmin delete shadows) is a standard ransomware step. Check Event IDs 524/8224 in the Event Log module and vssadmin/wmic usage in command histories.")"
    else
        ok "$(L "Contenitori shadow copy:" "Shadow copy containers:") ${BOLD}$NSTORE"
        info "$(L "Spazio occupato:" "Space used:") ${BOLD}$(numfmt --to=iec "$TOTBYTES" 2>/dev/null || echo "$TOTBYTES B")"
        echo ""
        info "$(L "Le date reali degli snapshot stanno nel catalogo VSS, leggibile solo dal device o dall'immagine raw:" \
                 "Real snapshot dates live in the VSS catalogue, readable only from the device or raw image:")"
        if $HAVE_VSHADOW; then
            echo -e "      ${CYAN}vshadowinfo <device|immagine.raw>${RESET}"
        else
            info "$(L "vshadowinfo non installato (pacchetto libvshadow-utils)." \
                     "vshadowinfo not installed (libvshadow-utils package).")"
        fi
    fi

    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    if [[ "$NSTORE" -eq 0 ]]; then
        BODY+="<b>$(L "Nessuna shadow copy presente" "No shadow copy present")</b><br>"
        BODY+="$(L "L'assenza va interpretata, non archiviata: la cancellazione delle shadow copy e' un passo standard del ransomware prima della cifratura. Da correlare con gli Event ID 524 e 8224 e con l'uso di vssadmin, wmic shadowcopy o Win32_ShadowCopy nelle history dei comandi." \
            "Absence needs interpreting, not filing: deleting shadow copies is a standard ransomware step before encryption. Correlate with Event IDs 524 and 8224 and with use of vssadmin, wmic shadowcopy or Win32_ShadowCopy in command histories.")"
    else
        BODY+="<b>$(L "Come sfruttarle" "How to use them")</b><br>"
        BODY+="$(L "Ogni shadow copy e' una versione precedente del volume: contiene file poi cancellati e hive di registro anteriori alla compromissione. Il vero valore e' il confronto differenziale fra due istanti." \
            "Each shadow copy is an earlier version of the volume: it holds files later deleted and registry hives predating the compromise. The real value is the differential comparison between two points in time.")<br><br>"
        BODY+="<b>$(L "Procedura consigliata" "Suggested procedure")</b><br>"
        BODY+="<code>vshadowinfo /dev/sdX1</code> — $(L "elenca gli snapshot e le loro date" "list snapshots and their dates")<br>"
        BODY+="<code>vshadowmount /dev/sdX1 /mnt/vss</code> — $(L "espone gli snapshot come file vssN" "expose snapshots as vssN files")<br>"
        BODY+="<code>mount -o ro,loop /mnt/vss/vss1 /mnt/snap1</code><br>"
        BODY+="<code>./fiuto.sh /mnt/snap1 --all</code> — $(L "rilancia FIUTO sullo snapshot" "re-run FIUTO on the snapshot")<br><br>"
        BODY+="$(L "Confrontando i report di due snapshot si isola cosa e' cambiato nella finestra di compromissione." \
            "Comparing the reports of two snapshots isolates what changed during the compromise window.")"
    fi
    BODY+="</div></div>"

    if [[ "$NSTORE" -gt 0 ]]; then
        local TABLE; TABLE=$(_rows_to_table "$ROWS" \
            "$(L "Ultima modifica" "Last modified")" "$(L "Contenitore" "Container")" "$(L "Byte" "Bytes")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Contenitori shadow copy" "Shadow copy containers")" "${SVI:-System Volume Information}" "$NSTORE" "$TABLE" "◫")</div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Contenitori" "Containers")" "$NSTORE" "$([[ "$NSTORE" -eq 0 ]] && echo warn || echo info)")"
    [[ "$NSTORE" -gt 0 ]] && STATS+="$(stat_box "$(L "Spazio" "Space")" "$(numfmt --to=iec "$TOTBYTES" 2>/dev/null || echo "$TOTBYTES")" "info")"
    finish_report "vss" "Volume Shadow Copies" "VSS" "System Volume Information" "$STATS" "$BODY"
}
