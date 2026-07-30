
# ================================================================
#  macOS 18 — Time Machine e snapshot
#
#  Come per le shadow copy su Windows, gli snapshot contengono versioni
#  precedenti dei file: contenuto poi cancellato o cifrato, e configurazioni
#  anteriori alla compromissione.
#
#  Valgono anche al contrario: il ransomware su macOS cancella gli snapshot
#  locali (tmutil deletelocalsnapshots) prima di cifrare, quindi la loro
#  scomparsa su una macchina che li aveva e' un indicatore.
#
#  Limite dichiarato: l'elenco degli snapshot APFS vive nei metadati del
#  container, non nel filesystem montato. Da qui si leggono la configurazione
#  di Time Machine, le destinazioni note e le tracce degli snapshot, non
#  l'elenco autoritativo — che richiede `tmutil listlocalsnapshots` o
#  `diskutil apfs listSnapshots` sul volume vivo.
# ================================================================
module_macos_backups() {
    section_header "macOS — Time Machine / Snapshot" "$BLUE"
    check_target_root || return 1

    local ROWS="" NDEST=0 NTRACE=0

    # ---------- Configurazione Time Machine ----------
    local TMP_PLIST
    for TMP_PLIST in "Library/Preferences/com.apple.TimeMachine.plist" \
                     "private/var/db/com.apple.xpc.launchd/disabled.plist"; do
        local F; F=$(ci_find_file "$WIN_ROOT" "$TMP_PLIST")
        [[ -s "$F" ]] || continue
        local TXT; TXT=$(read_plist "$F" 2>/dev/null)
        [[ -z "$TXT" ]] && continue
        if [[ "$TMP_PLIST" == *TimeMachine* ]]; then
            # Destinazioni, ultimo backup, esclusioni.
            local DESTS; DESTS=$(printf '%s\n' "$TXT" | grep -aoE '(BackupAlias|DestinationID|LastKnownVolumeName|LastDestinationID|SnapshotDates?|LastKnownEncryptionState|AutoBackup)[^<]*<?[^<>]{0,120}' | head -40)
            local LN
            while IFS= read -r LN; do
                [[ -z "$LN" ]] && continue
                NDEST=$((NDEST + 1))
                ROWS+="Time Machine	${LN}	$(basename "$F")
"
            done <<< "$DESTS"
            ok "$(L "Configurazione Time Machine trovata:" "Time Machine configuration found:") ${DIM}${F}${RESET}"
        fi
    done

    # ---------- Tracce di snapshot locali ----------
    local D
    for D in ".MobileBackups" "Volumes/.timemachine" "private/var/db/com.apple.TimeMachine.SnapshotCount"; do
        local X; X=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$X" ]] && X=$(ci_find_file "$WIN_ROOT" "$D")
        [[ -z "$X" ]] && continue
        NTRACE=$((NTRACE + 1))
        ROWS+="Snapshot	${D}	$(stat -c %y "$X" 2>/dev/null | cut -d. -f1)
"
        info "$(L "Traccia di snapshot:" "Snapshot trace:") ${DIM}${X}${RESET}"
    done

    # ---------- Backup montati ----------
    local NBK=0
    local BK; BK=$(ci_find_dir "$WIN_ROOT" "Backups.backupdb")
    if [[ -n "$BK" ]]; then
        while IFS= read -r F; do
            NBK=$((NBK + 1))
            ROWS+="Backup	$(basename "$F")	$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
"
        done < <(find "$BK" -mindepth 2 -maxdepth 2 -type d 2>/dev/null | head -100)
        ok "Backups.backupdb — ${BOLD}${NBK}${RESET} $(L "backup" "backups")"
    fi

    separator
    if [[ -z "$ROWS" ]]; then
        warn "$(L "Nessun artefatto Time Machine o snapshot trovato." "No Time Machine or snapshot artefact found.")"
        echo ""
        warn "$(L "L'assenza va interpretata: su macOS il ransomware cancella gli snapshot locali prima di cifrare (tmutil deletelocalsnapshots). Verificare la history della shell e i log per l'uso di tmutil." \
                 "Absence needs interpreting: on macOS, ransomware deletes local snapshots before encrypting (tmutil deletelocalsnapshots). Check shell history and logs for tmutil usage.")"
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Quello che questo modulo NON puo' dirti" "What this module cannot tell you")</b><br>"
    BODY+="$(L "L'elenco autoritativo degli snapshot APFS sta nei metadati del container, non nel filesystem montato: da un'analisi offline si vedono la configurazione di Time Machine e le tracce lasciate sul volume, non gli snapshot effettivamente esistenti. Per l'elenco reale servono, sul volume vivo o sull'immagine del container: tmutil listlocalsnapshots / e diskutil apfs listSnapshots." \
        "The authoritative list of APFS snapshots lives in the container metadata, not in the mounted filesystem: offline analysis shows the Time Machine configuration and the traces left on the volume, not the snapshots that actually exist. For the real list, on the live volume or the container image: tmutil listlocalsnapshots / and diskutil apfs listSnapshots.")<br><br>"
    BODY+="<b>$(L "Se ci sono, sfruttali" "If they exist, use them")</b><br>"
    BODY+="$(L "Uno snapshot montato e' un volume analizzabile: rilancia FIUTO su di esso e confronta i report per isolare cosa e' cambiato nella finestra di compromissione." \
        "A mounted snapshot is an analysable volume: re-run FIUTO on it and compare the reports to isolate what changed during the compromise window.")"
    BODY+="</div></div>"
    BODY+="<div class='cards'>$(generic_card_html "$(L "Artefatti di backup" "Backup artefacts")" "$WIN_ROOT" "$((NDEST + NTRACE + NBK))" \
        "$(_rows_to_table "$ROWS" "$(L "Tipo" "Type")" "$(L "Dato" "Item")" "$(L "Origine / data" "Source / date")")" "◫")"

    local STATS
    STATS="$(stat_box "$(L "Config TM" "TM config")" "$NDEST" "info")"
    STATS+="$(stat_box "$(L "Tracce snapshot" "Snapshot traces")" "$NTRACE" "$([[ "$NTRACE" -eq 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Backup" "$NBK" "info")"
    finish_report "macos_backups" "macOS Time Machine / Snapshot" "TM" "com.apple.TimeMachine.plist · Backups.backupdb" "$STATS" "$BODY"
}
