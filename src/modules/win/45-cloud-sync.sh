
# ================================================================
#  MODULO 45 — Cloud Sync (OneDrive / Dropbox / Google Drive / Box)
#
#  L'esfiltrazione moderna raramente passa da una chiavetta: passa da una
#  cartella sincronizzata. Un file trascinato in OneDrive esce dal perimetro
#  senza generare traffico riconoscibile come esfiltrazione e senza toccare
#  nessuno degli artefatti USB.
#
#  Il modulo ricostruisce quali provider erano configurati, con quali account,
#  e — dove il formato lo consente — l'elenco dei file sincronizzati.
#
#  Nota sui formati: il database di Google Drive e' SQLite e si legge per
#  intero. I log .odl di OneDrive sono binari con stringhe offuscate: qui se ne
#  estraggono solo i nomi di file leggibili, e il report lo dichiara.
# ================================================================
module_cloud_sync() {
    section_header "Cloud Sync — $(L "Provider e file sincronizzati" "Providers and synced files")" "$BLUE"
    check_win_root || return 1

    local BODY="" NPROV=0 NFILES=0 NACC=0
    local ACCROWS="" FILEROWS=""
    local HOME_DIR

    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local LOCAL ROAM
        LOCAL=$(ci_find_dir "$HOME_DIR" "AppData/Local")
        ROAM=$(ci_find_dir "$HOME_DIR" "AppData/Roaming")

        # ---------- OneDrive ----------
        local OD; OD=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "Microsoft/OneDrive")
        if [[ -n "$OD" ]]; then
            NPROV=$((NPROV + 1))
            # settings/<Personal|Business1>/*.ini contiene account e cartella locale
            local INI
            while IFS= read -r INI; do
                [[ -s "$INI" ]] || continue
                local EMAIL FOLDER
                EMAIL=$(grep -aoE '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' "$INI" 2>/dev/null | head -1)
                FOLDER=$(grep -aoE '[A-Z]:\\[^"|]{3,120}' "$INI" 2>/dev/null | head -1)
                [[ -z "$EMAIL$FOLDER" ]] && continue
                NACC=$((NACC + 1))
                ACCROWS+="OneDrive	${U}	${EMAIL:-?}	${FOLDER:-?}	$(basename "$INI")
"
            done < <(find "$OD" -maxdepth 3 -type f \( -iname '*.ini' -o -iname 'global.ini' \) 2>/dev/null)

            # Log .odl: estrazione euristica dei nomi di file
            local ODLDIR; ODLDIR=$(ci_find_dir "$OD" "logs")
            if [[ -n "$ODLDIR" ]]; then
                local NODL; NODL=$(find "$ODLDIR" -type f \( -iname '*.odl' -o -iname '*.odlgz' -o -iname '*.aodl' \) 2>/dev/null | wc -l)
                if [[ "$NODL" -gt 0 ]]; then
                    info "OneDrive ($U) — ${BOLD}${NODL}${RESET} $(L "file di log" "log files")"
                    local NAMES
                    NAMES=$(find "$ODLDIR" -type f \( -iname '*.odl' -o -iname '*.aodl' \) -print0 2>/dev/null \
                            | xargs -0 strings -n 6 2>/dev/null \
                            | grep -aoE '[A-Za-z0-9 _().+-]{3,80}\.(docx?|xlsx?|pptx?|pdf|zip|7z|rar|txt|csv|jpg|jpeg|png|eml|msg|pst|ost|key|pem|sql|bak|exe|dll|ps1|vbs|js)' \
                            | sort -u | head -3000)
                    local N; N=$(printf '%s\n' "$NAMES" | grep -c . || true)
                    if [[ "$N" -gt 0 ]]; then
                        NFILES=$((NFILES + N))
                        while IFS= read -r FN; do
                            [[ -n "$FN" ]] && FILEROWS+="OneDrive	${U}	${FN}
"
                        done <<< "$NAMES"
                    fi
                fi
            fi
        fi

        # ---------- Dropbox ----------
        local DBX; DBX=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "Dropbox")
        local DBXR; DBXR=$(ci_find_dir "${ROAM:-$HOME_DIR}" "Dropbox")
        if [[ -n "$DBX" || -n "$DBXR" ]]; then
            NPROV=$((NPROV + 1))
            local IJ; IJ=$(ci_find_file "${DBX:-$DBXR}" "info.json")
            if [[ -s "$IJ" ]]; then
                local PATHS
                PATHS=$("$PY3" -c "
import json,sys
try:
    d=json.load(open(sys.argv[1],encoding='utf-8',errors='replace'))
    for k,v in d.items():
        print(f\"{k}\t{v.get('path','')}\t{v.get('subscription_type','')}\")
except Exception:
    pass" "$IJ" 2>/dev/null)
                while IFS=$'\t' read -r KIND P SUB; do
                    [[ -z "$KIND" ]] && continue
                    NACC=$((NACC + 1))
                    ACCROWS+="Dropbox	${U}	${KIND} ${SUB}	${P}	info.json
"
                done <<< "$PATHS"
            fi
        fi

        # ---------- Google Drive ----------
        local GD; GD=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "Google/DriveFS")
        if [[ -n "$GD" ]]; then
            NPROV=$((NPROV + 1))
            local MDB
            while IFS= read -r MDB; do
                [[ -s "$MDB" ]] || continue
                # Schema DriveFS: items(stable_id, ..., local_title, modified_date, size)
                local ROWS
                ROWS=$(query_sqlite "$MDB" "SELECT local_title, datetime(modified_date,'unixepoch'), file_size FROM items WHERE local_title IS NOT NULL ORDER BY modified_date DESC LIMIT 20000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                local N; N=$(printf '%s\n' "$ROWS" | grep -c . || true)
                NFILES=$((NFILES + N))
                info "Google Drive ($U) — ${BOLD}${N}${RESET} $(L "elementi" "items")"
                while IFS=$'\t' read -r T D S; do
                    [[ -n "$T" ]] && FILEROWS+="Google Drive	${U}	${T}  (${D:-?}, ${S:-?} B)
"
                done <<< "$ROWS"
            done < <(find "$GD" -maxdepth 3 -type f -name 'metadata_sqlite_db' 2>/dev/null)
        fi

        # ---------- Box / iCloud: solo presenza ----------
        local P
        for P in "Box/Box" "Apple/CloudDocs" "iCloudDrive"; do
            local X; X=$(ci_find_dir "${LOCAL:-$HOME_DIR}" "$P")
            [[ -n "$X" ]] && { NPROV=$((NPROV + 1)); ACCROWS+="${P%%/*}	${U}	-	${X}	$(L "solo presenza" "presence only")
"; }
        done
    done < <(get_user_homes)

    separator
    if [[ "$NPROV" -eq 0 ]]; then
        warn "$(L "Nessun provider di sincronizzazione cloud rilevato." "No cloud sync provider detected.")"
        return 0
    fi
    ok "$(L "Provider rilevati:" "Providers detected:") ${BOLD}$NPROV"
    info "$(L "Account/configurazioni:" "Accounts/configurations:") ${BOLD}$NACC"
    info "$(L "Nomi di file recuperati:" "File names recovered:") ${BOLD}$NFILES"

    ask_yn "Generare report HTML?" || return 0

    BODY="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Perche' conta" "Why it matters")</b><br>"
    BODY+="$(L "Un file trascinato in una cartella sincronizzata esce dal perimetro senza toccare nessun artefatto USB e senza generare traffico riconoscibile come esfiltrazione. Questi elenchi vanno confrontati con i dati che l'organizzazione considera riservati." \
        "A file dropped into a synced folder leaves the perimeter without touching any USB artefact and without generating traffic recognisable as exfiltration. Cross-check these lists against the data the organisation treats as confidential.")<br><br>"
    BODY+="<b>$(L "Attendibilita' delle fonti" "Source reliability")</b><br>"
    BODY+="$(L "Google Drive: database SQLite, elenco completo e datato. Dropbox: configurazione JSON, account e percorsi. OneDrive: i log .odl sono binari con stringhe offuscate, qui se ne estraggono solo i nomi leggibili — sono indizi, non un inventario completo, e non hanno data." \
        "Google Drive: SQLite database, complete and dated listing. Dropbox: JSON configuration, accounts and paths. OneDrive: .odl logs are binary with obfuscated strings; only readable names are extracted here — these are leads, not a complete inventory, and carry no timestamp.")"
    BODY+="</div></div>"

    if [[ -n "$ACCROWS" ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "Account e configurazioni" "Accounts and configurations")" "AppData" "$NACC" \
            "$(_rows_to_table "$ACCROWS" "Provider" "$(L "Utente Windows" "Windows user")" "Account" "$(L "Cartella / dettaglio" "Folder / detail")" "$(L "Origine" "Source")")" "☁")</div>"
    fi
    if [[ -n "$FILEROWS" ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "File sincronizzati" "Synced files")" "$(L "Google Drive completo · OneDrive euristico" "Google Drive complete · OneDrive heuristic")" "$NFILES" \
            "$(_rows_to_table "$FILEROWS" "Provider" "$(L "Utente Windows" "Windows user")" "$(L "File" "File")")" "⇪")</div>"
    fi

    local STATS
    STATS="$(stat_box "Provider" "$NPROV")"
    STATS+="$(stat_box "Account" "$NACC" "info")"
    STATS+="$(stat_box "$(L "File" "Files")" "$NFILES" "warn")"
    finish_report "cloud_sync" "Cloud Sync" "CLD" "OneDrive · Dropbox · Google Drive · Box" "$STATS" "$BODY"
}
