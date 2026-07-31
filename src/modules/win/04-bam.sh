
# ================================================================
#  MODULO 3 — BAM (Background Activity Moderator)
# ================================================================
module_bam() {
    section_header "BAM - Background Activity Moderator" "$BLUE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE
    HIVE=$(get_hive "SYSTEM")
    if [[ -z "$HIVE" || ! -f "$HIVE" ]]; then
        err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found")"
        return 1
    fi

    info "$(L "Parsing hive SYSTEM:" "Parsing SYSTEM hive:") $HIVE"

    # Tenta entrambi i ControlSet
    local BAM_DATA
    BAM_DATA=$("$PY3" - "$HIVE" << 'PYEOF'
import sys, datetime

try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])

    bam_key = None
    for cs in ['ControlSet001', 'ControlSet002', 'CurrentControlSet']:
        try:
            bam_key = hive.get_key(f'{cs}\\Services\\bam\\State\\UserSettings')
            break
        except:
            try:
                bam_key = hive.get_key(f'{cs}\\Services\\bam\\UserSettings')
                break
            except:
                continue

    if not bam_key:
        print("ERROR\tChiave BAM non trovata (Windows < 10 1709 o BAM disabilitato)")
        sys.exit(0)

    for sid_key in bam_key.iter_subkeys():
        sid = sid_key.name
        for v in sid_key.get_values():
            name = v.name
            if name.startswith('\\') or name.startswith('\\Device'):
                val = v.value
                ts = None
                if isinstance(val, bytes) and len(val) >= 8:
                    try:
                        ft = int.from_bytes(val[:8], 'little')
                        if ft > 0:
                            ts = datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ft//10)
                    except:
                        pass
                ts_str = ts.strftime('%Y-%m-%d %H:%M:%S UTC') if ts else '-'
                print(f"{sid}\t{name}\t{ts_str}")
except Exception as e:
    print(f"ERROR\t{e}\t")
PYEOF
)

    declare -a BAM_ENTRIES=()
    local ERROR_MSG=""
    while IFS= read -r LINE; do
        if [[ "$LINE" == ERROR* ]]; then
            ERROR_MSG="${LINE#ERROR }"
        else
            BAM_ENTRIES+=("$LINE")
        fi
    done <<< "$BAM_DATA"

    if [[ -n "$ERROR_MSG" ]]; then
        warn "$ERROR_MSG"
    fi

    local TOTAL=${#BAM_ENTRIES[@]}
    if [[ $TOTAL -eq 0 ]]; then
        warn "$(L "Nessuna voce BAM trovata." "No BAM entries found.")"
        return 0
    fi

    # Raggruppa per SID e stampa
    local CUR_SID=""
    local SID_COUNT=0
    echo ""
    for ENTRY in "${BAM_ENTRIES[@]}"; do
        IFS=$'\t' read -r SID EXEPATH TIMESTAMP <<< "$ENTRY"
        if [[ "$SID" != "$CUR_SID" ]]; then
            [[ -n "$CUR_SID" ]] && separator
            echo -e "  ${GREEN}${BOLD}SID: $SID${RESET}"
            CUR_SID="$SID"
            SID_COUNT=$((SID_COUNT + 1))
        fi
        printf "  ${DIM}%-25s${RESET}  %s\n" "$TIMESTAMP" "$EXEPATH"
    done

    separator
    info "$(L "Totale eseguibili tracciati:" "Total tracked executables:") ${BOLD}$TOTAL${RESET}  |  SID distinti: ${BOLD}$SID_COUNT"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "bam")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per timestamp decrescente (colonna 3) e poi per SID (colonna 1)
    mapfile -t BAM_ENTRIES < <(printf '%s\n' "${BAM_ENTRIES[@]}" | sort -t$'\t' -k3,3r -k1,1)

    # Raggruppa per SID in HTML
    local CARDS_HTML=""
    local PREV_SID="" SID_ROWS="" SID_EXEC_COUNT=0
    for ENTRY in "${BAM_ENTRIES[@]}"; do
        IFS=$'\t' read -r SID EXEPATH TIMESTAMP <<< "$ENTRY"
        if [[ "$SID" != "$PREV_SID" ]]; then
            if [[ -n "$PREV_SID" ]]; then
                CARDS_HTML+="<div class='card'>
                  <div class='card-header'><div class='uicon'>▣</div>
                    <div><div class='uname'>SID: $(html_esc "$PREV_SID")</div></div>
                    <div class='badge'>${SID_EXEC_COUNT} eseguibili</div>
                  </div>
                  <table><thead><tr><th>Timestamp (UTC)</th><th>$(L "Eseguibile" "Executable")</th></tr></thead>
                  <tbody>${SID_ROWS}</tbody></table></div>"
            fi
            PREV_SID="$SID"; SID_ROWS=""; SID_EXEC_COUNT=0
        fi
        SID_EXEC_COUNT=$((SID_EXEC_COUNT + 1))
        local EXE_ESC; EXE_ESC=$(html_esc "$EXEPATH")
        local EXE_HTML="<span class='mono'>${EXE_ESC}</span>"
        # Evidenzia percorsi sospetti
        if echo "$EXEPATH" | grep -qi "temp\|appdata\|public\|programdata\|downloads"; then
            EXE_HTML="<span class='mono warn'>${EXE_ESC}</span>"
        fi
        SID_ROWS+="<tr><td class='mono ok' style='white-space:nowrap'>$(html_esc "$TIMESTAMP")</td><td>${EXE_HTML}</td></tr>"
    done
    # Ultima card
    if [[ -n "$PREV_SID" ]]; then
        CARDS_HTML+="<div class='card'>
          <div class='card-header'><div class='uicon'>▣</div>
            <div><div class='uname'>SID: $(html_esc "$PREV_SID")</div></div>
            <div class='badge'>${SID_EXEC_COUNT} eseguibili</div>
          </div>
          <table><thead><tr><th>Timestamp (UTC)</th><th>$(L "Eseguibile" "Executable")</th></tr></thead>
          <tbody>${SID_ROWS}</tbody></table></div>"
    fi

    {
        html_header "BAM"
        html_page_header "BA" "BAM - <span>Background Activity</span> Moderator" \
            "SYSTEM\\ControlSet001\\Services\\bam\\State\\UserSettings" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Eseguibili</div><div class='value'>${TOTAL}</div></div>
          <div class='stat info'><div class='label'>SID</div><div class='value'>${SID_COUNT}</div></div>
        </div>
        <main><div class='stitle'>Eseguibili tracciati per SID</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
