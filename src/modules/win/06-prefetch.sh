
# ================================================================
#  MODULO 6 — Prefetch
# ================================================================
module_prefetch() {
    section_header "$(L "Prefetch — Eseguibili Tracciati" "Prefetch — Execution History")" "$GREEN"
    check_win_root || return 1

    local PREF_DIR
    PREF_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/Prefetch")
    if [[ -z "$PREF_DIR" || ! -d "$PREF_DIR" ]]; then
        warn "$(L "Directory Prefetch non trovata (potrebbe essere disabilitato o sistema su SSD)" "Prefetch directory not found (may be disabled or SSD system)")"
        return 0
    fi

    mapfile -t PF_FILES < <(find "$PREF_DIR" -maxdepth 1 -iname "*.pf" -type f 2>/dev/null | sort)
    local TOTAL=${#PF_FILES[@]}
    if [[ $TOTAL -eq 0 ]]; then
        warn "$(L "Nessun file .pf trovato" "No .pf files found")"
        return 0
    fi

    info "$(L "Trovati" "Found") $TOTAL file .pf in: $PREF_DIR"

    # Parser Python per Prefetch (formato MAM/Uncompressed)
    declare -a PF_ENTRIES=()
    for PF in "${PF_FILES[@]}"; do
        local FNAME; FNAME=$(basename "$PF")
        local FSIZE; FSIZE=$(stat -c "%s" "$PF" 2>/dev/null || echo "?")
        local FMTIME; FMTIME=$(stat -c "%y" "$PF" 2>/dev/null | cut -d'.' -f1 || echo "?")
        # Estrai nome eseguibile dal nome file (es. CMD.EXE-XXXXXXXX.pf)
        local EXENAME="${FNAME%-*}"
        PF_ENTRIES+=("${EXENAME}|${FNAME}|${FSIZE}|${FMTIME}")
    done

    # Ordina per data decrescente (più recente prima) — FMTIME è campo 4 in formato YYYY-MM-DD HH:MM:SS
    mapfile -t PF_ENTRIES < <(printf '%s\n' "${PF_ENTRIES[@]}" | sort -t'|' -k4 -r)

    for E in "${PF_ENTRIES[@]}"; do
        IFS='|' read -r EXENAME FNAME FSIZE FMTIME <<< "$E"
        printf "  ${DIM}%-30s${RESET}  %-30s  %s\n" "$EXENAME" "$FMTIME" "$FSIZE B"
    done

    separator
    info "$(L "File Prefetch:" "Prefetch Files:") ${BOLD}$TOTAL"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "prefetch")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${PF_ENTRIES[@]}"; do
        IFS='|' read -r EXENAME FNAME FSIZE FMTIME <<< "$E"
        local CSS=""
        echo "$EXENAME" | grep -qi "powershell\|cmd\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin\|mimikatz\|procdump\|psexec\|wce\|fgdump" && CSS="class='warn'"
        ROWS+="<tr><td class='mono' ${CSS}>$(html_esc "$EXENAME")</td>
               <td class='mono dim'>$(html_esc "$FNAME")</td>
               <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
               <td class='mono mid'>${FSIZE} B</td></tr>"
    done

    {
        html_header "Prefetch"
        html_page_header "PF" "Prefetch <span>Forensics</span>" \
            "Windows\\Prefetch\\*.pf" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File .pf</div><div class='value'>${TOTAL}</div></div>
        </div>
        <main><div class='stitle'>Eseguibili tracciati dal Prefetcher</div>
        <div class='card'><table>
          <thead><tr><th>$(L "Eseguibile" "Executable")</th><th>File .pf</th><th>$(L "Ultima esecuzione (mtime)" "Last run (mtime)")</th><th>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
