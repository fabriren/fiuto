
# ================================================================
#  MODULO 26 — WER Files (Windows Error Reporting)
# ================================================================
module_wer_files() {
    section_header "WER Files - Windows Error Reporting" "$MAGENTA"
    check_win_root || return 1

    local TOTAL_FILES=0
    declare -a WER_FILES=()

    # Per-system WER (ProgramData) - ordina per data decrescente
    local SYSPATHS=(
        "ProgramData/Microsoft/Windows/WER/ReportArchive"
        "ProgramData/Microsoft/Windows/WER/ReportQueue"
        "Windows/System32/WER/ReportArchive"
        "Windows/System32/WER/ReportQueue"
    )
    for REL in "${SYSPATHS[@]}"; do
        local DIR
        DIR=$(ci_find_dir "$WIN_ROOT" "$REL")
        [[ -n "$DIR" && -d "$DIR" ]] || continue
        while IFS= read -r FDIR; do
            [[ -d "$FDIR" ]] || continue
            local WERF
            WERF=$(find "$FDIR" -maxdepth 1 -iname "Report.wer" -type f 2>/dev/null | head -1)
            [[ -n "$WERF" ]] && WER_FILES+=( "$WERF|System" )
        done < <(find "$DIR" -maxdepth 1 -mindepth 1 -type d -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
    done

    # Per-user WER (AppData/Local)
    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local USERPATHS=(
            "AppData/Local/Microsoft/Windows/WER/ReportArchive"
            "AppData/Local/Microsoft/Windows/WER/ReportQueue"
        )
        for REL in "${USERPATHS[@]}"; do
            local DIR
            DIR=$(ci_find_dir "$USER_DIR" "$REL")
            [[ -n "$DIR" && -d "$DIR" ]] || continue
            while IFS= read -r FDIR; do
                [[ -d "$FDIR" ]] || continue
                local WERF
                WERF=$(find "$FDIR" -maxdepth 1 -iname "Report.wer" -type f 2>/dev/null | head -1)
                [[ -n "$WERF" ]] && WER_FILES+=( "$WERF|$USERNAME" )
            done < <(find "$DIR" -maxdepth 1 -mindepth 1 -type d -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        done
    done < <(get_user_homes)

    TOTAL_FILES=${#WER_FILES[@]}
    info "$(L "File WER trovati:" "WER files found:") ${BOLD}$TOTAL_FILES"

    if [[ $TOTAL_FILES -eq 0 ]]; then
        ok "$(L "Nessun file WER trovato." "No WER files found.")"
        return 0
    fi

    local SCAN_COUNT=0
    for ENTRY in "${WER_FILES[@]}"; do
        SCAN_COUNT=$((SCAN_COUNT+1))
        local F="${ENTRY%|*}"
        local OWNER="${ENTRY#*|}"
        local FNAME; FNAME=$(basename "$(dirname "$F")")
        local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
        echo -e "      ${DIM}[$(printf "%2d" $SCAN_COUNT)] • $FNAME  (${OWNER}) - $FMTIME${RESET}"
    done

    separator
    info "$(L "Processati" "Processed") ${BOLD}$TOTAL_FILES${RESET} report di errore."

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wer_files")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""

    for ENTRY in "${WER_FILES[@]}"; do
        local F="${ENTRY%|*}"
        local OWNER="${ENTRY#*|}"
        local FDIR_NAME; FDIR_NAME=$(basename "$(dirname "$F")")
        local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
        
        # Caricamento contenuto file .wer (solitamente UTF-16LE con BOM)
        local CONTENT
        CONTENT=$("$PY3" - "$F" << 'PYEOF'
import sys, html
try:
    with open(sys.argv[1], 'rb') as f:
        raw = f.read()
    # Rilevamento encoding
    if raw.startswith(b'\xff\xfe'):
        text = raw[2:].decode('utf-16-le', errors='replace')
    elif raw.startswith(b'\xfe\xff'):
        text = raw[2:].decode('utf-16-be', errors='replace')
    elif raw.startswith(b'\xef\xbb\xbf'):
        text = raw[3:].decode('utf-8', errors='replace')
    else:
        # Tenta UTF-16-LE comunque se ci sono molti null
        if len(raw) >= 2 and raw[1] == 0:
            text = raw.decode('utf-16-le', errors='replace')
        else:
            text = raw.decode('utf-8', errors='replace')
    print(html.escape(text.strip()))
except Exception as e:
    print(f"Errore caricamento: {html.escape(str(e))}")
PYEOF
)
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon' style='background:linear-gradient(135deg,var(--accent2),#c0392b)'>!</div>
            <div>
              <div class='uname'>$(html_esc "$FDIR_NAME")</div>
              <div class='upath'>$(L "Sorgente:" "Source:") $(html_esc "$OWNER") &nbsp;·&nbsp; $(L "Data:" "Date:") ${FMTIME}</div>
            </div>
            <div class='badge warn'>Report.wer</div>
          </div>
          <div style='padding:0.8rem 1.2rem; background:var(--bg3); border-top:1px solid var(--border)'>
            <pre style='font-family:var(--mono); font-size:0.75rem; color:var(--text); line-height:1.5; white-space:pre-wrap; max-height:350px; overflow-y:auto'>${CONTENT}</pre>
          </div>
        </div>"
    done

    {
        html_header "Windows Error Reporting (WER)"
        html_page_header "WER" "Windows <span>Error Reporting</span> Analysis" \
            "WER\\ReportArchive | WER\\ReportQueue" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Totale Report</div><div class='value'>${TOTAL_FILES}</div></div>
        </div><main>
        <div class='stitle'>Dati rilevati dalle code di errore di sistema</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
