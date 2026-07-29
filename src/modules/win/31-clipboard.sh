

# ================================================================
#  MODULO 31 — Clipboard History (Win10+)
# ================================================================
module_clipboard() {
    section_header "$(L "Clipboard History — Cronologia Appunti" "Clipboard History")" "$YELLOW"
    check_win_root || return 1


    declare -a ALL_CLIPS=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local CB_DIR
        CB_DIR=$(ci_find_dir "$USER_DIR" "AppData/Local/Microsoft/Windows/Clipboard")
        [[ -z "$CB_DIR" || ! -d "$CB_DIR" ]] && continue
        info "$(L "Clipboard dir trovata:" "Clipboard dir found:") $CB_DIR"

        while IFS= read -r CFILE; do
            [[ -f "$CFILE" ]] || continue
            local FSIZE; FSIZE=$(stat -c "%s" "$CFILE" 2>/dev/null || echo "0")
            [[ "$FSIZE" -lt 4 ]] && continue
            local CLIP_TEXT
            CLIP_TEXT=$("$PY3" - "$CFILE" << 'PYEOF' 2>/dev/null || true
import sys, re
path = sys.argv[1]
try:
    with open(path,'rb') as f:
        data = f.read()
    # Try UTF-16-LE first (common for clipboard .dat)
    for enc in ('utf-16-le','utf-8','latin-1'):
        try:
            text = data.decode(enc, errors='ignore')
            text = text.replace('\x00','').strip()
            if len(text) > 3:
                print(text[:500])
                break
        except: continue
except Exception as e:
    pass
PYEOF
)
            [[ -z "$CLIP_TEXT" ]] && continue
            local FNAME; FNAME=$(basename "$CFILE")
            local FMTIME; FMTIME=$(stat -c "%y" "$CFILE" 2>/dev/null | cut -d'.' -f1 || echo "")
            local IS_SUSP="false"
            echo "$CLIP_TEXT" | grep -qiE '([0-9]{1,3}\.){3}[0-9]{1,3}|https?://|[A-Za-z0-9+/]{40,}={0,2}|password|token|secret' \
                && IS_SUSP="true"
            [[ "$IS_SUSP" == "true" ]] && printf "  ${RED}[!] [%-15s]${RESET}  %s\n" "$USERNAME" "${CLIP_TEXT:0:100}"
            ALL_CLIPS+=("${USERNAME}§${FMTIME}§${FNAME}§$(echo "$CLIP_TEXT" | head -1)§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            [[ "$IS_SUSP" == "true" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
        done < <(find "$CB_DIR" -maxdepth 3 -name "*.dat" -type f 2>/dev/null | sort)
    done < <(get_user_homes)

    separator
    info "Voci clipboard: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna voce clipboard trovata (feature Win10 1809+, va abilitata)." "No clipboard entries found (Win10 1809+ feature, must be enabled).")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "clipboard")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per timestamp decrescente globale
    mapfile -t ALL_CLIPS < <(printf '%s\n' "${ALL_CLIPS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+["",""])[1],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_CLIPS[@]}"; do
        IFS='§' read -r UNAME FMTIME FNAME CLIP_FIRST IS_SUSP <<< "$E"
        local RS="" TC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && TC="mono bad"
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${FMTIME}</td>
          <td class='${TC}' style='word-break:break-all;font-size:.72rem'>$(html_esc "${CLIP_FIRST:0:300}")</td>
        </tr>"
    done
    {
        html_header "Clipboard History"
        html_page_header "CB" "Clipboard History — <span>Cronologia Appunti</span>" \
            "AppData\\Local\\Microsoft\\Windows\\Clipboard" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Voci</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>$(L "Clipboard History — Utente · Timestamp · Contenuto (troncato a 300 char)" "Clipboard History — User · Timestamp · Content (truncated to 300 chars)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:14%'>Timestamp</th><th>$(L "Contenuto" "Content")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
