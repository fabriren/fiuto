
# ================================================================
#  MODULO 35 — JumpLists (AutomaticDestinations + CustomDestinations)
# ================================================================
module_jumplists() {
    section_header "JumpLists — AutomaticDestinations & CustomDestinations" "$GREEN"
    check_win_root || return 1

    declare -a ALL_ENTRIES=()
    local TOTAL=0 SUSP=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local AUTO_DIR
        AUTO_DIR=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Microsoft/Windows/Recent/AutomaticDestinations")
        local CUST_DIR
        CUST_DIR=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Microsoft/Windows/Recent/CustomDestinations")

        for JL_DIR in "$AUTO_DIR" "$CUST_DIR"; do
            [[ -z "$JL_DIR" || ! -d "$JL_DIR" ]] && continue
            local JL_TYPE="Auto"
            [[ "$JL_DIR" == *"CustomDestinations"* ]] && JL_TYPE="Custom"

            while IFS= read -r JFILE; do
                [[ -f "$JFILE" ]] || continue
                local FNAME; FNAME=$(basename "$JFILE")
                local FMTIME; FMTIME=$(stat -c "%y" "$JFILE" 2>/dev/null | cut -d'.' -f1 || echo "")
                local STRINGS_OUT
                STRINGS_OUT=$(strings -n 8 "$JFILE" 2>/dev/null | grep -E '\\\\|https?://|\.exe|\.dll|\.ps1|\.bat|\.vbs|\.lnk' | head -20 || true)
                [[ -z "$STRINGS_OUT" ]] && continue
                local FIRST_PATH; FIRST_PATH=$(echo "$STRINGS_OUT" | head -1)
                local IS_SUSP="false"
                echo "$STRINGS_OUT" | grep -qiE 'temp|appdata\\local\\temp|\.ps1|\.bat|\.vbs|programdata|public\\' && IS_SUSP="true"
                if [[ ${#IOC_LIST[@]} -gt 0 ]]; then
                    check_ioc "$STRINGS_OUT" && IS_SUSP="true"
                fi
                [[ "$IS_SUSP" == "true" ]] && printf "  ${RED}[!]${RESET} ${DIM}[%-10s]${RESET} [${JL_TYPE}] %s\n" "$USERNAME" "$FIRST_PATH"
                ALL_ENTRIES+=("${USERNAME}§${JL_TYPE}§${FMTIME}§${FNAME}§$(echo "$STRINGS_OUT" | head -3 | tr '\n' '|')§${IS_SUSP}")
                TOTAL=$((TOTAL+1))
                [[ "$IS_SUSP" == "true" ]] && SUSP=$((SUSP+1))
            done < <(find "$JL_DIR" -maxdepth 1 -type f 2>/dev/null | sort)
        done
    done < <(get_user_homes)

    separator
    info "JumpList entries: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna JumpList trovata." "No JumpLists found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "jumplists")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    # Ordina per timestamp decrescente (campo 3: FMTIME)
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+3*[""])[2],reverse=True); print("\n".join(L))')
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='§' read -r UNAME JT FMTIME FNAME PATHS IS_SUSP <<< "$E"
        local RS="" TC="mono"
        [[ "$IS_SUSP" == "true" ]] && RS="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" && TC="mono bad"
        local PATHS_DISP; PATHS_DISP=$(html_esc "${PATHS//|/<br>}")
        ROWS+="<tr ${RS}>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$JT")</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.68rem'>${FMTIME}</td>
          <td class='mono dim' style='font-size:.68rem'>$(html_esc "$FNAME")</td>
          <td class='${TC}' style='word-break:break-all;font-size:.72rem'>${PATHS_DISP}</td>
        </tr>"
    done
    {
        html_header "JumpLists"
        html_page_header "JL" "JumpLists — <span>AutomaticDestinations · CustomDestinations</span>" \
            "AppData\\Roaming\\Microsoft\\Windows\\Recent\\*Destinations" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Entries</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSP}</div></div>
        </div><main>
        <div class='stitle'>$(L "JumpLists — Utente · Tipo · Timestamp · File · Percorsi estratti (strings)" "JumpLists — User · Type · Timestamp · File · Extracted paths (strings)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:10%'>$(L "Utente" "User")</th><th style='width:7%'>$(L "Tipo" "Type")</th><th style='width:14%'>Modificato</th><th style='width:14%'>File JL</th><th>Percorsi estratti</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
