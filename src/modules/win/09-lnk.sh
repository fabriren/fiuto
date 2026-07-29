
# ================================================================
#  MODULO 10 — LNK & JumpList
# ================================================================
module_lnk() {
    section_header "LNK Files & JumpList" "$GREEN"
    check_win_root || return 1

    local TOTAL_LNK=0 TOTAL_USERS=0
    declare -a ALL_ENTRIES=()

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local RECENT_DIR
        RECENT_DIR=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Microsoft/Windows/Recent")
        [[ -z "$RECENT_DIR" ]] && { dim_msg "$USERNAME — $(L "Recent non trovata" "Recent not found")"; continue; }

        mapfile -t LNK_FILES < <(find "$RECENT_DIR" -maxdepth 1 -iname "*.lnk" -type f -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2- | head -50)
        if [[ ${#LNK_FILES[@]} -eq 0 ]]; then
            mapfile -t LNK_FILES < <(find "$RECENT_DIR" -maxdepth 1 -iname "*.lnk" -type f -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null | head -50)
        fi
        local COUNT=${#LNK_FILES[@]}
        [[ $COUNT -eq 0 ]] && { dim_msg "$USERNAME — $(L "nessun .lnk trovato" "no .lnk files found")"; continue; }

        ok "$USERNAME — $COUNT $(L "file .lnk recenti" "recent .lnk files")"
        TOTAL_USERS=$((TOTAL_USERS + 1))

        for LNK in "${LNK_FILES[@]}"; do
            local FNAME; FNAME=$(basename "$LNK")
            local FMTIME; FMTIME=$(stat -c "%y" "$LNK" 2>/dev/null | cut -d'.' -f1 || echo "?")
            # Prova a estrarre il target path con python-lnk o strings
            local TARGET
            TARGET=$("$PY3" - "$LNK" << 'PYEOF' 2>/dev/null || true
import sys, struct
try:
    with open(sys.argv[1], 'rb') as f: data = f.read()
    # LNK magic: 4C 00 00 00
    if data[:4] != b'\x4c\x00\x00\x00': sys.exit()
    # Header size = 76 bytes, flags at offset 20
    flags = struct.unpack_from('<I', data, 20)[0]
    offset = 76
    # Shell Item ID List
    if flags & 0x01:
        id_list_size = struct.unpack_from('<H', data, offset)[0]
        offset += 2 + id_list_size
    # Link Info
    if flags & 0x02:
        li_size = struct.unpack_from('<I', data, offset)[0]
        li_hdr_size = struct.unpack_from('<I', data, offset + 4)[0]
        # Local base path offset
        if li_hdr_size >= 28:
            lbp_off = struct.unpack_from('<I', data, offset + 16)[0]
            if lbp_off > 0:
                abs_off = offset + lbp_off
                end = data.index(b'\x00', abs_off)
                print(data[abs_off:end].decode('latin-1'))
                sys.exit()
        offset += li_size
except: pass
PYEOF
)
            [[ -z "$TARGET" ]] && TARGET=$(strings "$LNK" 2>/dev/null | grep -E "^[A-Za-z]:\\\\" | head -1 || echo "")
            printf "  ${DIM}%-45s${RESET}  ${CYAN}%-25s${RESET}  %s\n" "$FNAME" "$FMTIME" "${TARGET:--}"
            ALL_ENTRIES+=("${USERNAME}|${FNAME}|${FMTIME}|${TARGET}")
            TOTAL_LNK=$((TOTAL_LNK + 1))
        done
        echo ""
    done < <(get_user_homes)

    separator
    info "$(L "Utenti:" "Users:") ${BOLD}$TOTAL_USERS${RESET}  |  .lnk trovati: ${BOLD}$TOTAL_LNK"
    [[ $TOTAL_LNK -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "lnk")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | sort -t'|' -k3,3r)

    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r UNAME FNAME FMTIME TARGET <<< "$E"
        ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
               <td class='mono ok' style='white-space:nowrap'>${FMTIME}</td>
               <td class='mono'>$(html_esc "$FNAME")</td>
               <td class='mono mid' style='word-break:break-all;font-size:.72rem'>$([[ -n "$TARGET" ]] && html_esc "$TARGET" || echo "-")</td></tr>"
    done

    {
        html_header "LNK Files"
        html_page_header "LN" "LNK Files <span>&amp; JumpList</span> Forensics" \
            "%APPDATA%\\Microsoft\\Windows\\Recent\\*.lnk" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File .lnk</div><div class='value'>${TOTAL_LNK}</div></div>
          <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>
        </div>
        <main><div class='stitle'>File recenti (LNK)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:16%;white-space:nowrap'>$(L "Data accesso" "Access date")</th><th style='width:24%'>File .lnk</th><th>Target path</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
