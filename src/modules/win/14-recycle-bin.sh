
# ================================================================
#  MODULO 14 — Recycle Bin ($Recycle.Bin)
# ================================================================
module_recycle_bin() {
    section_header "$(L "Recycle Bin — File Eliminati" "Recycle Bin — Deleted Files")" "$GREEN"
    check_win_root || return 1


    # Cerca $Recycle.Bin sul volume (può stare nella root)
    local RB_DIR
    RB_DIR=$(find "$WIN_ROOT" -maxdepth 1 -iname "\$Recycle.Bin" -type d 2>/dev/null | head -1)
    if [[ -z "$RB_DIR" ]]; then
        warn "$(L "\$Recycle.Bin non trovato nella root del volume" "\$Recycle.Bin not found in volume root")"
        return 0
    fi

    info "Directory: $RB_DIR"
    echo ""

    declare -a ALL_ITEMS=()
    local TOTAL=0 SUSP_COUNT=0

    # Parser $I files — formato binario con path originale e timestamp
    local IPARSE=""
    IPARSE=$(mktemp /tmp/dfir_rb_XXXXXX.py)
    register_tmp "$IPARSE"
    cat > "$IPARSE" << 'PYEOF'
import sys, struct, datetime, os

def parse_ifile(path):
    try:
        with open(path, 'rb') as f:
            data = f.read()
        if len(data) < 24:
            return None, None, None
        ver = struct.unpack_from('<Q', data, 0)[0]
        size = struct.unpack_from('<Q', data, 8)[0]
        ts_raw = struct.unpack_from('<Q', data, 16)[0]
        ts = ''
        if ts_raw:
            try:
                ts = (datetime.datetime(1601,1,1) +
                      datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
            except: pass
        # Path: versione 2 (Win10) ha lunghezza a offset 24
        if ver == 2 and len(data) >= 28:
            plen = struct.unpack_from('<I', data, 24)[0]
            orig_path = data[28:28+plen*2].decode('utf-16-le', errors='replace').rstrip('\x00')
        else:
            orig_path = data[24:].decode('utf-16-le', errors='replace').rstrip('\x00')
        return orig_path, size, ts
    except Exception as e:
        return None, None, None

path = sys.argv[1]
orig, size, ts = parse_ifile(path)
if orig:
    print(f"{orig}\t{size or 0}\t{ts or ''}")
PYEOF

    # Scansiona ogni SID
    for SID_DIR in "$RB_DIR"/*/; do
        local SID; SID=$(basename "$SID_DIR")
        local SID_COUNT=0

        for IFILE in "$SID_DIR"/\$I*; do
            [[ -f "$IFILE" ]] || continue
            local RESULT
            RESULT=$("$PY3" "$IPARSE" "$IFILE" 2>/dev/null || true)
            [[ -z "$RESULT" ]] && continue

            IFS=$'\t' read -r ORIG_PATH FSIZE DEL_TIME <<< "$RESULT"
            [[ -z "$ORIG_PATH" ]] && continue

            local FNAME; FNAME=$(basename "${IFILE/$'\$I'/$'\$R'}")
            local IS_SUSP=false
            echo "$ORIG_PATH" | grep -qi "system32\|passwd\|shadow\|lsass\|sam\|ntds\|\.ps1\|\.bat\|\.vbs\|\.exe\|\.dll" && IS_SUSP=true

            ALL_ITEMS+=("${SID}§${ORIG_PATH}§${FSIZE}§${DEL_TIME}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
            $IS_SUSP && SUSP_COUNT=$((SUSP_COUNT + 1))
            SID_COUNT=$((SID_COUNT + 1))

            if $IS_SUSP; then
                printf "  ${RED}[!] %s${RESET}\n" "$ORIG_PATH"
                printf "      ${DIM}SID: %-40s  Eliminato: %s  Size: %s B${RESET}\n" "$SID" "$DEL_TIME" "$FSIZE"
            else
                printf "  ${DIM}%-60s${RESET}  %s\n" "${ORIG_PATH: -60}" "${DEL_TIME:--}"
            fi
        done
        [[ $SID_COUNT -gt 0 ]] && ok "  $(L "SID $SID: $SID_COUNT file" "SID $SID: $SID_COUNT files")"
    done

    separator
    info "File nel cestino: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Cestino vuoto o nessun \$I file trovato." "Recycle bin empty or no \$I file found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "recycle_bin")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data eliminazione decrescente globale
    mapfile -t ALL_ITEMS < <(printf '%s\n' "${ALL_ITEMS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+4*[""])[3],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_ITEMS[@]}"; do
        IFS='§' read -r SID ORIG_PATH FSIZE DEL_TIME IS_SUSP <<< "$E"
        local ROW_STYLE="" PATH_CSS=""
        [[ "$IS_SUSP" == "true" ]] && {
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
            PATH_CSS="bad"
        }
        local SIZE_HR
        if [[ "$FSIZE" -gt 1073741824 ]] 2>/dev/null; then
            SIZE_HR="$(( FSIZE / 1073741824 )) GB"
        elif [[ "$FSIZE" -gt 1048576 ]] 2>/dev/null; then
            SIZE_HR="$(( FSIZE / 1048576 )) MB"
        elif [[ "$FSIZE" -gt 1024 ]] 2>/dev/null; then
            SIZE_HR="$(( FSIZE / 1024 )) KB"
        else
            SIZE_HR="${FSIZE} B"
        fi
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono dim' style='font-size:.68rem;white-space:nowrap'>$(html_esc "$SID")</td>
          <td class='mono ${PATH_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$ORIG_PATH")</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${DEL_TIME:--}</td>
          <td class='mono mid' style='white-space:nowrap;font-size:.72rem'>${SIZE_HR}</td>
        </tr>"
    done

    {
        html_header "Recycle Bin"
        html_page_header "RB" "Recycle Bin <span>Forensics</span>" \
            "\$Recycle.Bin\\\<SID>\\\$I*" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File eliminati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>$(L "File nel Cestino — SID · Path originale · Data eliminazione · Dimensione" "Recycle Bin Files — SID · Original path · Deletion date · Size")</div>
        <div class='card'><table>
          <thead><tr><th style='width:16%'>SID</th><th>$(L "Path originale" "Original path")</th><th style='width:14%'>$(L "Eliminato" "Deleted")</th><th style='width:8%'>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
