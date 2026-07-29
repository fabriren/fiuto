
# ================================================================
#  MODULO 22 — OpenSaveMRU / LastVisitedMRU (dialoghi file)
# ================================================================
module_opensave() {
    section_header "$(L "OpenSaveMRU / LastVisitedMRU — Dialoghi File" "OpenSaveMRU / LastVisitedMRU — File Dialogs")" "$GREEN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    declare -a OS_ENTRIES=()   # OpenSaveMRU
    declare -a LV_ENTRIES=()   # LastVisitedMRU
    local TOTAL_OS=0 TOTAL_LV=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue

        # ── OpenSavePidlMRU / OpenSaveMRU ───────────────────────────
        local OS_DATA
        OS_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    # Prova entrambe le chiavi (Win7 usa PidlMRU, versioni precedenti MRU)
    for key_path in [
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\OpenSavePidlMRU',
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\OpenSaveMRU',
    ]:
        try:
            root = hive.get_key(key_path)
            for ext_key in root.iter_subkeys():
                ext = ext_key.name  # es. "exe", "ps1", "*"
                for v in ext_key.get_values():
                    if v.name.lower() in ('mrulistex','mrulist'):
                        continue
                    val = v.value
                    # PidlMRU: valore binario con path embedded
                    if isinstance(val, bytes):
                        # Cerca stringa UTF-16 nel blob
                        try:
                            decoded = val.decode('utf-16-le', errors='replace').rstrip('\x00')
                            # Prendi solo la parte che sembra un path
                            parts = [p for p in decoded.split('\x00') if '\\' in p and len(p) > 3]
                            for p in parts[:1]:
                                print(f"{user}\t{ext}\t{p.strip()}")
                        except:
                            pass
                    elif isinstance(val, str) and '\\' in val:
                        print(f"{user}\t{ext}\t{val}")
        except:
            continue
except:
    pass
PYEOF
)
        while IFS=$'\t' read -r U EXT PATH_V; do
            [[ -z "$PATH_V" ]] && continue
            OS_ENTRIES+=("${U}§${EXT}§${PATH_V}")
            TOTAL_OS=$((TOTAL_OS + 1))
        done <<< "$OS_DATA"

        # ── LastVisitedPidlMRU / LastVisitedMRU ─────────────────────
        local LV_DATA
        LV_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    for key_path in [
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\LastVisitedPidlMRU',
        'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\LastVisitedMRU',
    ]:
        try:
            key = hive.get_key(key_path)
            for v in key.get_values():
                if v.name.lower() in ('mrulistex','mrulist'):
                    continue
                val = v.value
                if isinstance(val, bytes):
                    try:
                        # Prima stringa terminata da null = nome app
                        decoded = val.decode('utf-16-le', errors='replace')
                        parts = decoded.split('\x00')
                        app = parts[0].strip() if parts else ''
                        # Cerca path nelle parti successive
                        path = next((p for p in parts[1:] if '\\' in p and len(p) > 3), '')
                        if app:
                            print(f"{user}\t{app}\t{path}")
                    except:
                        pass
                elif isinstance(val, str):
                    print(f"{user}\t{val}\t")
        except:
            continue
except:
    pass
PYEOF
)
        while IFS=$'\t' read -r U APP PATH_V; do
            [[ -z "$APP" ]] && continue
            LV_ENTRIES+=("${U}§${APP}§${PATH_V}")
            TOTAL_LV=$((TOTAL_LV + 1))
        done <<< "$LV_DATA"

    done < <(get_user_homes)

    # Stampa terminale
    echo ""
    echo -e "  ${GREEN}${BOLD}OpenSaveMRU${RESET} (${BOLD}$TOTAL_OS${RESET} file aperti/salvati)"
    for E in "${OS_ENTRIES[@]}"; do
        IFS='§' read -r U EXT P <<< "$E"
        local COLOR="$DIM"
        echo "$P" | grep -qi "\.ps1\|\.bat\|\.exe\|\.vbs\|\.hta\|temp\|public" && COLOR="$RED"
        printf "  [${DIM}%-12s${RESET}]  ${CYAN}%-8s${RESET}  ${COLOR}%s${RESET}\n" "$U" ".$EXT" "$P"
    done | head -30

    echo ""
    echo -e "  ${GREEN}${BOLD}LastVisitedMRU${RESET} (${BOLD}$TOTAL_LV${RESET} applicazioni/cartelle)"
    for E in "${LV_ENTRIES[@]}"; do
        IFS='§' read -r U APP P <<< "$E"
        printf "  [${DIM}%-12s${RESET}]  ${YELLOW}%-25s${RESET}  ${DIM}%s${RESET}\n" "$U" "$APP" "$P"
    done | head -20

    separator
    info "OpenSave: ${BOLD}$TOTAL_OS${RESET}  |  LastVisited: ${BOLD}$TOTAL_LV"
    local GRAND=$((TOTAL_OS + TOTAL_LV))
    [[ $GRAND -eq 0 ]] && { warn "$(L "Nessun dato trovato." "No data found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "opensave")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local OS_ROWS=""
    for E in "${OS_ENTRIES[@]}"; do
        IFS='§' read -r U EXT P <<< "$E"
        local SUSP_CSS="" ROW_STYLE=""
        echo "$P" | grep -qi "\.ps1\|\.bat\|\.exe\|\.vbs\|\.hta\|temp\|public\|programdata" && {
            SUSP_CSS="bad"
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        }
        OS_ROWS+="<tr ${ROW_STYLE}>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='white-space:nowrap;color:var(--accent4)'>.$(html_esc "$EXT")</td>
          <td class='mono ${SUSP_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
        </tr>"
    done

    local LV_ROWS=""
    for E in "${LV_ENTRIES[@]}"; do
        IFS='§' read -r U APP P <<< "$E"
        LV_ROWS+="<tr>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='color:var(--accent4);white-space:nowrap'>$(html_esc "$APP")</td>
          <td class='mono mid' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
        </tr>"
    done

    {
        html_header "OpenSaveMRU"
        html_page_header "OS" "OpenSave / <span>LastVisited</span> MRU" \
            "NTUSER.DAT → ComDlg32" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File aperti/salvati</div><div class='value'>${TOTAL_OS}</div></div>
          <div class='stat info'><div class='label'>App + cartelle</div><div class='value'>${TOTAL_LV}</div></div>
        </div><main>
        <div class='stitle'>OpenSaveMRU — File aperti o salvati tramite dialogo di Windows</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:8%'>$(L "Estensione" "Extension")</th><th>Path</th></tr></thead>
          <tbody>${OS_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>LastVisitedMRU — Ultima cartella visitata per applicazione</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:20%'>Applicazione</th><th>$(L "Ultima cartella" "Last folder")</th></tr></thead>
          <tbody>${LV_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
