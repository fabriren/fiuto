
# ================================================================
#  MODULO 18 — UserAssist + RunMRU + TypedPaths + WordWheelQuery
# ================================================================
module_userassist() {
    section_header "UserAssist / RunMRU / TypedPaths - Attività Utente" "$MAGENTA"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    local TOTAL_UA=0 TOTAL_RUN=0 TOTAL_TP=0 TOTAL_WW=0
    declare -a UA_ENTRIES=()   # UserAssist
    declare -a RUN_ENTRIES=()  # RunMRU
    declare -a TP_ENTRIES=()   # TypedPaths
    declare -a WW_ENTRIES=()   # WordWheelQuery

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && { dim_msg "$USERNAME - NTUSER.DAT $(L "non trovato" "not found")"; continue; }

        info "Parsing NTUSER.DAT: $USERNAME"

        # ── UserAssist ───────────────────────────────────────────────
        local UA_DATA
        UA_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, json, codecs, struct, datetime
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    UA_KEY = 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\UserAssist'
    try:
        ua_root = hive.get_key(UA_KEY)
    except:
        sys.exit(0)
    for guid_key in ua_root.iter_subkeys():
        try:
            count_key = hive.get_key(f'{UA_KEY}\\{guid_key.name}\\Count')
        except:
            continue
        for v in count_key.get_values():
            # Decodifica ROT13 del nome
            name = codecs.decode(v.name, 'rot_13')
            val  = v.value
            ts   = ''
            count = 0
            if isinstance(val, bytes) and len(val) >= 16:
                try:
                    count = struct.unpack_from('<I', val, 4)[0]
                    ft    = struct.unpack_from('<Q', val, 8)[0]
                    if ft > 0:
                        ts = (datetime.datetime(1601,1,1) +
                              datetime.timedelta(microseconds=ft//10)).strftime('%Y-%m-%d %H:%M:%S')
                except: pass
            if name and '{' not in name and count > 0:
                print(f"{user}\t{name}\t{count}\t{ts}")
except Exception as e:
    pass
PYEOF
)
        while IFS=$'\t' read -r U N C T; do
            [[ -z "$N" ]] && continue
            UA_ENTRIES+=("${U}§${N}§${C}§${T}")
            TOTAL_UA=$((TOTAL_UA + 1))
        done <<< "$UA_DATA"

        # ── RunMRU ───────────────────────────────────────────────────
        local RUN_DATA
        RUN_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    try:
        key = hive.get_key('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU')
        mru_order = ''
        entries = {}
        for v in key.get_values():
            if v.name.lower() == 'mrulist':
                mru_order = str(v.value)
            else:
                entries[v.name] = str(v.value).rstrip('\\1')
        # Stampa in ordine MRU
        for ch in mru_order:
            if ch in entries:
                print(f"{user}\t{entries[ch]}")
    except: pass
except: pass
PYEOF
)
        while IFS=$'\t' read -r U CMD; do
            [[ -z "$CMD" ]] && continue
            RUN_ENTRIES+=("${U}§${CMD}")
            TOTAL_RUN=$((TOTAL_RUN + 1))
        done <<< "$RUN_DATA"

        # ── TypedPaths ───────────────────────────────────────────────
        local TP_DATA
        TP_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    try:
        key = hive.get_key('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\TypedPaths')
        for v in key.get_values():
            print(f"{user}\t{v.name}\t{v.value}")
    except: pass
except: pass
PYEOF
)
        while IFS=$'\t' read -r U K V; do
            [[ -z "$V" ]] && continue
            TP_ENTRIES+=("${U}§${K}§${V}")
            TOTAL_TP=$((TOTAL_TP + 1))
        done <<< "$TP_DATA"

        # ── WordWheelQuery (ricerche Start) ──────────────────────────
        local WW_DATA
        WW_DATA=$("$PY3" - "$NTUSER" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    user = sys.argv[2]
    try:
        key = hive.get_key('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\WordWheelQuery')
        mru_order = ''
        entries = {}
        for v in key.get_values():
            if v.name.lower() == 'mrulistex':
                continue
            val = v.value
            if isinstance(val, bytes):
                try:
                    val = val.decode('utf-16-le', errors='replace').rstrip('\x00')
                except:
                    val = repr(val)
            entries[v.name] = str(val)
        for k,v in sorted(entries.items()):
            if v:
                print(f"{user}\t{v}")
    except: pass
except: pass
PYEOF
)
        while IFS=$'\t' read -r U TERM; do
            [[ -z "$TERM" ]] && continue
            WW_ENTRIES+=("${U}§${TERM}")
            TOTAL_WW=$((TOTAL_WW + 1))
        done <<< "$WW_DATA"

    done < <(get_user_homes)

    # ── Stampa terminale ─────────────────────────────────────────────
    echo ""
    echo -e "  ${MAGENTA}${BOLD}UserAssist${RESET} (${BOLD}$TOTAL_UA${RESET} voci)"
    for E in "${UA_ENTRIES[@]}"; do
        IFS='§' read -r U N C T <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  cnt=${CYAN}%-3s${RESET}  ${DIM}%s${RESET}  %s\n" \
            "$U" "$C" "$T" "${N: -70}"
    done | head -30
    [[ $TOTAL_UA -gt 30 ]] && echo -e "  ${DIM}... e altri $((TOTAL_UA-30)) (vedi report HTML)${RESET}"

    echo ""
    echo -e "  ${MAGENTA}${BOLD}RunMRU${RESET} (${BOLD}$TOTAL_RUN${RESET} voci)"
    for E in "${RUN_ENTRIES[@]}"; do
        IFS='§' read -r U CMD <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  %s\n" "$U" "$CMD"
    done

    echo ""
    echo -e "  ${MAGENTA}${BOLD}TypedPaths${RESET} (${BOLD}$TOTAL_TP${RESET} voci)"
    for E in "${TP_ENTRIES[@]}"; do
        IFS='§' read -r U K V <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  %s\n" "$U" "$V"
    done

    echo ""
    echo -e "  ${MAGENTA}${BOLD}WordWheelQuery${RESET} (${BOLD}$TOTAL_WW${RESET} ricerche Start)"
    for E in "${WW_ENTRIES[@]}"; do
        IFS='§' read -r U TERM <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  ${CYAN}%s${RESET}\n" "$U" "$TERM"
    done

    separator
    info "UserAssist: ${BOLD}$TOTAL_UA${RESET}  RunMRU: ${BOLD}$TOTAL_RUN${RESET}  TypedPaths: ${BOLD}$TOTAL_TP${RESET}  WordWheel: ${BOLD}$TOTAL_WW"
    local GRAND=$((TOTAL_UA + TOTAL_RUN + TOTAL_TP + TOTAL_WW))
    [[ $GRAND -eq 0 ]] && { warn "$(L "Nessuna voce trovata." "No entries found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "userassist")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina UserAssist per timestamp decrescente
    mapfile -t UA_ENTRIES < <(printf '%s\n' "${UA_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+4*[""])[3],reverse=True); print("\n".join(L))')

    # UserAssist rows
    local UA_ROWS=""
    for E in "${UA_ENTRIES[@]}"; do
        IFS='§' read -r U N C T <<< "$E"
        local SUSP_CSS=""
        echo "$N" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|\\\\temp\\\\" && SUSP_CSS="class='bad'"
        UA_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='word-break:break-all;font-size:.72rem'><span ${SUSP_CSS}>$(html_esc "$N")</span></td>
          <td class='mono ok' style='white-space:nowrap;text-align:center'>${C}</td>
          <td class='mono mid' style='white-space:nowrap'>${T:--}</td></tr>"
    done

    # RunMRU rows
    local RUN_ROWS=""
    for E in "${RUN_ENTRIES[@]}"; do
        IFS='§' read -r U CMD <<< "$E"
        local SUSP_CSS=""
        echo "$CMD" | grep -qi "powershell\|cmd\|wscript\|mshta\|certutil\|bitsadmin\|-enc\|base64" && SUSP_CSS="class='bad'"
        RUN_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='word-break:break-all'><span ${SUSP_CSS}>$(html_esc "$CMD")</span></td></tr>"
    done

    # TypedPaths rows
    local TP_ROWS=""
    for E in "${TP_ENTRIES[@]}"; do
        IFS='§' read -r U K V <<< "$E"
        TP_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono mid' style='white-space:nowrap'>$(html_esc "$K")</td>
          <td class='mono' style='word-break:break-all'>$(html_esc "$V")</td></tr>"
    done

    # WordWheelQuery rows
    local WW_ROWS=""
    for E in "${WW_ENTRIES[@]}"; do
        IFS='§' read -r U TERM <<< "$E"
        WW_ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono' style='color:var(--accent)'>$(html_esc "$TERM")</td></tr>"
    done

    {
        html_header "UserAssist"
        html_page_header "UA" "UserAssist / RunMRU / <span>TypedPaths</span>" \
            "NTUSER.DAT - Attività interattiva utente" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>UserAssist</div><div class='value'>${TOTAL_UA}</div></div>
          <div class='stat info'><div class='label'>RunMRU</div><div class='value'>${TOTAL_RUN}</div></div>
          <div class='stat ok'><div class='label'>TypedPaths</div><div class='value'>${TOTAL_TP}</div></div>
          <div class='stat'><div class='label'>WordWheel</div><div class='value' style='color:var(--accent4)'>${TOTAL_WW}</div></div>
        </div><main>
        <div class='stitle'>UserAssist - Programmi avviati dalla GUI (ROT13 decodificato)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Eseguibile" "Executable")</th><th style='width:6%'>$(L "Avvii" "Runs")</th><th style='width:16%'>$(L "Ultimo uso" "Last used")</th></tr></thead>
          <tbody>${UA_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>RunMRU - Comandi digitati in Esegui (Win+R)</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Comando" "Command")</th></tr></thead>
          <tbody>${RUN_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>TypedPaths - Percorsi digitati in Explorer</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th style='width:8%'>Slot</th><th>$(L "Percorso" "Path")</th></tr></thead>
          <tbody>${TP_ROWS:-<tr><td colspan='3' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>WordWheelQuery - Ricerche in Start Menu</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Termine cercato" "Search term")</th></tr></thead>
          <tbody>${WW_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
