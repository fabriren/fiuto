
# ================================================================
#  MODULO 5 — Run Keys & Persistenza nel Registro
# ================================================================
module_run_keys() {
    section_header "$(L "Run Keys & Persistenza Registro" "Run Keys & Registry Persistence")" "$ORANGE"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE_SW
    HIVE_SW=$(get_hive "SOFTWARE")
    [[ -z "$HIVE_SW" ]] && { err "$(L "Hive SOFTWARE non trovato" "SOFTWARE hive not found")"; return 1; }

    declare -a ALL_ENTRIES=()

    # Chiavi da controllare su SOFTWARE (HKLM)
    local SW_KEYS=(
        "Microsoft\\Windows\\CurrentVersion\\Run"
        "Microsoft\\Windows\\CurrentVersion\\RunOnce"
        "Microsoft\\Windows\\CurrentVersion\\RunOnceEx"
        "Microsoft\\Windows NT\\CurrentVersion\\Winlogon"
        "Microsoft\\Windows NT\\CurrentVersion\\Windows"
        "Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run"
    )

    info "$(L "Scansione hive SOFTWARE (HKLM)..." "Scanning SOFTWARE hive (HKLM)...")"
    for KEY in "${SW_KEYS[@]}"; do
        local VALS
        VALS=$("$PY3" - "$HIVE_SW" "$KEY" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for v in key.get_values():
        print(f"{v.name}\t{v.value}")
except Exception as e:
    pass
PYEOF
)
        if [[ -n "$VALS" ]]; then
            while IFS=$'\t' read -r VNAME VVAL; do
                ALL_ENTRIES+=("HKLM\\${KEY}|${VNAME}|${VVAL}")
            done <<< "$VALS"
        fi
    done

    # Chiavi NTUSER.DAT per ogni utente (HKCU)
    local HKCU_KEYS=(
        "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        "Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce"
    )

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")
        [[ -z "$NTUSER" ]] && continue
        for KEY in "${HKCU_KEYS[@]}"; do
            local VALS
            VALS=$("$PY3" - "$NTUSER" "$KEY" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for v in key.get_values():
        print(f"{v.name}\t{v.value}")
except: pass
PYEOF
)
            if [[ -n "$VALS" ]]; then
                while IFS=$'\t' read -r VNAME VVAL; do
                    ALL_ENTRIES+=("HKCU[${USERNAME}]\\${KEY}|${VNAME}|${VVAL}")
                done <<< "$VALS"
            fi
        done
    done < <(get_user_homes)

    local TOTAL=${#ALL_ENTRIES[@]}
    echo ""
    printf "  ${BOLD}%-60s %-30s %s${RESET}\n" "Chiave" "Nome valore" "Dati"
    separator
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r HKEY VNAME VVAL <<< "$E"
        local VVAL_CLEAN="${VVAL}"
        local COLOR="$DIM"
        # Percorsi sospetti
        if echo "$VVAL" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|powershell\|cmd.exe\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin"; then
            COLOR="$RED"
        fi
        printf "  ${DIM}%-60s${RESET} %-30s ${COLOR}%s${RESET}\n" "$HKEY" "$VNAME" "$VVAL_CLEAN"
    done

    separator
    info "$(L "Voci totali trovate:" "Total entries found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { ok "$(L "Nessuna voce Run trovata." "No Run entries found.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "run_keys")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='|' read -r HKEY VNAME VVAL <<< "$E"
        local CSS=""
        echo "$VVAL" | grep -qi "temp\|appdata\\\\local\|programdata\|public\|downloads\|powershell\|cmd.exe\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin" && CSS="class='bad'"
        ROWS+="<tr><td class='mono dim'>$(html_esc "$HKEY")</td><td class='mono'>$(html_esc "$VNAME")</td>
               <td class='mono' ${CSS}>$(html_esc "$VVAL")</td></tr>"
    done

    {
        html_header "Run Keys"
        html_page_header "RK" "Run Keys <span>&amp; Persistenza</span>" \
            "HKLM/HKCU Run, RunOnce, Winlogon, AppInit_DLLs" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Voci trovate</div><div class='value'>${TOTAL}</div></div>
        </div>
        <main><div class='stitle'>Voci di autorun nel registro</div>
        <div class='card'><table>
          <thead><tr><th>$(L "Chiave" "Key")</th><th>$(L "Nome valore" "Value name")</th><th>$(L "Dati (comando)" "Data (command)")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
