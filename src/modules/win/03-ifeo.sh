
# ================================================================
#  MODULO 2 — IFEO (Image File Execution Options)
# ================================================================
module_ifeo() {
    section_header "IFEO - Image File Execution Options" "$RED"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local HIVE
    HIVE=$(get_hive "SOFTWARE")
    if [[ -z "$HIVE" || ! -f "$HIVE" ]]; then
        err "$(L "Hive SOFTWARE non trovato in" "SOFTWARE hive not found in") $WIN_ROOT"
        return 1
    fi

    info "Parsing hive: $HIVE"

    # Chiave IFEO
    local IFEO_KEY="Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options"

    # Eseguibili sensibili tipicamente abusati
    local SENSITIVE_EXES="sethc.exe utilman.exe osk.exe magnify.exe narrator.exe displayswitch.exe atbroker.exe"

    declare -a ENTRIES_EXEC=()   # eseguibili con Debugger impostato
    declare -a ENTRIES_GFLAGS=() # GFlags sospetti

    mapfile -t IFEO_ENTRIES < <("$PY3" - "$HIVE" "$IFEO_KEY" << 'PYEOF'
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for subkey in key.iter_subkeys():
        name = subkey.name
        vals = {}
        for v in subkey.get_values():
            vals[v.name.lower()] = str(v.value)
        debugger = vals.get('debugger','')
        gflags   = vals.get('globalflag','') or vals.get('globalflags','')
        verifier = vals.get('verifierdlls','')
        if debugger or gflags or verifier:
            print(f"{name}\t{debugger}\t{gflags}\t{verifier}")
except Exception as e:
    print(f"ERROR\t{e}\t\t", file=sys.stderr)
PYEOF
)

    local FOUND_COUNT=0
    local SUSPICIOUS_COUNT=0

    echo ""
    printf "  %-35s %-45s %-15s %s\n" "Eseguibile" "Debugger" "GlobalFlag" "VerifierDLLs"
    separator

    for ENTRY in "${IFEO_ENTRIES[@]}"; do
        IFS=$'\t' read -r EXE DBG GFLAG VDLL <<< "$ENTRY"
        [[ "$EXE" == "ERROR" ]] && { err "$DBG"; continue; }
        FOUND_COUNT=$((FOUND_COUNT + 1))
        local IS_SENS=false
        for S in $SENSITIVE_EXES; do
            [[ "${EXE,,}" == "${S,,}" ]] && IS_SENS=true && break
        done
        if $IS_SENS || [[ -n "$DBG" ]]; then
            SUSPICIOUS_COUNT=$((SUSPICIOUS_COUNT + 1))
            printf "  ${RED}%-35s${RESET} ${YELLOW}%-45s${RESET} %-15s %s\n" "$EXE" "${DBG:-(nessuno)}" "${GFLAG:--}" "${VDLL:--}"
        else
            printf "  ${DIM}%-35s${RESET} %-45s %-15s %s\n" "$EXE" "${DBG:-(nessuno)}" "${GFLAG:--}" "${VDLL:--}"
        fi
        ENTRIES_EXEC+=("$EXE|${DBG}|${GFLAG}|${VDLL}")
    done

    separator
    info "$(L "Voci con Debugger/GFlags:" "Entries with Debugger/GFlags:") ${BOLD}$FOUND_COUNT${RESET}  |  Sospette: ${RED}${BOLD}$SUSPICIOUS_COUNT${RESET}"

    [[ $FOUND_COUNT -eq 0 ]] && { ok "$(L "Nessuna voce IFEO con Debugger/GFlags trovata." "No IFEO entries with Debugger/GFlags found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ifeo")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local ROWS=""
    for E in "${ENTRIES_EXEC[@]}"; do
        IFS='|' read -r EXE DBG GFLAG VDLL <<< "$E"
        local IS_SENS=false
        for S in $SENSITIVE_EXES; do
            [[ "${EXE,,}" == "${S,,}" ]] && IS_SENS=true && break
        done
        local ROW_CLASS=""
        $IS_SENS && ROW_CLASS="style='background:rgba(255,123,114,.07)'"
        local DBG_HTML="-"
        [[ -n "$DBG" ]] && DBG_HTML="<span class='bad'>$(html_esc "$DBG")</span>"
        ROWS+="<tr ${ROW_CLASS}>
          <td class='mono'>$(html_esc "$EXE")$(${IS_SENS} && echo " <span class='badge warn'>⚠ sensibile</span>" || true)</td>
          <td class='mono'>${DBG_HTML}</td>
          <td class='mono mid'>$([[ -n "$GFLAG" ]] && html_esc "$GFLAG" || echo "-")</td>
          <td class='mono dim'>$([[ -n "$VDLL" ]] && html_esc "$VDLL" || echo "-")</td>
        </tr>"
    done

    {
        html_header "IFEO"
        html_page_header "IE" "IFEO <span>Hijacking</span> Forensics" \
            "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Voci con Debugger/Flags</div><div class='value'>${FOUND_COUNT}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSPICIOUS_COUNT}</div></div>
        </div>
        <main>
        <div class='stitle'>Voci IFEO con Debugger o GlobalFlags impostati</div>
        <div class='card'><table>
          <thead><tr><th>$(L "Eseguibile" "Executable")</th><th>Debugger</th><th>GlobalFlag</th><th>VerifierDLLs</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.82rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>ℹ Eseguibili sensibili monitorati</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>${SENSITIVE_EXES// / &nbsp;·&nbsp; }</div>
          <div style='margin-top:.8rem;font-size:.78rem;color:var(--text-mid)'>Questi eseguibili sono accessibili dalla schermata di login (prima dell&apos;autenticazione). Un Debugger su di essi consente shell SYSTEM senza credenziali. Ref: MITRE ATT&amp;CK T1546.012</div>
        </div>
        </main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
