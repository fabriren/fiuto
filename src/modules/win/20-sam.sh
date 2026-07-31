
# ================================================================
#  MODULO 20 — SAM: Hash password account locali
# ================================================================
module_sam() {
    section_header "$(L "SAM - Hash Account Locali" "SAM - Local Account Hashes")" "$RED"
    check_win_root || return 1

    local SAM_HIVE;  SAM_HIVE=$(get_hive "SAM")
    local SYS_HIVE;  SYS_HIVE=$(get_hive "SYSTEM")

    if [[ -z "$SAM_HIVE" || ! -f "$SAM_HIVE" ]]; then
        err "Hive SAM $(L "non trovato" "not found")"
        return 1
    fi
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found") (necessario per la SysKey)"
        return 1
    fi

    # Verifica impacket
    if ! "$PY3" -c "from impacket.examples.secretsdump import LocalOperations" 2>/dev/null; then
        warn "impacket $(L "non trovato" "not found")."
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install impacket"
        info "$(L "Estrazione parziale tramite regipy (senza hash decifrati)..." "Partial extraction via regipy (without decrypted hashes)...")"

        if ! check_regipy; then
            return 1
        fi
        # Fallback: solo lista account senza hash
        local ACC_DATA
        ACC_DATA=$("$PY3" - "$SAM_HIVE" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    try:
        users_key = hive.get_key('SAM\\Domains\\Account\\Users\\Names')
        for sk in users_key.iter_subkeys():
            print(f"{sk.name}\t(hash non disponibile senza impacket)")
    except Exception as e:
        print(f"ERROR\t{e}")
except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
        echo ""
        while IFS=$'\t' read -r NAME NOTE; do
            [[ "$NAME" == "ERROR" ]] && { err "$NOTE"; continue; }
            printf "  ${CYAN}%-25s${RESET}  ${DIM}%s${RESET}\n" "$NAME" "$NOTE"
        done <<< "$ACC_DATA"
        warn "$(L "Per gli hash installa impacket e rilancia il modulo." "For hashes install impacket and relaunch the module.")"
        return 0
    fi

    info "$(L "Estrazione hash con impacket secretsdump..." "Extracting hashes with impacket secretsdump...")"
    echo ""

    local TMP_DIR; TMP_DIR=$(mktemp -d /tmp/dfir_sam_XXXXXX)
    register_tmp "$TMP_DIR"
    cp "$SAM_HIVE" "$TMP_DIR/SAM" 2>/dev/null || { err "Copia hive fallita"; return 1; }
    cp "$SYS_HIVE" "$TMP_DIR/SYSTEM" 2>/dev/null || { err "Copia hive fallita"; return 1; }
    chmod 600 "$TMP_DIR/SAM" "$TMP_DIR/SYSTEM" 2>/dev/null || true

    local DUMP_OUT
    DUMP_OUT=$("$PY3" - "$TMP_DIR/SAM" "$TMP_DIR/SYSTEM" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from impacket.examples.secretsdump import LocalOperations, SAMHashes
    from impacket.smbconnection import SessionError

    sys_ops = LocalOperations(sys.argv[2])
    boot_key = sys_ops.getBootKey()

    sam_hashes = SAMHashes(sys.argv[1], boot_key, isRemote=False)
    sam_hashes.dump()
    sam_hashes.export()
except Exception as e:
    # Fallback: usa secretsdump direttamente
    import subprocess, os
    try:
        result = subprocess.run(
            [sys.executable, '-c',
             f'''
from impacket.examples.secretsdump import LocalOperations, SAMHashes
import sys
sys_ops = LocalOperations("{sys.argv[2]}")
boot_key = sys_ops.getBootKey()
sam = SAMHashes("{sys.argv[1]}", boot_key, isRemote=False)
sam.dump()
'''],
            capture_output=True, text=True, timeout=30
        )
        if result.stdout:
            print(result.stdout)
        else:
            print(f"ERROR\t{str(e)}")
    except:
        print(f"ERROR\t{str(e)}")
PYEOF
)

    declare -a ACC_ENTRIES=()
    local TOTAL=0

    local KNOWN_HASHES=(
        "31d6cfe0d16ae931b73c59d7e0c089c0"   # Password vuota
        "aad3b435b51404eeaad3b435b51404ee"   # LM vuoto
    )

    while IFS= read -r LINE; do
        [[ -z "$LINE" ]] && continue
        [[ "$LINE" == "ERROR"* ]] && { err "${LINE#ERROR    }"; continue; }
        # Formato impacket: username:RID:LMhash:NThash:::
        if echo "$LINE" | grep -qE '^[^:]+:[0-9]+:[a-fA-F0-9]{32}:[a-fA-F0-9]{32}'; then
            IFS=':' read -r UNAME RID LM NT REST <<< "$LINE"
            local IS_EMPTY=false IS_SUSP=false
            [[ "${NT,,}" == "31d6cfe0d16ae931b73c59d7e0c089c0" ]] && IS_EMPTY=true
            # RID 500=Admin, 501=Guest, altri potrebbero essere account nascosti
            [[ "$RID" -gt 1000 ]] 2>/dev/null && IS_SUSP=true
            printf "  ${CYAN}%-20s${RESET}  RID:${DIM}%-6s${RESET}  NT:${YELLOW}%s${RESET}" \
                "$UNAME" "$RID" "$NT"
            $IS_EMPTY && printf "  ${DIM}(password vuota)${RESET}"
            $IS_SUSP  && printf "  ${RED}[utente non default]${RESET}"
            echo ""
            ACC_ENTRIES+=("${UNAME}§${RID}§${LM}§${NT}§${IS_EMPTY}§${IS_SUSP}")
            TOTAL=$((TOTAL + 1))
        else
            echo -e "  ${DIM}$LINE${RESET}"
        fi
    done <<< "$DUMP_OUT"

    separator
    info "$(L "Account trovati:" "Accounts found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun account estratto." "No accounts extracted.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "sam")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ACC_ENTRIES[@]}"; do
        IFS='§' read -r UNAME RID LM NT IS_EMPTY IS_SUSP <<< "$E"
        local ROW_STYLE="" NT_CSS="mono"
        [[ "$IS_SUSP" == "true" ]] && ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        [[ "$IS_EMPTY" == "true" ]] && NT_CSS="mono warn"
        local NOTES=""
        [[ "$IS_EMPTY" == "true" ]] && NOTES+="<span class='dim' style='font-size:.68rem'>password vuota</span> "
        [[ "$IS_SUSP"  == "true" ]] && NOTES+="<span style='color:var(--accent2);font-size:.68rem'>account non default</span>"
        [[ "$RID" == "500" ]]       && NOTES+="<span style='color:var(--accent4);font-size:.68rem'>Administrator</span>"
        [[ "$RID" == "501" ]]       && NOTES+="<span class='dim' style='font-size:.68rem'>Guest</span>"
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono'>$(html_esc "$UNAME")</td>
          <td class='mono mid' style='white-space:nowrap'>${RID}</td>
          <td class='mono dim' style='font-size:.65rem;word-break:break-all'>$(html_esc "$LM")</td>
          <td class='${NT_CSS}' style='font-size:.72rem;word-break:break-all'>$(html_esc "$NT")</td>
          <td style='font-size:.72rem'>${NOTES}</td>
        </tr>"
    done

    {
        html_header "SAM"
        html_page_header "SM" "SAM - <span>Hash Account Locali</span>" \
            "Windows\\System32\\config\\SAM" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Account</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>Hash NTLM account locali - craccabili offline con hashcat/john</div>
        <div class='card'><table>
          <thead><tr><th>Username</th><th style='width:7%'>RID</th><th style='width:27%'>LM Hash</th><th style='width:27%'>NT Hash</th><th>Note</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Cracking offline</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            hashcat -m 1000 hashes.txt wordlist.txt<br>
            john --format=NT hashes.txt
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
