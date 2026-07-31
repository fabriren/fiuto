
# ================================================================
#  MODULO 24 — NTDS.dit (Active Directory hashes)
# ================================================================
module_ntds() {
    section_header "NTDS.dit - Active Directory Hashes" "$RED"
    check_win_root || return 1


    local NTDS_FILE
    NTDS_FILE=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f \
        -not -path "*/WinSxS/*" -not -path "*/WinSXS/*" \
        2>/dev/null | head -1)
    # Fallback: try canonical path via ci_find_dir if find returned nothing
    if [[ -z "$NTDS_FILE" ]]; then
        local _NTDS_DIR; _NTDS_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/NTDS")
        [[ -n "$_NTDS_DIR" ]] && NTDS_FILE=$(find "$_NTDS_DIR" -maxdepth 1 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    fi
    if [[ -z "$NTDS_FILE" ]]; then
        warn "$(L "ntds.dit non trovato - questo non è un Domain Controller (o la struttura NTDS/ è assente)" "ntds.dit not found - this is not a Domain Controller (or NTDS/ structure is absent)")"
        return 0
    fi
    local SYS_HIVE; SYS_HIVE=$(get_hive "SYSTEM")
    if [[ -z "$SYS_HIVE" || ! -f "$SYS_HIVE" ]]; then
        err "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found") (necessario per il BootKey)"; return 1
    fi
    info "NTDS.dit: $NTDS_FILE"
    info "SYSTEM:   $SYS_HIVE"
    echo ""

    if ! "$PY3" -c "from impacket.examples.secretsdump import NTDSHashes" 2>/dev/null; then
        warn "$(L "impacket non trovato - installalo con:" "impacket not found - install it with:") ${PY3} -m pip install impacket"
        info "$(L "Fallback: estrazione account senza hash tramite strings..." "Fallback: account extraction without hashes via strings...")"
        local ACC_STRINGS
        ACC_STRINGS=$(strings "$NTDS_FILE" 2>/dev/null | grep -E '^[A-Za-z0-9_.-]{3,20}\$?$' | sort -u | head -100 || true)
        echo "$ACC_STRINGS" | while IFS= read -r L; do
            printf "  ${DIM}%s${RESET}\n" "$L"
        done
        warn "$(L "Installa impacket per l'estrazione degli hash NTLM." "Install impacket for NTLM hash extraction.")"
        return 0
    fi

    info "$(L "Estrazione hash con impacket NTDSHashes..." "Extracting hashes with impacket NTDSHashes...")"
    local TMP_DIR; TMP_DIR=$(mktemp -d /tmp/dfir_ntds_XXXXXX)
    register_tmp "$TMP_DIR"

    # Copia con timeout: su share SMB da DC live la copia può bloccarsi indefinitamente
    info "$(L "Copia ntds.dit in /tmp (timeout 180s - file di rete, attendere)..." "Copying ntds.dit to /tmp (timeout 180s - network file, please wait)...")"
    if ! portable_timeout 180 cp "$NTDS_FILE" "$TMP_DIR/ntds.dit" 2>/dev/null; then
        err "$(L "Timeout o errore nella copia di ntds.dit (share lenta o file bloccato da Windows)" "Timeout or error copying ntds.dit (slow share or file locked by Windows)")"
        info "$(L "Suggerimento: esegui il modulo su un'immagine montata offline oppure tramite VSS snapshot" "Hint: run the module on an offline mounted image or via VSS snapshot")"
        return 1
    fi
    info "$(L "Copia SYSTEM hive..." "Copying SYSTEM hive...")"
    if ! portable_timeout 60 cp "$SYS_HIVE" "$TMP_DIR/SYSTEM" 2>/dev/null; then
        err "$(L "Timeout o errore nella copia del SYSTEM hive" "Timeout or error copying SYSTEM hive")"
        return 1
    fi
    chmod 600 "$TMP_DIR/ntds.dit" "$TMP_DIR/SYSTEM" 2>/dev/null || true

    # Patch dirty state ESE: NTDSHashes usa ESENT_DB internamente, che rifiuta
    # database in stato DirtyShutdown (3) — condizione normale su DC live.
    # Offset 52 (0x34): campo dbstate nello standard ESE/JET Blue.
    info "$(L "Verifica/patch dirty state ESE (necessario per DC live)..." "Checking/patching ESE dirty state (required for live DC)...")"
    "$PY3" - "$TMP_DIR/ntds.dit" << 'PYEOF_PATCH' 2>/dev/null || true
import sys, struct
path = sys.argv[1]
try:
    with open(path, 'r+b') as f:
        f.seek(4)
        if struct.unpack('<I', f.read(4))[0] == 0x89ABCDEF:
            f.seek(52)
            state = struct.unpack('<I', f.read(4))[0]
            if state in (3, 4, 5):
                f.seek(52)
                f.write(struct.pack('<I', 2))
                sys.stderr.write(f"INFO: dirty state {state}→2 (CleanShutdown) patchato\n")
except Exception as ex:
    sys.stderr.write(f"WARN: patch dirty state fallita: {ex}\n")
PYEOF_PATCH

    info "$(L "Dump hash NTLM (può richiedere diversi minuti per domini grandi)..." "Dumping NTLM hashes (may take several minutes for large domains)...")"
    local DUMP_OUT
    DUMP_OUT=$(portable_timeout 600 "$PY3" - "$TMP_DIR/ntds.dit" "$TMP_DIR/SYSTEM" << 'PYEOF' 2>/dev/null || true
import sys, io
try:
    from impacket.examples.secretsdump import LocalOperations, NTDSHashes
    sys_ops = LocalOperations(sys.argv[2])
    boot_key = sys_ops.getBootKey()
    class _CB:
        def __init__(self): self.lines = []
        def __call__(self, secret): self.lines.append(secret); print(secret)
    cb = _CB()
    ntds = NTDSHashes(sys.argv[1], boot_key, isRemote=False, history=False,
                      noLMHash=True, remoteOps=None, useVSSMethod=False,
                      justNTLM=True, pwdLastSet=False, resumeSession=None,
                      outputFileName=None, justUser=None, ldapFilter=None,
                      printUserStatus=True)
    ntds.dump()
    ntds.finish()
except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
    if [[ -z "$DUMP_OUT" ]]; then
        if grep -qiE 'Errno 35|EAGAIN|temporarily unavailable' "$TMP_DIR/ntds_err.log" 2>/dev/null \
        || [[ "$NTDS_FILE" == /Volumes/* || "$NTDS_FILE" == /mnt/* || "$NTDS_FILE" == /media/* ]]; then
            err "$(L "ntds.dit è bloccato dal servizio Active Directory (DC live, Errno 35 / EAGAIN)." "ntds.dit is locked by Active Directory service (live DC, Errno 35 / EAGAIN).")"
            info "$(L "Soluzioni: 1) ntdsutil IFM sul DC  2) impacket-secretsdump remoto  3) immagine disco offline" "Solutions: 1) ntdsutil IFM on DC  2) remote impacket-secretsdump  3) offline disk image")"
        else
            err "$(L "Dump vuoto - ntds.dit potrebbe essere in stato dirty o corrotto." "Empty dump - ntds.dit may be dirty or corrupted.")"
            info "$(L "Prova il modulo 38 (PAD Offline) che usa analisi ESE con recovery avanzato." "Try module 38 (PAD Offline) which uses ESE analysis with advanced recovery.")"
        fi
        return 1
    fi

    declare -a ACC_ENTRIES=()
    local TOTAL=0
    local KNOWN_EMPTY="31d6cfe0d16ae931b73c59d7e0c089c0"

    while IFS= read -r LINE; do
        [[ -z "$LINE" ]] && continue
        [[ "$LINE" == ERROR* ]] && { err "${LINE#ERROR	}"; continue; }
        if echo "$LINE" | grep -qE '^[^:]+:[0-9]+:[a-fA-F0-9]{32}:[a-fA-F0-9]{32}'; then
            IFS=':' read -r UNAME RID LM NT REST <<< "$LINE"
            local IS_EMPTY=false; [[ "${NT,,}" == "$KNOWN_EMPTY" ]] && IS_EMPTY=true
            local IS_DISABLED=false; echo "$LINE" | grep -qi "disabled\|account disabled\|(disabled)" && IS_DISABLED=true
            printf "  ${CYAN}%-30s${RESET}  RID:${DIM}%-6s${RESET}  NT:${YELLOW}%s${RESET}" "$UNAME" "$RID" "$NT"
            $IS_EMPTY    && printf "  ${DIM}(password vuota)${RESET}"
            $IS_DISABLED && printf "  ${RED}[DISABLED]${RESET}"
            echo ""
            ACC_ENTRIES+=("${UNAME}§${RID}§${NT}§${IS_EMPTY}§${IS_DISABLED}")
            TOTAL=$((TOTAL + 1))
        else
            echo -e "  ${DIM}$LINE${RESET}"
        fi
    done <<< "$DUMP_OUT"

    separator
    info "$(L "Account AD estratti:" "AD accounts extracted:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun account estratto." "No accounts extracted.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ntds")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ACC_ENTRIES[@]}"; do
        IFS='§' read -r UNAME RID NT IS_EMPTY IS_DISABLED <<< "$E"
        local ROW_STYLE="" NT_CSS="mono"
        [[ "$IS_DISABLED" == "true" ]] && ROW_STYLE="style='opacity:.5'"
        [[ "$IS_EMPTY"    == "true" ]] && NT_CSS="mono warn"
        local NOTES=""
        [[ "$IS_EMPTY"    == "true" ]] && NOTES+="<span class='dim' style='font-size:.68rem'>password vuota</span> "
        [[ "$IS_DISABLED" == "true" ]] && NOTES+="<span style='color:var(--text-dim);font-size:.68rem'>DISABLED</span>"
        [[ "$RID" == "500" ]]          && NOTES+="<span style='color:var(--accent2);font-size:.68rem'>Administrator</span>"
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono'>$(html_esc "$UNAME")</td>
          <td class='mono mid' style='white-space:nowrap'>${RID}</td>
          <td class='${NT_CSS}' style='font-size:.72rem;word-break:break-all'>$(html_esc "$NT")</td>
          <td style='font-size:.72rem'>${NOTES}</td>
        </tr>"
    done
    {
        html_header "NTDS.dit"
        html_page_header "AD" "NTDS.dit - <span>Active Directory</span> Hashes" \
            "Windows\\NTDS\\ntds.dit" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Account</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>Hash NTLM account AD - craccabili offline con hashcat -m 1000</div>
        <div class='card'><table>
          <thead><tr><th>Username</th><th style='width:8%'>RID</th><th style='width:30%'>NT Hash</th><th>Note</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Cracking offline</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            hashcat -m 1000 ntds_hashes.txt wordlist.txt<br>
            john --format=NT ntds_hashes.txt
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
