
# ================================================================
#  MODULO 33 — Defender Quarantine
# ================================================================
module_defender_quarantine() {
    section_header "$(L "Defender Quarantine — File in Quarantena" "Defender Quarantine — Quarantined Files")" "$RED"
    check_win_root || return 1


    local QENTRIES_DIR
    QENTRIES_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Windows Defender/Quarantine/Entries")
    local QRESOURCE_DIR
    QRESOURCE_DIR=$(ci_find_dir "$WIN_ROOT" "ProgramData/Microsoft/Windows Defender/Quarantine/ResourceData")

    if [[ -z "$QENTRIES_DIR" || ! -d "$QENTRIES_DIR" ]]; then
        warn "$(L "Directory quarantena Defender non trovata" "Defender quarantine directory not found")"
        info "$(L "Percorso atteso: ProgramData/Microsoft/Windows Defender/Quarantine/" "Expected path: ProgramData/Microsoft/Windows Defender/Quarantine/")"
        return 0
    fi

    info "Quarantine Entries: $QENTRIES_DIR"
    [[ -n "$QRESOURCE_DIR" ]] && info "ResourceData:       $QRESOURCE_DIR"
    echo ""

    declare -a ALL_ENTRIES=()
    local TOTAL=0

    while IFS= read -r EFILE; do
        [[ -f "$EFILE" ]] || continue
        local FNAME; FNAME=$(basename "$EFILE")
        local FSIZE; FSIZE=$(stat -c "%s" "$EFILE" 2>/dev/null || echo "0")
        local FMTIME; FMTIME=$(stat -c "%y" "$EFILE" 2>/dev/null | cut -d'.' -f1 || echo "")

        local ENTRY_INFO
        ENTRY_INFO=$("$PY3" - "$EFILE" << 'PYEOF' 2>/dev/null || true
import sys, struct, binascii, re

path = sys.argv[1]
try:
    with open(path,'rb') as f:
        data = f.read()
    # Cerca stringhe UTF-16LE leggibili: path file, nome minaccia, SHA256-like
    strings_utf16 = data.decode('utf-16-le', errors='replace').replace('\x00','')
    strings_ascii = data.decode('latin-1', errors='replace')

    # Cerca path Windows
    paths = re.findall(r'[A-Za-z]:\\[^\x00\n\r"<>|]{5,260}', strings_utf16)
    # Cerca nome minaccia (pattern Defender: Trojan:Win32/...)
    threats = re.findall(r'(?:Trojan|Ransom|Worm|Backdoor|Exploit|HackTool|Riskware|Generic|Behavior)[:\.\w\/\\]+', strings_utf16)
    # Cerca SHA256 hex
    hashes = re.findall(r'[0-9a-fA-F]{64}', strings_ascii)

    orig_path = paths[0] if paths else ''
    threat    = threats[0] if threats else ''
    sha256    = hashes[0] if hashes else ''
    print(f"{orig_path}\t{threat}\t{sha256}")
except Exception as e:
    print(f"\t\t")
PYEOF
)
        local ORIG_PATH THREAT HASH
        IFS=$'\t' read -r ORIG_PATH THREAT HASH <<< "$ENTRY_INFO"

        printf "  ${RED}%-45s${RESET}  ${YELLOW}%s${RESET}\n" "${THREAT:-sconosciuta}" "${ORIG_PATH:0:60}"
        [[ -n "$HASH" ]] && printf "  ${DIM}SHA256: %s${RESET}\n" "$HASH"
        ALL_ENTRIES+=("${FNAME}§${ORIG_PATH}§${THREAT}§${HASH}§${FSIZE}§${FMTIME}")
        TOTAL=$((TOTAL + 1))
    done < <(find "$QENTRIES_DIR" -maxdepth 2 -type f 2>/dev/null | sort)

    separator
    info "$(L "Voci in quarantena trovate:" "Quarantine entries found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna voce trovata nella quarantena." "No entries found in quarantine.")"; return 0; }

    if [[ -n "$QRESOURCE_DIR" ]]; then
        info "${BOLD}$(L "Payload cifrati disponibili in:" "Encrypted payloads available in:")${RESET} $QRESOURCE_DIR"
        info "Per ripristino (su VM isolata): MpCmdRun.exe -Restore -Name <ThreatName>"
    fi

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "defender_quarantine")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    # Ordina per timestamp decrescente (campo 6: FMTIME)
    mapfile -t ALL_ENTRIES < <(printf '%s\n' "${ALL_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+6*[""])[5],reverse=True); print("\n".join(L))')
    local ROWS=""
    for E in "${ALL_ENTRIES[@]}"; do
        IFS='§' read -r FNAME ORIG_PATH THREAT HASH FSIZE FMTIME <<< "$E"
        ROWS+="<tr style='background:rgba(255,123,114,.05);border-left:3px solid var(--accent2)'>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${FMTIME}</td>
          <td class='mono bad' style='word-break:break-all;font-size:.72rem'>$(html_esc "${THREAT:-Sconosciuta}")</td>
          <td class='mono' style='word-break:break-all;font-size:.72rem'>$(html_esc "$ORIG_PATH")</td>
          <td class='mono dim' style='word-break:break-all;font-size:.65rem'>$(html_esc "${HASH:0:64}")</td>
          <td class='mono mid' style='white-space:nowrap'>${FSIZE}B</td>
        </tr>"
    done
    {
        html_header "Defender Quarantine"
        html_page_header "DQ" "Defender — <span>Quarantine</span>" \
            "ProgramData\\Microsoft\\Windows Defender\\Quarantine" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>In quarantena</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>$(L "File in quarantena — recuperabili offline da ResourceData/ (cifrati XOR)" "Quarantined files — recoverable offline from ResourceData/ (XOR-encrypted)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:14%'>$(L "Rilevato" "Detected")</th><th style='width:22%'>$(L "Minaccia" "Threat")</th><th>$(L "Path originale" "Original path")</th><th style='width:18%'>SHA256</th><th style='width:6%'>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Analisi payload (ambiente isolato)</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            I file in ResourceData/ sono cifrati con RC4 (key: hardcoded Defender) — tool: defender-un-quarantine.py<br>
            Ripristino: MpCmdRun.exe -Restore -FilePath &lt;path&gt;
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
