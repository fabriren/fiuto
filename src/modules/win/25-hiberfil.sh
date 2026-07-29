

# ================================================================
#  MODULO 25 — Hibernation / Pagefile (artefatti memoria)
# ================================================================
module_hiberfil() {
    section_header "$(L "Hibernation / Pagefile — Artefatti Memoria" "Hibernation / Pagefile — Memory Artifacts")" "$BLUE"
    check_win_root || return 1


    local HFILE PFILE
    HFILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "hiberfil.sys" -type f 2>/dev/null | head -1)
    PFILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "pagefile.sys" -type f 2>/dev/null | head -1)

    [[ -z "$HFILE" && -z "$PFILE" ]] && {
        warn "$(L "hiberfil.sys e pagefile.sys non trovati nella root del volume" "hiberfil.sys and pagefile.sys not found in volume root")"
        info "$(L "Il volume potrebbe essere montato senza visibilità sui file di sistema" "The volume may be mounted without visibility on system files")"
        return 0
    }

    declare -a FILE_ENTRIES=()
    local TOTAL_STRINGS=0

    for MFILE in "$HFILE" "$PFILE"; do
        [[ -z "$MFILE" || ! -f "$MFILE" ]] && continue
        local MNAME; MNAME=$(basename "$MFILE")
        local MSIZE; MSIZE=$(du -h "$MFILE" 2>/dev/null | cut -f1 || echo "?")
        ok "$MNAME $(L "trovato:" "found:") $MFILE  ($MSIZE)"

        # Verifica magic bytes per hiberfil.sys
        local MAGIC=""
        if [[ "$MNAME" =~ hiberfil ]]; then
            MAGIC=$(dd if="$MFILE" bs=4 count=1 2>/dev/null | od -An -tx1 | tr -d ' \n' | cut -c1-8 || echo "")
            case "${MAGIC^^}" in
                "5041474D") info "  $(L "Formato: HIBR (hibernated)" "Format: HIBR (hibernated)")" ;;
                "52535452") info "  $(L "Formato: RSTR (restored)" "Format: RSTR (restored)")" ;;
                *) info "  $(L "Magic: 0x${MAGIC} (sconosciuto)" "Magic: 0x${MAGIC} (unknown)")" ;;
            esac
        fi

        info "  $(L "Estrazione stringhe significative..." "Extracting significant strings...")"
        local STR_OUT
        STR_OUT=$(python3 - "$MFILE" << 'PYEOF' 2>/dev/null || true
import sys, re
path = sys.argv[1]
PATTERNS = [
    (r'https?://[a-zA-Z0-9_/\-\.\?&=%~+#]+', 'URL'),
    (r'\b(?:\d{1,3}\.){3}\d{1,3}\b', 'IP'),
    (r'[A-Za-z]:\\[a-zA-Z0-9_\\\-\.\s\(\)~!@#%&\+,;=\[\]\{\}\^]{5,200}', 'PATH'),
    (r'(?i)(?:password|passwd|pwd|secret|token|apikey|authorization)[=:\s]+[a-zA-Z0-9_!@#\$\^&\*\-\.\?]{4,50}', 'CRED'),
    (r'(?i)(?:cmd\.exe|powershell|wscript|mshta|certutil|bitsadmin)[ -~]{0,100}', 'EXEC'),
    (r'[A-Za-z0-9+/]{40,}={0,2}', 'B64_CANDIDATE'),
]
results = []
try:
    CHUNK = 1024*1024*4  # 4MB at a time
    with open(path,'rb') as f:
        while True:
            chunk = f.read(CHUNK)
            if not chunk: break
            text = chunk.decode('utf-16-le',errors='replace') + chunk.decode('latin-1',errors='replace')
            for pat, label in PATTERNS:
                for m in re.findall(pat, text):
                    v = m.strip(' "''<>\\r\\n')[:200]
                    # Filter out purely non-ascii / bad decoding artifacts
                    if len(v) > 6 and v not in results and v.isascii():
                        results.append(v)
                        print(f"{label}\t{v}")
            if len(results) > 500: break
except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
        local COUNT; COUNT=$(echo "$STR_OUT" | wc -l)
        TOTAL_STRINGS=$((TOTAL_STRINGS + COUNT))
        echo "$STR_OUT" | grep -v "^ERROR" | head -20 | while IFS=$'\t' read -r LBL VAL; do
            local COLOR="$DIM"
            [[ "$LBL" == "CRED" || "$LBL" == "URL" ]] && COLOR="$RED"
            printf "  ${CYAN}%-16s${RESET} ${COLOR}%s${RESET}\n" "$LBL" "${VAL:0:100}"
        done
        echo ""
        FILE_ENTRIES+=("${MNAME}§${MFILE}§${MSIZE}§${MAGIC}§$(echo "$STR_OUT" | grep -v "^ERROR" | head -100 | base64 -w0 2>/dev/null || true)")
    done

    separator
    info "$(L "Stringhe estratte totali:" "Total extracted strings:") ${BOLD}$TOTAL_STRINGS"
    info "${BOLD}$(L "Per analisi completa della memoria usa Volatility3:" "For complete memory analysis use Volatility3:")${RESET}"
    echo -e "    ${DIM}vol -f \"${HFILE:-<hiberfil.sys>}\" windows.pstree${RESET}"
    echo -e "    ${DIM}vol -f \"${HFILE:-<hiberfil.sys>}\" windows.netscan${RESET}"
    echo -e "    ${DIM}vol -f \"${HFILE:-<hiberfil.sys>}\" windows.cmdline${RESET}"

    [[ ${#FILE_ENTRIES[@]} -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "hiberfil")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""
    for E in "${FILE_ENTRIES[@]}"; do
        IFS='§' read -r MNAME MFILE MSIZE MAGIC B64_STRINGS <<< "$E"
        local STR_ROWS=""
        local DECODED; DECODED=$(echo "$B64_STRINGS" | base64 -d 2>/dev/null || true)
        while IFS=$'\t' read -r LBL VAL; do
            [[ -z "$LBL" ]] && continue
            local CSS="mono fld"; [[ "$LBL" == "CRED" || "$LBL" == "URL" ]] && CSS="mono fld bad"
            STR_ROWS+="<tr><td class='mono' style='white-space:nowrap;color:var(--accent4)'>$(html_esc "$LBL")</td>
              <td><span class='${CSS}' style='font-size:.72rem'>$(html_esc "${VAL:0:200}")</span></td></tr>"
        done <<< "$DECODED"
        CARDS_HTML+="<div class='card'>
          <div class='card-header'><div class='uicon'>MEM</div>
            <div><div class='uname'>$(html_esc "$MNAME")</div><div class='upath'>$(html_esc "$MFILE")</div></div>
            <div class='badge'>${MSIZE}</div>
          </div>
          <table><thead><tr><th style='width:14%'>$(L "Tipo" "Type")</th><th>$(L "Valore estratto" "Extracted value")</th></tr></thead>
          <tbody>${STR_ROWS:-<tr><td colspan='2' class='dim' style='padding:.8rem'>Nessuna stringa significativa estratta</td></tr>}</tbody></table>
          <div style='padding:.8rem 1.5rem;font-family:var(--mono);font-size:.7rem;color:var(--text-dim)'>
            Magic: <span style='color:var(--accent)'>0x${MAGIC}</span> &nbsp;·&nbsp;
            Per analisi completa: <span style='color:var(--accent4)'>vol -f \"${MFILE}\" windows.pstree</span>
          </div></div>"
    done
    {
        html_header "Hibernation"
        html_page_header "HB" "Hibernation / <span>Pagefile</span> Forensics" \
            "hiberfil.sys · pagefile.sys" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File trovati</div><div class='value'>${#FILE_ENTRIES[@]}</div></div>
          <div class='stat info'><div class='label'>Stringhe estratte</div><div class='value'>${TOTAL_STRINGS}</div></div>
        </div><main>
        <style>.fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}</style>
        <div class='stitle'>Artefatti in memoria — analisi tramite strings (parziale)</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
