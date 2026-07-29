
# ================================================================
#  MODULO 15 — WMI Subscriptions (persistenza invisibile)
# ================================================================
module_wmi() {
    section_header "$(L "WMI Subscriptions — Persistenza Invisibile" "WMI Subscriptions — Fileless Persistence")" "$RED"
    check_win_root || return 1


    # Repository WMI
    local WMI_DIR
    WMI_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/wbem/Repository")
    if [[ -z "$WMI_DIR" || ! -d "$WMI_DIR" ]]; then
        warn "$(L "Repository WMI non trovato" "WMI repository not found")"
        return 0
    fi

    info "Repository: $WMI_DIR"
    echo ""

    # Cerca i file OBJECTS.DATA e INDEX.BTR
    local OBJECTS_FILE
    OBJECTS_FILE=$(ci_find_file "$WMI_DIR" "OBJECTS.DATA")
    [[ -z "$OBJECTS_FILE" ]] && \
        OBJECTS_FILE=$(find "$WMI_DIR" -iname "OBJECTS.DATA" 2>/dev/null | head -1)

    if [[ -z "$OBJECTS_FILE" ]]; then
        warn "$(L "OBJECTS.DATA non trovato — analisi tramite strings sul repository" "OBJECTS.DATA not found — analysis via strings on repository")"
    fi

    declare -a WMI_ENTRIES=()
    local TOTAL=0 SUSP_COUNT=0

    # Parsing: usa strings per estrarre pattern WMI sospetti dal repository
    # (parsing completo del formato CIM richiede librerie specializzate)
    local WMI_DATA
    if [[ -n "$OBJECTS_FILE" ]]; then
        WMI_DATA=$("$PY3" - "$OBJECTS_FILE" << 'PYEOF' 2>/dev/null || true
import sys, re

# Legge il file come binario e cerca pattern WMI subscription
with open(sys.argv[1], 'rb') as f:
    raw = f.read()

text = raw.decode('utf-16-le', errors='replace')
text2 = raw.decode('latin-1', errors='replace')

found = set()

# Cerca CommandLineEventConsumer (esegue comandi)
for m in re.finditer(r'CommandLineTemplate\s*=\s*"([^"]{4,})"', text + text2, re.I):
    found.add(('CommandLineConsumer', m.group(1)[:300]))

# Cerca ActiveScriptEventConsumer (script VBS/PS)
for m in re.finditer(r'ScriptText\s*=\s*"([^"]{4,})"', text + text2, re.I):
    found.add(('ScriptConsumer', m.group(1)[:300]))

# Cerca __EventFilter queries
for m in re.finditer(r'Query\s*=\s*"([^"]{10,})"', text + text2, re.I):
    q = m.group(1)
    if any(k in q.lower() for k in ['win32_process','logon','startup','timer','permanentevent']):
        found.add(('EventFilter', q[:300]))

# Cerca nomi consumer e filter
for m in re.finditer(r'Name\s*=\s*"([A-Za-z0-9_\-\.]{3,64})"', text + text2, re.I):
    name = m.group(1)
    if any(k in name.lower() for k in ['persist','backdoor','evil','shell','exec','powershell','cmd','script','mof']):
        found.add(('SuspiciousName', name))

for ctype, val in sorted(found):
    print(f"{ctype}\t{val}")
PYEOF
)
    fi

    # Fallback: strings su tutto il repository
    if [[ -z "$WMI_DATA" || $(echo "$WMI_DATA" | wc -l) -lt 2 ]]; then
        info "$(L "Fallback: strings sul repository..." "Fallback: strings on repository...")"
        WMI_DATA=$(find "$WMI_DIR" -type f -print0 2>/dev/null | \
            xargs -0 strings 2>/dev/null | \
            grep -iE "CommandLineTemplate|ScriptText|ActiveScript|EventFilter|EventConsumer|FilterToConsumer" | \
            head -100 | \
            awk '{print "StringsMatch\t" $0}' || true)
    fi

    while IFS=$'\t' read -r WMI_TYPE WMI_VAL; do
        [[ -z "$WMI_TYPE" ]] && continue
        local IS_SUSP=false
        echo "$WMI_VAL" | grep -qi "powershell\|cmd\|wscript\|mshta\|certutil\|base64\|http\|download\|-enc\|bypass" && IS_SUSP=true
        [[ "$WMI_TYPE" == "CommandLineConsumer" || "$WMI_TYPE" == "ScriptConsumer" ]] && IS_SUSP=true

        WMI_ENTRIES+=("${WMI_TYPE}§${WMI_VAL}§${IS_SUSP}")
        TOTAL=$((TOTAL + 1))
        $IS_SUSP && SUSP_COUNT=$((SUSP_COUNT + 1))

        if $IS_SUSP; then
            echo -e "  ${RED}[!] ${WMI_TYPE}${RESET}"
            echo -e "      ${YELLOW}${WMI_VAL:0:120}${RESET}"
        else
            printf "  ${DIM}%-25s${RESET}  %s\n" "$WMI_TYPE" "${WMI_VAL:0:80}"
        fi
    done <<< "$WMI_DATA"

    separator
    info "Pattern WMI trovati: ${BOLD}$TOTAL${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna subscription WMI sospetta trovata." "No suspicious WMI subscription found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "wmi")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    local ROWS=""
    for E in "${WMI_ENTRIES[@]}"; do
        IFS='§' read -r WMI_TYPE WMI_VAL IS_SUSP <<< "$E"
        local ROW_STYLE="" VAL_CSS="mono fld"
        [[ "$IS_SUSP" == "true" ]] && {
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
            VAL_CSS="mono fld bad"
        }
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono' style='white-space:nowrap;vertical-align:top;padding-top:.45rem;color:var(--accent4)'>$(html_esc "$WMI_TYPE")</td>
          <td style='padding:.35rem .9rem'><span class='${VAL_CSS}'>$(html_esc "$WMI_VAL")</span></td>
        </tr>"
    done

    {
        html_header "WMI"
        html_page_header "WM" "WMI <span>Subscriptions</span> Forensics" \
            "Windows\\System32\\wbem\\Repository\\OBJECTS.DATA" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Pattern trovati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <style>
          .fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}
        </style>
        <div class='stitle'>$(L "WMI Event Subscriptions — Tipo · Contenuto" "WMI Event Subscriptions — Type · Content")</div>
        <div class='card'><table>
          <thead><tr><th style='width:20%'>$(L "Tipo" "Type")</th><th>$(L "Valore / Query / Script" "Value / Query / Script")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>ℹ ATT&amp;CK T1546.003 — WMI Event Subscription</div>
          <div style='color:var(--text-mid);line-height:1.8'>
            Meccanismo di persistenza invisibile: non compare in Run keys, Task, né Services.<br>
            Classi chiave: <span class='mono'>__EventFilter</span> · <span class='mono'>CommandLineEventConsumer</span> · <span class='mono'>ActiveScriptEventConsumer</span> · <span class='mono'>__FilterToConsumerBinding</span>
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
