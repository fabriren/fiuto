
# ================================================================
#  MODULO 37 — Master Timeline (aggregazione cross-moduli)
# ================================================================
module_master_timeline() {
    section_header "$(L "Master Timeline — Aggregazione Cross-Moduli" "Master Timeline — Cross-Module Aggregation")" "$YELLOW"

    if [[ ${#GENERATED_REPORTS[@]} -eq 0 ]]; then
        warn "$(L "Nessun report generato in questa sessione." "No reports generated in this session.")"
        info "$(L "Esegui prima i moduli che ti interessano, poi richiama la Master Timeline." "Run the desired modules first, then call the Master Timeline.")"
        return 0
    fi

    info "Report disponibili in sessione: ${#GENERATED_REPORTS[@]}"
    echo ""

    # Python estrae le righe con timestamp da tutti i report HTML.
    # Usa re.DOTALL per matchare <tr>...</tr> multiriga (grep -E non può farlo).
    # Output: YYYY-MM-DD HH:MM:SS \t modulo \t descrizione (una riga per evento)
    local TL_RAW
    TL_RAW=$("$PY3" - "${GENERATED_REPORTS[@]}" << 'PYEOF' 2>/dev/null || true
import sys, re, html as html_mod

TS_PAT  = re.compile(r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}')
TD_PAT  = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL | re.IGNORECASE)
TAG_PAT = re.compile(r'<[^>]+>')

def strip_tags(s):
    return TAG_PAT.sub('', html_mod.unescape(s)).strip()

seen = set()
for rpath in sys.argv[1:]:
    mod = rpath.rstrip('/').split('/')[-2]
    mod = re.sub(r'_\d{8}_\d{6}$', '', mod)
    try:
        with open(rpath, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        for tr_m in re.finditer(r'<tr[^>]*>(.*?)</tr>', content, re.DOTALL | re.IGNORECASE):
            cells = [strip_tags(m.group(1)) for m in TD_PAT.finditer(tr_m.group(1))]
            if not cells:
                continue
            ts = ''
            for cell in cells:
                m = TS_PAT.search(cell)
                if m:
                    ts = m.group(0).replace('T', ' ')[:19]
                    break
            if not ts:
                continue
            # Descrizione: celle che non contengono solo il timestamp né un numero puro
            desc_parts = [c[:80] for c in cells if c and not TS_PAT.fullmatch(c.strip()) and not re.fullmatch(r'\d+', c.strip())]
            desc = ' | '.join(desc_parts)[:200]
            desc = ' '.join(desc.split())   # normalizza spazi
            key = (ts, mod, desc[:40])
            if key in seen:
                continue
            seen.add(key)
            sys.stdout.buffer.write(f"{ts}\t{mod}\t{desc}\n".encode('utf-8', errors='replace'))
        # Report AI Chat (modulo ai_chat): eventi nei blocchi <div class='msg'>
        for mm in re.finditer(
                r"<span class='role'>(.*?)</span>.*?<span class='mts'>(.*?)</span>"
                r".*?<pre class='msg-t'>(.*?)</pre>", content, re.DOTALL):
            role = strip_tags(mm.group(1))
            ts_m = TS_PAT.search(html_mod.unescape(mm.group(2)))
            if not ts_m:
                continue
            ts = ts_m.group(0).replace('T', ' ')[:19]
            text = ' '.join(strip_tags(mm.group(3)).split())[:200]
            desc = (role + ': ' + text) if role and role != '—' else text
            key = (ts, mod, desc[:40])
            if key in seen:
                continue
            seen.add(key)
            sys.stdout.buffer.write(f"{ts}\t{mod}\t{desc}\n".encode('utf-8', errors='replace'))
    except Exception:
        pass
PYEOF
)

    local TOTAL_EVENTS; TOTAL_EVENTS=$(printf '%s' "$TL_RAW" | grep -c $'\t' 2>/dev/null || echo 0)

    # Ordina per timestamp decrescente, scrive su file temp per evitare
    # problemi con here-string su variabili di grandi dimensioni
    local _TL_TMP; _TL_TMP=$(mktemp)
    printf '%s\n' "$TL_RAW" | sort -t$'\t' -k1 -r | head -100000 > "$_TL_TMP"
    local SHOW_COUNT; SHOW_COUNT=$(wc -l < "$_TL_TMP")

    separator
    info "$(L "Eventi con timestamp estratti:" "Events with extracted timestamps:") ${BOLD}$TOTAL_EVENTS${RESET} ($(L "mostrati" "showing") $SHOW_COUNT)"
    [[ $TOTAL_EVENTS -eq 0 ]] && { rm -f "$_TL_TMP"; warn "$(L "Nessun timestamp trovato nei report. Verifica che i moduli abbiano generato dati." "No timestamps found in reports. Verify that modules have generated data.")"; return 0; }
    ask_yn "$(L "Generare report HTML Master Timeline?" "Generate Master Timeline HTML report?")" || { rm -f "$_TL_TMP"; return 0; }

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "master_timeline")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local NR="${#GENERATED_REPORTS[@]}"

    {
        html_header "Master Timeline"
        html_page_header "TL" "Master Timeline — <span>$(L "Aggregazione Cross-Moduli" "Cross-Module Aggregation")</span>" \
            "$(L "Aggregazione di" "Aggregation of") ${NR} $(L "report" "reports")" "$SCAN" "$WIN_ROOT"
        cat << STATSEOF
<div class='statsbar'>
  <div class='stat'><div class='label'>$(L "Report aggregati" "Aggregated reports")</div><div class='value'>${NR}</div></div>
  <div class='stat info'><div class='label'>$(L "Eventi con timestamp" "Events with timestamp")</div><div class='value'>${TOTAL_EVENTS}</div></div>
  <div class='stat ok'><div class='label'>$(L "Mostrati nel report" "Shown in report")</div><div class='value'>${SHOW_COUNT}</div></div>
</div>
<main>
<div class='stitle'>$(L "Timeline cronologica (decrescente) — ordinata per data" "Chronological timeline (descending) — sorted by date")</div>
<div class='card'>
<table>
  <thead><tr>
    <th style='width:16%'>Timestamp</th>
    <th style='width:14%'>$(L "Modulo" "Module")</th>
    <th>$(L "Descrizione" "Description")</th>
  </tr></thead>
  <tbody>
STATSEOF

        # Genera righe HTML direttamente da file temp — escaping inline senza subshell
        while IFS=$'\t' read -r _TS _MOD _DESC; do
            [[ -z "$_TS" ]] && continue
            # Escaping inline: nessun subshell, nessun fork
            local _TE="${_TS//&/&amp;}"; _TE="${_TE//</&lt;}"; _TE="${_TE//>/&gt;}"
            local _ME="${_MOD//&/&amp;}"; _ME="${_ME//</&lt;}"; _ME="${_ME//>/&gt;}"
            local _D="${_DESC:0:150}"
            local _DE="${_D//&/&amp;}"; _DE="${_DE//</&lt;}"; _DE="${_DE//>/&gt;}"
            local _BC="var(--accent)"
            case "${_MOD,,}" in
                *evtx*|*service*|*ifeo*|*wmi*|*sam*|*ntds*) _BC="var(--accent2)" ;;
                *ps_hist*|*script*|*prefetch*|*amcache*)     _BC="var(--accent4)" ;;
                *browser*|*lnk*|*shell*|*userassist*)        _BC="var(--accent3)" ;;
                *ai_chat*)                                    _BC="#a371f7" ;;
            esac
            printf '<tr><td class="mono ok" style="white-space:nowrap;font-size:.72rem">%s</td>' "$_TE"
            printf '<td><span style="background:rgba(88,166,255,.1);border:1px solid rgba(88,166,255,.2);'
            printf 'color:%s;font-family:var(--mono);font-size:.65rem;padding:.15rem .5rem;border-radius:2px">%s</span></td>' "$_BC" "$_ME"
            printf '<td class="mono" style="font-size:.72rem;color:var(--text)">%s</td></tr>\n' "$_DE"
        done < "$_TL_TMP"

        cat << 'TABLEEOF'
  </tbody>
</table>
</div>
</main>
TABLEEOF

        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    rm -f "$_TL_TMP"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
