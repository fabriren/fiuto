
# ================================================================
#  MODULO 21 — MFT Timeline (Master File Table)
# ================================================================
module_mft() {
    section_header "MFT - Master File Table Timeline" "$YELLOW"
    check_win_root || return 1

    local MFT_FILE=""
    MFT_FILE=$(find "$WIN_ROOT" -maxdepth 1 -iname "\$mft" -type f 2>/dev/null | head -1)
    [[ -z "$MFT_FILE" ]] && \
        MFT_FILE=$(find "$WIN_ROOT" -maxdepth 1 \( -iname "mft" -o -iname "\$mft" \) 2>/dev/null | head -1)

    # Se non trovato, prova a estrarlo con ntfscat (non richiede rimontaggio)
    if [[ -z "$MFT_FILE" ]]; then
        # Recupera il device block associato a WIN_ROOT
        local DEV=""
        while IFS=' ' read -r RD RM _ _; do
            local DM; DM=$(printf '%b' "$RM")
            [[ "$DM" == "$WIN_ROOT" ]] && DEV=$(printf '%b' "$RD") && break
        done < /proc/mounts

        local TMP_MFT=""
        if [[ -n "$DEV" ]] && command -v ntfscat &>/dev/null; then
            info "$(L "Estrazione \$MFT tramite ntfscat dal device" "Extracting \$MFT via ntfscat from device") ${DEV}..."
            TMP_MFT=$(mktemp /tmp/dfir_MFT_XXXXXX)
            if ntfscat -f "$DEV" '$MFT' > "$TMP_MFT" 2>/dev/null && [[ -s "$TMP_MFT" ]]; then
                MFT_FILE="$TMP_MFT"
                ok "$(L "\$MFT estratto con ntfscat" "\$MFT extracted with ntfscat") ($(du -h "$MFT_FILE" | cut -f1))"
                register_tmp "$TMP_MFT"
            else
                rm -f "$TMP_MFT"
                TMP_MFT=""
            fi
        fi

        if [[ -z "$MFT_FILE" ]]; then
            warn "$(L "\$MFT non accessibile - il volume è montato con ntfs3 (driver kernel)" "\$MFT not accessible - volume mounted with ntfs3 (kernel driver)")"
            echo ""
            info "$(L "Il driver ntfs3 non espone i file di sistema (\$MFT, \$LogFile, ecc.)." "The ntfs3 driver does not expose system files (\$MFT, \$LogFile, etc.).")"
            info "$(L "Per accedere all'\$MFT hai due opzioni:" "To access \$MFT you have two options:")"
            echo ""
            if [[ -n "$DEV" ]]; then
                echo -e "  ${CYAN}Opzione A${RESET} - rimonta con ntfs-3g (supporta file di sistema):"
                echo -e "  ${DIM}sudo umount \"$WIN_ROOT\"${RESET}"
                echo -e "  ${DIM}sudo mount -t ntfs-3g -o ro,noload,show_sys_files \"$DEV\" \"$WIN_ROOT\"${RESET}"
                echo ""
                echo -e "  ${CYAN}Opzione B${RESET} - estrai \$MFT senza rimontare (richiede ntfs-3g-tools):"
                echo -e "  ${DIM}sudo apt install ntfs-3g${RESET}"
                echo -e "  ${DIM}ntfscat -f \"$DEV\" '\$MFT' > /tmp/MFT && sudo chmod a+r /tmp/MFT${RESET}"
                echo -e "  ${DIM}# poi imposta WIN_ROOT e riavvia il modulo${RESET}"
            else
                echo -e "  ${CYAN}Opzione A${RESET} - rimonta con ntfs-3g:"
                echo -e "  ${DIM}sudo umount \"$WIN_ROOT\"${RESET}"
                echo -e "  ${DIM}sudo mount -t ntfs-3g -o ro,noload,show_sys_files /dev/DEVICE \"$WIN_ROOT\"${RESET}"
                echo ""
                echo -e "  ${CYAN}Opzione B${RESET} - estrai direttamente:"
                echo -e "  ${DIM}ntfscat -f /dev/DEVICE '\$MFT' > /tmp/MFT${RESET}"
            fi
            echo ""
            return 0
        fi
    fi

    info "$(L "MFT trovato:" "MFT found:") $MFT_FILE  ($(du -h "$MFT_FILE" 2>/dev/null | cut -f1 || echo '?'))"

    # Verifica python-mft
    local HAS_PYMFT=false
    "$PY3" -c "import mft" 2>/dev/null && HAS_PYMFT=true

    if ! $HAS_PYMFT; then
        warn "python-mft non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install mft"
        info "$(L "Analisi rapida tramite strings (parziale)..." "Quick analysis via strings (partial)...")"
        echo ""
        # Estrai nomi file e timestamp con strings come fallback
        local STRINGS_OUT
        STRINGS_OUT=$(strings "$MFT_FILE" 2>/dev/null | \
            grep -iE '\.(exe|dll|bat|ps1|vbs|py|zip|rar|7z|iso)$' | \
            sort -u | head -200 || true)
        echo "$STRINGS_OUT" | while IFS= read -r L; do
            printf "  ${DIM}%s${RESET}\n" "$L"
        done | head -50
        warn "$(L "Per l'analisi completa installa python-mft." "For complete analysis install python-mft.")"
        return 0
    fi

    # Calcola il numero approssimativo di record (ogni record MFT è 1KB)
    local MFT_SIZE; MFT_SIZE=$(stat -c "%s" "$MFT_FILE" 2>/dev/null || echo 0)
    local MFT_RECORDS=$((MFT_SIZE / 1024))
    
    info "$(L "Parsing MFT con python-mft (~${MFT_RECORDS} record presenti)..." "Parsing MFT with python-mft (~${MFT_RECORDS} records present)...")"
    if [[ $MFT_SIZE -gt 500000000 ]]; then
        warn "$(L "⏳ L'MFT supera i 500MB! Il parsing completo può richiedere vari minuti." "⏳ MFT exceeds 500MB! Full parsing may take several minutes.")"
        info "   $(L "(Specialmente su volumi montati con BitLocker/FUSE). Attendi senza interrompere..." "(Especially on volumes mounted with BitLocker/FUSE). Wait without interrupting...")"
    fi

    local TMP_MFT_CSV; TMP_MFT_CSV=$(mktemp /tmp/dfir_mft_XXXXXX.csv)
    local TMP_MFT_ROWS; TMP_MFT_ROWS=$(mktemp /tmp/dfir_mft_rows_XXXXXX.html)
    register_tmp "$TMP_MFT_CSV"; register_tmp "$TMP_MFT_ROWS"

    "$PY3" - "$MFT_FILE" "$TMP_MFT_CSV" << 'PYEOF' 2>/dev/null
import sys, json, datetime
try:
    from mft import PyMftParser
except ImportError:
    sys.exit(1)

mft_file = sys.argv[1]
outf     = sys.argv[2]

SKIP_EXTS = {'.lnk','.log','.dat','.mui','.dll','.sys','.manifest',''}
SUSP_EXTS = {'.exe','.ps1','.bat','.vbs','.hta','.js','.cmd',
             '.zip','.rar','.7z','.iso','.img',
             '.py','.rb','.php','.aspx','.jsp'}

try:
    parser = PyMftParser(mft_file)
    with open(outf,'w', encoding='utf-8') as out:
        count = 0
        for chunk_bytes in parser.entries_csv():
            for line in chunk_bytes.decode('utf-8', 'replace').split('\n'):
                if not line or line.startswith('Signature'):
                    continue
                parts = line.split(',', 21)
                if len(parts) < 22:
                    continue
                
                if parts[10] == 'true': # IsADirectory
                    continue
                
                fpath = parts[21].strip('"\r')
                if '\\' in fpath:
                    name = fpath.rsplit('\\', 1)[-1]
                else:
                    name = fpath
                
                ext = ''
                if '.' in name:
                    ext = '.' + name.rsplit('.', 1)[-1].lower()
                if ext in SKIP_EXTS:
                    continue

                si_cre = parts[16][:19]
                si_mod = parts[14][:19]
                fn_cre = parts[20][:19]

                stomped = ''
                if si_cre and fn_cre and si_cre != fn_cre:
                    try:
                        d_si = datetime.datetime.fromisoformat(si_cre)
                        d_fn = datetime.datetime.fromisoformat(fn_cre)
                        if abs((d_si - d_fn).total_seconds()) > 3600:
                            stomped = 'TIMESTOMP'
                    except:
                        pass
                
                susp = '1' if ext in SUSP_EXTS or stomped else '0'
                si_cre_s = si_cre.replace('T', ' ')
                si_mod_s = si_mod.replace('T', ' ')
                fn_cre_s = fn_cre.replace('T', ' ')
                
                out.write(f"{name}\t{ext}\t{si_cre_s}\t{si_mod_s}\t\t{fn_cre_s}\t{stomped}\t{susp}\n")
                count += 1
    print(f"Records processati: {count}", file=sys.stderr)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF

    local TOTAL_LINES
    TOTAL_LINES=$(wc -l < "$TMP_MFT_CSV" 2>/dev/null || echo 0)
    local STOMP_COUNT
    STOMP_COUNT=$(grep -c "TIMESTOMP" "$TMP_MFT_CSV" 2>/dev/null || echo 0)
    local SUSP_COUNT
    SUSP_COUNT=$(awk -F'\t' '$8=="1"' "$TMP_MFT_CSV" 2>/dev/null | wc -l || echo 0)

    info "File processati: ${BOLD}$TOTAL_LINES${RESET}  |  $(L "Sospetti:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT${RESET}  |  Timestomping: ${RED}${BOLD}$STOMP_COUNT"

    # Stampa terminale: solo sospetti e timestomped
    echo ""
    awk -F'\t' '$8=="1" || $7=="TIMESTOMP" {print}' "$TMP_MFT_CSV" 2>/dev/null | head -30 | \
    while IFS=$'\t' read -r NAME EXT CREATED MODIFIED MFTMOD FNCREATED STOMPED SUSP; do
        local MARK=""
        [[ "$STOMPED" == "TIMESTOMP" ]] && MARK="${RED}[STOMP]${RESET} "
        printf "  %s${DIM}%-40s${RESET}  ${CYAN}%s${RESET}  %s\n" \
            "$MARK" "${NAME: -40}" "$CREATED" "$EXT"
    done

    [[ $TOTAL_LINES -eq 0 ]] && { warn "$(L "Nessun record MFT processato." "No MFT records processed.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "mft")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Genera righe HTML dal CSV (solo sospetti + primi 2000 per non gonfiare)
    "$PY3" - "$TMP_MFT_CSV" "$TMP_MFT_ROWS" << 'PYEOF' 2>/dev/null || true
import sys, html as H

rows_data = []
with open(sys.argv[1]) as f:
    for line in f:
        parts = line.rstrip('\n').split('\t')
        if len(parts) < 8:
            continue
        susp = parts[7]
        stomped = parts[6]
        if susp != '1' and stomped != 'TIMESTOMP':
            continue
        rows_data.append(parts)

# Ordina per data di modifica (SI Modified - colonna indice 3) decrescente
rows_data.sort(key=lambda x: x[3], reverse=True)

rows = []
for parts in rows_data[:2000]:
    name,ext,created,modified,mftmod,fncreated,stomped,susp = parts[:8]
    stomp_badge = ("<span style='color:var(--accent2);font-weight:700;font-size:.68rem;"
                   "margin-right:.3rem'>⚠ TIMESTOMP</span>") if stomped == 'TIMESTOMP' else ''
    row_style = ("style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
                 if stomped == 'TIMESTOMP' else "")
    rows.append(
        f"<tr {row_style}>"
        f"<td class='mono' style='word-break:break-all;font-size:.72rem'>{stomp_badge}{H.escape(name)}</td>"
        f"<td class='mono mid' style='white-space:nowrap'>{H.escape(ext)}</td>"
        f"<td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{H.escape(created)}</td>"
        f"<td class='mono' style='white-space:nowrap;font-size:.7rem;color:var(--accent4)'>{H.escape(modified)}</td>"
        f"<td class='mono dim' style='white-space:nowrap;font-size:.7rem'>{H.escape(fncreated)}</td>"
        f"</tr>"
    )

with open(sys.argv[2],'w') as f:
    f.write('\n'.join(rows))
PYEOF

    local TABLE_ROWS; TABLE_ROWS=$(cat "$TMP_MFT_ROWS" 2>/dev/null || echo "")

    {
        html_header "MFT"
        html_page_header "MF" "MFT - <span>Master File Table</span> Timeline" \
            "\$MFT - radice volume NTFS" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File analizzati</div><div class='value'>${TOTAL_LINES}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
          <div class='stat'><div class='label'>Timestomping</div><div class='value' style='color:var(--accent2)'>${STOMP_COUNT}</div></div>
        </div>
        <div class='notice' style='margin:1rem;padding:1rem;background:rgba(255,255,255,0.05);border-radius:6px;font-size:0.85rem'>
          <b>Nota:</b> Per prevenire blocchi del browser, vengono mostrati a schermo solo i primi <b>2000</b> record sospetti/timestomped ordinati per data decrescente. <br>
          Il report completo (TUTTI i record analizzati) è salvato in formato CSV nella stessa cartella: <br><code style='color:var(--accent4)'>$(basename "${REPORT_HTML%.html}").csv</code>
        </div>
        <main>
        <div class='stitle'>File sospetti e con possibile timestomping (STANDARD_INFO ≠ FILENAME &gt;1h)</div>
        <div class='card'><table style='table-layout: fixed; width: 100%;'>
          <thead><tr>
            <th style='width:57%'>$(L "Nome file" "Filename")</th>
            <th style='width:7%'>Ext</th>
            <th style='width:12%'>SI Created</th>
            <th style='width:12%'>SI Modified</th>
            <th style='width:12%'>FN Created</th>
          </tr></thead>
          <tbody>${TABLE_ROWS:-<tr><td colspan='5' class='dim' style='padding:1rem'>Nessun file sospetto trovato</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    # Salva il CSV completo accanto al report
    cp "$TMP_MFT_CSV" "${REPORT_HTML%.html}.csv" 2>/dev/null || echo ""

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    ok "Dati completi salvati: ${BOLD}${REPORT_HTML%.html}.csv"
    open_report_prompt "$REPORT_HTML"
}
