
# ================================================================
#  MODULO 1 — PowerShell PSReadLine History
# ================================================================
module_ps_history() {
    section_header "PowerShell PSReadLine History" "$MAGENTA"
    check_win_root || return 1

    local PSREADLINE_REL="AppData/Roaming/Microsoft/Windows/PowerShell/PSReadLine"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_ROOTS=() RES_FILES=() RES_SIZES=()
    declare -a RES_CTIMES=() RES_MTIMES=() RES_PATHS=()

    while IFS= read -r USER_DIR; do
        local USERNAME
        USERNAME=$(basename "$USER_DIR")
        local PSRL_DIR
        PSRL_DIR=$(ci_find_dir "$USER_DIR" "$PSREADLINE_REL")
        if [[ -z "$PSRL_DIR" || ! -d "$PSRL_DIR" ]]; then
            dim_msg "$USERNAME — $(L "PSReadLine non trovata" "PSReadLine not found")"
            continue
        fi
        mapfile -t HIST_FILES < <(find "$PSRL_DIR" -maxdepth 1 -iname "*_history.txt" -type f 2>/dev/null)
        if [[ ${#HIST_FILES[@]} -eq 0 ]]; then
            warn "$USERNAME — $(L "PSReadLine trovata ma nessun history" "PSReadLine found but no history")"
            continue
        fi
        declare -a SORT_LIST=()
        for F in "${HIST_FILES[@]}"; do
            local CT
            CT=$(stat -c "%W" "$F" 2>/dev/null || echo "0")
            [[ "$CT" == "0" ]] && CT=$(stat -c "%Y" "$F" 2>/dev/null || echo "0")
            SORT_LIST+=("${CT}|${F}")
        done
        mapfile -t SORTED < <(printf '%s\n' "${SORT_LIST[@]}" | sort -t'|' -k1 -rn)
        local COUNT=${#SORTED[@]}
        ok "$USERNAME — ${BOLD}$COUNT file history"
        local FILE_NAMES="" FILE_SIZES="" FILE_CTIMES="" FILE_MTIMES="" FILE_PATHS=""
        for ENTRY in "${SORTED[@]}"; do
            local F="${ENTRY#*|}"
            local FNAME; FNAME=$(basename "$F")
            local FSIZE; FSIZE=$(stat -c "%s" "$F" 2>/dev/null || echo "?")
            local FMTIME; FMTIME=$(stat -c "%y" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
            local RAW_CT="${ENTRY%%|*}"
            local FCTIME
            if [[ "$RAW_CT" != "0" ]]; then
                FCTIME=$(date -d "@${RAW_CT}" "+%Y-%m-%d %H:%M:%S" 2>/dev/null || echo "?")
            else
                FCTIME=$(stat -c "%z" "$F" 2>/dev/null | cut -d'.' -f1 || echo "?")
            fi
            local LINES; LINES=$(wc -l < "$F" 2>/dev/null || echo "?")
            echo -e "      ${DIM}• $FNAME  (${FSIZE} bytes, ${LINES} righe — creato: $FCTIME)${RESET}"
            # Stampa contenuto con highlight
            local DECODED
            DECODED=$("$PY3" -c "
import sys
path=sys.argv[1]
try:
    with open(path,'rb') as f: raw=f.read()
    if raw.startswith(b'\xff\xfe'):
        text=raw[2:].decode('utf-16-le',errors='replace')
    elif raw.startswith(b'\xfe\xff'):
        text=raw[2:].decode('utf-16-be',errors='replace')
    elif raw.startswith(b'\xef\xbb\xbf'):
        text=raw[3:].decode('utf-8',errors='replace')
    elif len(raw)>=2 and raw[1]==0:
        text=raw.decode('utf-16-le',errors='replace').lstrip('\ufeff')
    else:
        for enc in ('utf-8','latin-1'):
            try: text=raw.decode(enc,errors='strict'); break
            except: continue
        else: text=raw.decode('latin-1',errors='replace')
    text=text.lstrip('\ufeff').replace('\r\n','\n').replace('\r','\n')
    sys.stdout.buffer.write(text.encode('utf-8',errors='replace'))
except: pass
" "$F" 2>/dev/null || true)
            local LN=0
            while IFS= read -r LINE || [[ -n "$LINE" ]]; do
                LN=$((LN+1))
                if echo "$LINE" | grep -qi "password\|passwd\|secret\|credential\|token\|invoke-expression\|iex\|downloadstring\|bypass\|encodedcommand\|-enc\|mimikatz\|base64\|frombase64\|webclient\|net user\|net localgroup"; then
                    printf "      ${RED}%4d  %s${RESET}\n" "$LN" "$LINE"
                else
                    printf "      ${DIM}%4d${RESET}  %s\n" "$LN" "$LINE"
                fi
            done <<< "$DECODED"
            echo ""
            FILE_NAMES="${FILE_NAMES}${FNAME}|"
            FILE_SIZES="${FILE_SIZES}${FSIZE}|"
            FILE_CTIMES="${FILE_CTIMES}${FCTIME}|"
            FILE_MTIMES="${FILE_MTIMES}${FMTIME}|"
            FILE_PATHS="${FILE_PATHS}${F}|"
        done
        unset SORT_LIST SORTED
        RES_USERS+=("$USERNAME"); RES_ROOTS+=("$PSRL_DIR")
        RES_FILES+=("${FILE_NAMES%|}"); RES_SIZES+=("${FILE_SIZES%|}")
        RES_CTIMES+=("${FILE_CTIMES%|}"); RES_MTIMES+=("${FILE_MTIMES%|}")
        RES_PATHS+=("${FILE_PATHS%|}")
        TOTAL_FILES=$((TOTAL_FILES + COUNT)); TOTAL_USERS=$((TOTAL_USERS + 1))
    done < <(get_user_homes)

    separator
    info "$(L "Utenti con history:" "Users with history:") ${BOLD}$TOTAL_USERS${RESET}  |  File totali: ${BOLD}$TOTAL_FILES"

    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun file history trovato." "No history file found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "ps_history")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""

    for i in "${!RES_USERS[@]}"; do
        local USER="${RES_USERS[$i]}"
        local PSRL_PATH="${RES_ROOTS[$i]}"
        IFS='|' read -ra FNAMES  <<< "${RES_FILES[$i]}"
        IFS='|' read -ra FSIZES  <<< "${RES_SIZES[$i]}"
        IFS='|' read -ra FCTIMES <<< "${RES_CTIMES[$i]}"
        IFS='|' read -ra FMTIMES <<< "${RES_MTIMES[$i]}"
        IFS='|' read -ra FPATHS  <<< "${RES_PATHS[$i]}"
        # Combina gli array paralleli in un unico array ordinabile per FMTIME desc
        declare -a _FENTRIES=()
        for j in "${!FNAMES[@]}"; do
            _FENTRIES+=("${FMTIMES[$j]:-?}§${FCTIMES[$j]:-?}§${FNAMES[$j]}§${FSIZES[$j]:-?}§${FPATHS[$j]:-}")
        done
        mapfile -t _FENTRIES < <(printf '%s\n' "${_FENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+[""])[0],reverse=True); print("\n".join(L))')
        local FILE_BLOCKS=""
        for _FE in "${_FENTRIES[@]}"; do
            IFS='§' read -r FMTIME FCTIME FNAME FSIZE FPATH <<< "$_FE"
            local LINES_COUNT=0
            local CONTENT_HTML=""
            if [[ -f "$FPATH" && -s "$FPATH" ]]; then
                local RAW_CONTENT
                RAW_CONTENT=$("$PY3" - "$FPATH" << 'PYEOF'
import sys, html
path = sys.argv[1]
try:
    with open(path, 'rb') as f:
        raw = f.read()
    # Rileva encoding dal BOM o dalla presenza di null byte (UTF-16-LE senza BOM)
    if raw.startswith(b'\xff\xfe'):
        text = raw[2:].decode('utf-16-le', errors='replace')
    elif raw.startswith(b'\xfe\xff'):
        text = raw[2:].decode('utf-16-be', errors='replace')
    elif raw.startswith(b'\xef\xbb\xbf'):
        text = raw[3:].decode('utf-8', errors='replace')
    elif len(raw) >= 2 and raw[1] == 0:   # UTF-16-LE senza BOM (byte pari = null)
        text = raw.decode('utf-16-le', errors='replace')
    elif len(raw) >= 2 and raw[0] == 0:   # UTF-16-BE senza BOM
        text = raw.decode('utf-16-be', errors='replace')
    else:
        for enc in ('utf-8', 'latin-1'):
            try: text = raw.decode(enc); break
            except: continue
        else: text = raw.decode('latin-1', errors='replace')
    text = text.lstrip('\ufeff').replace('\r\n', '\n').replace('\r', '\n')
    lines = text.splitlines()
    SENS=['password','passwd','secret','credential','token','key','invoke-webrequest',
          'invoke-expression','iex','downloadstring','bypass','encodedcommand','-enc',
          'net user','net localgroup','whoami','mimikatz','base64','frombase64','webclient','hidden']
    out=[]
    for i,line in enumerate(lines,1):
        esc=html.escape(line)
        css='line sensitive' if any(k in line.lower() for k in SENS) else 'line'
        out.append(f'<span class="{css}"><span class="lnum">{i:4d}</span> {esc}</span>')
    print(f'LINES:{len(lines)}')
    print('\n'.join(out))
except Exception as e:
    print(f'LINES:0')
    print(f'<span class="line bad">Errore: {html.escape(str(e))}</span>')
PYEOF
)
                LINES_COUNT=$(echo "$RAW_CONTENT" | head -1 | sed 's/LINES://')
                RAW_CONTENT=$(echo "$RAW_CONTENT" | tail -n +2)
                CONTENT_HTML="<div class='hist-content'><pre class='hist-pre'>${RAW_CONTENT}</pre></div>"
            else
                CONTENT_HTML="<div style='padding:.8rem 1rem;font-family:var(--mono);font-size:.72rem;color:var(--text-dim);font-style:italic'>File vuoto</div>"
            fi
            FILE_BLOCKS+="<div class='card' style='margin-bottom:.8rem'>
              <div class='card-header'>
                <div class='uicon' style='font-size:.7rem'>PS</div>
                <div>
                  <div class='uname' style='font-size:.85rem'>$(html_esc "$FNAME")</div>
                  <div class='upath'>$(html_esc "$FPATH")</div>
                </div>
                <div style='margin-left:auto;text-align:right;font-family:var(--mono);font-size:.65rem;color:var(--text-dim)'>
                  <div class='ok'>creato: ${FCTIME}</div>
                  <div class='mid'>modificato: ${FMTIME}</div>
                  <div>${LINES_COUNT} righe &nbsp;·&nbsp; ${FSIZE} B</div>
                </div>
              </div>
              ${CONTENT_HTML}
            </div>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>▣</div>
            <div class='user-info'>
              <div class='uname'>$(html_esc "$USER")</div>
              <div class='upath'>$(html_esc "$PSRL_PATH")</div>
            </div>
            <div class='badge'>${#FNAMES[@]} file</div>
          </div>
          <div style='padding:1rem 1.5rem'>${FILE_BLOCKS}</div>
        </div>"
    done

    {
        html_header "PowerShell History"
        html_page_header "PS" "PowerShell <span>History</span> Forensics" \
            "%APPDATA%\\Microsoft\\Windows\\PowerShell\\PSReadLine\\*_history.txt" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>"
        echo "  <div class='stat'><div class='label'>File trovati</div><div class='value'>${TOTAL_FILES}</div></div>"
        echo "  <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>"
        echo "</div>"
        echo "<main><div class='stitle'>Risultati per utente</div>"
        echo "<style>
          .hist-pre{font-family:var(--mono);font-size:.75rem;line-height:1.7;padding:.8rem 1rem;
            overflow-x:auto;max-height:500px;overflow-y:auto;}
          .hist-pre::-webkit-scrollbar{width:4px;height:4px}
          .hist-pre::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
          .line{display:block;color:var(--text);white-space:pre;padding:.05rem .5rem;border-radius:2px}
          .line:hover{background:rgba(88,166,255,.05)}
          .line.sensitive{color:var(--accent2);background:rgba(255,123,114,.07);
            border-left:2px solid rgba(255,123,114,.5);padding-left:calc(.5rem - 2px)}
          .lnum{color:var(--text-dim);user-select:none;margin-right:1rem;font-size:.7rem}
        </style>"
        echo "<div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
