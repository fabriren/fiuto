
# ================================================================
#  MODULO 7 — Notepad TabState
# ================================================================
module_notepad_tabstate() {
    section_header "$(L "Notepad TabState — Tab Rimasti Aperti" "Notepad TabState — Open Tabs")" "$MAGENTA"
    check_win_root || return 1

    local NOTEPAD_PKG="Microsoft.WindowsNotepad_8wekyb3d8bbwe"
    local TABSTATE_REL="AppData/Local/Packages"
    local TOTAL_FILES=0 TOTAL_USERS=0
    declare -a RES_USERS=() RES_PATHS=() RES_COUNTS=()
    declare -a RES_FILES=() RES_SIZES=() RES_MTIMES=() RES_CTIMES=()

    # Parser .bin inline (stessa logica di notepad_tabstate.sh)
    local PARSER_PY=""
    PARSER_PY=$(mktemp /tmp/npad_parse_XXXXXX.py)
    register_tmp "$PARSER_PY"
    cat > "$PARSER_PY" << 'PYEOF'
import sys, json, html as html_mod

def read_varint(data, offset):
    result = 0; shift = 0
    while offset < len(data):
        byte = data[offset]; offset += 1
        result |= (byte & 0x7F) << shift
        if not (byte & 0x80): break
        shift += 7
    return result, offset

def is_good_text(txt, thr=0.85):
    if not txt: return False
    ok = sum(1 for c in txt if c.isprintable() or c in '\n\r\t ')
    return ok / len(txt) >= thr

def parse_bin(path):
    try:
        with open(path, 'rb') as f: data = f.read()
    except Exception as e:
        return {'error': str(e), 'text': '', 'type': 'error', 'saved_path': '', 'char_count': 0}
    if len(data) < 4:
        return {'error': 'File troppo piccolo', 'text': '', 'type': 'empty', 'saved_path': '', 'char_count': 0}
    saved_path = ''; ftype = 0
    if data[0] == 0x4E and data[1] == 0x50:
        offset = 2; offset += 1
        ftype = data[offset]; offset += 1
        if ftype == 1 and offset < len(data):
            path_len = data[offset]; offset += 1
            if path_len > 0 and offset + path_len * 2 <= len(data):
                saved_path = data[offset:offset + path_len * 2].decode('utf-16-le', errors='replace')
                offset += path_len * 2
        best_text = ''; best_score = 0
        for skip in range(0, min(120, len(data) - offset)):
            o = offset + skip
            text_len, text_start = read_varint(data, o)
            if not (1 <= text_len <= 200000): continue
            end = text_start + text_len * 2
            if end > len(data): continue
            try:
                txt = data[text_start:end].decode('utf-16-le', errors='strict')
                if not is_good_text(txt): continue
                ok = sum(1 for c in txt if c.isprintable() or c in '\n\r\t ')
                score = len(txt) * (ok / len(txt))
                if score > best_score: best_text = txt; best_score = score
            except: pass
        if best_text:
            return {'text': best_text, 'saved_path': saved_path,
                    'type': 'saved' if ftype == 1 else 'unsaved',
                    'error': '', 'char_count': len(best_text)}
    # Fallback scan UTF-16LE
    start = 2 if (len(data) >= 2 and data[0] == 0x4E and data[1] == 0x50) else 0
    runs = []; i = start; run = []
    while i < len(data) - 1:
        b1, b2 = data[i], data[i+1]
        if b2 == 0 and (0x20 <= b1 <= 0x7e or b1 in (0x09, 0x0a, 0x0d)):
            run.append(chr(b1)); i += 2
        else:
            if len(run) >= 6: runs.append(''.join(run).strip())
            run = []; i += 1
    if len(run) >= 6: runs.append(''.join(run).strip())
    text = '\n'.join(r for r in runs if r)
    if text:
        return {'text': text, 'saved_path': saved_path, 'type': 'scan',
                'error': '', 'char_count': len(text)}
    return {'text': '', 'saved_path': '', 'type': 'empty', 'error': 'Nessun testo', 'char_count': 0}

if __name__ == '__main__':
    path = sys.argv[1]
    result = parse_bin(path)
    print(json.dumps({
        'text': html_mod.escape(result.get('text', '')),
        'saved_path': html_mod.escape(result.get('saved_path', '')),
        'type': result.get('type', ''),
        'error': result.get('error', ''),
        'char_count': result.get('char_count', 0)
    }))
PYEOF

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local PACKAGES_DIR
        PACKAGES_DIR=$(ci_find_dir "$USER_DIR" "AppData/Local/Packages")
        [[ -z "$PACKAGES_DIR" ]] && { dim_msg "$USERNAME — $(L "AppData\\Local\\Packages non trovata" "AppData\\Local\\Packages not found")"; continue; }
        local NOTEPAD_DIR
        NOTEPAD_DIR=$(find "$PACKAGES_DIR" -maxdepth 1 -iname "${NOTEPAD_PKG}*" -type d 2>/dev/null | head -1)
        [[ -z "$NOTEPAD_DIR" ]] && { dim_msg "$USERNAME — $(L "Notepad UWP non installato" "Notepad UWP not installed")"; continue; }
        local TABSTATE_DIR
        TABSTATE_DIR=$(ci_find_dir "$NOTEPAD_DIR" "LocalState/TabState")
        [[ -z "$TABSTATE_DIR" || ! -d "$TABSTATE_DIR" ]] && { warn "$USERNAME — $(L "TabState non trovata" "TabState not found")"; continue; }
        mapfile -t BIN_FILES < <(find "$TABSTATE_DIR" -maxdepth 1 -iname "*.bin" -type f -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
        if [[ ${#BIN_FILES[@]} -eq 0 ]]; then
            # Fallback se printf %T@ non è supportato (BSD/macOS)
            mapfile -t BIN_FILES < <(find "$TABSTATE_DIR" -maxdepth 1 -iname "*.bin" -type f -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
        fi
        local COUNT=${#BIN_FILES[@]}
        [[ $COUNT -eq 0 ]] && { warn "$USERNAME — $(L "TabState vuota" "TabState empty")"; continue; }
        ok "$USERNAME — $COUNT $(L "file .bin trovati" ".bin files found")"
        local FILE_NAMES="" FILE_SIZES="" FILE_MTIMES="" FILE_CTIMES=""
        for BIN in "${BIN_FILES[@]}"; do
            local FNAME; FNAME=$(basename "$BIN")
            local FSIZE; FSIZE=$(stat -c "%s" "$BIN" 2>/dev/null || echo "?")
            local FMTIME; FMTIME=$(stat -c "%y" "$BIN" 2>/dev/null | cut -d'.' -f1 || echo "?")
            local RAW_CT; RAW_CT=$(stat -c "%W" "$BIN" 2>/dev/null || echo "0")
            [[ "$RAW_CT" == "0" ]] && RAW_CT=$(stat -c "%Y" "$BIN" 2>/dev/null || echo "0")
            local FCTIME
            [[ "$RAW_CT" != "0" ]] && FCTIME=$(date -d "@${RAW_CT}" "+%Y-%m-%d %H:%M:%S" 2>/dev/null || echo "?") || FCTIME="?"
            if [[ "$FNAME" =~ \.[0-9]+\.bin$ ]]; then
                echo -e "      ${DIM}• $FNAME  (metadato sessione)${RESET}"
            else
                local TERM_TEXT
                TERM_TEXT=$("$PY3" "$PARSER_PY" "$BIN" 2>/dev/null \
                    | "$PY3" -c "import sys,json,html; d=json.load(sys.stdin); print(html.unescape(d.get('text','')))" 2>/dev/null || echo "")
                if [[ -n "$TERM_TEXT" ]]; then
                    echo -e "      ${GREEN}• $FNAME${RESET}  ${DIM}($FSIZE bytes)${RESET}"
                    local LN=0
                    while IFS= read -r LINE || [[ -n "$LINE" ]]; do
                        LN=$((LN+1))
                        printf "        ${DIM}%4d${RESET}  %s\n" "$LN" "$LINE"
                    done <<< "$TERM_TEXT"
                else
                    echo -e "      ${DIM}• $FNAME  ($FSIZE bytes — nessun testo)${RESET}"
                fi
            fi
            FILE_NAMES="${FILE_NAMES}${FNAME}|"
            FILE_SIZES="${FILE_SIZES}${FSIZE}|"
            FILE_MTIMES="${FILE_MTIMES}${FMTIME}|"
            FILE_CTIMES="${FILE_CTIMES}${FCTIME}|"
        done
        RES_USERS+=("$USERNAME"); RES_PATHS+=("$TABSTATE_DIR")
        RES_COUNTS+=("$COUNT")
        RES_FILES+=("${FILE_NAMES%|}"); RES_SIZES+=("${FILE_SIZES%|}")
        RES_MTIMES+=("${FILE_MTIMES%|}"); RES_CTIMES+=("${FILE_CTIMES%|}")
        TOTAL_FILES=$((TOTAL_FILES + COUNT)); TOTAL_USERS=$((TOTAL_USERS + 1))
    done < <(get_user_homes)

    separator
    info "$(L "Utenti:" "Users:") ${BOLD}$TOTAL_USERS${RESET}  |  File .bin: ${BOLD}$TOTAL_FILES"
    [[ $TOTAL_FILES -eq 0 ]] && { warn "$(L "Nessun tab Notepad trovato." "No Notepad tab found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "notepad_tabstate")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local CARDS_HTML=""
    for i in "${!RES_USERS[@]}"; do
        local USER="${RES_USERS[$i]}"
        local PATH_FULL="${RES_PATHS[$i]}"
        local COUNT="${RES_COUNTS[$i]}"
        IFS='|' read -ra FNAMES  <<< "${RES_FILES[$i]}"
        IFS='|' read -ra FSIZES  <<< "${RES_SIZES[$i]}"
        IFS='|' read -ra FMTIMES <<< "${RES_MTIMES[$i]}"
        IFS='|' read -ra FCTIMES <<< "${RES_CTIMES[$i]}"
        local ROWS=""
        for j in "${!FNAMES[@]}"; do
            local FNAME="${FNAMES[$j]}" FSIZE="${FSIZES[$j]:-?}"
            local FMTIME="${FMTIMES[$j]:-?}" FCTIME="${FCTIMES[$j]:-?}"
            local BIN_PATH="${PATH_FULL}/${FNAME}"
            local CONTENT_HTML=""
            if [[ "$FNAME" =~ \.[0-9]+\.bin$ ]]; then
                CONTENT_HTML="<span class='dim' style='font-size:.7rem'>metadato sessione</span>"
            elif [[ -f "$BIN_PATH" ]]; then
                local POUT
                POUT=$("$PY3" "$PARSER_PY" "$BIN_PATH" 2>/dev/null || echo '{}')
                local PTEXT; PTEXT=$("$PY3" -c "import sys,json; d=json.load(sys.stdin); print(d.get('text',''))" 2>/dev/null <<< "$POUT" || echo "")
                local PTYPE; PTYPE=$("$PY3" -c "import sys,json; d=json.load(sys.stdin); print(d.get('type',''))" 2>/dev/null <<< "$POUT" || echo "")
                local PPATH; PPATH=$("$PY3" -c "import sys,json; d=json.load(sys.stdin); print(d.get('saved_path',''))" 2>/dev/null <<< "$POUT" || echo "")
                local TBADGE=""
                case "$PTYPE" in
                    unsaved) TBADGE="<span style='color:var(--accent2);font-size:.7rem'>● non salvato</span>" ;;
                    saved)   TBADGE="<span style='color:var(--accent3);font-size:.7rem'>● salvato</span>" ;;
                    scan)    TBADGE="<span style='color:var(--accent);font-size:.7rem'>● scan</span>" ;;
                esac
                local SPATH_HTML=""
                [[ -n "$PPATH" ]] && SPATH_HTML="<div class='dim mono' style='font-size:.65rem;margin:.3rem 0'>📄 $(html_esc "$PPATH")</div>"
                if [[ -n "$PTEXT" ]]; then
                    CONTENT_HTML="${TBADGE}${SPATH_HTML}<pre style='font-family:var(--mono);font-size:.75rem;white-space:pre-wrap;max-height:250px;overflow-y:auto;margin-top:.4rem;color:var(--text)'>$(html_esc "$PTEXT")</pre>"
                else
                    CONTENT_HTML="${TBADGE}<span class='dim' style='margin-left:.5rem;font-size:.7rem'>nessun testo leggibile</span>"
                fi
            fi
            ROWS+="<tr>
              <td class='mono ok' style='white-space:nowrap;font-size:.7rem'>${FCTIME}</td>
              <td class='mono' style='font-size:.75rem'>$(html_esc "$FNAME")</td>
              <td class='mono mid' style='white-space:nowrap'>${FMTIME}</td>
              <td class='mono mid'>${FSIZE} B</td>
            </tr>
            <tr><td colspan='4' style='padding:.5rem 1rem 1rem'>${CONTENT_HTML}</td></tr>"
        done
        CARDS_HTML+="<div class='card'>
          <div class='card-header'>
            <div class='uicon'>NT</div>
            <div><div class='uname'>$(html_esc "$USER")</div><div class='upath'>$(html_esc "$PATH_FULL")</div></div>
            <div class='badge'>${COUNT} tab</div>
          </div>
          <table><thead><tr><th>$(L "Creato" "Created")</th><th>GUID / File</th><th>$(L "Modificato" "Modified")</th><th>Dim.</th></tr></thead>
          <tbody>${ROWS}</tbody></table>
        </div>"
    done

    {
        html_header "Notepad TabState"
        html_page_header "NT" "Notepad <span>TabState</span> Forensics" \
            "%LOCALAPPDATA%\\Packages\\Microsoft.WindowsNotepad_*\\LocalState\\TabState" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>File .bin</div><div class='value'>${TOTAL_FILES}</div></div>
          <div class='stat info'><div class='label'>Utenti</div><div class='value'>${TOTAL_USERS}</div></div>
        </div>
        <main><div class='stitle'>Tab Notepad per utente</div>
        <div class='cards'>${CARDS_HTML}</div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
