
# ================================================================
#  MODULO 23 — USN Journal ($UsnJrnl:$J)
# ================================================================
module_usn() {
    section_header "$(L "USN Journal - Change Log NTFS" "USN Journal - NTFS Change Log")" "$CYAN"
    check_win_root || return 1


    # Cerca $UsnJrnl nella root del volume (ntfs-3g show_sys_files) o via ntfscat
    local UJFILE=""
    UJFILE=$(find "$WIN_ROOT" -maxdepth 3 \( -iname "\$j" -o -iname "usnjrnl" \) 2>/dev/null | head -1)
    [[ -z "$UJFILE" ]] && \
        UJFILE=$(find "$WIN_ROOT" -maxdepth 2 \( -iname "\$usnjrnl" -o -iname "\$j" \) 2>/dev/null | head -1)

    if [[ -z "$UJFILE" ]]; then
        local DEV=""
        while IFS=' ' read -r RD RM _ _; do
            local DM; DM=$(printf '%b' "$RM")
            [[ "$DM" == "$WIN_ROOT" ]] && DEV=$(printf '%b' "$RD") && break
        done < /proc/mounts
        if [[ -n "$DEV" ]] && command -v ntfscat &>/dev/null; then
            info "$(L "Estrazione \$UsnJrnl via ntfscat da" "Extracting \$UsnJrnl via ntfscat from") $DEV..."
            local TMP_USN; TMP_USN=$(mktemp /tmp/dfir_usn_XXXXXX)
            if ntfscat -f "$DEV" '$Extend/$UsnJrnl:$J' > "$TMP_USN" 2>/dev/null && [[ -s "$TMP_USN" ]]; then
                UJFILE="$TMP_USN"
                ok "$(L "\$UsnJrnl estratto" "\$UsnJrnl extracted") ($(du -h "$UJFILE" | cut -f1))"
                register_tmp "$TMP_USN"
            else
                rm -f "$TMP_USN"
                warn "$(L "\$UsnJrnl non accessibile. Monta con: mount -t ntfs-3g -o ro,show_sys_files" "\$UsnJrnl not accessible. Mount with: mount -t ntfs-3g -o ro,show_sys_files")"
                return 0
            fi
        else
            warn "$(L "\$UsnJrnl non accessibile - volume montato con ntfs3 (driver kernel)" "\$UsnJrnl not accessible - volume mounted with ntfs3 (kernel driver)")"
            [[ -n "$DEV" ]] && info "Prova: ntfscat -f \"$DEV\" '\$Extend/\$UsnJrnl:\$J' > /tmp/J"
            return 0
        fi
    fi

    info "USN Journal: $UJFILE  ($(du -h "$UJFILE" 2>/dev/null | cut -f1 || echo '?'))"

    local TMP_USN_CSV; TMP_USN_CSV=$(mktemp /tmp/dfir_usn_XXXXXX.csv)
    local TMP_USN_ROWS; TMP_USN_ROWS=$(mktemp /tmp/dfir_usn_rows_XXXXXX.html)
    register_tmp "$TMP_USN_CSV"; register_tmp "$TMP_USN_ROWS"

    "$PY3" - "$UJFILE" "$TMP_USN_CSV" << 'PYEOF' 2>/dev/null || true
import sys, struct, datetime, os

REASONS = {
    0x00000001:'DATA_OVERWRITE', 0x00000002:'DATA_EXTEND', 0x00000004:'DATA_TRUNCATION',
    0x00000010:'NAMED_DATA_OVERWRITE', 0x00000020:'NAMED_DATA_EXTEND', 0x00000040:'NAMED_DATA_TRUNCATION',
    0x00000100:'FILE_CREATE', 0x00000200:'FILE_DELETE', 0x00000400:'EA_CHANGE',
    0x00000800:'SECURITY_CHANGE', 0x00001000:'RENAME_OLD_NAME', 0x00002000:'RENAME_NEW_NAME',
    0x00004000:'INDEXABLE_CHANGE', 0x00008000:'BASIC_INFO_CHANGE', 0x00010000:'HARD_LINK_CHANGE',
    0x00020000:'COMPRESSION_CHANGE', 0x00040000:'ENCRYPTION_CHANGE', 0x00080000:'OBJECT_ID_CHANGE',
    0x00100000:'REPARSE_POINT_CHANGE', 0x00200000:'STREAM_CHANGE', 0x80000000:'CLOSE',
}
SUSP_EXTS = {'.exe','.dll','.ps1','.bat','.vbs','.hta','.cmd','.js','.py','.msi','.iso','.zip','.rar','.7z'}
SUSP_PATHS_KWORDS = ['temp','tmp','appdata','public','programdata','downloads','users\\']
SUSP_REASONS = {0x00000100,0x00000200,0x00001000,0x00002000}  # CREATE DELETE RENAME

path = sys.argv[1]
outf = sys.argv[2]
count = 0

def decode_reason(r):
    parts = [n for mask,n in REASONS.items() if r & mask]
    return '|'.join(parts) if parts else hex(r)

try:
    with open(path, 'rb') as f, open(outf, 'w') as out:
        raw = f.read()
    offset = 0
    while offset < len(raw) - 60:
        # Skip nulls (padding between records)
        if raw[offset:offset+4] == b'\x00\x00\x00\x00':
            offset += 8; continue
        try:
            rec_len = struct.unpack_from('<I', raw, offset)[0]
            if rec_len < 60 or rec_len > 65536:
                offset += 8; continue
            major = struct.unpack_from('<H', raw, offset+4)[0]
            if major not in (2, 3):
                offset += 8; continue
            reason = struct.unpack_from('<I', raw, offset+40)[0]
            ts_raw = struct.unpack_from('<q', raw, offset+24)[0]
            fname_off = struct.unpack_from('<H', raw, offset+58)[0]
            fname_len = struct.unpack_from('<H', raw, offset+56)[0]
            ts = ''
            if ts_raw > 0:
                try:
                    ts = (datetime.datetime(1601,1,1) +
                          datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
                except: pass
            abs_off = offset + fname_off
            name = raw[abs_off:abs_off+fname_len].decode('utf-16-le', errors='replace')
            ext = ''
            if '.' in name:
                ext = '.' + name.rsplit('.',1)[-1].lower()
            reason_str = decode_reason(reason)
            susp = '1' if (ext in SUSP_EXTS or any(reason & m for m in SUSP_REASONS)) else '0'
            with open(outf, 'a') as out:
                out.write(f"{ts}\t{name}\t{ext}\t{reason_str}\t{susp}\n")
            count += 1
            if count >= 200000:
                break
            offset += rec_len
            if offset % 8 != 0:
                offset += 8 - (offset % 8)
        except Exception:
            offset += 8
    print(f"Records: {count}", file=sys.stderr)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF

    local TOTAL_LINES; TOTAL_LINES=$(wc -l < "$TMP_USN_CSV" 2>/dev/null || echo 0)
    local SUSP_COUNT; SUSP_COUNT=$(awk -F'\t' '$5=="1"' "$TMP_USN_CSV" 2>/dev/null | wc -l || echo 0)

    info "$(L "Record USN:" "USN Records:") ${BOLD}$TOTAL_LINES${RESET}  |  Sospetti (ext+azione): ${RED}${BOLD}$SUSP_COUNT"
    echo ""
    awk -F'\t' '$5=="1"' "$TMP_USN_CSV" 2>/dev/null | head -25 | \
    while IFS=$'\t' read -r TS NAME EXT REASON SUSP; do
        printf "  ${RED}%-20s${RESET}  ${CYAN}%-30s${RESET}  %s\n" "$TS" "$NAME" "$REASON"
    done

    [[ $TOTAL_LINES -eq 0 ]] && { warn "$(L "Nessun record USN processato." "No USN records processed.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "usn")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    "$PY3" - "$TMP_USN_CSV" "$TMP_USN_ROWS" << 'PYEOF' 2>/dev/null || true
import sys, html as H
rows = []
with open(sys.argv[1]) as f:
    for line in f:
        parts = line.rstrip('\n').split('\t')
        if len(parts) < 5: continue
        ts, name, ext, reason, susp = parts[0], parts[1], parts[2], parts[3], parts[4]
        is_susp = susp == '1'
        rs = "style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if is_susp else ""
        ec = "bad" if is_susp else "ok"
        rows.append(f"<tr {rs}><td class='mono {ec}' style='white-space:nowrap;font-size:.7rem'>{H.escape(ts)}</td>"
                    f"<td class='mono' style='word-break:break-all;font-size:.72rem'>{H.escape(name)}</td>"
                    f"<td class='mono mid' style='white-space:nowrap'>{H.escape(ext)}</td>"
                    f"<td class='mono dim' style='font-size:.68rem;word-break:break-all'>{H.escape(reason)}</td></tr>")
        if len(rows) >= 5000: break
with open(sys.argv[2], 'w') as f:
    f.write('\n'.join(rows))
PYEOF

    local TABLE_ROWS; TABLE_ROWS=$(cat "$TMP_USN_ROWS" 2>/dev/null || echo "")
    {
        html_header "USN Journal"
        html_page_header "UJ" "USN Journal - <span>Change Log</span> NTFS" \
            "\$Extend\\\$UsnJrnl:\$J" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Record totali</div><div class='value'>${TOTAL_LINES}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>$(L "Change Journal NTFS - Timestamp · File · Ext · Azione" "NTFS Change Journal - Timestamp · File · Ext · Action")</div>
        <div class='card'><table>
          <thead><tr><th style='width:14%'>Timestamp</th><th>$(L "Nome file" "Filename")</th><th style='width:6%'>Ext</th><th>$(L "Motivo" "Reason")</th></tr></thead>
          <tbody>${TABLE_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun record</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
