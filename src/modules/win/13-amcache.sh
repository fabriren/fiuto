

# ================================================================
#  MODULO 13 — Amcache + Shimcache (timeline esecuzione binari)
# ================================================================
module_amcache() {
    section_header "$(L "Amcache + Shimcache - Timeline Binari" "Amcache + Shimcache - Binary Timeline")" "$YELLOW"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    # ── Amcache ──────────────────────────────────────────────────────
    local AMC_PATH
    AMC_PATH=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "Windows/appcompat/Programs")" "Amcache.hve")
    [[ -z "$AMC_PATH" ]] && \
        AMC_PATH=$(find "$WIN_ROOT" -maxdepth 6 -iname "Amcache.hve" 2>/dev/null | head -1)

    declare -a AMC_ENTRIES=()
    local AMC_COUNT=0

    if [[ -n "$AMC_PATH" ]]; then
        info "$(L "Amcache trovato:" "Amcache found:") $AMC_PATH"
        local AMC_DATA
        AMC_DATA=$("$PY3" - "$AMC_PATH" << 'PYEOF' 2>/dev/null || true
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    # Win10+: InventoryApplicationFile
    try:
        key = hive.get_key('Root\\InventoryApplicationFile')
        for sk in key.iter_subkeys():
            vals = {v.name.lower(): str(v.value) for v in sk.get_values()}
            name    = vals.get('name','')
            path    = vals.get('lowercaselong path', vals.get('filepath',''))
            sha1    = vals.get('fileid','').lstrip('0000')
            pubname = vals.get('publisher','')
            ltime   = vals.get('linkdate','')
            print(f"{name}\t{path}\t{sha1}\t{pubname}\t{ltime}\tInventoryApplicationFile")
    except Exception:
        pass
    # Fallback Win7/8: File entries
    try:
        for root_sub in ['Root\\File','Root\\Programs']:
            try:
                key = hive.get_key(root_sub)
                for vol in key.iter_subkeys():
                    for entry in vol.iter_subkeys():
                        vals = {v.name.lower(): str(v.value) for v in entry.get_values()}
                        path = vals.get('15','') or vals.get('filepath','')
                        sha1 = vals.get('101','').lstrip('0000')
                        ltime = vals.get('17','')
                        print(f"\t{path}\t{sha1}\t\t{ltime}\t{root_sub}")
            except Exception:
                pass
    except Exception:
        pass
except Exception as e:
    print(f"ERROR\t{e}\t\t\t\t")
PYEOF
)
        while IFS=$'\t' read -r AMC_NAME AMC_PATH_F AMC_SHA1 AMC_PUB AMC_TIME AMC_SRC; do
            [[ -z "$AMC_PATH_F" && -z "$AMC_NAME" ]] && continue
            [[ "$AMC_NAME" == "ERROR" ]] && { err "$AMC_PATH_F"; continue; }
            AMC_ENTRIES+=("${AMC_NAME}§${AMC_PATH_F}§${AMC_SHA1}§${AMC_PUB}§${AMC_TIME}§${AMC_SRC}")
            AMC_COUNT=$((AMC_COUNT + 1))
            printf "  ${DIM}%-40s${RESET}  ${CYAN}%-15s${RESET}  %s\n" \
                "${AMC_PATH_F: -40}" "$AMC_TIME" "${AMC_SHA1:0:16}"
        done <<< "$AMC_DATA"
        ok "$(L "Amcache: $AMC_COUNT entry" "Amcache: $AMC_COUNT entries")"
    else
        warn "$(L "Amcache.hve non trovato" "Amcache.hve not found")"
    fi

    echo ""

    # ── Shimcache ─────────────────────────────────────────────────────
    local SYS_HIVE
    SYS_HIVE=$(get_hive "SYSTEM")
    declare -a SHIM_ENTRIES=()
    local SHIM_COUNT=0

    if [[ -n "$SYS_HIVE" ]]; then
        info "$(L "Shimcache da hive SYSTEM:" "Shimcache from SYSTEM hive:") $SYS_HIVE"
        local SHIM_DATA
        SHIM_DATA=$("$PY3" - "$SYS_HIVE" << 'PYEOF' 2>/dev/null || true
import sys, struct, datetime
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])

    shim_key = None
    for cs in ['ControlSet001','ControlSet002']:
        for sub in ['Session Manager\\AppCompatCache',
                    'Session Manager\\AppCompatibility']:
            try:
                shim_key = hive.get_key(f'{cs}\\{sub}'); break
            except: continue
        if shim_key: break

    if not shim_key:
        print("ERROR\tChiave AppCompatCache non trovata")
        sys.exit(0)

    for v in shim_key.get_values():
        if v.name.lower() in ('appcompatcache','cachemainkey') and isinstance(v.value, bytes):
            data = v.value
            break
    else:
        print("ERROR\tValore cache non trovato"); sys.exit(0)

    # Parse Win10 format (signature 10ts / 10ba)
    sig = data[:4]
    entries = []
    if sig in (b'\x30\x00\x00\x00', b'\x80\x00\x00\x00'):  # Win10
        offset = 52 if sig == b'\x30\x00\x00\x00' else 52
        while offset < len(data):
            try:
                esig = data[offset:offset+4]
                if esig != b'\x10\x00\x00\x00': break
                path_size = struct.unpack_from('<H', data, offset+8)[0]
                path = data[offset+12:offset+12+path_size].decode('utf-16-le', errors='replace')
                ts_raw = struct.unpack_from('<Q', data, offset+12+path_size)[0]
                ts = ''
                if ts_raw:
                    try:
                        ts = (datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
                    except: pass
                entries.append((path, ts))
                entry_size = struct.unpack_from('<I', data, offset+4)[0]
                offset += 12 + entry_size
            except: break
    else:
        # Win7/8 simple format
        try:
            num = struct.unpack_from('<I', data, 4)[0]
            off = 128
            for _ in range(min(num, 1024)):
                plen = struct.unpack_from('<H', data, off)[0]
                path = data[off+4:off+4+plen].decode('utf-16-le', errors='replace')
                ts_raw = struct.unpack_from('<Q', data, off+4+plen)[0]
                ts = ''
                if ts_raw:
                    try:
                        ts = (datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ts_raw//10)).strftime('%Y-%m-%d %H:%M:%S')
                    except: pass
                entries.append((path, ts))
                rec_size = struct.unpack_from('<I', data, off+4+plen+8)[0]
                off += 8 + plen + rec_size
        except: pass

    for path, ts in entries:
        print(f"{path}\t{ts}")

except Exception as e:
    print(f"ERROR\t{e}")
PYEOF
)
        while IFS=$'\t' read -r SHIM_PATH SHIM_TIME; do
            [[ -z "$SHIM_PATH" ]] && continue
            [[ "$SHIM_PATH" == "ERROR" ]] && { err "$SHIM_TIME"; continue; }
            SHIM_ENTRIES+=("${SHIM_PATH}§${SHIM_TIME}")
            SHIM_COUNT=$((SHIM_COUNT + 1))
            printf "  ${DIM}%-50s${RESET}  %s\n" "${SHIM_PATH: -50}" "${SHIM_TIME:--}"
        done <<< "$SHIM_DATA"
        ok "$(L "Shimcache: $SHIM_COUNT entry" "Shimcache: $SHIM_COUNT entries")"
    else
        warn "$(L "Hive SYSTEM non trovato" "SYSTEM hive not found") - Shimcache saltato"
    fi

    separator
    info "Amcache: ${BOLD}$AMC_COUNT${RESET}  |  Shimcache: ${BOLD}$SHIM_COUNT"
    [[ $((AMC_COUNT + SHIM_COUNT)) -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "amcache")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente (Amcache=k5 LinkDate, Shimcache=k2)
    mapfile -t AMC_ENTRIES  < <(printf '%s\n' "${AMC_ENTRIES[@]}"  | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+5*[""])[4],reverse=True); print("\n".join(L))')
    mapfile -t SHIM_ENTRIES < <(printf '%s\n' "${SHIM_ENTRIES[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+["",""])[1],reverse=True); print("\n".join(L))')

    # Amcache rows
    local AMC_ROWS=""
    for E in "${AMC_ENTRIES[@]}"; do
        IFS='§' read -r N P S PUB T SRC <<< "$E"
        local PESC; PESC=$(html_esc "$P")
        local SUSP_S=""
        echo "$P" | grep -qi "temp\|public\|appdata\\\\local\|programdata\|downloads\|\\\\users\\\\" && SUSP_S="class='bad'"
        AMC_ROWS+="<tr>
          <td class='mono' style='font-size:.72rem;word-break:break-all'><span ${SUSP_S}>${PESC}</span></td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${T:--}</td>
          <td class='mono dim' style='font-size:.68rem;word-break:break-all'>${S:0:20}</td>
          <td class='mono dim' style='font-size:.68rem'>$(html_esc "$PUB")</td>
        </tr>"
    done

    # Shimcache rows
    local SHIM_ROWS=""
    for E in "${SHIM_ENTRIES[@]}"; do
        IFS='§' read -r P T <<< "$E"
        local PESC; PESC=$(html_esc "$P")
        local SUSP_S=""
        echo "$P" | grep -qi "temp\|public\|appdata\|programdata\|downloads" && SUSP_S="class='bad'"
        SHIM_ROWS+="<tr>
          <td class='mono' style='font-size:.72rem;word-break:break-all'><span ${SUSP_S}>${PESC}</span></td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${T:--}</td>
        </tr>"
    done

    {
        html_header "Amcache"
        html_page_header "AC" "Amcache + <span>Shimcache</span> Forensics" \
            "Windows\\appcompat\\Programs\\Amcache.hve" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Amcache entry</div><div class='value'>${AMC_COUNT}</div></div>
          <div class='stat info'><div class='label'>Shimcache entry</div><div class='value'>${SHIM_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Amcache - Eseguibili con hash SHA1</div>
        <div class='card'><table>
          <thead><tr><th>Path</th><th>LinkDate</th><th>$(L "SHA1 (troncato)" "SHA1 (truncated)")</th><th>Publisher</th></tr></thead>
          <tbody>${AMC_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>Shimcache - Eseguibili visti dal sistema</div>
        <div class='card'><table>
          <thead><tr><th>Path</th><th>Timestamp</th></tr></thead>
          <tbody>${SHIM_ROWS:-<tr><td colspan='2' class='dim' style='padding:1rem'>$(L "Nessun dato" "No data")</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
