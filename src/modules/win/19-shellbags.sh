
# ================================================================
#  MODULO 19 — ShellBags (navigazione cartelle, anche cancellate)
# ================================================================
module_shellbags() {
    section_header "$(L "ShellBags - Navigazione Cartelle" "ShellBags - Folder Navigation")" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install regipy"
        return 1
    fi


    declare -a ALL_BAGS=()
    local TOTAL=0 SUSP_COUNT=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        # UsrClass.dat contiene le ShellBags principali (Win7+)
        local USRCLASS; USRCLASS=$(get_user_hive \
            "$(ci_find_dir "$USER_DIR" "AppData/Local/Microsoft/Windows")" \
            "UsrClass.dat")
        # Fallback: anche in NTUSER.DAT ci sono bag per desktop/drive locali
        local NTUSER; NTUSER=$(get_user_hive "$USER_DIR" "NTUSER.DAT")

        local FOUND_ANY=false

        for HIVE_PATH in "$USRCLASS" "$NTUSER"; do
            [[ -z "$HIVE_PATH" || ! -f "$HIVE_PATH" ]] && continue
            local HIVE_LABEL; HIVE_LABEL=$(basename "$HIVE_PATH")

            local BAG_DATA
            BAG_DATA=$("$PY3" - "$HIVE_PATH" "$USERNAME" "$HIVE_LABEL" << 'PYEOF' 2>/dev/null || true
import sys, struct, datetime
try:
    from regipy.registry import RegistryHive
    hive  = RegistryHive(sys.argv[1])
    user  = sys.argv[2]
    label = sys.argv[3]

    # Chiavi ShellBags in UsrClass.dat e NTUSER.DAT
    BAGMRU_KEYS = [
        'Local Settings\\Software\\Microsoft\\Windows\\Shell\\BagMRU',
        'Software\\Microsoft\\Windows\\Shell\\BagMRU',
        'Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\BagMRU',
    ]

    def parse_shell_item(data):
        """
        Estrae il nome leggibile da uno ShellItem binario (SHITEMID).
        Struttura Microsoft: offset 14=short name ASCII, ext block con
        wVersion>=7: long name UTF-16 a ext+4+24. Network items: ASCII a offset 4.
        """
        import struct as _struct
        if not data or len(data) < 4:
            return ''
        item_type = data[2]

        if item_type & 0x70 == 0x10:
            return ''

        if item_type in (0x2F, 0x23, 0x25, 0x29):
            try:
                ch = chr(data[3])
                if ch.isalpha():
                    return ch + ':'
            except Exception:
                pass
            return ''

        if item_type in (0x41, 0x42, 0x46, 0x47, 0x4C, 0x61):
            try:
                end = data.index(0, 4)
                s = data[4:end].decode('ascii', errors='replace').strip()
                if len(s) >= 1:
                    return s
            except Exception:
                pass
            return ''

        if item_type in (0x30, 0x31, 0x32, 0x35, 0x36, 0xB1, 0x74):
            short_name = ''
            ext_start = 14
            try:
                sn_end = data.index(0, 14)
                short_name = data[14:sn_end].decode('ascii', errors='replace').strip()
                ext_start = sn_end + 1
                if ext_start % 2 != 0:
                    ext_start += 1
            except (ValueError, Exception):
                return ''

            if ext_start + 4 > len(data):
                return short_name

            cb_size   = _struct.unpack_from('<H', data, ext_start)[0]
            w_version = _struct.unpack_from('<H', data, ext_start + 2)[0]

            if w_version >= 7 and cb_size >= 32:
                ln_off = ext_start + 4 + 24
                if ln_off < len(data) - 2:
                    try:
                        end = ln_off
                        while end + 1 < len(data) and not (data[end] == 0 and data[end+1] == 0):
                            end += 2
                        s = data[ln_off:end].decode('utf-16-le', errors='strict').rstrip('\x00')
                        if (1 <= len(s) <= 260 and
                                not any('\u2E80' <= c <= '\u9FFF' for c in s) and
                                any(c.isalnum() or c in r'._-~$! ()[]' for c in s)):
                            return s
                    except Exception:
                        pass

            if short_name and not short_name.startswith('1SPS'):
                return short_name

        off = 4
        while off < min(len(data) - 1, 256):
            if 0x20 <= data[off] <= 0x7e:
                end = off
                while end < len(data) and 0x20 <= data[end] <= 0x7e:
                    end += 1
                s = data[off:end].decode('ascii', errors='replace')
                if len(s) >= 2 and not s.startswith(('1SPS', 'SPS', '{', '\x00')):
                    return s
                off = end
            else:
                off += 1
        return ''


    def walk_bagmru(key_path, path_prefix=''):
        try:
            key = hive.get_key(key_path)
        except:
            return
        for v in key.get_values():
            if not isinstance(v.value, bytes) or v.name == 'MRUListEx':
                continue
            name = parse_shell_item(v.value)
            full_path = (path_prefix + '\\' + name).lstrip('\\') if name else path_prefix
            if full_path:
                print(f"{user}\t{full_path}\t{label}")
            try:
                subkey_path = f'{key_path}\\{v.name}'
                walk_bagmru(subkey_path, full_path)
            except:
                pass

    for key_name in BAGMRU_KEYS:
        walk_bagmru(key_name)

except Exception as e:
    pass
PYEOF
)
            if [[ -n "$BAG_DATA" ]]; then
                FOUND_ANY=true
                while IFS=$'\t' read -r U P L; do
                    [[ -z "$P" ]] && continue
                    ALL_BAGS+=("${U}§${P}§${L}")
                    TOTAL=$((TOTAL + 1))
                    local IS_SUSP=false
                    echo "$P" | grep -qi "\\\\temp\\\\\|appdata\\\\local\\\\\|programdata\\\\\|\\\\public\\\\\|network\|\\\\c\$\|admin\$\|ipc\$" \
                        && IS_SUSP=true
                    $IS_SUSP && SUSP_COUNT=$((SUSP_COUNT + 1))
                done <<< "$BAG_DATA"
            fi
        done

        if $FOUND_ANY; then
            ok "$USERNAME - $TOTAL shellbags totali"
        else
            dim_msg "$USERNAME - $(L "nessuna ShellBag trovata" "no ShellBag found")"
        fi

    done < <(get_user_homes)

    # Stampa terminale (prime 40)
    echo ""
    local CNT=0
    for E in "${ALL_BAGS[@]}"; do
        IFS='§' read -r U P L <<< "$E"
        printf "  ${DIM}[%-15s]${RESET}  %s\n" "$U" "$P"
        CNT=$((CNT+1))
        [[ $CNT -ge 40 ]] && break
    done
    [[ $TOTAL -gt 40 ]] && echo -e "  ${DIM}... e altri $((TOTAL-40)) (vedi report HTML)${RESET}"

    separator
    info "ShellBags totali: ${BOLD}$TOTAL${RESET}  |  $(L "Sospette:" "Suspicious:") ${RED}${BOLD}$SUSP_COUNT"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna ShellBag trovata." "No ShellBags found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "shellbags")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local ROWS=""
    for E in "${ALL_BAGS[@]}"; do
        IFS='§' read -r U P L <<< "$E"
        local SUSP_CSS="" ROW_STYLE=""
        echo "$P" | grep -qi "\\\\temp\|programdata\|\\\\public\|network\|\\\\c\$\|admin\$" && {
            SUSP_CSS="bad"; ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
        }
        ROWS+="<tr ${ROW_STYLE}>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$U")</td>
          <td class='mono ${SUSP_CSS}' style='word-break:break-all;font-size:.72rem'>$(html_esc "$P")</td>
          <td class='mono dim' style='white-space:nowrap;font-size:.68rem'>$(html_esc "$L")</td>
        </tr>"
    done

    {
        html_header "ShellBags"
        html_page_header "SB" "ShellBags - <span>Navigazione Cartelle</span>" \
            "UsrClass.dat → BagMRU" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>ShellBags</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospette</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div><main>
        <div class='stitle'>Cartelle visitate - anche se poi cancellate o smontate</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>$(L "Utente" "User")</th><th>$(L "Percorso ricostruito" "Reconstructed path")</th><th style='width:12%'>Hive</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
