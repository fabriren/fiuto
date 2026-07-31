

# ================================================================
#  MODULO 27 — Credential Manager (DPAPI blobs)
# ================================================================
module_credential_manager() {
    section_header "Credential Manager - DPAPI Blobs" "$MAGENTA"
    check_win_root || return 1


    declare -a ALL_CREDS=()
    local TOTAL=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        for REL in "AppData/Local/Microsoft/Credentials" "AppData/Roaming/Microsoft/Credentials" \
                   "AppData/Local/Microsoft/Vault" "AppData/Roaming/Microsoft/Vault"; do
            local CRED_DIR
            CRED_DIR=$(ci_find_dir "$USER_DIR" "$REL")
            [[ -z "$CRED_DIR" || ! -d "$CRED_DIR" ]] && continue
            while IFS= read -r CFILE; do
                [[ -f "$CFILE" ]] || continue
                local FNAME; FNAME=$(basename "$CFILE")
                local FSIZE; FSIZE=$(stat -c "%s" "$CFILE" 2>/dev/null || echo "?")
                local FMTIME; FMTIME=$(stat -c "%y" "$CFILE" 2>/dev/null | cut -d'.' -f1 || echo "?")
                # Parse DPAPI blob header
                local DPAPI_INFO
                DPAPI_INFO=$("$PY3" - "$CFILE" << 'PYEOF' 2>/dev/null || true
import sys, struct, binascii
path = sys.argv[1]
try:
    with open(path,'rb') as f: data = f.read()
    # DPAPI blob: version(4) + provider_guid(16) + mk_guid(16) + ...
    if len(data) < 36: sys.exit()
    # Cerca DPAPI magic 01 00 00 00 44 50 41 50 49
    idx = data.find(b'\x01\x00\x00\x00\x44\x50\x41\x50\x49')
    if idx == -1: idx = 0
    mk_off = idx + 20  # after version + provider guid
    if mk_off + 16 <= len(data):
        mk = data[mk_off:mk_off+16]
        mk_guid = '{%s-%s-%s-%s-%s}' % (
            binascii.hexlify(mk[0:4]).decode(), binascii.hexlify(mk[4:6]).decode(),
            binascii.hexlify(mk[6:8]).decode(), binascii.hexlify(mk[8:10]).decode(),
            binascii.hexlify(mk[10:16]).decode())
        print(f"DPAPI\t{mk_guid}\t{len(data)}")
    else:
        print(f"BLOB\t-\t{len(data)}")
except Exception as e:
    print(f"ERROR\t{e}\t0")
PYEOF
)
                local DTYPE MKGUID DSIZE
                IFS=$'\t' read -r DTYPE MKGUID DSIZE <<< "$DPAPI_INFO"
                [[ "$DTYPE" == "ERROR" ]] && continue
                local LABEL; LABEL=$(basename "$(dirname "$CFILE")")
                printf "  ${MAGENTA}[%-15s]${RESET}  ${CYAN}%-30s${RESET}  ${DIM}MK: %s${RESET}\n" \
                    "$USERNAME" "$FNAME" "${MKGUID:--}"
                ALL_CREDS+=("${USERNAME}§${FNAME}§${CFILE}§${REL}§${MKGUID}§${FSIZE}§${FMTIME}")
                TOTAL=$((TOTAL + 1))
            done < <(find "$CRED_DIR" -maxdepth 2 -type f 2>/dev/null)
        done
    done < <(get_user_homes)

    separator
    info "$(L "Blob DPAPI trovati:" "DPAPI blobs found:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun blob DPAPI trovato." "No DPAPI blobs found.")"; return 0; }
    warn "$(L "I blob sono cifrati con DPAPI - richiedono la master key utente per la decifratura" "Blobs are DPAPI-encrypted - require the user master key for decryption")"
    info "$(L "Tool utili: dpapick3, mimikatz (sekurlsa::dpapi), impacket dpapi" "Useful tools: dpapick3, mimikatz (sekurlsa::dpapi), impacket dpapi")"
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "credential_manager")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Ordina per data decrescente
    mapfile -t ALL_CREDS < <(printf '%s\n' "${ALL_CREDS[@]}" | "$PY3" -c 'import sys; L=[l.rstrip("\n") for l in sys.stdin if l.strip()]; L.sort(key=lambda x:(x.split("§")+7*[""])[6],reverse=True); print("\n".join(L))')

    local ROWS=""
    for E in "${ALL_CREDS[@]}"; do
        IFS='§' read -r UNAME FNAME CPATH REL MKGUID FSIZE FMTIME <<< "$E"
        ROWS+="<tr>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$UNAME")</td>
          <td class='mono' style='font-size:.72rem;word-break:break-all'>$(html_esc "$FNAME")</td>
          <td class='mono mid' style='font-size:.68rem;word-break:break-all'>$(html_esc "$REL")</td>
          <td class='mono dim' style='font-size:.65rem;word-break:break-all'>$([[ -n "$MKGUID" ]] && html_esc "$MKGUID" || echo "-")</td>
          <td class='mono mid' style='white-space:nowrap'>${FSIZE} B</td>
          <td class='mono ok' style='white-space:nowrap;font-size:.72rem'>${FMTIME}</td>
        </tr>"
    done
    {
        html_header "Credential Manager"
        html_page_header "CM" "Credential Manager - <span>DPAPI Blobs</span>" \
            "AppData\\Local\\Microsoft\\Credentials" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Blob trovati</div><div class='value'>${TOTAL}</div></div>
        </div><main>
        <div class='stitle'>$(L "Blob DPAPI - Utente · File · Path · MasterKey GUID · Dim · Timestamp" "DPAPI Blobs - User · File · Path · MasterKey GUID · Size · Timestamp")</div>
        <div class='card'><table>
          <thead><tr><th>$(L "Utente" "User")</th><th>File</th><th>Store</th><th>MK GUID</th><th>Dim.</th><th>$(L "Modificato" "Modified")</th></tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div>
        <div style='margin-top:1.5rem;padding:1rem 1.5rem;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:.8rem'>
          <div style='font-family:var(--mono);color:var(--accent);margin-bottom:.5rem'>Decifratura offline</div>
          <div style='font-family:var(--mono);font-size:.72rem;color:var(--text-dim)'>
            impacket: dpapi.py masterkey -file &lt;MK&gt; -password &lt;pwd&gt;<br>
            mimikatz: sekurlsa::dpapi (su sistema live)
          </div>
        </div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
