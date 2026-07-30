
# ================================================================
#  macOS 15 — Cookie Safari e cronologia download
#
#  Due artefatti che il modulo Browser History non copre:
#
#  - Cookies.binarycookies: formato proprietario Safari. I cookie di sessione
#    dicono a quali servizi l'utente era autenticato e quando; un cookie di un
#    servizio mai usato dall'utente e' un indizio di sessione altrui.
#  - Downloads.plist: la cronologia dei download di Safari, con l'URL DI
#    ORIGINE oltre al nome del file. Il modulo Quarantine copre gli eventi di
#    quarantena, ma non tutti i download vi finiscono.
# ================================================================
module_macos_cookies_downloads() {
    section_header "macOS — Cookie & Download" "$YELLOW"
    check_target_root || return 1

    local -a COOKIES=() PLISTS=()
    local HOME_DIR F
    while IFS= read -r HOME_DIR; do
        while IFS= read -r F; do
            [[ -s "$F" ]] && COOKIES+=("$F")
        done < <(find "$HOME_DIR" -maxdepth 6 -type f -name 'Cookies.binarycookies' 2>/dev/null)
        while IFS= read -r F; do
            [[ -s "$F" ]] && PLISTS+=("$F")
        done < <(find "$HOME_DIR" -maxdepth 6 -type f -name 'Downloads.plist' 2>/dev/null)
    done < <(get_macos_user_homes)

    if [[ ${#COOKIES[@]} -eq 0 && ${#PLISTS[@]} -eq 0 ]]; then
        warn "$(L "Nessun cookie Safari o Downloads.plist trovato." "No Safari cookies or Downloads.plist found.")"
        return 0
    fi

    # ---------- Cookie ----------
    local COUT; COUT=$(mktemp); register_tmp "$COUT"
    local NCOOK=0
    if [[ ${#COOKIES[@]} -gt 0 ]]; then
        "$PY3" - "$COUT" "${COOKIES[@]}" << 'PYEOF' 2>/dev/null
import sys, struct, os, datetime

out_path = sys.argv[1]
files = sys.argv[2:]

def mac_time(v):
    """Secondi dal 2001-01-01 (epoca Apple) -> stringa UTC."""
    try:
        if not v or v <= 0:
            return ''
        return (datetime.datetime(2001, 1, 1) + datetime.timedelta(seconds=float(v))).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ''

def cstr(buf, off):
    end = buf.find(b'\x00', off)
    return buf[off:end if end >= 0 else len(buf)].decode('utf-8', 'replace')

rows = []
for path in files:
    src = os.path.basename(os.path.dirname(path)) or os.path.basename(path)
    try:
        with open(path, 'rb') as fh:
            data = fh.read()
    except Exception:
        continue
    if data[:4] != b'cook':
        continue
    try:
        npages = struct.unpack_from('>I', data, 4)[0]
        sizes = [struct.unpack_from('>I', data, 8 + 4 * i)[0] for i in range(npages)]
    except Exception:
        continue
    off = 8 + 4 * npages
    for psize in sizes:
        page = data[off:off + psize]
        off += psize
        if len(page) < 12:
            continue
        try:
            ncook = struct.unpack_from('<I', page, 4)[0]
            offsets = [struct.unpack_from('<I', page, 8 + 4 * i)[0] for i in range(ncook)]
        except Exception:
            continue
        for co in offsets:
            if co + 56 > len(page):
                continue
            try:
                (_size, _u1, flags, _u2, url_o, name_o, path_o, val_o) = struct.unpack_from('<8I', page, co)
                expiry, creation = struct.unpack_from('<dd', page, co + 40)
            except Exception:
                continue
            try:
                dom  = cstr(page, co + url_o)
                name = cstr(page, co + name_o)
                cpath = cstr(page, co + path_o)
                val  = cstr(page, co + val_o)
            except Exception:
                continue
            if not dom:
                continue
            attrs = []
            if flags & 0x1: attrs.append('Secure')
            if flags & 0x4: attrs.append('HttpOnly')
            # Il VALORE del cookie e' materiale di sessione: se ne riporta solo
            # la lunghezza, perche' un cookie valido nel report e' una credenziale
            # riutilizzabile da chiunque legga il report.
            rows.append((mac_time(creation), dom, name, f"{len(val)} byte",
                         mac_time(expiry), ';'.join(attrs), cpath, src))

rows.sort(key=lambda r: r[0], reverse=True)
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')
PYEOF
        [[ -s "$COUT" ]] && NCOOK=$(wc -l < "$COUT")
        ok "$(L "Cookie Safari:" "Safari cookies:") ${BOLD}$NCOOK"
    fi

    # ---------- Download ----------
    local DOUT; DOUT=$(mktemp); register_tmp "$DOUT"
    local NDL=0
    local P
    for P in "${PLISTS[@]}"; do
        local TXT; TXT=$(read_plist "$P" 2>/dev/null)
        [[ -z "$TXT" ]] && continue
        # Dalla rappresentazione testuale si estraggono le coppie URL/percorso.
        printf '%s\n' "$TXT" | "$PY3" -c "
import sys, re
txt = sys.stdin.read()
urls  = re.findall(r'DownloadEntryURL\W+([^\n<]+)', txt) or re.findall(r'(https?://[^\s<\"]+)', txt)
paths = re.findall(r'DownloadEntryPath\W+([^\n<]+)', txt)
for i, u in enumerate(urls):
    p = paths[i] if i < len(paths) else ''
    print(f\"{u.strip()}\t{p.strip()}\")
" >> "$DOUT" 2>/dev/null || true
    done
    [[ -s "$DOUT" ]] && NDL=$(wc -l < "$DOUT")
    [[ "$NDL" -gt 0 ]] && ok "$(L "Voci di download:" "Download entries:") ${BOLD}$NDL"

    separator
    if [[ "$NCOOK" -eq 0 && "$NDL" -eq 0 ]]; then
        warn "$(L "Nessun dato estratto." "No data extracted.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Il valore dei cookie non viene riportato" "Cookie values are not reported")</b><br>"
    BODY+="$(L "Della parte segreta si indica solo la lunghezza. Un cookie di sessione ancora valido e' una credenziale a tutti gli effetti: stamparlo nel report significherebbe consegnare a chiunque lo legga la possibilita' di impersonare l'utente. Dominio, nome, date e attributi bastano a stabilire a quali servizi l'utente era autenticato e quando." \
        "Only the length of the secret part is shown. A still-valid session cookie is a credential: printing it in the report would hand anyone reading it the ability to impersonate the user. Domain, name, dates and attributes are enough to establish which services the user was authenticated to, and when.")"
    BODY+="</div></div>"

    if [[ "$NCOOK" -gt 0 ]]; then
        local CT; CT=$(_rows_to_table "$(head -20000 "$COUT")" \
            "$(L "Creato" "Created")" "$(L "Dominio" "Domain")" "$(L "Nome" "Name")" \
            "$(L "Valore" "Value")" "$(L "Scadenza" "Expiry")" "$(L "Attributi" "Attributes")" "Path" "$(L "Origine" "Source")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Cookie Safari" "Safari cookies")" "Cookies.binarycookies" "$NCOOK" "$CT" "🍪")</div>"
    fi
    if [[ "$NDL" -gt 0 ]]; then
        local DT; DT=$(_rows_to_table "$(cat "$DOUT")" "$(L "URL di origine" "Source URL")" "$(L "Percorso locale" "Local path")")
        BODY+="<div class='cards'>$(generic_card_html "$(L "Download Safari" "Safari downloads")" "Downloads.plist" "$NDL" "$DT" "⤓")</div>"
    fi

    local STATS
    STATS="$(stat_box "Cookie" "$NCOOK")"
    STATS+="$(stat_box "Download" "$NDL" "info")"
    finish_report "macos_cookies_downloads" "macOS Cookie & Download" "CKD" "Cookies.binarycookies · Downloads.plist" "$STATS" "$BODY"
}
