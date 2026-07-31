
# ================================================================
#  MODULO 49 — WebCacheV01.dat (Internet Explorer / Edge Legacy)
#
#  Il modulo 17 legge la cronologia dai database SQLite di Chrome, Edge
#  Chromium e Firefox. Prima di Edge Chromium, pero', Windows registrava la
#  navigazione in WebCacheV01.dat, un database ESE che contiene:
#
#  - la cronologia di IE e Edge Legacy;
#  - i download;
#  - i cookie;
#  - e soprattutto la cronologia di TUTTO cio' che usa WinINET, quindi anche
#    la navigazione fatta da applicazioni, script e malware che si appoggiano
#    alle API di sistema invece che a un browser.
#
#  Quest'ultimo punto e' il motivo per cui il file resta rilevante anche su
#  macchine dove nessuno usa piu' Internet Explorer.
# ================================================================
module_webcache() {
    section_header "WebCacheV01 - IE / Edge Legacy / WinINET" "$CYAN"
    check_win_root || return 1

    local -a DBS=()
    local HOME_DIR F
    while IFS= read -r HOME_DIR; do
        local D; D=$(ci_find_dir "$HOME_DIR" "AppData/Local/Microsoft/Windows/WebCache")
        [[ -z "$D" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] && DBS+=("$F")
        done < <(find "$D" -maxdepth 1 -type f -iname 'WebCacheV*.dat' 2>/dev/null)
    done < <(get_user_homes)

    if [[ ${#DBS[@]} -eq 0 ]]; then
        warn "$(L "Nessun WebCacheV01.dat trovato." "No WebCacheV01.dat found.")"
        return 0
    fi
    info "$(L "Database trovati:" "Databases found:") ${BOLD}${#DBS[@]}"

    local HAVE_ESE=false
    "$PY3" -c "import pyesedb" 2>/dev/null && HAVE_ESE=true
    $HAVE_ESE || info "$(L "libesedb (pyesedb) non disponibile: si usera' l'estrazione dalle stringhe." \
                          "libesedb (pyesedb) unavailable: string extraction will be used.")"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    # Il parser effettivamente riuscito viene riportato dallo script.
    local METHOD; METHOD=$("$PY3" - "$OUT" "$($HAVE_ESE && echo 1 || echo 0)" "${DBS[@]}" << 'PYEOF' 2>/dev/null
import sys, os, re, datetime

out_path = sys.argv[1]
have_ese = sys.argv[2] == '1'
dbs = sys.argv[3:]

def filetime(v):
    """FILETIME (100ns dal 1601) -> stringa UTC."""
    try:
        v = int(v)
        if v <= 0:
            return ''
        return (datetime.datetime(1601, 1, 1)
                + datetime.timedelta(microseconds=v / 10)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ''

rows = []
seen = set()

def add(kind, when, url, extra, src):
    key = (kind, url[:150], when)
    if key in seen:
        return
    seen.add(key)
    rows.append((when, kind, url[:500], extra[:200], src))

def parse_ese():
    import pyesedb
    for db in dbs:
        src = os.path.basename(db)
        try:
            f = pyesedb.file()
            f.open(db)
        except Exception:
            continue
        for i in range(f.get_number_of_tables()):
            try:
                t = f.get_table(i)
                name = t.get_name()
            except Exception:
                continue
            # Le tabelle utili sono Container_N (cronologia, cookie, download).
            if not name.startswith('Container_'):
                continue
            try:
                cols = [t.get_column(c).get_name() for c in range(t.get_number_of_columns())]
            except Exception:
                continue
            idx = {c: n for n, c in enumerate(cols)}
            for r in range(t.get_number_of_records()):
                try:
                    rec = t.get_record(r)
                except Exception:
                    continue
                def val(col):
                    n = idx.get(col)
                    if n is None:
                        return ''
                    try:
                        v = rec.get_value_data_as_string(n)
                        return v if v else ''
                    except Exception:
                        try:
                            v = rec.get_value_data_as_integer(n)
                            return str(v) if v is not None else ''
                        except Exception:
                            return ''
                url = val('Url')
                if not url:
                    continue
                # WinINET prefissa le voci con "Visited:", "Cookie:", "iecompat:"...
                kind = 'history'
                low = url.lower()
                if low.startswith('cookie:'):
                    kind = 'cookie'
                elif low.startswith('visited:'):
                    kind = 'visited'
                elif low.startswith('iedownload') or 'download' in name.lower():
                    kind = 'download'
                when = filetime(val('AccessedTime')) or filetime(val('ModifiedTime')) or filetime(val('CreationTime'))
                extra = ' '.join(x for x in (val('Filename'), val('AccessCount') and f"hits={val('AccessCount')}") if x)
                add(kind, when, url, extra, src)
        try:
            f.close()
        except Exception:
            pass


def parse_strings():
    """Gli URL restano in chiaro nel file anche quando l'ESE non e' apribile."""
    URL = re.compile(r'(?:https?|ftp)://[!-~]{4,400}')
    for db in dbs:
        src = os.path.basename(db)
        try:
            with open(db, 'rb') as fh:
                raw = fh.read(1024 * 1024 * 1024)
        except Exception:
            continue
        for enc in ('utf-16-le', 'latin-1'):
            try:
                text = raw.decode(enc, 'ignore')
            except Exception:
                continue
            for m in URL.finditer(text):
                add('string', '', m.group(0).rstrip('\x00').strip(), '', src)

# Un WebCacheV01 acquisito da macchina accesa e' quasi sempre dirty: pyesedb
# non lo apre. Ripiegare sulle stringhe invece di restituire un report vuoto
# e' la differenza fra "nessun dato" e "dati parziali ma reali".
method = 'strings'
if have_ese:
    parse_ese()
    if rows:
        method = 'ESE'
if not rows:
    parse_strings()
print(method if rows else 'nessuno')

rows.sort(key=lambda r: (r[0] == '', r[0]), reverse=False)
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(str(x).replace('\t', ' ') for x in r) + '\n')
PYEOF
    )

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessuna voce estratta da WebCacheV01." "No entry extracted from WebCacheV01.")"
        return 0
    fi
    ok "$(L "Voci estratte:" "Entries extracted:") ${BOLD}$TOTAL${RESET} ($(L "parser" "parser"): ${METHOD:-strings})"
    if [[ "$METHOD" != "ESE" ]] && $HAVE_ESE; then
        warn "$(L "Database non apribile con libesedb (probabilmente dirty): estrazione dalle stringhe, senza date." \
                 "Database not openable with libesedb (likely dirty): string extraction, without dates.")"
    fi
    if [[ ${#IOC_LIST[@]} -gt 0 ]]; then
        local NIOC=0 LINE
        while IFS= read -r LINE; do
            check_ioc "$LINE" && NIOC=$((NIOC + 1))
        done < "$OUT"
        [[ $NIOC -gt 0 ]] && warn "$(L "Voci con match IoC:" "Entries matching IoC:") ${BOLD}$NIOC"
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(head -30000 "$OUT")
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Data (UTC)" "Date (UTC)")" "$(L "Tipo" "Type")" "URL" "$(L "Dettaglio" "Detail")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' guardarlo anche senza Internet Explorer" "Why look at it even without Internet Explorer")</b><br>"
    NOTE+="$(L "WebCacheV01 non registra solo la navigazione del browser: raccoglie tutto cio' che passa dalle API WinINET, quindi anche le richieste fatte da applicazioni, script e malware che si appoggiano alle librerie di sistema. Su una macchina dove nessuno usa piu' IE, una voce qui e' spesso proprio codice non-browser che ha contattato la rete." \
        "WebCacheV01 does not only record browser activity: it collects everything going through the WinINET APIs, including requests made by applications, scripts and malware relying on system libraries. On a machine where nobody uses IE any more, an entry here is often exactly the non-browser code that reached out to the network.")"
    if [[ "$METHOD" != "ESE" ]]; then
        NOTE+="<br><br><b>$(L "Estrazione parziale" "Partial extraction")</b><br>"
        if $HAVE_ESE; then
            NOTE+="$(L "libesedb era disponibile ma non ha potuto aprire il database (file dirty o danneggiato, tipico di un'acquisizione a caldo): si e' ripiegato sull'estrazione degli URL dalle stringhe, senza date, tipo di voce e conteggi." \
                "libesedb was available but could not open the database (dirty or damaged file, typical of a live acquisition): extraction fell back to URLs from strings, without dates, entry type and counts.")"
        else
            NOTE+="$(L "libesedb non era disponibile: sono stati estratti solo gli URL dalle stringhe, senza date, tipo di voce e conteggi di accesso." \
                "libesedb was unavailable: only URLs were extracted from strings, without dates, entry type and access counts.")"
        fi
    fi
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Voci" "Entries")" "$TOTAL")"
    STATS+="$(stat_box "Database" "${#DBS[@]}" "info")"
    STATS+="$(stat_box "$(L "Parser" "Parser")" "${METHOD:-strings}" "$([[ "$METHOD" == "ESE" ]] && echo info || echo warn)")"
    finish_report "webcache" "WebCacheV01" "WEB" "AppData/Local/Microsoft/Windows/WebCache" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Voci WinINET" "WinINET entries")" "${DBS[0]}" "$TOTAL" "$TABLE" "🌐")</div>"
}
