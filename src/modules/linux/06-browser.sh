
# --- LINUX 6 — Browser History ---
module_linux_browser() {
    section_header "Linux — Browser History" "$CYAN"
    check_target_root || return 1
    local BODY="" TOTAL=0 USERS=0
    while IFS= read -r HOME_DIR; do
        local UNAME; UNAME=$(basename "$HOME_DIR")
        local CARDS="" UCOUNT=0
        # Chromium-like: History (SQLite) sotto vari profili
        local CHROME_ROOTS=(".config/google-chrome" ".config/chromium" ".config/BraveSoftware/Brave-Browser"
                            ".config/microsoft-edge" ".config/vivaldi"
                            "snap/chromium/common/chromium" ".var/app/com.google.Chrome/config/google-chrome")
        for CR in "${CHROME_ROOTS[@]}"; do
            local BASE="$HOME_DIR/$CR"; [[ -d "$BASE" ]] || continue
            while IFS= read -r HISTDB; do
                [[ -f "$HISTDB" ]] || continue
                local ROWS; ROWS=$(query_sqlite "$HISTDB" "SELECT datetime(last_visit_time/1000000-11644473600,'unixepoch'), url, title FROM urls ORDER BY last_visit_time DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
                ok "$UNAME — $(basename "$(dirname "$HISTDB")") (Chromium)"
                CARDS+=$(_browser_table_card "$HISTDB" "$ROWS")
            done < <(find "$BASE" -maxdepth 2 -name "History" -type f 2>/dev/null)
        done
        # Firefox: places.sqlite
        for FR in ".mozilla/firefox" "snap/firefox/common/.mozilla/firefox" ".var/app/org.mozilla.firefox/.mozilla/firefox"; do
            local FBASE="$HOME_DIR/$FR"; [[ -d "$FBASE" ]] || continue
            while IFS= read -r PLACES; do
                local ROWS; ROWS=$(query_sqlite "$PLACES" "SELECT datetime(last_visit_date/1000000,'unixepoch'), url, title FROM moz_places WHERE last_visit_date IS NOT NULL ORDER BY last_visit_date DESC LIMIT 100000")
                [[ -z "$ROWS" || "$ROWS" == ERROR* ]] && continue
                UCOUNT=$((UCOUNT + 1)); TOTAL=$((TOTAL + 1))
                ok "$UNAME — $(basename "$(dirname "$PLACES")") (Firefox)"
                CARDS+=$(_browser_table_card "$PLACES" "$ROWS")
            done < <(find "$FBASE" -maxdepth 2 -name "places.sqlite" -type f 2>/dev/null)
        done
        [[ $UCOUNT -eq 0 ]] && continue
        USERS=$((USERS + 1))
        BODY+=$(generic_card_html "$UNAME" "$HOME_DIR" "$UCOUNT profili" "$CARDS" "◐")
    done < <(get_linux_user_homes)
    separator
    info "$(L "Profili browser con history:" "Browser profiles with history:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessuna history browser trovata." "No browser history found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS="$(stat_box "$(L "Profili" "Profiles")" "$TOTAL")$(stat_box "$(L "Utenti" "Users")" "$USERS" "info")"
    finish_report "linux_browser" "Linux Browser History" "WEB" "Firefox · Chrome · Chromium · Brave · Edge" "$STATS" "<div class='cards'>$BODY</div>"
}

# Tabella HTML per le righe (data\turl\ttitolo) di un DB browser
_browser_table_card() {
    local DB="$1" ROWS="$2"
    local _RTMP; _RTMP=$(mktemp); printf '%s\n' "$ROWS" > "$_RTMP"
    local TABLE; TABLE=$("$PY3" - "$_RTMP" << 'PYEOF'
import sys, html
print("<table><tr><th>Data</th><th>URL</th><th>Titolo</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    parts=line.rstrip('\n').split('\t')
    if len(parts)<2: continue
    d=html.escape(parts[0]); u=html.escape(parts[1]); t=html.escape(parts[2] if len(parts)>2 else '')
    print(f"<tr><td class='mono dim'>{d}</td><td class='mono'>{u}</td><td>{t}</td></tr>")
print("</table>")
PYEOF
)
    rm -f "$_RTMP"
    local N; N=$(printf '%s\n' "$ROWS" | grep -c . || true)
    generic_card_html "$(basename "$(dirname "$DB")")" "$DB" "$N URL" "$TABLE" "◐"
}
