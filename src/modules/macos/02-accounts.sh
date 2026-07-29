
# Estrae i campi salienti da un plist utente dslocal
_parse_dslocal_user() {
    "$PY3" - "$1" << 'PYEOF'
import sys, plistlib
try:
    d = plistlib.load(open(sys.argv[1], 'rb'))
    def g(k):
        v = d.get(k)
        return str(v[0]) if isinstance(v, list) and v else (str(v) if v is not None else '')
    has_hash = 'yes' if d.get('ShadowHashData') else 'no'
    print(f"{g('name')}\t{g('uid')}\t{g('gid')}\t{g('home')}\t{g('shell')}\t{has_hash}")
except Exception as e:
    print(f"ERROR\t{e}\t\t\t\t")
PYEOF
}

# --- macOS 2 — User Accounts (dslocal) ---
module_macos_accounts() {
    section_header "macOS — User Accounts" "$RED"
    check_target_root || return 1
    local UDIR
    UDIR=$(ci_find_dir "$WIN_ROOT" "var/db/dslocal/nodes/Default/users")
    [[ -z "$UDIR" ]] && UDIR=$(ci_find_dir "$WIN_ROOT" "private/var/db/dslocal/nodes/Default/users")
    [[ -z "$UDIR" ]] && { warn "$(L "Database dslocal non trovato." "dslocal database not found.")"; return 0; }
    local ROWS="" COUNT=0
    while IFS= read -r PL; do
        local R; R=$(_parse_dslocal_user "$PL")
        [[ "$R" == ERROR* ]] && continue
        local UNAME; UNAME=$(printf '%s' "$R" | cut -f1)
        [[ "$UNAME" == _* ]] && continue   # account di servizio
        COUNT=$((COUNT + 1)); ROWS+="$R"$'\n'
        ok "$UNAME $(printf '%s' "$R" | awk -F'\t' '{print "(uid "$2", hash:"$6")"}')"
    done < <(find "$UDIR" -maxdepth 1 -type f -iname "*.plist" 2>/dev/null)
    separator
    info "$(L "Account utente:" "User accounts:") ${BOLD}$COUNT"
    [[ $COUNT -eq 0 ]] && { warn "$(L "Nessun account utente reale trovato." "No real user account found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local TABLE; TABLE=$(_rows_to_table "${ROWS%$'\n'}" "User" "UID" "GID" "Home" "Shell" "Hash")
    local BODY; BODY=$(generic_card_html "dslocal users" "$UDIR" "$COUNT account" "$TABLE" "◉")
    local STATS; STATS=$(stat_box "$(L "Account" "Accounts")" "$COUNT" "info")
    finish_report "macos_accounts" "macOS User Accounts" "USR" "/var/db/dslocal/.../users/*.plist" "$STATS" "<div class='cards'>$BODY</div>"
}
