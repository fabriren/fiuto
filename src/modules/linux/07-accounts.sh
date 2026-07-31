
# --- LINUX 7 — User Accounts ---
module_linux_accounts() {
    section_header "Linux - User Accounts" "$RED"
    check_target_root || return 1
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    [[ -z "$ETC" ]] && { warn "$(L "Directory etc non trovata." "etc directory not found.")"; return 0; }
    local BODY="" FOUND=0
    # passwd: evidenzia UID 0 e shell interattive
    local PASSWD; PASSWD=$(ci_find_file "$ETC" "passwd")
    if [[ -n "$PASSWD" ]]; then
        FOUND=$((FOUND + 1))
        ok "passwd"
        local TABLE; TABLE=$("$PY3" - "$PASSWD" << 'PYEOF'
import sys, html
print("<table><tr><th>User</th><th>UID</th><th>GID</th><th>Home</th><th>Shell</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    line=line.rstrip('\n')
    if not line or line.startswith('#'): continue
    p=line.split(':')
    if len(p)<7: continue
    uid=p[2]; sh=p[6]
    bad = uid=='0' or (sh and not sh.endswith(('nologin','false','sync')))
    cls=" class='bad'" if uid=='0' else ""
    print(f"<tr{cls}><td class='mono'>{html.escape(p[0])}</td><td class='mono'>{uid}</td><td class='mono'>{p[3]}</td><td class='mono dim'>{html.escape(p[5])}</td><td class='mono'>{html.escape(sh)}</td></tr>")
print("</table>")
PYEOF
)
        awk -F: '$3==0{print "      \033[0;31m[UID 0] "$1" → "$7"\033[0m"}' "$PASSWD"
        BODY+=$(generic_card_html "passwd" "$PASSWD" "$(grep -cv '^#' "$PASSWD" 2>/dev/null) account" "$TABLE" "◉")
    fi
    # shadow: stato password (hash presente / vuoto / lock)
    local SHADOW; SHADOW=$(ci_find_file "$ETC" "shadow")
    if [[ -n "$SHADOW" && -r "$SHADOW" ]]; then
        FOUND=$((FOUND + 1)); ok "shadow"
        local STAB; STAB=$("$PY3" - "$SHADOW" << 'PYEOF'
import sys, html
print("<table><tr><th>User</th><th>Stato password</th></tr>")
for line in open(sys.argv[1], errors='replace'):
    line=line.rstrip('\n')
    if not line or line.startswith('#'): continue
    p=line.split(':')
    if len(p)<2: continue
    h=p[1]
    if h in ('!','*','!!',''): st='bloccata / nessuna'
    elif h.startswith('!'): st='bloccata (hash presente)'
    else: st='hash impostato'
    cls=" class='bad'" if h=='' else ""
    print(f"<tr{cls}><td class='mono'>{html.escape(p[0])}</td><td>{st}</td></tr>")
print("</table>")
PYEOF
)
        BODY+=$(generic_card_html "shadow" "$SHADOW" "" "$STAB" "◉")
    elif [[ -n "$SHADOW" ]]; then
        warn "shadow $(L "presente ma non leggibile" "present but unreadable")"
    fi
    # group, sudoers, sudoers.d
    for NF in group sudoers; do
        local F; F=$(ci_find_file "$ETC" "$NF")
        [[ -n "$F" && -s "$F" ]] || continue
        FOUND=$((FOUND + 1)); ok "$NF"
        BODY+=$(file_card_html "$F" "wheel|sudo|admin|ALL|NOPASSWD|root" "◉")
    done
    local SUDOERSD; SUDOERSD=$(ci_find_dir "$ETC" "sudoers.d")
    if [[ -n "$SUDOERSD" ]]; then
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue; FOUND=$((FOUND + 1))
            BODY+=$(file_card_html "$F" "NOPASSWD|ALL|root" "◉")
        done < <(find "$SUDOERSD" -maxdepth 1 -type f 2>/dev/null)
    fi
    separator
    info "$(L "File analizzati:" "Files analysed:") ${BOLD}$FOUND"
    [[ $FOUND -eq 0 ]] && return 0
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "File account" "Account files")" "$FOUND" "info")
    finish_report "linux_accounts" "Linux User Accounts" "USR" "/etc/passwd · shadow · group · sudoers" "$STATS" "<div class='cards'>$BODY</div>"
}
