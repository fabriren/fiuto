
# --- LINUX 13 — Filesystem Timeline ---
module_linux_timeline() {
    section_header "Linux — Filesystem Timeline" "$YELLOW"
    check_target_root || return 1
    info "$(L "Raccolta timestamp MAC nelle aree sensibili (può richiedere tempo)..." "Collecting MAC timestamps in sensitive areas (may take a while)...")"
    # aree di interesse forense
    local AREAS=(tmp var/tmp dev/shm root home etc/cron.d etc/systemd/system usr/local/bin var/www)
    local TMPF; TMPF=$(mktemp)
    local SCANNED=0
    for A in "${AREAS[@]}"; do
        local D; D=$(ci_find_dir "$WIN_ROOT" "$A")
        [[ -z "$D" ]] && continue
        SCANNED=$((SCANNED + 1))
        find "$D" -xdev -type f -printf '%T+\t%A+\t%C+\t%s\t%p\n' 2>/dev/null | head -100000 >> "$TMPF"
    done
    local N; N=$(wc -l < "$TMPF" 2>/dev/null || echo 0)
    if [[ "$N" -eq 0 ]]; then warn "$(L "Nessun file nelle aree sensibili." "No files in sensitive areas.")"; rm -f "$TMPF"; return 0; fi
    ok "$(L "File raccolti:" "Files collected:") ${BOLD}$N"
    # ordina per mtime desc, mostra anteprima
    sort -r "$TMPF" -o "$TMPF"
    head -25 "$TMPF" | while IFS=$'\t' read -r MT AT CT SZ P; do
        echo -e "      ${DIM}$MT${RESET}  ${P#$WIN_ROOT}"
    done
    separator
    ask_yn "Generare report HTML?" || { rm -f "$TMPF"; return 0; }
    local TABLE; TABLE=$("$PY3" - "$TMPF" "$WIN_ROOT" << 'PYEOF'
import sys, html
root=sys.argv[2]
print("<table><tr><th>Modified</th><th>Accessed</th><th>Changed</th><th>Size</th><th>Path</th></tr>")
with open(sys.argv[1]) as f:
    for line in f:
        c=line.rstrip('\n').split('\t')
        if len(c)<5: continue
        mt,at,ct,sz,p=c[0],c[1],c[2],c[3],c[4]
        rel=p[len(root):] if p.startswith(root) else p
        print(f"<tr><td class='mono dim'>{html.escape(mt)}</td><td class='mono dim'>{html.escape(at)}</td><td class='mono dim'>{html.escape(ct)}</td><td class='mono'>{sz}</td><td class='mono'>{html.escape(rel)}</td></tr>")
print("</table>")
PYEOF
)
    rm -f "$TMPF"
    local BODY; BODY=$(generic_card_html "$(L "Timeline aree sensibili" "Sensitive areas timeline")" "$WIN_ROOT" "$N file" "$TABLE" "◷")
    local STATS; STATS="$(stat_box "$(L "File" "Files")" "$N")$(stat_box "$(L "Aree" "Areas")" "$SCANNED" "info")"
    finish_report "linux_timeline" "Linux Filesystem Timeline" "TML" "MAC times (find/stat)" "$STATS" "<div class='cards'>$BODY</div>"
}
