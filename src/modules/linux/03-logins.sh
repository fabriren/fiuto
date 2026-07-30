
# --- LINUX 3 — Login History (wtmp/btmp/lastlog) ---
module_linux_logins() {
    section_header "Linux — Login History" "$CYAN"
    check_target_root || return 1
    local LOGDIR; LOGDIR=$(ci_find_dir "$WIN_ROOT" "var/log")
    local VARRUN; VARRUN=$(ci_find_dir "$WIN_ROOT" "var/run")
    local WTMP BTMP LASTLOG
    [[ -n "$LOGDIR" ]] && { WTMP=$(ci_find_file "$LOGDIR" "wtmp"); BTMP=$(ci_find_file "$LOGDIR" "btmp"); LASTLOG=$(ci_find_file "$LOGDIR" "lastlog"); }
    if [[ -z "$WTMP" && -z "$BTMP" ]]; then
        warn "$(L "wtmp/btmp non trovati." "wtmp/btmp not found.")"; return 0
    fi
    local BODY="" TOTAL=0
    for PAIR in "wtmp|$WTMP|$(L "Login riusciti" "Successful logins")" "btmp|$BTMP|$(L "Login FALLITI" "FAILED logins")"; do
        IFS='|' read -r TAG FILE LABEL <<< "$PAIR"
        [[ -z "$FILE" || ! -s "$FILE" ]] && continue
        echo -e "  ${BOLD}${LABEL}${RESET}  ${DIM}($FILE)${RESET}"
        local OUT; OUT=$(_parse_utmp "$FILE")
        local N; N=$(printf '%s\n' "$OUT" | grep -c . || true)
        TOTAL=$((TOTAL + N))
        printf '%s\n' "$OUT" | head -40 | while IFS= read -r R; do echo -e "      ${DIM}$R${RESET}"; done
        local TMPF; TMPF=$(mktemp); printf '%s\n' "$OUT" > "$TMPF"
        local KW; [[ "$TAG" == "btmp" ]] && KW=".*" || KW="root|0\\.0\\.0\\.0"
        BODY+=$(generic_card_html "$LABEL" "$FILE" "$N record" "<pre class='hist-pre'>$(render_pre_block "$TMPF" "$KW")</pre>" "⇆")
        rm -f "$TMPF"
        echo ""
    done
    separator
    info "$(L "Record totali:" "Total records:") ${BOLD}$TOTAL"
    [[ $TOTAL -eq 0 ]] && { warn "$(L "Nessun record di login leggibile." "No readable login record.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0
    local STATS; STATS=$(stat_box "$(L "Record login" "Login records")" "$TOTAL" "info")
    finish_report "linux_logins" "Linux Login History" "LOG" "wtmp / btmp / lastlog" "$STATS" "<div class='cards'>$BODY</div>"
}

# Parser binario utmp/wtmp/btmp (struct standard glibc, 384 byte/record)
_parse_utmp() {
    "$PY3" - "$1" << 'PYEOF'
import sys, struct, datetime
# struct utmp (Linux x86-64): ut_type(h) pad(2x) ut_pid(i) ut_line(32s) ut_id(4s)
# ut_user(32s) ut_host(256s) ut_exit(4s) ut_session(i) tv_sec(i) tv_usec(i) ut_addr_v6(16s) unused(20s)
REC = 384
TYPES = {0:'EMPTY',1:'RUNLVL',2:'BOOT',3:'NEWTIME',4:'OLDTIME',5:'INIT',6:'LOGIN',7:'USER',8:'DEAD'}
try:
    with open(sys.argv[1],'rb') as f:
        data=f.read()
    rows=[]
    for off in range(0,len(data)-REC+1,REC):
        rec=data[off:off+REC]
        try:
            ut_type=struct.unpack('<h',rec[0:2])[0]
            pid=struct.unpack('<i',rec[4:8])[0]
            line=rec[8:40].split(b'\x00')[0].decode('utf-8','replace')
            user=rec[44:76].split(b'\x00')[0].decode('utf-8','replace')
            host=rec[76:332].split(b'\x00')[0].decode('utf-8','replace')
            tv_sec=struct.unpack('<i',rec[340:344])[0]
        except Exception:
            continue
        if tv_sec<=0 and not user: continue
        try: ts=datetime.datetime.utcfromtimestamp(tv_sec).strftime('%Y-%m-%d %H:%M:%S')
        except Exception: ts='?'
        t=TYPES.get(ut_type,str(ut_type))
        if not user and t in ('EMPTY','DEAD'): continue
        rows.append(f"{ts}  {t:7s}  {user:16s}  {line:12s}  {host}")
    print('\n'.join(rows))
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF
}
