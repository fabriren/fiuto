
# ================================================================
#  MODULO 30 — Browser Downloads & Login Data
# ================================================================
module_browser_extra() {
    section_header "Browser Downloads & Login Data" "$CYAN"
    check_win_root || return 1

    if ! "$PY3" -c "import sqlite3" 2>/dev/null; then
        err "$(L "sqlite3 non disponibile" "sqlite3 not available")"; return 1
    fi


    declare -a DL_PROFILES=(
        "Chrome|AppData/Local/Google/Chrome/User Data/Default/History|chromium"
        "Edge|AppData/Local/Microsoft/Edge/User Data/Default/History|chromium"
        "Brave|AppData/Local/BraveSoftware/Brave-Browser/User Data/Default/History|chromium"
        "Firefox|AppData/Roaming/Mozilla/Firefox/Profiles|firefox"
    )
    declare -a LOGIN_PROFILES=(
        "Chrome|AppData/Local/Google/Chrome/User Data/Default/Login Data|chromium"
        "Edge|AppData/Local/Microsoft/Edge/User Data/Default/Login Data|chromium"
        "Brave|AppData/Local/BraveSoftware/Brave-Browser/User Data/Default/Login Data|chromium"
    )

    local TMP_DL_ROWS; TMP_DL_ROWS=$(mktemp /tmp/dfir_dlrows_XXXXXX.html)
    local TMP_LG_ROWS; TMP_LG_ROWS=$(mktemp /tmp/dfir_lgrows_XXXXXX.html)
    register_tmp "$TMP_DL_ROWS"; register_tmp "$TMP_LG_ROWS"
    local TOTAL_DL=0 TOTAL_LG=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")

        # === Downloads ===
        for SPEC in "${DL_PROFILES[@]}"; do
            IFS='|' read -r BLABEL BPATH BTYPE <<< "$SPEC"
            local DB_PATH
            if [[ "$BTYPE" == "firefox" ]]; then
                local FF_BASE
                FF_BASE=$(ci_find_dir "$USER_DIR" "AppData/Roaming/Mozilla/Firefox/Profiles")
                [[ -z "$FF_BASE" ]] && continue
                for FF_PROF in "$FF_BASE"/*/; do
                    local FF_DB="${FF_PROF}places.sqlite"
                    [[ -f "$FF_DB" ]] || continue
                    local DL_OUT
                    DL_OUT=$("$PY3" - "$FF_DB" "firefox" "$BLABEL" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, sqlite3, shutil, os, tempfile, datetime, html as H
db_path, mode, browser, user = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("""
        SELECT p.url, a.content, h.visit_date
        FROM moz_annos a
        JOIN moz_places p ON a.place_id=p.id
        LEFT JOIN moz_historyvisits h ON h.place_id=p.id
        WHERE a.anno_attribute_id IN (
            SELECT id FROM moz_anno_attributes WHERE name='downloads/destinationFileName'
        )
        ORDER BY h.visit_date DESC LIMIT 100000
    """)
    for row in cur.fetchall():
        url = (row[0] or '')[:300]
        dest = (row[1] or '')[:200]
        ts_raw = row[2]
        ts = ''
        if ts_raw:
            try: ts = (datetime.datetime(1970,1,1)+datetime.timedelta(microseconds=ts_raw)).strftime('%Y-%m-%d %H:%M:%S')
            except: pass
        susp = any(p in url.lower() for p in ['.exe','.ps1','.bat','.dll','.hta','.msi','transfer.sh','mega.nz','pastebin'])
        rs = "style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if susp else ""
        uc = "mono fld bad" if susp else "mono fld"
        print(f"<tr {rs}><td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{H.escape(ts)}</td>"
              f"<td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{H.escape(user)} / {H.escape(browser)}</td>"
              f"<td><span class='{uc}' style='font-size:.72rem'>{H.escape(url)}</span></td>"
              f"<td class='mono mid' style='font-size:.7rem;word-break:break-word'>{H.escape(dest)}</td></tr>")
    conn.close()
except Exception as e:
    pass
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF
)
                    if [[ -n "$DL_OUT" ]]; then
                        echo "$DL_OUT" >> "$TMP_DL_ROWS"
                        TOTAL_DL=$((TOTAL_DL + $(echo "$DL_OUT" | wc -l)))
                        ok "$USERNAME / $BLABEL — $(L "download trovati" "downloads found")"
                    fi
                done
            else
                DB_PATH=$(ci_find_dir "$USER_DIR" "$(dirname "$BPATH")")
                [[ -z "$DB_PATH" ]] && continue
                DB_PATH="${DB_PATH}/$(basename "$BPATH")"
                [[ -f "$DB_PATH" ]] || continue
                local DL_OUT
                DL_OUT=$("$PY3" - "$DB_PATH" "chromium" "$BLABEL" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, sqlite3, shutil, os, tempfile, datetime, html as H
db_path, mode, browser, user = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("""
        SELECT tab_url, target_path, start_time, total_bytes
        FROM downloads ORDER BY start_time DESC LIMIT 100000
    """)
    def ct(t):
        if not t: return ''
        try: return (datetime.datetime(1601,1,1)+datetime.timedelta(microseconds=t)).strftime('%Y-%m-%d %H:%M:%S')
        except: return ''
    for row in cur.fetchall():
        url  = (row[0] or '')[:300]
        dest = (row[1] or '')[:200]
        ts   = ct(row[2])
        susp = any(p in url.lower() for p in ['.exe','.ps1','.bat','.dll','.hta','.msi','transfer.sh','mega.nz','pastebin'])
        rs = "style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'" if susp else ""
        uc = "mono fld bad" if susp else "mono fld"
        print(f"<tr {rs}><td class='mono ok' style='white-space:nowrap;font-size:.7rem'>{H.escape(ts)}</td>"
              f"<td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{H.escape(user)} / {H.escape(browser)}</td>"
              f"<td><span class='{uc}' style='font-size:.72rem'>{H.escape(url)}</span></td>"
              f"<td class='mono mid' style='font-size:.7rem;word-break:break-word'>{H.escape(dest)}</td></tr>")
    conn.close()
except Exception as e:
    pass
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF
)
                if [[ -n "$DL_OUT" ]]; then
                    echo "$DL_OUT" >> "$TMP_DL_ROWS"
                    TOTAL_DL=$((TOTAL_DL + $(echo "$DL_OUT" | wc -l)))
                    ok "$USERNAME / $BLABEL — $(L "download trovati" "downloads found")"
                fi
            fi
        done

        # === Login Data ===
        for SPEC in "${LOGIN_PROFILES[@]}"; do
            IFS='|' read -r BLABEL BPATH BTYPE <<< "$SPEC"
            local LDB_PATH
            LDB_PATH=$(ci_find_dir "$USER_DIR" "$(dirname "$BPATH")")
            [[ -z "$LDB_PATH" ]] && continue
            LDB_PATH="${LDB_PATH}/$(basename "$BPATH")"
            [[ -f "$LDB_PATH" ]] || continue
            local LG_OUT
            LG_OUT=$("$PY3" - "$LDB_PATH" "$BLABEL" "$USERNAME" << 'PYEOF' 2>/dev/null || true
import sys, sqlite3, shutil, os, tempfile, html as H
db_path, browser, user = sys.argv[1], sys.argv[2], sys.argv[3]
_fd, tmp = tempfile.mkstemp(suffix='.db'); os.close(_fd)
try:
    shutil.copy2(db_path, tmp)
    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("SELECT origin_url, username_value, length(password_value), date_password_changed FROM logins ORDER BY date_password_changed DESC LIMIT 100000")
    for row in cur.fetchall():
        url  = (row[0] or '')[:200]
        uname= (row[1] or '???')[:80]
        pwlen= row[2] or 0
        print(f"<tr><td class='mono ok' style='white-space:nowrap;font-size:.7rem'></td>"
              f"<td class='mono dim' style='white-space:nowrap;font-size:.68rem'>{H.escape(user)} / {H.escape(browser)}</td>"
              f"<td class='mono fld' style='font-size:.72rem'>{H.escape(url)}</td>"
              f"<td class='mono' style='white-space:nowrap'>{H.escape(uname)}</td>"
              f"<td class='mono dim' style='white-space:nowrap'>{pwlen}B (DPAPI)</td></tr>")
    conn.close()
except Exception as e:
    pass
finally:
    try: os.unlink(tmp)
    except: pass
PYEOF
)
            if [[ -n "$LG_OUT" ]]; then
                echo "$LG_OUT" >> "$TMP_LG_ROWS"
                TOTAL_LG=$((TOTAL_LG + $(echo "$LG_OUT" | wc -l)))
                ok "$USERNAME / $BLABEL — $(L "credenziali salvate trovate" "saved credentials found")"
            fi
        done
    done < <(get_user_homes)

    separator
    info "$(L "Download:" "Downloads:") ${BOLD}$TOTAL_DL${RESET}  |  Credenziali salvate: ${BOLD}$TOTAL_LG"
    [[ $((TOTAL_DL + TOTAL_LG)) -eq 0 ]] && { warn "$(L "Nessun dato trovato." "No data found.")"; return 0; }
    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "browser_extra")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local DL_ROWS; DL_ROWS=$(cat "$TMP_DL_ROWS" 2>/dev/null || echo "")
    local LG_ROWS; LG_ROWS=$(cat "$TMP_LG_ROWS" 2>/dev/null || echo "")
    {
        html_header "Browser Downloads"
        html_page_header "BD" "Browser <span>Downloads &amp; Login Data</span>" \
            "Chrome · Edge · Firefox · Brave" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Download</div><div class='value'>${TOTAL_DL}</div></div>
          <div class='stat info'><div class='label'>Login salvati</div><div class='value'>${TOTAL_LG}</div></div>
        </div><main>
        <style>.fld{word-break:break-all;overflow-wrap:anywhere;white-space:pre-wrap}</style>
        <div class='stitle'>$(L "Download — Timestamp · Utente/Browser · URL · Destinazione" "Downloads — Timestamp · User/Browser · URL · Destination")</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Timestamp</th><th style='width:14%'>$(L "Utente/Browser" "User/Browser")</th><th>URL</th><th style='width:20%'>$(L "Destinazione" "Destination")</th></tr></thead>
          <tbody>${DL_ROWS:-<tr><td colspan='4' class='dim' style='padding:1rem'>Nessun download trovato</td></tr>}</tbody>
        </table></div>
        <div class='stitle' style='margin-top:2rem'>$(L "Login Data — Credenziali salvate (password cifrate DPAPI)" "Login Data — Saved credentials (DPAPI-encrypted passwords)")</div>
        <div class='card'><table>
          <thead><tr><th style='width:12%'>Timestamp</th><th style='width:14%'>$(L "Utente/Browser" "User/Browser")</th><th>URL</th><th style='width:15%'>Username</th><th style='width:10%'>Password</th></tr></thead>
          <tbody>${LG_ROWS:-<tr><td colspan='5' class='dim' style='padding:1rem'>Nessuna credenziale trovata</td></tr>}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"
    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
