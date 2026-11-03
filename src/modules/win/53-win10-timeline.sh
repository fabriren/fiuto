
# ================================================================
#  MODULO 53 — Windows 10 Timeline (ActivitiesCache.db)
#
#  Dalla 1803, Windows registra in ActivitiesCache.db (SQLite, per utente)
#  quali app sono state in primo piano e quando: non solo "e' stata eseguita"
#  come Prefetch o Amcache, ma una finestra StartTime/EndTime di uso attivo.
#  Utile per correlare un'esecuzione con QUANTO a lungo e' durata, e perche'
#  copre anche attivita' che altri artefatti non vedono (documenti aperti in
#  app UWP, pagine Edge con solo il titolo, senza URL completo).
#
#  Le righe piu' recenti possono stare SOLO nel -wal: SQLite consolida nel
#  file principale alla chiusura pulita del database, e un sistema spento a
#  caldo o un'immagine acquisita mentre l'app e' aperta lascia le ultime
#  scritture li'. Si copiano quindi .db, .db-wal e .db-shm insieme: aprendo
#  la copia, SQLite le fonde da solo.
# ================================================================
module_win10_timeline() {
    section_header "Windows 10 Timeline" "$CYAN"
    check_win_root || return 1

    if ! "$PY3" -c "import sqlite3" 2>/dev/null; then
        err "$(L "sqlite3 non disponibile" "sqlite3 not available") nel Python rilevato"
        return 1
    fi

    local PARSER
    PARSER=$(mktemp /tmp/dfir_timeline_XXXXXX.py 2>/dev/null) || {
        err "$(L "Impossibile creare il file temporaneo per il parser" "Could not create temp file for the parser")"
        return 1
    }
    register_tmp "$PARSER"
    cat > "$PARSER" << 'PYEOF'
import sys, sqlite3, shutil, os, tempfile, json

db_path = sys.argv[1]
SEP = '\x1f'

TMP_SUSP = ('\\temp\\', '\\appdata\\local\\temp\\', '\\programdata\\', '\\users\\public\\',
            '\\downloads\\')


def as_text(v):
    """AppId e Payload sono dichiarati TEXT ma SQLite e' dinamicamente
    tipizzato: una riga scritta male puo' farli arrivare come BLOB. Senza
    questa normalizzazione un bytes finito in app_label()/description()
    rompe il confronto 'in' piu' sotto (str contro bytes) A META' STREAM,
    dopo che righe precedenti sono gia' state stampate correttamente."""
    if isinstance(v, bytes):
        return v.decode('utf-8', errors='replace')
    return v


def clean(s):
    return s.replace(SEP, ' ').replace('\n', ' ').replace('\r', ' ') if s else ''


def app_label(raw):
    """AppId e' un JSON array di descrittori (win32/UWP/package); si prende
    il primo percorso o id leggibile invece di stampare l'array intero."""
    raw = as_text(raw)
    if not raw:
        return ''
    try:
        items = json.loads(raw)
        for it in items:
            v = it.get('id') or it.get('application') or ''
            if v:
                return v
    except Exception:
        pass
    return raw[:200]


def description(raw):
    """Payload e' quasi sempre JSON con un campo descrittivo; nelle app UWP
    compresse o con schema diverso si rinuncia, non si tenta di indovinare."""
    raw = as_text(raw)
    if not raw:
        return ''
    try:
        data = json.loads(raw)
        for key in ('displayText', 'description', 'appDisplayName'):
            if data.get(key):
                return str(data[key])[:300]
    except Exception:
        pass
    return ''


# Le righe si accumulano in memoria e si stampano SOLO a fine lettura: un
# errore a meta' dello stream (una riga malformata dopo N righe buone) non
# deve far scambiare per dato quello che e' in realta' un fallimento
# parziale. Uscita non-zero sul fallimento, cosi' bash puo' distinguere
# "0 attivita'" da "lettura fallita" anche quando lo stdout risulta vuoto
# in entrambi i casi.
_out = []
tmp = tmp_wal = tmp_shm = None
try:
    _fd, tmp = tempfile.mkstemp(suffix='.db')
    os.close(_fd)
    tmp_wal, tmp_shm = tmp + '-wal', tmp + '-shm'
    shutil.copy2(db_path, tmp)
    os.chmod(tmp, 0o600)
    for suffix, dest in ((db_path + '-wal', tmp_wal), (db_path + '-shm', tmp_shm)):
        if os.path.isfile(suffix):
            shutil.copy2(suffix, dest)
            os.chmod(dest, 0o600)

    conn = sqlite3.connect(tmp)
    cur = conn.cursor()
    cur.execute("""
        SELECT AppId, ActivityType, Payload, StartTime, EndTime, LastModifiedTime
        FROM Activity
        ORDER BY LastModifiedTime DESC LIMIT 5000
    """)
    for appid, atype, payload, start, end, lastmod in cur.fetchall():
        app = app_label(appid)
        desc = description(payload)
        ts = as_text(start) or as_text(lastmod) or as_text(end) or ''
        susp = '1' if app and any(p in app.lower() for p in TMP_SUSP) else '0'
        _out.append(SEP.join([clean(ts), clean(app), clean(str(atype or '')),
                               clean(desc), clean(as_text(end) or ''), susp]))
    conn.close()
    sys.stdout.write('\n'.join(_out) + ('\n' if _out else ''))
except Exception as e:
    print('ERROR' + SEP + str(e))
    sys.exit(1)
finally:
    for f in (tmp, tmp_wal, tmp_shm):
        if not f:
            continue
        try:
            os.unlink(f)
        except Exception:
            pass
PYEOF

    declare -a ALL_ENTRIES=()  # campi \x1f: ts|app|type|desc|end|susp|user
    local TOTAL=0 SUSP_COUNT=0 NUSERS=0 READ_ERRORS=0

    while IFS= read -r USER_DIR; do
        local USERNAME; USERNAME=$(basename "$USER_DIR")
        local CDP_DIR
        CDP_DIR=$(ci_find_dir "$USER_DIR" "AppData/Local/ConnectedDevicesPlatform")
        [[ -z "$CDP_DIR" ]] && continue

        local -a DBS=()
        while IFS= read -r DB; do
            [[ -n "$DB" ]] && DBS+=("$DB")
        done < <(find "$CDP_DIR" -maxdepth 2 -type f -iname "ActivitiesCache.db" 2>/dev/null)
        [[ ${#DBS[@]} -eq 0 ]] && continue
        NUSERS=$((NUSERS + 1))

        local DB
        for DB in "${DBS[@]}"; do
            evidence_note "$DB" "ActivitiesCache"
            [[ -f "${DB}-wal" ]] && evidence_note "${DB}-wal" "ActivitiesCache WAL"
            local OUT RC
            OUT=$("$PY3" "$PARSER" "$DB" 2>/dev/null); RC=$?
            # RC != 0 copre sia l'ERROR esplicito sia un crash dell'interprete
            # che non arriva a stampare nulla (stdout vuoto in entrambi i
            # casi): senza controllare l'exit code sarebbe indistinguibile da
            # un database genuinamente senza righe.
            if [[ "$RC" -ne 0 ]]; then
                READ_ERRORS=$((READ_ERRORS + 1))
                warn "$(L "Impossibile leggere" "Could not read") $DB"
                continue
            fi
            [[ -z "$OUT" ]] && continue
            while IFS=$'\x1f' read -r TS APP ATYPE DESC END SUSP; do
                [[ -z "$TS$APP" ]] && continue
                TOTAL=$((TOTAL + 1))
                [[ "$SUSP" == "1" ]] && SUSP_COUNT=$((SUSP_COUNT + 1))
                ALL_ENTRIES+=("${TS}"$'\x1f'"${APP}"$'\x1f'"${ATYPE}"$'\x1f'"${DESC}"$'\x1f'"${END}"$'\x1f'"${SUSP}"$'\x1f'"${USERNAME}")
            done <<< "$OUT"
        done
    done < <(get_user_homes)

    separator
    info "$(L "Utenti con ActivitiesCache.db:" "Users with ActivitiesCache.db:") ${BOLD}$NUSERS${RESET}  |  $(L "Attività:" "Activities:") ${BOLD}$TOTAL${RESET}  |  $(L "Da percorso sospetto:" "From suspicious path:") ${RED}${BOLD}${SUSP_COUNT}${RESET}"
    [[ "$READ_ERRORS" -gt 0 ]] && warn "$(L "Database trovati ma non leggibili:" "Databases found but unreadable:") ${BOLD}${READ_ERRORS}${RESET}"
    if [[ "$TOTAL" -eq 0 ]]; then
        if [[ "$READ_ERRORS" -gt 0 ]]; then
            warn "$(L "Nessuna attività estratta: i database trovati non erano leggibili, non e' detto che siano vuoti." "No activity extracted: the databases found could not be read, they are not necessarily empty.")"
        else
            warn "$(L "Nessuna Timeline trovata (ActivitiesCache.db assente o vuoto)." "No Timeline found (ActivitiesCache.db missing or empty).")"
        fi
        return 0
    fi

    ask_yn "$(L "Generare report HTML?" "Generate HTML report?")" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="$(L "La Timeline registra finestre di uso attivo (inizio/fine), non solo l'esecuzione: copre app UWP e documenti che altri artefatti non vedono, ma e' limitata dalla finestra di retention (ExpirationTime), tipicamente alcuni giorni." \
        "The Timeline records active-use windows (start/end), not just execution: it covers UWP apps and documents other artefacts miss, but is bounded by its retention window (ExpirationTime), typically a few days.")"
    BODY+="</div></div>"

    local ROWS=""
    local E
    for E in "${ALL_ENTRIES[@]}"; do
        IFS=$'\x1f' read -r TS APP ATYPE DESC END SUSP USERNAME <<< "$E"
        local APP_CSS="class='mono'"
        [[ "$SUSP" == "1" ]] && APP_CSS="class='mono bad'"
        ROWS+="<tr><td class='mono dim' style='white-space:nowrap'>$(html_esc "$TS")</td>
          <td class='mono dim' style='white-space:nowrap'>$(html_esc "$USERNAME")</td>
          <td ${APP_CSS} style='word-break:break-all'>$(html_esc "$APP")</td>
          <td class='mono dim'>$(html_esc "$ATYPE")</td>
          <td class='mono'>$(html_esc "$DESC")</td></tr>"
    done

    BODY+="<div class='stitle'>$(L "Attività registrate" "Recorded activities")</div>
    <div class='card'><table>
      <thead><tr><th>$(L "Inizio" "Start")</th><th>$(L "Utente" "User")</th><th>$(L "App / percorso" "App / path")</th><th>$(L "Tipo" "Type")</th><th>$(L "Descrizione" "Description")</th></tr></thead>
      <tbody>${ROWS}</tbody>
    </table></div>"

    local STATS
    STATS="$(stat_box "$(L "Attività" "Activities")" "$TOTAL" "info")"
    STATS+="$(stat_box "$(L "Sospette" "Suspicious")" "$SUSP_COUNT" "$([[ "$SUSP_COUNT" -gt 0 ]] && echo warn || echo info)")"
    finish_report "win10_timeline" "Windows 10 Timeline" "TL" "AppData/Local/ConnectedDevicesPlatform/*/ActivitiesCache.db" "$STATS" "$BODY"
}
