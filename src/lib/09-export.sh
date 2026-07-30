
# ----------------------------------------------------------------
#  Export JSONL (schema Timesketch / plaso)
#
#  I report HTML sono ottimi per l'analista e inutilizzabili per una
#  pipeline: non si correlano con altre sorgenti e non si caricano in un
#  SIEM. Qui gli stessi eventi vengono riemessi in JSON Lines con i campi
#  attesi da Timesketch (datetime, timestamp_desc, message), cosi' il
#  risultato di FIUTO entra direttamente in una super-timeline.
#
#  L'estrazione lavora sull'HTML gia' prodotto invece che sui dati grezzi
#  dei singoli moduli: e' l'unico punto in cui il formato e' omogeneo per
#  tutti e 91 i moduli, e non richiede di toccarli uno per uno. Il codice
#  che la esegue sta in pylib_timeline, condiviso con l'executive summary.
# ----------------------------------------------------------------
export_report_jsonl() {
    local HTML="$1"
    local DIR; DIR=$(dirname "$HTML")
    local SLUG; SLUG=$(basename "$DIR" | sed -E 's/_[0-9]{8}_[0-9]{6}$//')
    local OUT="${DIR}/report.jsonl"

    local DROPF; DROPF=$(mktemp); register_tmp "$DROPF"
    FIUTO_SINCE="${TIME_SINCE:-}" FIUTO_UNTIL="${TIME_UNTIL:-}" FIUTO_TZ="${VOLUME_TZ:-}" \
    FIUTO_DROPFILE="$DROPF" \
    run_py_with_lib pylib_timeline "$HTML" "$SLUG" "${WIN_ROOT:-}" "${HOST_NAME:-}" "${OS_TYPE:-}" > "$OUT" << 'PYEOF' 2>/dev/null
import sys, os, json

html_path, slug = sys.argv[1], sys.argv[2]
volume  = sys.argv[3] if len(sys.argv) > 3 else ''
host    = sys.argv[4] if len(sys.argv) > 4 else ''
os_type = sys.argv[5] if len(sys.argv) > 5 else ''

# La stessa finestra applicata ai report HTML: se l'export non la rispettasse,
# la timeline caricata in Timesketch conterrebbe eventi che il report esclude.
since = os.environ.get('FIUTO_SINCE', '')
until = os.environ.get('FIUTO_UNTIL', '')
volume_tz = os.environ.get('FIUTO_TZ', '')

dropped = 0

for ev in extract_events(html_path):
    dt = ev['datetime']
    if (since and dt < since) or (until and dt > until):
        dropped += 1
        continue
    rec = {
        # Campi richiesti da Timesketch
        "datetime": dt,
        "timestamp_desc": f"FIUTO {slug}",
        "message": ev['message'],
        # Contesto aggiuntivo
        "data_type": f"fiuto:{slug}",
        "module": slug,
        "source_volume": volume,
        "hostname": host,
        "os": os_type,
    }
    if volume_tz:
        # Il campo datetime NON e' riportato a UTC: dichiarare il fuso del
        # volume e' l'unico modo perche' chi carica la timeline sappia in che
        # riferimento sono gli eventi presi dagli artefatti in ora locale.
        rec["volume_timezone"] = volume_tz
    if ev['year_inferred']:
        # L'anno non era nel dato di origine: va dichiarato, non nascosto.
        rec["year_inferred"] = True
    print(json.dumps(rec, ensure_ascii=False))

drop_file = os.environ.get('FIUTO_DROPFILE')
if drop_file and dropped:
    try:
        with open(drop_file, 'w') as fh:
            fh.write(str(dropped))
    except Exception:
        pass
PYEOF

    if [[ -s "$DROPF" ]]; then
        log_msg "[JSONL] $(cat "$DROPF") eventi esclusi dalla finestra --since/--until"
    fi
    rm -f "$DROPF"

    local N=0
    [[ -s "$OUT" ]] && N=$(wc -l < "$OUT")
    if [[ "$N" -eq 0 ]]; then
        rm -f "$OUT"
        return 0
    fi

    # Timeline unica di sessione: e' il file da caricare in Timesketch.
    # Con --jobs l'append arriva da piu' processi e le righe superano PIPE_BUF:
    # senza lock si intreccerebbero, producendo JSON non parsabile proprio nel
    # file destinato a essere caricato in un altro strumento.
    local COMBINED="${REPORT_BASE_DIR}/fiuto_timeline.jsonl"
    if command -v flock > /dev/null 2>&1; then
        flock "$COMBINED.lock" -c "cat '$OUT' >> '$COMBINED'" 2>/dev/null || \
            cat "$OUT" >> "$COMBINED" 2>/dev/null || true
    else
        cat "$OUT" >> "$COMBINED" 2>/dev/null || true
    fi
    ok "$(L "Export JSONL:" "JSONL export:") ${BOLD}${N}$(L " eventi" " events")${RESET} → $(basename "$OUT")"
    log_msg "[JSONL] $OUT — $N eventi"
}
