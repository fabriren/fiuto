#!/usr/bin/env bats
#
# Test della finestra temporale (--since/--until) e del rilevamento del fuso
# del volume.
#
# Il filtro è pericoloso per costruzione: una riga scartata a torto sparisce
# dal report senza che nessuno se ne accorga. I test presidiano soprattutto i
# casi in cui NON deve scartare — righe senza data, righe con più date, riga
# esattamente sul bordo della finestra.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    TIME_SINCE=""; TIME_UNTIL=""
    VOLUME_TZ=""; VOLUME_TZ_SOURCE=""
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
    rm -rf "${TMPDIR:-/tmp}/fiuto_time_$$"
}

# ---------------------------------------------------------------- parsing ---

@test "una data senza orario diventa inizio giornata come --since" {
    run parse_time_bound "2026-03-01" start
    [ "$status" -eq 0 ]
    [ "$output" = "2026-03-01T00:00:00" ]
}

@test "una data senza orario diventa fine giornata come --until" {
    # Chi scrive --until 2026-03-01 intende includere il 1 marzo per intero.
    run parse_time_bound "2026-03-01" end
    [ "$status" -eq 0 ]
    [ "$output" = "2026-03-01T23:59:59" ]
}

@test "sono accettate le forme con spazio, con T e con Z finale" {
    fail=0
    for IN in "2026-03-01 14:30:00" "2026-03-01T14:30:00" "2026-03-01T14:30:00Z"; do
        OUT=$(parse_time_bound "$IN" start) || { echo "rifiutato: $IN"; fail=1; continue; }
        [ "$OUT" = "2026-03-01T14:30:00" ] || { echo "$IN -> $OUT"; fail=1; }
    done
    [ "$fail" -eq 0 ]
}

@test "una data con ore e minuti viene completata ai secondi" {
    run parse_time_bound "2026-03-01 14:30" start
    [ "$output" = "2026-03-01T14:30:00" ]
    run parse_time_bound "2026-03-01 14:30" end
    [ "$output" = "2026-03-01T14:30:59" ]
}

@test "le forme relative producono un istante nel passato" {
    OUT=$(parse_time_bound "-7d" start)
    [[ "$OUT" =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}$ ]]
    NOW=$(date "+%Y-%m-%dT%H:%M:%S")
    [[ "$OUT" < "$NOW" ]]
}

@test "un limite malformato viene rifiutato invece di essere ignorato" {
    fail=0
    # Una data assurda che passasse silenziosamente filtrerebbe tutto.
    for BAD in "ieri" "01/03/2026" "2026-3-1" "" "-7settimane"; do
        if parse_time_bound "$BAD" start > /dev/null 2>&1; then
            echo "accettato a torto: '$BAD'"
            fail=1
        fi
    done
    [ "$fail" -eq 0 ]
}

# ----------------------------------------------------------------- stato ----

@test "la finestra è inattiva finché non si imposta un limite" {
    run time_window_active
    [ "$status" -ne 0 ]
    TIME_SINCE="2026-03-01T00:00:00"
    run time_window_active
    [ "$status" -eq 0 ]
}

@test "l'etichetta della finestra riporta entrambi i limiti" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-08T23:59:59"
    run time_window_label
    [[ "$output" == *"2026-03-01 00:00:00"* ]]
    [[ "$output" == *"2026-03-08 23:59:59"* ]]
}

# ---------------------------------------------------------------- filtro ----

_rows() {
    printf '%s\n' \
        "2026-02-01 10:00:00	prima della finestra" \
        "2026-03-05 11:22:33	dentro la finestra" \
        "2026-04-10 09:00:00	dopo la finestra" \
        "nessuna data	riga non datata"
}

@test "senza finestra attiva nessuna riga viene toccata" {
    OUT=$(_rows_to_table "$(_rows)" "Data" "Nota")
    [ "$(grep -c "<tr>" <<< "$OUT")" -eq 5 ]   # intestazione + 4 righe
    [[ "$OUT" != *"--since/--until"* ]]
}

@test "le righe fuori dalla finestra non compaiono nella tabella" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    OUT=$(_rows_to_table "$(_rows)" "Data" "Nota")
    [[ "$OUT" == *"dentro la finestra"* ]]
    [[ "$OUT" != *"prima della finestra"* ]]
    [[ "$OUT" != *"dopo la finestra"* ]]
}

@test "una riga senza date viene mantenuta perché non è valutabile" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    OUT=$(_rows_to_table "$(_rows)" "Data" "Nota")
    [[ "$OUT" == *"riga non datata"* ]]
}

@test "il numero di righe nascoste è dichiarato sotto la tabella" {
    # Un filtro silenzioso è indistinguibile da un'assenza di dati.
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    OUT=$(_rows_to_table "$(_rows)" "Data" "Nota")
    [[ "$OUT" == *">2 "*"--since/--until"* ]]
}

@test "una riga con più date sopravvive se almeno una cade nella finestra" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    ROWS="2026-01-01 00:00:00	2026-03-15 12:00:00	creato prima, usato dentro"
    OUT=$(_rows_to_table "$ROWS" "Creato" "Usato" "Nota")
    [[ "$OUT" == *"creato prima, usato dentro"* ]]
}

@test "i limiti sono inclusivi" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-01T23:59:59"
    ROWS=$(printf '%s\n%s' "2026-03-01 00:00:00	bordo inferiore" "2026-03-01 23:59:59	bordo superiore")
    OUT=$(_rows_to_table "$ROWS" "Data" "Nota")
    [[ "$OUT" == *"bordo inferiore"* ]]
    [[ "$OUT" == *"bordo superiore"* ]]
}

@test "una data senza orario è confrontata come inizio giornata" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    ROWS=$(printf '%s\n%s' "2026-03-15	dentro" "2026-05-15	fuori")
    OUT=$(_rows_to_table "$ROWS" "Data" "Nota")
    [[ "$OUT" == *"dentro"* ]]
    [[ "$OUT" != *"fuori"* ]]
}

@test "il conteggio delle righe escluse si accumula sulla sessione" {
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    _rows_to_table "$(_rows)" "Data" "Nota" > /dev/null
    _rows_to_table "$(_rows)" "Data" "Nota" > /dev/null
    [ "$(time_filtered_total)" -eq 4 ]
}

@test "la nota di finestra attiva dichiara i limiti nel report" {
    run time_window_html
    [ -z "$output" ]
    TIME_SINCE="2026-03-01T00:00:00"
    OUT=$(time_window_html)
    [[ "$OUT" == *"2026-03-01 00:00:00"* ]]
}

# ------------------------------------------------------ blocchi <pre> -------

_logfile() {
    cat > "$FIXTURE/auth.log" <<'EOF'
2026-02-01 10:00:00 sshd: Failed password for root
2026-03-05 11:22:33 sshd: Accepted password for mrossi
riga di configurazione senza data
EOF
    echo "$FIXTURE/auth.log"
}

@test "anche i blocchi pre rispettano la finestra" {
    # Senza questo, l'HTML mostrerebbe righe che l'export JSONL esclude:
    # due viste dello stesso modulo che si contraddicono.
    F=$(_logfile)
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    OUT=$(render_pre_block "$F" "")
    [[ "$OUT" == *"Accepted password"* ]]
    [[ "$OUT" != *"Failed password"* ]]
}

@test "nei blocchi pre le righe senza data restano" {
    F=$(_logfile)
    TIME_SINCE="2026-03-01T00:00:00"
    OUT=$(render_pre_block "$F" "")
    [[ "$OUT" == *"configurazione senza data"* ]]
}

@test "i numeri di riga originali sono conservati nei blocchi filtrati" {
    # Il salto di numerazione è il segnale visibile che qualcosa manca.
    F=$(_logfile)
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    OUT=$(render_pre_block "$F" "")
    [[ "$OUT" == *">    2<"* ]]
    [[ "$OUT" != *">    1<"* ]]
    [[ "$OUT" == *"--since/--until"* ]]
}

@test "le righe syslog senza anno sono valutate sull'anno corrente" {
    printf 'Mar  5 11:22:33 srv sshd: dentro\nFeb  1 10:00:00 srv sshd: fuori\n' > "$FIXTURE/syslog"
    Y=$(date +%Y)
    TIME_SINCE="${Y}-03-01T00:00:00"
    TIME_UNTIL="${Y}-03-31T23:59:59"
    OUT=$(render_pre_block "$FIXTURE/syslog" "")
    [[ "$OUT" == *"dentro"* ]]
    [[ "$OUT" != *"fuori"* ]]
}

@test "senza finestra attiva i blocchi pre sono intatti" {
    F=$(_logfile)
    OUT=$(render_pre_block "$F" "")
    [[ "$OUT" == *"Failed password"* ]]
    [[ "$OUT" == *"Accepted password"* ]]
    [[ "$OUT" != *"--since/--until"* ]]
}

# -------------------------------------------------------------- timezone ----

@test "il fuso di un volume Linux è letto da /etc/timezone" {
    mkdir -p "$FIXTURE/etc"
    echo "Europe/Rome" > "$FIXTURE/etc/timezone"
    WIN_ROOT="$FIXTURE"; OS_TYPE="linux"
    detect_volume_timezone
    [ "$VOLUME_TZ" = "Europe/Rome" ]
    [ "$VOLUME_TZ_SOURCE" = "/etc/timezone" ]
}

@test "senza /etc/timezone si ricade sul symlink /etc/localtime" {
    # Su systemd /etc/timezone spesso non esiste.
    mkdir -p "$FIXTURE/etc" "$FIXTURE/usr/share/zoneinfo/Europe"
    touch "$FIXTURE/usr/share/zoneinfo/Europe/Berlin"
    ln -s "../usr/share/zoneinfo/Europe/Berlin" "$FIXTURE/etc/localtime"
    WIN_ROOT="$FIXTURE"; OS_TYPE="linux"
    detect_volume_timezone
    [ "$VOLUME_TZ" = "Europe/Berlin" ]
}

@test "un volume senza indicazioni di fuso non ne inventa uno" {
    WIN_ROOT="$FIXTURE"; OS_TYPE="linux"
    detect_volume_timezone
    [ -z "$VOLUME_TZ" ]
}

# ----------------------------------------------------------------- jsonl ----

@test "l'export JSONL rispetta la stessa finestra dei report" {
    REPORT_BASE_DIR="$FIXTURE"
    RDIR="$FIXTURE/evtx_20260729_120000"; mkdir -p "$RDIR"
    EXPORT_JSONL=true; WIN_ROOT="/mnt/disk"; HOST_NAME="WS-01"; OS_TYPE="windows"
    cat > "$RDIR/report.html" <<'EOF'
<html><body><table>
<tr><td>2026-02-01 10:00:00</td><td>fuori</td></tr>
<tr><td>2026-03-05 11:22:33</td><td>dentro</td></tr>
</table></body></html>
EOF
    TIME_SINCE="2026-03-01T00:00:00"
    TIME_UNTIL="2026-03-31T23:59:59"
    register_report "$RDIR/report.html"
    [ "$(wc -l < "$RDIR/report.jsonl")" -eq 1 ]
    grep -q "dentro" "$RDIR/report.jsonl"
}

@test "il fuso del volume è dichiarato in ogni record JSONL" {
    REPORT_BASE_DIR="$FIXTURE"
    RDIR="$FIXTURE/evtx_20260729_120000"; mkdir -p "$RDIR"
    EXPORT_JSONL=true; WIN_ROOT="/mnt/disk"; OS_TYPE="linux"
    VOLUME_TZ="Europe/Rome"
    cat > "$RDIR/report.html" <<'EOF'
<html><body><table>
<tr><td>2026-03-05 11:22:33</td><td>evento</td></tr>
</table></body></html>
EOF
    register_report "$RDIR/report.html"
    grep -q '"volume_timezone": "Europe/Rome"' "$RDIR/report.jsonl"
}
