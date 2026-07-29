#!/usr/bin/env bats
#
# Test dell'export JSONL (schema Timesketch).
#
# L'export lavora sull'HTML già prodotto dai moduli: qui si costruisce un
# report sintetico che riproduce le due strutture riconosciute (tabelle e
# blocchi <pre>) e si verifica lo schema in uscita.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    REPORT_BASE_DIR="$FIXTURE"
    RDIR="$FIXTURE/evtx_20260729_120000"
    mkdir -p "$RDIR"
    EXPORT_JSONL=true
    WIN_ROOT="/mnt/disk"
    HOST_NAME="WS-01"
    OS_TYPE="windows"
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

_write_report() {
    cat > "$RDIR/report.html" <<'EOF'
<html><body>
<table>
<tr><td>2026-03-15 08:45:32</td><td>4688</td><td>powershell.exe -enc AAA</td></tr>
<tr><td>2026-03-15T14:22:15</td><td>4768</td><td>Kerberos failure</td></tr>
<tr><td>nessuna data</td><td>0</td><td>riga senza timestamp</td></tr>
</table>
<pre>
    1  Mar 15 09:01:02 host sshd[123]: Failed password for root
    2  riga senza data
</pre>
</body></html>
EOF
}

@test "l'export produce un record per ogni evento con timestamp" {
    _write_report
    run register_report "$RDIR/report.html"
    [ -f "$RDIR/report.jsonl" ]
    [ "$(wc -l < "$RDIR/report.jsonl")" -eq 3 ]
}

@test "ogni riga è JSON valido" {
    _write_report
    register_report "$RDIR/report.html"
    while IFS= read -r line; do
        echo "$line" | python3 -c 'import json,sys; json.loads(sys.stdin.read())'
    done < "$RDIR/report.jsonl"
}

@test "sono presenti i campi richiesti da Timesketch" {
    _write_report
    register_report "$RDIR/report.html"
    head -1 "$RDIR/report.jsonl" | python3 -c '
import json, sys
rec = json.loads(sys.stdin.read())
for field in ("datetime", "timestamp_desc", "message"):
    assert field in rec and rec[field], f"campo mancante: {field}"
'
}

@test "i timestamp sono normalizzati in ISO8601" {
    _write_report
    register_report "$RDIR/report.html"
    python3 -c '
import json, re, sys
pat = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$")
for line in open(sys.argv[1]):
    rec = json.loads(line)
    assert pat.match(rec["datetime"]), rec["datetime"]
' "$RDIR/report.jsonl"
}

@test "gli eventi sono ordinati cronologicamente" {
    _write_report
    register_report "$RDIR/report.html"
    python3 -c '
import json, sys
ts = [json.loads(l)["datetime"] for l in open(sys.argv[1])]
assert ts == sorted(ts), ts
' "$RDIR/report.jsonl"
}

# Un timestamp syslog non contiene l'anno: viene dedotto, e questo va
# dichiarato nel record invece di essere presentato come dato certo.
@test "l'anno dedotto dai log syslog è marcato esplicitamente" {
    _write_report
    register_report "$RDIR/report.html"
    grep -q '"year_inferred": true' "$RDIR/report.jsonl"
}

@test "le righe senza timestamp non finiscono nella timeline" {
    _write_report
    register_report "$RDIR/report.html"
    ! grep -q "riga senza timestamp" "$RDIR/report.jsonl"
    ! grep -q "riga senza data" "$RDIR/report.jsonl"
}

@test "gli eventi confluiscono nella timeline unica di sessione" {
    _write_report
    register_report "$RDIR/report.html"
    [ -f "$REPORT_BASE_DIR/fiuto_timeline.jsonl" ]
    [ "$(wc -l < "$REPORT_BASE_DIR/fiuto_timeline.jsonl")" -eq 3 ]
}

@test "il contesto del volume è riportato in ogni record" {
    _write_report
    register_report "$RDIR/report.html"
    head -1 "$RDIR/report.jsonl" | python3 -c '
import json, sys
rec = json.loads(sys.stdin.read())
assert rec["source_volume"] == "/mnt/disk", rec
assert rec["hostname"] == "WS-01", rec
assert rec["module"] == "evtx", rec
'
}

@test "senza --jsonl non viene prodotto alcun file" {
    _write_report
    EXPORT_JSONL=false
    register_report "$RDIR/report.html"
    [ ! -f "$RDIR/report.jsonl" ]
}

@test "un report senza eventi datati non lascia un jsonl vuoto" {
    cat > "$RDIR/report.html" <<'EOF'
<html><body><table><tr><td>nessuna</td><td>data</td></tr></table></body></html>
EOF
    register_report "$RDIR/report.html"
    [ ! -f "$RDIR/report.jsonl" ]
}

@test "register_report continua a registrare il report anche con export attivo" {
    _write_report
    GENERATED_REPORTS=()
    register_report "$RDIR/report.html"
    [ "${#GENERATED_REPORTS[@]}" -eq 1 ]
}
