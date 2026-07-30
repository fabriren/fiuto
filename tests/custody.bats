#!/usr/bin/env bats
#
# Test della catena di custodia.
#
# Il manifesto e' il documento che rende verificabile un report forense:
# questi test presidiano che nulla di cio' che viene letto sfugga al
# tracciamento e che i percorsi di errore non lo interrompano.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    REPORT_BASE_DIR="$FIXTURE/report"
    mkdir -p "$REPORT_BASE_DIR" "$FIXTURE/vol/etc"
    printf 'root:x:0:0::/root:/bin/bash\n' > "$FIXTURE/vol/etc/passwd"
    WIN_ROOT="$FIXTURE/vol"
    OS_TYPE="linux"
    CUSTODY=true
    CUSTODY_HASH=true
    rm -rf "${TMPDIR:-/tmp}/fiuto_custody_$$"
}

teardown() {
    rm -rf "$FIXTURE" "${TMPDIR:-/tmp}/fiuto_custody_$$"
}

@test "evidence_note registra un file esistente" {
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    [ "$(custody_count)" -eq 1 ]
}

@test "evidence_note ignora i file inesistenti" {
    evidence_note "$FIXTURE/vol/etc/non-esiste" "letto"
    [ "$(custody_count)" -eq 0 ]
}

@test "evidence_note ignora un percorso vuoto" {
    evidence_note "" "letto"
    [ "$(custody_count)" -eq 0 ]
}

# Regressione: l'esclusione dei file di lavoro usava /tmp/*, che scartava
# qualunque evidenza su un volume montato sotto /tmp — caso legittimo.
@test "un volume montato sotto /tmp viene comunque tracciato" {
    tmpvol="$(mktemp -d -p /tmp)"
    printf 'x' > "$tmpvol/artefatto.log"
    evidence_note "$tmpvol/artefatto.log" "letto"
    n="$(custody_count)"
    rm -rf "$tmpvol"
    [ "$n" -eq 1 ]
}

@test "i file di lavoro di FIUTO non sono trattati come evidenza" {
    d="$(_custody_dir)"
    printf 'x' > "$d/lavoro.tmp"
    evidence_note "$d/lavoro.tmp" "letto"
    [ "$(custody_count)" -eq 0 ]
}

@test "ci_find_file registra automaticamente il file trovato" {
    run ci_find_file "$FIXTURE/vol" "etc/passwd"
    [ -n "$output" ]
    [ "$(custody_count)" -eq 1 ]
}

@test "con CUSTODY=false non viene registrato nulla" {
    CUSTODY=false
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    [ "$(custody_count)" -eq 0 ]
}

@test "il manifesto e' JSON valido e contiene i campi attesi" {
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    write_evidence_manifest > /dev/null 2>&1
    [ -f "$REPORT_BASE_DIR/evidence_manifest.json" ]
    python3 -c '
import json, sys
m = json.load(open(sys.argv[1]))
for k in ("manifest_version", "tool", "session", "subject", "integrity_policy",
          "evidence_files", "generated_reports", "counts"):
    assert k in m, f"campo mancante: {k}"
assert m["counts"]["evidence_files"] == 1, m["counts"]
assert m["evidence_files"][0]["sha256"], "hash mancante"
assert len(m["evidence_files"][0]["sha256"]) == 64
' "$REPORT_BASE_DIR/evidence_manifest.json"
}

@test "l'hash nel manifesto corrisponde al file" {
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    write_evidence_manifest > /dev/null 2>&1
    atteso="$(sha256sum "$FIXTURE/vol/etc/passwd" | awk '{print $1}')"
    trovato="$(python3 -c '
import json, sys
print(json.load(open(sys.argv[1]))["evidence_files"][0]["sha256"])' "$REPORT_BASE_DIR/evidence_manifest.json")"
    [ "$atteso" = "$trovato" ]
}

# Con --no-hash il file deve comparire lo stesso: sparire dal manifesto
# sarebbe peggio che comparire senza impronta.
@test "con --no-hash il file resta nel manifesto, dichiarato senza hash" {
    CUSTODY_HASH=false
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    write_evidence_manifest > /dev/null 2>&1
    python3 -c '
import json, sys
e = json.load(open(sys.argv[1]))["evidence_files"][0]
assert e["sha256"] is None, e
assert e["sha256_omitted"], "manca la motivazione"
' "$REPORT_BASE_DIR/evidence_manifest.json"
}

@test "oltre la soglia di dimensione l'hash e' omesso con motivazione" {
    CUSTODY_HASH_LIMIT_MB=0
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    write_evidence_manifest > /dev/null 2>&1
    python3 -c '
import json, sys
e = json.load(open(sys.argv[1]))["evidence_files"][0]
assert e["sha256"] is None, e
assert "soglia" in e["sha256_omitted"], e
' "$REPORT_BASE_DIR/evidence_manifest.json"
}

@test "lo stesso file annotato piu' volte compare una volta sola" {
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    evidence_note "$FIXTURE/vol/etc/passwd" "plist"
    write_evidence_manifest > /dev/null 2>&1
    python3 -c '
import json, sys
m = json.load(open(sys.argv[1]))
assert m["counts"]["evidence_files"] == 1, m["counts"]
assert len(m["evidence_files"][0]["roles"]) == 2, m["evidence_files"][0]
' "$REPORT_BASE_DIR/evidence_manifest.json"
}

@test "senza evidenze non viene scritto alcun manifesto" {
    write_evidence_manifest > /dev/null 2>&1
    [ ! -f "$REPORT_BASE_DIR/evidence_manifest.json" ]
}

@test "write_evidence_manifest non emette diagnostica su stdout non pertinente" {
    evidence_note "$FIXTURE/vol/etc/passwd" "letto"
    out="$(write_evidence_manifest 2>/dev/null)"
    [[ "$out" != *"	"* ]]
}
