#!/usr/bin/env bats
#
# Test dell'esecuzione parallela di --all (--jobs N).
#
# Il criterio è uno solo: **l'esito non deve dipendere dal grado di
# parallelismo**. Stessi report, stesso ordine nel riepilogo, e la Master
# Timeline — che aggrega gli altri — deve continuare a vederli tutti. Un
# risultato che cambia con --jobs è un risultato di cui non ci si può fidare,
# e in un contesto peritale è peggio di un tool lento.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    REPORT_BASE_DIR="$FIXTURE"
    BATCH_MODE=true
    GENERATED_REPORTS=()
    SUMMARY_TABLE=()
    JOBS=1
    REDACT=false; REDACT_DEFANG=false; EXPORT_JSONL=false
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

# Moduli finti: scrivono un report e lo registrano, come i veri.
_mk_module() {
    eval "mod_$1() {
        local d=\"\$REPORT_BASE_DIR/$1_20260301_1200$2\"
        mkdir -p \"\$d\"
        printf '<html><body>modulo $1</body></html>' > \"\$d/report.html\"
        register_report \"\$d/report.html\"
    }"
}

# Modulo aggregatore: come la Master Timeline, legge GENERATED_REPORTS.
_mk_aggregator() {
    eval "mod_agg() {
        local d=\"\$REPORT_BASE_DIR/agg_20260301_120099\"
        mkdir -p \"\$d\"
        printf '<html><body>visti: %d</body></html>' \"\${#GENERATED_REPORTS[@]}\" > \"\$d/report.html\"
        register_report \"\$d/report.html\"
    }"
}

_queue() {
    _mk_module alfa 01; _mk_module beta 02; _mk_module gamma 03; _mk_module delta 04
    QUEUE=("1|mod_alfa|Alfa|" "2|mod_beta|Beta|" "3|mod_gamma|Gamma|" "4|mod_delta|Delta|")
}

# ------------------------------------------------------------------ pool ----

@test "il pool produce gli stessi report dell'esecuzione sequenziale" {
    _queue
    JOBS=3
    run_batch_pool 4 "${QUEUE[@]}"
    [ "${#GENERATED_REPORTS[@]}" -eq 4 ]
    for m in alfa beta gamma delta; do
        [ -f "$REPORT_BASE_DIR"/${m}_*/report.html ]
    done
}

@test "il riepilogo resta nell'ordine dei moduli, non di completamento" {
    # Un riepilogo che cambia ordine a ogni esecuzione non è confrontabile fra
    # due analisi della stessa macchina.
    _queue
    JOBS=4
    run_batch_pool 4 "${QUEUE[@]}"
    [ "$(cut -d'|' -f1 <<< "${SUMMARY_TABLE[0]}")" = "1" ]
    [ "$(cut -d'|' -f1 <<< "${SUMMARY_TABLE[1]}")" = "2" ]
    [ "$(cut -d'|' -f1 <<< "${SUMMARY_TABLE[2]}")" = "3" ]
    [ "$(cut -d'|' -f1 <<< "${SUMMARY_TABLE[3]}")" = "4" ]
}

@test "un modulo senza risultati è registrato come tale, non perso" {
    _mk_module alfa 01
    mod_vuoto() { return 0; }
    JOBS=2
    run_batch_pool 2 "1|mod_alfa|Alfa|" "2|mod_vuoto|Vuoto|"
    [ "${#SUMMARY_TABLE[@]}" -eq 2 ]
    [[ "${SUMMARY_TABLE[1]}" == *"|NO|"* ]]
}

@test "un modulo che fallisce non blocca gli altri" {
    _mk_module alfa 01; _mk_module beta 02
    mod_rotto() { return 1; }
    JOBS=3
    run_batch_pool 3 "1|mod_alfa|Alfa|" "2|mod_rotto|Rotto|" "3|mod_beta|Beta|"
    [ "${#GENERATED_REPORTS[@]}" -eq 2 ]
    [[ "${SUMMARY_TABLE[1]}" == *"|NO|"* ]]
}

# ----------------------------------------------------------- deferred -------

@test "il modulo defer gira per ultimo e vede tutti gli altri report" {
    # È il difetto che la prima versione aveva: la Master Timeline girava
    # mentre GENERATED_REPORTS era ancora vuoto e produceva una timeline vuota
    # senza segnalare nulla.
    _queue
    _mk_aggregator
    QUEUE+=("5|mod_agg|Aggregatore|defer")
    JOBS=4
    run_batch_pool 5 "${QUEUE[@]}"
    H=$(ls "$REPORT_BASE_DIR"/agg_*/report.html)
    grep -q "visti: 4" "$H"
}

@test "il defer resta ultimo anche se dichiarato per primo nella coda" {
    _mk_module alfa 01
    _mk_aggregator
    JOBS=4
    run_batch_pool 2 "1|mod_agg|Aggregatore|defer" "2|mod_alfa|Alfa|"
    H=$(ls "$REPORT_BASE_DIR"/agg_*/report.html)
    grep -q "visti: 1" "$H"
}

# --------------------------------------------------------- lock sugli hive --

@test "due richieste concorrenti dello stesso hive non si sovrappongono" {
    # Senza lock il secondo processo leggerebbe la copia ricostruita mentre il
    # primo la sta ancora scrivendo: un hive troncato non dà errore, dà
    # risultati parziali.
    HIVE="$FIXTURE/SYSTEM"
    printf 'regf-finto' > "$HIVE"
    HIVE_REPLAY=true
    # Nessun .LOG1/.LOG2: recover_hive marca l'hive come pulito e restituisce
    # l'originale. Serve solo a esercitare acquisizione e rilascio del lock.
    A=$(recover_hive "$HIVE" 2>/dev/null) &
    B=$(recover_hive "$HIVE" 2>/dev/null) &
    wait
    OUT1=$(recover_hive "$HIVE" 2>/dev/null)
    [ "$OUT1" = "$HIVE" ]
    # Il lock deve essere stato rilasciato: se restasse, la sessione successiva
    # aspetterebbe due minuti a vuoto.
    P=$(_hive_recovered_path "$HIVE")
    [ ! -d "${P}.lock" ]
}

@test "il lock viene rilasciato anche quando l'hive risulta pulito" {
    HIVE="$FIXTURE/SOFTWARE"
    printf 'regf-finto' > "$HIVE"
    HIVE_REPLAY=true
    recover_hive "$HIVE" > /dev/null 2>&1
    P=$(_hive_recovered_path "$HIVE")
    [ ! -d "${P}.lock" ]
    [ -f "${P}.skip" ]
}

# ------------------------------------------------------------------ flag ----

@test "--jobs rifiuta un valore non numerico invece di ignorarlo" {
    # Ignorarlo farebbe girare in sequenza chi crede di avere il parallelo.
    run env FIUTO_LIB_ONLY= bash "$REPO_ROOT/fiuto.sh" /tmp --jobs abc --all <<< "1"
    [ "$status" -ne 0 ]
    [[ "$output" == *"--jobs"* ]]
}

# --------------------------------------------- effetti collaterali batch ----

@test "in batch sequenziale gli effetti collaterali del report avvengono" {
    # Regressione: l'esecuzione batch ridefinisce register_report per far
    # risalire i percorsi dal figlio, e per un intero rilascio quella
    # ridefinizione ha silenziosamente saltato export JSONL e --redact.
    # `--all --jsonl` non produceva alcun JSONL, senza un errore.
    _mk_module alfa 01
    REDACT=true; REDACT_DEFANG=false; EXPORT_JSONL=false; CUSTODY=false
    run_batch_module 1 mod_alfa Alfa 1
    [ -f "$REPORT_BASE_DIR"/alfa_*/report.redacted.html ]
}

@test "nel pool parallelo gli effetti collaterali avvengono per ogni modulo" {
    _mk_module alfa 01; _mk_module beta 02
    REDACT=true; REDACT_DEFANG=false; EXPORT_JSONL=false; CUSTODY=false
    JOBS=2
    run_batch_pool 2 "1|mod_alfa|Alfa|" "2|mod_beta|Beta|"
    [ -f "$REPORT_BASE_DIR"/alfa_*/report.redacted.html ]
    [ -f "$REPORT_BASE_DIR"/beta_*/report.redacted.html ]
}
