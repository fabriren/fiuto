#!/usr/bin/env bats
#
# Test del modulo YARA.
#
# Il rischio qui non è tecnico ma interpretativo: "nessun match" su un ambito
# ristretto verrà letto come "il disco è pulito". I test presidiano quindi,
# oltre al funzionamento, il fatto che il report dichiari sempre cosa ha
# guardato e cosa ha saltato.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    REPORT_BASE_DIR="$FIXTURE/reports"; mkdir -p "$REPORT_BASE_DIR"
    VOL="$FIXTURE/vol"; mkdir -p "$VOL/tmp" "$VOL/etc" "$VOL/home/mrossi"
    WIN_ROOT="$VOL"; OS_TYPE="linux"
    BATCH_MODE=true                    # niente prompt
    YARA_RULES=""; YARA_SCAN_PATH=""
    YARA_MAX_MB=64; YARA_MAX_FILES=200000
    GENERATED_REPORTS=()

    cat > "$FIXTURE/rules.yar" <<'EOF'
rule fiuto_test_marker {
    meta:
        description = "stringa piantata dai test"
    strings:
        $a = "MARCATORE_FIUTO_TEST"
    condition:
        $a
}
EOF
    printf 'innocuo\n' > "$VOL/etc/hosts"
    printf 'payload con MARCATORE_FIUTO_TEST dentro\n' > "$VOL/tmp/sospetto.bin"
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

# Saltare i test quando yara-python manca tiene la suite eseguibile su una
# workstation senza dipendenze. In CI lo stesso comportamento farebbe sparire
# nove test senza che nulla lo segnali: FIUTO_TEST_REQUIRE_DEPS lo vieta.
_need_yara() {
    python3 -c "import yara" 2>/dev/null && return 0
    if [[ -n "${FIUTO_TEST_REQUIRE_DEPS:-}" ]]; then
        echo "yara-python assente ma FIUTO_TEST_REQUIRE_DEPS è impostata"
        return 1
    fi
    skip "yara-python non installato"
}

# La risposta a "Generare report HTML?" arriva da stdin: in BATCH_MODE ask_yn
# non chiede, quindi il report viene sempre prodotto.
_run_module() { module_xplat_yara; }

_report() { ls -d "$REPORT_BASE_DIR"/yara_*/report.html 2>/dev/null | head -1; }

# ---------------------------------------------------------------- guardia ---

@test "senza --yara la guardia salta il modulo con un motivo" {
    run _guard_yara
    [ "$status" -ne 0 ]
    [ -n "$output" ]
}

@test "con --yara la guardia lascia passare" {
    YARA_RULES="$FIXTURE/rules.yar"
    run _guard_yara
    [ "$status" -eq 0 ]
}

@test "un percorso di regole inesistente è un errore, non un report vuoto" {
    YARA_RULES="$FIXTURE/non-esiste.yar"
    run _run_module
    [ "$status" -ne 0 ]
    [ -z "$(_report)" ]
}

# --------------------------------------------------------------- scansione --

@test "una regola che corrisponde produce il match" {
    _need_yara
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    _run_module
    H=$(_report)
    [ -n "$H" ]
    grep -q "fiuto_test_marker" "$H"
    grep -q "sospetto.bin" "$H"
}

@test "i file che non corrispondono non compaiono fra i match" {
    _need_yara
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    _run_module
    H=$(_report)
    # etc/hosts è stato scansionato ma non ha fatto match: non deve comparire
    # nella tabella dei riscontri.
    ! grep -q "etc/hosts" "$H"
}

@test "il report elenca sempre le posizioni effettivamente scansionate" {
    _need_yara
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    _run_module
    H=$(_report)
    grep -qi "scope of the scan\|ambito effettivo" "$H"
    grep -q "$VOL" "$H"
}

@test "il report avverte che nessun match non equivale a volume pulito" {
    _need_yara
    rm -f "$VOL/tmp/sospetto.bin"
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    _run_module
    H=$(_report)
    grep -qi "whole volume was not scanned\|intero volume" "$H"
}

@test "i file oltre il tetto vengono dichiarati saltati, non ignorati" {
    _need_yara
    # Un payload dentro un file enorme non verrebbe visto: chi legge deve
    # sapere che quel file esiste ed è stato escluso.
    head -c 2000000 /dev/zero > "$VOL/tmp/grosso.bin"
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    YARA_MAX_MB=1
    _run_module
    H=$(_report)
    grep -qi "over the per-file cap: <b>1</b>\|tetto per file: <b>1</b>" "$H"
}

@test "raggiunto il tetto di file la scansione si dichiara parziale" {
    _need_yara
    local i
    for i in $(seq 1 12); do printf 'x\n' > "$VOL/tmp/f$i"; done
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    YARA_MAX_FILES=3
    _run_module
    H=$(_report)
    grep -qi "PARTIAL\|PARZIALE" "$H"
}

@test "un file di regole non compilabile non impedisce agli altri di girare" {
    _need_yara
    mkdir -p "$FIXTURE/rulesdir"
    cp "$FIXTURE/rules.yar" "$FIXTURE/rulesdir/buone.yar"
    printf 'rule rotta { condition: $inesistente }\n' > "$FIXTURE/rulesdir/rotte.yar"
    YARA_RULES="$FIXTURE/rulesdir"
    YARA_SCAN_PATH="$VOL"
    _run_module
    H=$(_report)
    grep -q "fiuto_test_marker" "$H"
    grep -qi "rejected\|scartate" "$H"
}

@test "i symlink non vengono seguiti" {
    _need_yara
    # Seguirli su un volume montato porterebbe fuori dall'evidenza, fino al
    # file system della workstation di analisi.
    mkdir -p "$FIXTURE/fuori"
    printf 'MARCATORE_FIUTO_TEST fuori dal volume\n' > "$FIXTURE/fuori/esterno.bin"
    ln -s "$FIXTURE/fuori/esterno.bin" "$VOL/tmp/link.bin"
    YARA_RULES="$FIXTURE/rules.yar"
    YARA_SCAN_PATH="$VOL"
    _run_module
    H=$(_report)
    ! grep -q "link.bin" "$H"
}

# --------------------------------------------------------------- registri ---

@test "il modulo è presente in tutti e tre i registri con la sua guardia" {
    for reg in MODULES_WIN MODULES_LINUX MODULES_MACOS; do
        declare -n R="$reg"
        printf '%s\n' "${R[@]}" | grep -q "module_xplat_yara|.*|_guard_yara" \
            || { echo "manca o senza guardia in $reg"; false; }
        unset -n R
    done
}
