#!/usr/bin/env bats
#
# Test strutturali sul sorgente di fiuto.sh.
#
# Non esercitano la logica forense: verificano invarianti del codice che, se
# violate, producono fallimenti silenziosi a runtime (moduli che non partono,
# funzioni sovrascritte, voci di menu che puntano nel vuoto).

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    SCRIPT="$REPO_ROOT/fiuto.sh"
    # shellcheck disable=SC1091
    source "$SCRIPT"
}

# Regressione: ci_find_file era definita due volte con semantiche incompatibili.
# In bash vince l'ultima, quindi il bug è invisibile fino all'uso in campo.
@test "nessuna funzione è definita due volte" {
    dups="$(grep -oE '^[a-zA-Z_][a-zA-Z0-9_]*\(\)' "$SCRIPT" | sort | uniq -d)"
    [ -z "$dups" ] || {
        echo "Funzioni duplicate: $dups"
        false
    }
}

@test "ogni voce del registro Linux punta a una funzione esistente" {
    for entry in "${MODULES_LINUX[@]}"; do
        fn="${entry%%|*}"
        declare -F "$fn" > /dev/null || {
            echo "Funzione mancante nel registro Linux: $fn"
            false
        }
    done
}

@test "ogni voce del registro macOS punta a una funzione esistente" {
    for entry in "${MODULES_MACOS[@]}"; do
        fn="${entry%%|*}"
        declare -F "$fn" > /dev/null || {
            echo "Funzione mancante nel registro macOS: $fn"
            false
        }
    done
}

@test "ogni voce di registro ha i 4 campi attesi" {
    for entry in "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}"; do
        IFS='|' read -r fn name color desc <<< "$entry"
        [ -n "$fn" ]
        [ -n "$name" ]
        [ -n "$color" ]
        [ -n "$desc" ]
    done
}

@test "i nomi dei moduli sono univoci dentro ogni registro" {
    for reg in MODULES_LINUX MODULES_MACOS; do
        declare -n _r="$reg"
        names="$(printf '%s\n' "${_r[@]}" | cut -d'|' -f2 | sort)"
        [ "$(echo "$names" | wc -l)" -eq "$(echo "$names" | sort -u | wc -l)" ] || {
            echo "Nomi duplicati nel registro $reg"
            false
        }
        unset -n _r
    done
}

@test "ogni numero del menu Windows invoca una funzione esistente" {
    # Estrae le righe "N)  module_xxx ;;" dal dispatcher run_module_by_number.
    while read -r fn; do
        declare -F "$fn" > /dev/null || {
            echo "Funzione Windows mancante: $fn"
            false
        }
    done < <(sed -n '/^run_module_by_number()/,/^}/p' "$SCRIPT" \
             | grep -oE '\bmodule_[a-z0-9_]+' | sort -u)
}

@test "lo script non contiene marcatori di conflitto git" {
    ! grep -qE '^(<<<<<<<|=======|>>>>>>>)$' "$SCRIPT"
}

@test "lo script dichiara bash come interprete e set -uo pipefail" {
    head -1 "$SCRIPT" | grep -q 'bash'
    grep -q 'set -uo pipefail' "$SCRIPT"
}
