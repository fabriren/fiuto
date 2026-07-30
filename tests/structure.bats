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

@test "ogni voce dei registri punta a una funzione esistente" {
    for reg in MODULES_WIN MODULES_LINUX MODULES_MACOS; do
        declare -n _r="$reg"
        for entry in "${_r[@]}"; do
            fn="${entry%%|*}"
            declare -F "$fn" > /dev/null || {
                echo "Funzione mancante nel registro $reg: $fn"
                false
            }
        done
        unset -n _r
    done
}

@test "ogni voce di registro ha almeno i 4 campi attesi" {
    for entry in "${MODULES_WIN[@]}" "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}"; do
        IFS='|' read -r fn name color desc _guard <<< "$entry"
        [ -n "$fn" ]
        [ -n "$name" ]
        [ -n "$color" ]
        [ -n "$desc" ]
    done
}

@test "i colori dei registri sono variabili definite" {
    for entry in "${MODULES_WIN[@]}" "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}"; do
        IFS='|' read -r _fn _name color _rest <<< "$entry"
        [ -n "${!color:-}" ] || {
            echo "Colore non definito: $color"
            false
        }
    done
}

@test "le guardie dichiarate nei registri esistono" {
    for entry in "${MODULES_WIN[@]}" "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}"; do
        IFS='|' read -r _fn _name _color _desc guard <<< "$entry"
        [ -z "${guard:-}" ] && continue
        declare -F "$guard" > /dev/null || {
            echo "Guardia mancante: $guard"
            false
        }
    done
}

@test "i nomi dei moduli sono univoci dentro ogni registro" {
    for reg in MODULES_WIN MODULES_LINUX MODULES_MACOS; do
        declare -n _r="$reg"
        names="$(printf '%s\n' "${_r[@]}" | cut -d'|' -f2 | sort)"
        [ "$(echo "$names" | wc -l)" -eq "$(echo "$names" | sort -u | wc -l)" ] || {
            echo "Nomi duplicati nel registro $reg"
            false
        }
        unset -n _r
    done
}

# La numerazione e' un contratto con l'utente: --module N e --modules 1,4,6-8
# sono documentati nel README e usati negli script. Riordinare un registro
# senza aggiornare il README rompe entrambi in silenzio.
@test "il numero di moduli per OS coincide con quanto dichiarato nel README" {
    readme="$REPO_ROOT/README.md"
    grep -q "The ${#MODULES_WIN[@]} Windows analysis modules" "$readme"
    grep -q "The ${#MODULES_LINUX[@]} Linux analysis modules" "$readme"
    grep -q "The ${#MODULES_MACOS[@]} macOS analysis modules" "$readme"
}

# reg_text estrae la variante linguistica dalla forma "italiano§english".
@test "reg_text seleziona la lingua corretta" {
    LANG=it
    [ "$(reg_text 'Persistenza§Persistence')" = "Persistenza" ]
    LANG=en
    [ "$(reg_text 'Persistenza§Persistence')" = "Persistence" ]
}

@test "reg_text lascia intatto il testo non bilingue" {
    LANG=it
    [ "$(reg_text 'SRUM')" = "SRUM" ]
    LANG=en
    [ "$(reg_text 'SRUM')" = "SRUM" ]
}

# Le descrizioni contengono $Recycle.Bin e $UsnJrnl:$J: dentro le virgolette
# doppie del registro il dollaro va escapato, o bash espande una variabile.
@test "i dollari letterali nelle descrizioni sopravvivono" {
    printf '%s\n' "${MODULES_WIN[@]}" | grep -q 'Recycle.Bin'
    printf '%s\n' "${MODULES_WIN[@]}" | grep -q 'UsnJrnl'
}

@test "i dispatcher duplicati sono stati rimossi" {
    ! grep -q '^run_module_by_number()' "$SCRIPT"
    ! grep -q '^print_menu()' "$SCRIPT"
    ! grep -q '^run_all_modules()' "$SCRIPT"
}

@test "lo script non contiene marcatori di conflitto git" {
    ! grep -qE '^(<<<<<<<|=======|>>>>>>>)$' "$SCRIPT"
}

@test "lo script dichiara bash come interprete e set -uo pipefail" {
    head -1 "$SCRIPT" | grep -q 'bash'
    grep -q 'set -uo pipefail' "$SCRIPT"
}
