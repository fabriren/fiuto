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

# NB: read deve leggere TUTTI i campi. Con una variabile in meno l'ultima
# assorbe il resto della riga, quindi una voce "...|desc||defer" darebbe
# guard="|defer" invece di guard vuota.
@test "le guardie dichiarate nei registri esistono" {
    fail=0
    for entry in "${MODULES_WIN[@]}" "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}"; do
        IFS='|' read -r _fn _name _color _desc guard _flags <<< "$entry"
        [ -z "${guard:-}" ] && continue
        if ! declare -F "$guard" > /dev/null; then
            echo "Guardia inesistente: '$guard' nella voce: $entry"
            fail=1
        fi
    done
    [ "$fail" -eq 0 ]
}

@test "i flag dichiarati nei registri sono riconosciuti" {
    fail=0
    for entry in "${MODULES_WIN[@]}" "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}"; do
        IFS='|' read -r _fn _name _color _desc _guard flags <<< "$entry"
        [ -z "${flags:-}" ] && continue
        case "$flags" in
            defer) ;;
            *) echo "Flag sconosciuto: '$flags' nella voce: $entry"; fail=1 ;;
        esac
    done
    [ "$fail" -eq 0 ]
}

# La Master Timeline aggrega gli altri moduli: se girasse a meta' elenco
# produrrebbe una timeline incompleta senza segnalare nulla.
@test "la Master Timeline e' marcata defer in ogni registro che la contiene" {
    fail=0
    for reg in MODULES_WIN MODULES_LINUX MODULES_MACOS; do
        declare -n _r="$reg"
        for entry in "${_r[@]}"; do
            IFS='|' read -r fn _name _color _desc _guard flags <<< "$entry"
            case "$fn" in
                *master_timeline*)
                    if [[ "${flags:-}" != *defer* ]]; then
                        echo "$reg: $fn non e' marcata defer"
                        fail=1
                    fi
                    ;;
            esac
        done
        unset -n _r
    done
    [ "$fail" -eq 0 ]
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

# L'elenco dei moduli in --help era scritto a mano ed era rimasto fermo a 39
# voci mentre i registri ne contavano 54, 24 e 22. Ora si rende dai registri:
# questi test impediscono che qualcuno lo riscriva a mano.
@test "--help dichiara per ogni OS il numero di moduli del suo registro" {
    out=$(print_module_list)
    fail=0
    for reg in MODULES_WIN MODULES_LINUX MODULES_MACOS; do
        declare -n R="$reg"
        grep -q "(1-${#R[@]})" <<< "$out" || { echo "conteggio errato per $reg: ${#R[@]}"; fail=1; }
        unset -n R
    done
    [ "$fail" -eq 0 ]
}

@test "--help non contiene un conteggio di moduli scritto a mano" {
    ! grep -qE '(Moduli disponibili|Available modules).*\([0-9]+-[0-9]+\)' "$REPO_ROOT/fiuto.sh"
}

@test "ogni nome di modulo dei registri compare in --help" {
    fail=0
    out=$(print_module_list)
    for reg in MODULES_WIN MODULES_LINUX MODULES_MACOS; do
        declare -n R="$reg"
        for entry in "${R[@]}"; do
            IFS='|' read -r _f nm _rest <<< "$entry"
            label=$(reg_text "$nm")
            # I nomi lunghi vengono troncati dalla colonna: basta il prefisso.
            grep -qF "${label:0:20}" <<< "$out" || { echo "assente da --help: $label"; fail=1; }
        done
        unset -n R
    done
    [ "$fail" -eq 0 ]
}

# --report-dir: la cartella dei report deve poter essere fissata da riga di
# comando. Il default e' la directory di invocazione, che analizzando un
# sistema vivo finisce DENTRO il volume analizzato.
@test "--report-dir usa la cartella indicata senza chiedere nulla" {
    d=$(mktemp -d)/report
    REPORT_DIR_FIXED="$d"
    mkdir -p "$d"
    run setup_report_dir
    [ "$status" -eq 0 ]
    [ "$REPORT_BASE_DIR" = "$d" ] || {
        REPORT_DIR_FIXED="$d"; setup_report_dir > /dev/null
        [ "$REPORT_BASE_DIR" = "$d" ]
    }
    rm -rf "$(dirname "$d")"
}

@test "--report-dir non scrivibile e' un errore, non un ripiego sul default" {
    # Ripiegare in silenzio significherebbe scrivere i report da un'altra parte
    # senza dirlo, e su un sistema vivo proprio dentro il volume analizzato.
    run env FIUTO_LIB_ONLY= bash "$REPO_ROOT/fiuto.sh" /tmp --report-dir /proc/non-creabile --all <<< "1"
    [ "$status" -ne 0 ]
    [[ "$output" == *"report"* ]]
}
