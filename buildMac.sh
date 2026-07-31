#!/usr/bin/env bash
#
# Genera la build macOS di FIUTO: uno script che gira SU un Mac usato come
# workstation di analisi.
#
# Non e' un fork. Concatena esattamente gli stessi sorgenti della build Linux,
# nell'ordine di src/build.order.macos, con due sole differenze:
#
#   1. src/macos-header.sh al posto di src/header.sh — shebang su bash di
#      Homebrew, perche' macOS spedisce ancora la 3.2 del 2007 e FIUTO usa
#      array associativi, ${var,,}, mapfile, declare -n e wait -n;
#   2. src/compat/macos.sh inserito prima di main.sh — le funzioni che
#      traducono gli strumenti GNU verso l'userland BSD.
#
# Il compat sta in FONDO e non in testa deliberatamente. In bash i nomi di
# funzione si risolvono alla chiamata, non alla definizione, quindi mettere le
# ridefinizioni alla fine e' sufficiente perche' valgano ovunque. Metterle in
# testa non basterebbe, anzi: alcune sarebbero sovrascritte dai file
# successivi che definiscono le stesse funzioni (_find_user_cwd, image_open,
# _find_windows_mounts). L'ultimo che parla vince, e deve essere il compat.
#
# Uso:
#   ./buildMac.sh                genera ./fiuto-macos.sh
#   ./buildMac.sh --check        verifica che sia allineato ai sorgenti
#   ./buildMac.sh -o FILE        genera su un percorso diverso

set -uo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ORDER_FILE="${REPO_DIR}/src/build.order.macos"
OUTPUT="${REPO_DIR}/fiuto-macos.sh"
MODE="build"

while [[ $# -gt 0 ]]; do
    case "$1" in
        --check) MODE="check" ;;
        -o)      OUTPUT="$2"; shift ;;
        -h|--help)
            sed -n '2,25p' "$0" | sed 's/^# \{0,1\}//'
            exit 0 ;;
        *) echo "buildMac.sh: opzione sconosciuta: $1" >&2; exit 2 ;;
    esac
    shift
done

[[ -f "$ORDER_FILE" ]] || { echo "buildMac.sh: manca $ORDER_FILE" >&2; exit 1; }

# Ogni sorgente elencato deve esistere: un percorso sbagliato produrrebbe uno
# script monco che parte comunque, con i moduli mancanti spariti in silenzio.
MISSING=()
while IFS= read -r SRC; do
    [[ -z "$SRC" || "$SRC" == \#* ]] && continue
    [[ -f "${REPO_DIR}/${SRC}" ]] || MISSING+=("$SRC")
done < "$ORDER_FILE"
if [[ ${#MISSING[@]} -gt 0 ]]; then
    echo "buildMac.sh: sorgenti elencati ma inesistenti:" >&2
    printf '  %s\n' "${MISSING[@]}" >&2
    exit 1
fi

# E il contrario: un sorgente che esiste ma non e' elencato resterebbe fuori
# dalla build macOS pur essendo in quella Linux. E' la deriva piu' facile da
# introdurre aggiungendo un modulo, e la piu' difficile da notare.
UNLISTED=()
while IFS= read -r F; do
    REL="${F#"$REPO_DIR"/}"
    [[ "$REL" == "src/header.sh" ]] && continue   # sostituito da macos-header.sh
    grep -qxF "$REL" "$ORDER_FILE" || UNLISTED+=("$REL")
done < <(find "${REPO_DIR}/src" -name '*.sh' -type f | sort)
if [[ ${#UNLISTED[@]} -gt 0 ]]; then
    echo "buildMac.sh: sorgenti non elencati in src/build.order.macos:" >&2
    printf '  %s\n' "${UNLISTED[@]}" >&2
    exit 1
fi

TMP=$(mktemp) || exit 1
trap 'rm -f "$TMP"' EXIT

while IFS= read -r SRC; do
    [[ -z "$SRC" || "$SRC" == \#* ]] && continue
    cat "${REPO_DIR}/${SRC}" >> "$TMP"
done < "$ORDER_FILE"

N=$(grep -cve '^[[:space:]]*$' "$ORDER_FILE")

if [[ "$MODE" == "check" ]]; then
    if [[ ! -f "$OUTPUT" ]]; then
        echo "buildMac.sh: $OUTPUT non esiste — esegui ./buildMac.sh" >&2
        exit 1
    fi
    if diff -q "$TMP" "$OUTPUT" > /dev/null; then
        echo "buildMac.sh: $(basename "$OUTPUT") è allineato ai sorgenti ($N file)"
        exit 0
    fi
    echo "buildMac.sh: $(basename "$OUTPUT") NON è allineato ai sorgenti." >&2
    diff -u "$OUTPUT" "$TMP" | head -40 >&2
    exit 1
fi

cp "$TMP" "$OUTPUT"
chmod +x "$OUTPUT"
echo "buildMac.sh: generato $(basename "$OUTPUT") da $N sorgenti ($(wc -l < "$OUTPUT") righe)"
