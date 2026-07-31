#!/usr/bin/env bash
#
# Genera fiuto.sh concatenando i sorgenti elencati in src/build.order.
#
# FIUTO si distribuisce come file singolo: si copia su una workstation
# forense e funziona, senza installare nulla. Il sorgente però è diviso per
# sistema operativo in src/, altrimenti un file da oltre 13.000 righe
# diventa impossibile da mantenere.
#
# Il build è una concatenazione pura: nessuna trasformazione, nessuna
# sostituzione. Così il file generato resta leggibile e verificabile, e un
# diff fra due build dice esattamente cosa è cambiato nei sorgenti.
#
# Uso:
#   ./build.sh            genera ./fiuto.sh
#   ./build.sh --check    verifica che ./fiuto.sh sia allineato ai sorgenti
#   ./build.sh -o FILE    genera su un percorso diverso

set -uo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ORDER_FILE="${REPO_DIR}/src/build.order"
OUTPUT="${REPO_DIR}/fiuto.sh"
MODE="build"

while [[ $# -gt 0 ]]; do
    case "$1" in
        --check) MODE="check" ;;
        -o)      OUTPUT="$2"; shift ;;
        -h|--help)
            sed -n '3,22p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
            exit 0
            ;;
        *) echo "Opzione sconosciuta: $1" >&2; exit 2 ;;
    esac
    shift
done

if [[ ! -f "$ORDER_FILE" ]]; then
    echo "build.sh: src/build.order non trovato" >&2
    exit 1
fi

TMP="$(mktemp)"
trap 'rm -f "$TMP"' EXIT

COUNT=0
while IFS= read -r REL; do
    [[ -z "$REL" || "$REL" == \#* ]] && continue
    PART="${REPO_DIR}/${REL}"
    if [[ ! -f "$PART" ]]; then
        echo "build.sh: sorgente mancante: $REL" >&2
        exit 1
    fi
    cat "$PART" >> "$TMP"
    COUNT=$((COUNT + 1))
done < "$ORDER_FILE"

# Ogni sorgente elencato dev'essere davvero usato: un file dimenticato
# fuori da build.order sparirebbe dal prodotto finale senza errori.
#
# I sorgenti della build macOS sono esclusi dal controllo: appartengono a
# src/build.order.macos e in questa build non devono comparire. L'esclusione e'
# per percorso esatto e non per pattern generico, cosi' un file nuovo sotto
# src/compat/ continua a far scattare il guardiano invece di scivolare via.
MACOS_ONLY='src/macos-header.sh
src/compat/macos.sh'
ORPHANS="$(comm -23 \
    <(cd "$REPO_DIR" && find src -name '*.sh' | sort) \
    <(printf '%s\n' "$MACOS_ONLY" | cat - <(grep -v '^\s*\(#\|$\)' "$ORDER_FILE") | sort))"
if [[ -n "$ORPHANS" ]]; then
    echo "build.sh: sorgenti non elencati in src/build.order:" >&2
    echo "$ORPHANS" >&2
    exit 1
fi

if ! bash -n "$TMP"; then
    echo "build.sh: il file generato non è sintatticamente valido" >&2
    exit 1
fi

if [[ "$MODE" == "check" ]]; then
    if diff -q "$TMP" "$OUTPUT" > /dev/null 2>&1; then
        echo "build.sh: $(basename "$OUTPUT") è allineato ai sorgenti ($COUNT file)"
        exit 0
    fi
    echo "build.sh: $(basename "$OUTPUT") NON è allineato ai sorgenti." >&2
    echo "Rigenera con ./build.sh e committa il risultato. Differenze:" >&2
    diff -u "$OUTPUT" "$TMP" | head -40 >&2
    exit 1
fi

cp "$TMP" "$OUTPUT"
chmod +x "$OUTPUT"
echo "build.sh: generato $(basename "$OUTPUT") da $COUNT sorgenti ($(wc -l < "$OUTPUT") righe)"
