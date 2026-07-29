#!/usr/bin/env bats
#
# Test dello script di build.
#
# fiuto.sh e' il prodotto della concatenazione di src/: il rischio concreto e'
# che qualcuno modifichi il file generato invece dei sorgenti e perda il lavoro
# al build successivo. Questi test presidiano l'allineamento.

setup() {
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    cd "$REPO_ROOT"
}

@test "fiuto.sh e' allineato ai sorgenti in src/" {
    run ./build.sh --check
    [ "$status" -eq 0 ]
}

@test "il build e' riproducibile (due esecuzioni danno lo stesso risultato)" {
    a="$(mktemp)"; b="$(mktemp)"
    ./build.sh -o "$a"
    ./build.sh -o "$b"
    run diff -q "$a" "$b"
    rm -f "$a" "$b"
    [ "$status" -eq 0 ]
}

@test "ogni sorgente in src/ e' elencato in build.order" {
    orfani="$(comm -23 <(find src -name '*.sh' | sort) <(grep -v '^\s*\(#\|$\)' src/build.order | sort))"
    [ -z "$orfani" ] || {
        echo "sorgenti non elencati: $orfani"
        false
    }
}

@test "ogni voce di build.order esiste sul disco" {
    while IFS= read -r rel; do
        [[ -z "$rel" || "$rel" == \#* ]] && continue
        [ -f "$rel" ] || {
            echo "sorgente mancante: $rel"
            false
        }
    done < src/build.order
}

@test "il file generato dichiara di non essere modificabile a mano" {
    head -30 fiuto.sh | grep -q "QUESTO FILE E' GENERATO"
}

@test "ogni sorgente e' sintatticamente valido da solo o in concatenazione" {
    run bash -n fiuto.sh
    [ "$status" -eq 0 ]
}
