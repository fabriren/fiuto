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

@test "ogni sorgente in src/ e' elencato in almeno un build.order" {
    # Da quando esiste la build macOS gli ordini sono due: un sorgente puo'
    # appartenere all'uno, all'altro o a entrambi, ma non a nessuno dei due —
    # in quel caso sparirebbe da ogni prodotto finale senza errori.
    orfani="$(comm -23 <(find src -name '*.sh' | sort) \
        <(cat src/build.order src/build.order.macos | grep -v '^\s*\(#\|$\)' | sort -u))"
    [ -z "$orfani" ] || {
        echo "sorgenti non elencati in nessun build.order: $orfani"
        false
    }
}

@test "i sorgenti solo-macOS non entrano nella build Linux" {
    # macos-header.sh sostituisce header.sh e compat/macos.sh ridefinisce
    # funzioni Linux: finirebbero per rompere la build Linux, non solo
    # appesantirla.
    fail=0
    for f in src/macos-header.sh src/compat/macos.sh; do
        grep -qxF "$f" src/build.order && { echo "$f non deve stare in build.order"; fail=1; }
        grep -qxF "$f" src/build.order.macos || { echo "$f manca da build.order.macos"; fail=1; }
    done
    [ "$fail" -eq 0 ]
}

@test "i sorgenti condivisi sono gli stessi nelle due build" {
    # Se un modulo entrasse in una sola delle due, su una piattaforma
    # semplicemente non esisterebbe, e nessuno se ne accorgerebbe.
    solo_linux="$(comm -23 <(grep -v '^\s*\(#\|$\)' src/build.order | sort) \
                           <(grep -v '^\s*\(#\|$\)' src/build.order.macos | sort))"
    [ "$solo_linux" = "src/header.sh" ] || {
        echo "presenti solo nella build Linux: $solo_linux"
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
