#!/usr/bin/env bats
#
# Test strutturali dell'immagine.
#
# Non costruiscono l'immagine — è compito della CI — ma presidiano la deriva
# che la costruirebbe sbagliata: FIUTO controlla `import X` per nove librerie,
# e se una sparisce da requirements.txt l'immagine si costruisce lo stesso e il
# modulo relativo degrada in silenzio. Che è esattamente ciò che l'immagine
# dovrebbe rendere impossibile.

setup() {
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
}

# Nomi dei moduli Python di cui FIUTO verifica la presenza a runtime.
_checked_imports() {
    grep -rhoE '"\$PY3" -c "import [a-zA-Z_]+"' "$REPO_ROOT/src" \
        | sed 's/.*import //; s/"//' | sort -u
}

# Il nome del pacchetto PyPI non coincide con quello del modulo importato.
_pypi_for() {
    case "$1" in
        Evtx)     echo "python-evtx" ;;
        pyesedb)  echo "libesedb-python" ;;
        pypff)    echo "libpff-python" ;;
        yaml)     echo "PyYAML" ;;
        yara)     echo "yara-python" ;;
        snappy)   echo "python-snappy" ;;
        *)        echo "$1" ;;
    esac
}

@test "ogni libreria che FIUTO controlla è in requirements.txt" {
    fail=0
    while IFS= read -r mod; do
        pkg=$(_pypi_for "$mod")
        grep -qiE "^${pkg}([=<>~]|\$)" "$REPO_ROOT/requirements.txt" \
            || { echo "manca in requirements.txt: $pkg (import $mod)"; fail=1; }
    done < <(_checked_imports)
    [ "$fail" -eq 0 ]
}

@test "ogni dipendenza ha una versione fissata" {
    # Un parser che cambia fra due build produce report diversi sullo stesso
    # disco, e in ambito peritale la differenza va spiegata.
    fail=0
    while IFS= read -r line; do
        [[ -z "${line//[[:space:]]/}" || "$line" == \#* ]] && continue
        [[ "$line" == *"=="* ]] || { echo "senza pin: $line"; fail=1; }
    done < "$REPO_ROOT/requirements.txt"
    [ "$fail" -eq 0 ]
}

@test "l'entrypoint elenca tutte le librerie controllate da FIUTO" {
    # Se --deps ne dimenticasse una, la CI non si accorgerebbe della sua
    # assenza nell'immagine: il controllo si basa su quell'elenco.
    fail=0
    while IFS= read -r mod; do
        grep -q "'${mod}'" "$REPO_ROOT/docker-entrypoint.sh" \
            || { echo "--deps non elenca: $mod"; fail=1; }
    done < <(_checked_imports)
    [ "$fail" -eq 0 ]
}

@test "il Dockerfile verifica che fiuto.sh sia allineato ai sorgenti" {
    # Un'immagine costruita a mano da un albero sporco conterrebbe altrimenti
    # un fiuto.sh diverso dai suoi sorgenti, senza modo di accorgersene.
    grep -q 'build.sh --check' "$REPO_ROOT/Dockerfile"
}

@test "il Dockerfile non lascia i compilatori nell'immagine finale" {
    grep -q 'apt-get purge -y --auto-remove build-essential' "$REPO_ROOT/Dockerfile"
}

@test "la directory dei report è scrivibile da un --user arbitrario" {
    # Senza, chi lancia con --user 1000:1000 non potrebbe creare nulla.
    grep -qE 'chmod 777 /report' "$REPO_ROOT/Dockerfile"
}

@test "l'immagine viene pubblicata solo dai tag" {
    # Un latest che cambia a ogni push renderebbe irriproducibile un'analisi
    # fatta la settimana prima.
    ci="$REPO_ROOT/.github/workflows/ci.yml"
    grep -q "startsWith(github.ref, 'refs/tags/v')" "$ci"
    [ "$(grep -c "startsWith(github.ref, 'refs/tags/v')" "$ci")" -eq 2 ]
}

@test "la CI fallisce se una dipendenza manca dall'immagine" {
    grep -q 'grep -q MANCANTE' "$REPO_ROOT/.github/workflows/ci.yml"
}

@test "il contesto di build esclude ciò che non serve all'immagine" {
    for p in .git tests docs; do
        grep -qx "$p" "$REPO_ROOT/.dockerignore" || { echo "manca in .dockerignore: $p"; false; }
    done
}
