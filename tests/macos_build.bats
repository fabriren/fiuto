#!/usr/bin/env bats
#
# Test della build macOS.
#
# Il rischio qui e' doppio e speculare:
#   - che la build macOS diverga da quella Linux, diventando un fork da
#     mantenere due volte;
#   - che una traduzione verso l'userland BSD sia sbagliata in modo silenzioso,
#     restituendo un dato plausibile ma falso. Il caso peggiore e' `stat -c %s`
#     (dimensione) che, tradotto nell'ordine sbagliato, restituisce una data.
#
# Non si puo' eseguire davvero su macOS da qui: cio' che si verifica e' la
# traduzione, con un userland BSD simulato, e la coerenza fra le due build.

setup() {
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    FIXTURE="$(mktemp -d)"
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

# Finto stat BSD che riporta gli argomenti ricevuti: e' il modo per vedere la
# traduzione senza avere un Mac.
_fake_bsd() {
    mkdir -p "$FIXTURE/bin"
    printf '#!/bin/sh\necho "ARGS: $*"\n' > "$FIXTURE/bin/stat"
    chmod +x "$FIXTURE/bin/stat"
}

# Carica il solo strato di compatibilita', senza il resto del tool.
_load_compat() {
    # shellcheck disable=SC1090
    source "$REPO_ROOT/src/compat/macos.sh"
}

# ------------------------------------------------------------- coerenza ----

@test "la build macOS è allineata ai sorgenti" {
    run "$REPO_ROOT/buildMac.sh" --check
    [ "$status" -eq 0 ]
}

@test "le due build condividono gli stessi sorgenti" {
    # L'unica differenza ammessa e' l'intestazione e lo strato di
    # compatibilita'. Se comparissero altre righe solo in una delle due, la
    # build macOS sarebbe diventata un fork.
    extra=$(comm -13 \
        <(sort "$REPO_ROOT/src/build.order") \
        <(sort "$REPO_ROOT/src/build.order.macos"))
    [ "$extra" = "src/compat/macos.sh
src/macos-header.sh" ]
    missing=$(comm -23 \
        <(sort "$REPO_ROOT/src/build.order") \
        <(sort "$REPO_ROOT/src/build.order.macos"))
    [ "$missing" = "src/header.sh" ]
}

@test "lo strato di compatibilità è l'ultimo prima di main.sh" {
    # In bash i nomi di funzione si risolvono alla chiamata, ma le
    # ridefinizioni valgono solo se nessuno le sovrascrive dopo. Il compat
    # ridefinisce _find_user_cwd, image_open e _find_windows_mounts, che sono
    # definite nei file di libreria: deve venire dopo di loro.
    [ "$(tail -2 "$REPO_ROOT/src/build.order.macos" | head -1)" = "src/compat/macos.sh" ]
    [ "$(tail -1 "$REPO_ROOT/src/build.order.macos")" = "src/main.sh" ]
}

@test "la build Linux non contiene lo strato macOS" {
    ! grep -q "STRATO DI COMPATIBILITA' macOS" "$REPO_ROOT/fiuto.sh"
    ! grep -q "opt/homebrew" "$REPO_ROOT/fiuto.sh"
}

@test "la build macOS punta al bash di Homebrew" {
    [ "$(head -1 "$REPO_ROOT/fiuto-macos.sh")" = "#!/opt/homebrew/bin/bash" ]
}

@test "la build macOS rifiuta bash 3.2 con l'istruzione per rimediare" {
    # macOS spedisce ancora la 3.2 del 2007: un errore di sintassi a meta' di
    # un modulo sarebbe incomprensibile.
    grep -q 'BASH_VERSINFO\[0\]' "$REPO_ROOT/fiuto-macos.sh"
    grep -q 'brew install bash' "$REPO_ROOT/fiuto-macos.sh"
}

# ---------------------------------------------------------- traduzioni -----

@test "stat -c %s resta la dimensione e non diventa una data" {
    # GNU %s (dimensione) -> BSD %z, ma GNU %z (ctime) -> BSD %Sc. Tradurre %s
    # per primo produrrebbe un %z che la regola dopo riconvertirebbe: la
    # chiamata restituirebbe una data al posto dei byte, senza errore.
    _fake_bsd; _load_compat
    PATH="$FIXTURE/bin:$PATH"
    run stat -c %s /etc/hosts
    [[ "$output" == *"-f %z"* ]]
    [[ "$output" != *"%Sc"* ]]
}

@test "stat -c %z resta il ctime" {
    _fake_bsd; _load_compat
    PATH="$FIXTURE/bin:$PATH"
    run stat -c '%z' /etc/hosts
    [[ "$output" == *"-f %Sc"* ]]
}

@test "le altre conversioni di stat sono quelle attese" {
    _fake_bsd; _load_compat
    PATH="$FIXTURE/bin:$PATH"
    fail=0
    for pair in "%y:%Sm" "%Y:%m" "%W:%B" "%a:%Lp"; do
        gnu="${pair%%:*}"; bsd="${pair##*:}"
        out=$(stat -c "$gnu" /etc/hosts)
        [[ "$out" == *"-f $bsd"* ]] || { echo "$gnu -> atteso $bsd, ottenuto: $out"; fail=1; }
    done
    [ "$fail" -eq 0 ]
}

@test "stat senza -c non viene tradotto" {
    _fake_bsd; _load_compat
    PATH="$FIXTURE/bin:$PATH"
    run stat /etc/hosts
    [[ "$output" != *"-f"* ]]
}

@test "numfmt --to=iec rende leggibili le dimensioni" {
    _load_compat
    [ "$(numfmt --to=iec 512)" = "512" ]
    [ "$(numfmt --to=iec 1024)" = "1.0K" ]
    [ "$(numfmt --to=iec 5368709120)" = "5.0G" ]
}

@test "realpath -m risolve anche un percorso inesistente" {
    # BSD non ha -m, e senza emulazione la validazione di --report-dir
    # rifiuterebbe qualunque cartella non ancora creata.
    _load_compat
    [ "$(realpath -m /tmp/non/esiste/mai)" = "/tmp/non/esiste/mai" ]
    [ "$(realpath /tmp)" = "/tmp" ]
}

@test "realpath senza -m fallisce su un percorso inesistente" {
    _load_compat
    run realpath /tmp/non/esiste/mai
    [ "$status" -ne 0 ]
}

@test "tutti gli strumenti GNU usati dai sorgenti hanno una traduzione" {
    # Se un modulo nuovo introduce un comando GNU non emulato, su macOS
    # fallisce a runtime: meglio che se ne accorga la suite.
    fail=0
    for tool in stat sha256sum numfmt tac realpath readlink flock; do
        grep -qE "^${tool}\(\)" "$REPO_ROOT/src/compat/macos.sh" \
            || { echo "manca la traduzione di: $tool"; fail=1; }
    done
    [ "$fail" -eq 0 ]
}

@test "le funzioni Linux-only sono ridefinite e non lasciate a metà" {
    # _find_user_cwd legge /proc, _find_windows_mounts legge /proc/mounts,
    # image_open usa losetup: su macOS nessuna delle tre funziona.
    fail=0
    for fn in _find_user_cwd _find_windows_mounts image_open image_cleanup; do
        grep -qE "^${fn}\(\)" "$REPO_ROOT/src/compat/macos.sh" \
            || { echo "non ridefinita per macOS: $fn"; fail=1; }
    done
    [ "$fail" -eq 0 ]
}

@test "la build macOS non contiene riferimenti a /proc nel codice attivo" {
    # Restano solo dentro i moduli Linux, che su un volume Linux montato da un
    # Mac leggono comunque i file del VOLUME, non quelli dell'host.
    ! grep -nE '^[^#]*\/proc\/(mounts|self)' "$REPO_ROOT/src/compat/macos.sh"
}
