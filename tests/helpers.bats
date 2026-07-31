#!/usr/bin/env bats
#
# Unit test degli helper di fiuto.sh.
#
# fiuto.sh viene caricato con FIUTO_LIB_ONLY=1, che definisce le funzioni senza
# avviare il menu interattivo.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

# ---------------------------------------------------------------- escaping

@test "html_esc neutralizza i metacaratteri HTML" {
    run html_esc '<script>alert(1)</script> & "x"'
    [ "$status" -eq 0 ]
    [[ "$output" != *"<script>"* ]]
    [[ "$output" == *"&lt;script&gt;"* ]]
    [[ "$output" == *"&amp;"* ]]
}

@test "html_attr neutralizza anche apici singoli e doppi" {
    run html_attr "onmouseover=\"evil()\" x='y'"
    [ "$status" -eq 0 ]
    [[ "$output" != *'"'* ]]
    [[ "$output" != *"'"* ]]
    [[ "$output" == *"&quot;"* ]]
    [[ "$output" == *"&#39;"* ]]
}

@test "html_esc preserva il testo non pericoloso" {
    run html_esc 'C:\Users\Mario\Documenti'
    [ "$output" = 'C:\Users\Mario\Documenti' ]
}

# ---------------------------------------------------------------- selezione moduli

@test "expand_module_list espande gli intervalli" {
    run expand_module_list "1,3,5-8"
    [ "$status" -eq 0 ]
    [ "$(echo "$output" | tr '\n' ' ')" = "1 3 5 6 7 8 " ]
}

@test "expand_module_list deduplica e ordina numericamente" {
    run expand_module_list "10,2,2,1-3"
    [ "$(echo "$output" | tr '\n' ' ')" = "1 2 3 10 " ]
}

@test "expand_module_list scarta i token non numerici" {
    run expand_module_list "1,pippo,3"
    [ "$(echo "$output" | tr '\n' ' ')" = "1 3 " ]
}

# ---------------------------------------------------------------- ricerca case-insensitive

@test "ci_find_dir risolve i path con case diverso" {
    mkdir -p "$FIXTURE/WINDOWS/system32/Config"
    run ci_find_dir "$FIXTURE" "Windows/System32/config"
    [ -n "$output" ]
    [ -d "$output" ]
}

@test "ci_find_file trova un file per nome semplice" {
    mkdir -p "$FIXTURE/var/log"
    echo "x" > "$FIXTURE/var/log/SysLog"
    run ci_find_file "$FIXTURE/var/log" "syslog"
    [ -n "$output" ]
    [ -f "$output" ]
}

# Regressione: una seconda definizione di ci_find_file con semantiche diverse
# (solo basename, -maxdepth 1) sovrascriveva la prima e rompeva silenziosamente
# le chiamate con percorso relativo multi-segmento.
@test "ci_find_file accetta un percorso relativo multi-segmento" {
    mkdir -p "$FIXTURE/.Local/Share"
    echo "x" > "$FIXTURE/.Local/Share/Recently-Used.xbel"
    run ci_find_file "$FIXTURE" ".local/share/recently-used.xbel"
    [ -n "$output" ]
    [ -f "$output" ]
}

@test "ci_find_file ritorna vuoto se il file non esiste" {
    run ci_find_file "$FIXTURE" "inesistente.txt"
    [ -z "$output" ]
}

# ---------------------------------------------------------------- rilevamento OS

@test "detect_os_type riconosce un volume Windows" {
    mkdir -p "$FIXTURE/Windows/System32" "$FIXTURE/Users"
    run detect_os_type "$FIXTURE"
    [ "$output" = "windows" ]
}

@test "detect_os_type riconosce un volume Linux" {
    mkdir -p "$FIXTURE/etc" "$FIXTURE/var" "$FIXTURE/bin"
    echo 'ID=debian' > "$FIXTURE/etc/os-release"
    run detect_os_type "$FIXTURE"
    [ "$output" = "linux" ]
}

# Un volume macOS contiene anche /etc: deve vincere il ramo macOS.
@test "detect_os_type preferisce macOS su un volume che ha anche /etc" {
    mkdir -p "$FIXTURE/System/Library/CoreServices" "$FIXTURE/etc" "$FIXTURE/var" "$FIXTURE/bin"
    touch "$FIXTURE/System/Library/CoreServices/SystemVersion.plist"
    run detect_os_type "$FIXTURE"
    [ "$output" = "macos" ]
}

@test "detect_os_type ritorna unknown su directory vuota" {
    run detect_os_type "$FIXTURE"
    [ "$output" = "unknown" ]
}

@test "detect_os_type ritorna unknown su path inesistente" {
    run detect_os_type "$FIXTURE/non/esiste"
    [ "$output" = "unknown" ]
}

# ---------------------------------------------------------------- IoC

# Gli indicatori vanno caricati dal file: dalla 2.3 check_ioc usa la regex
# tipizzata costruita da load_ioc_file, non piu' l'elenco grezzo IOC_LIST.
_load_ioc() {
    IOC_LIST=(); IOC_TYPES=(); IOC_REGEX=""
    local f; f=$(mktemp)
    printf '%s\n' "$@" > "$f"
    load_ioc_file "$f" > /dev/null
    rm -f "$f"
}

@test "check_ioc matcha in modo case-insensitive" {
    _load_ioc "evil.exe" "1.2.3.4"
    run check_ioc "C:\\Temp\\EVIL.EXE"
    [ "$status" -eq 0 ]
}

@test "check_ioc non matcha testo estraneo" {
    _load_ioc "evil.exe"
    run check_ioc "C:\\Windows\\explorer.exe"
    [ "$status" -ne 0 ]
}

@test "check_ioc con lista vuota non matcha" {
    IOC_LIST=(); IOC_TYPES=(); IOC_REGEX=""
    run check_ioc "qualsiasi cosa"
    [ "$status" -ne 0 ]
}

# ---------------------------------------------------------------- integrità

@test "sha256_file calcola l'hash del file" {
    printf 'fiuto' > "$FIXTURE/f.bin"
    run sha256_file "$FIXTURE/f.bin"
    [ "${#output}" -eq 64 ]
    [ "$output" = "$(sha256sum "$FIXTURE/f.bin" | awk '{print $1}')" ]
}

# ---------------------------------------------------------------- history

# I REPL basati su readline (python3, node, psql) scrivono le voci multi-riga
# con spazi e backslash codificati in ottale. Senza decodifica il report mostra
# "\040\040value = ..." al posto del codice: illeggibile, e non e' un dato
# dell'artefatto ma una codifica del formato.
@test "gli escape ottali di readline vengono decodificati" {
    f=$(mktemp)
    printf '\\040\\040\\040\\040value = str(raw).strip()\n' > "$f"
    out=$(render_pre_block "$f" "" "histrl")
    [[ "$out" == *"    value = str(raw).strip()"* ]]
    [[ "$out" != *'\040'* ]]
    rm -f "$f"
}

@test "il backslash codificato non viene ri-interpretato come escape" {
    # \134 e' il backslash: decodificarlo per primo trasformerebbe il risultato
    # in un nuovo escape da interpretare. La decodifica e' una sola passata.
    #
    # printf '%s\n' e non printf '<formato>': serve che nel file finiscano i
    # caratteri \134 e \040 letterali, non la loro espansione.
    f=$(mktemp)
    printf '%s\n' 're.sub(r"\134s+",\040"\040",\040key)' > "$f"
    render_pre_block "$f" "" "histrl" > "$f.out"
    # Le virgolette nel report sono entita' HTML, quindi si verifica il pezzo
    # che conta: il backslash e' tornato UNO e non e' stato riletto come
    # escape. grep -F perche' in un pattern di [[ == ]] il backslash e' a sua
    # volta un escape, e \s diventerebbe s facendo passare il test per il
    # motivo sbagliato.
    grep -qF 'r&quot;\s+&quot;' "$f.out"
    ! grep -qF '134' "$f.out"
    rm -f "$f" "$f.out"
}

@test "senza la modalita' readline gli ottali restano come sono" {
    # Una history di shell puo' contenere un \040 letterale in un printf: non
    # va toccato.
    f=$(mktemp)
    printf 'printf "a\\040b"\n' > "$f"
    out=$(render_pre_block "$f" "" "histts")
    [[ "$out" == *'\040'* ]]
    rm -f "$f"
}
