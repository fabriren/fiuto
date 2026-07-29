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

@test "check_ioc matcha in modo case-insensitive" {
    IOC_LIST=("evil.exe" "1.2.3.4")
    run check_ioc "C:\\Temp\\EVIL.EXE"
    [ "$status" -eq 0 ]
}

@test "check_ioc non matcha testo estraneo" {
    IOC_LIST=("evil.exe")
    run check_ioc "C:\\Windows\\explorer.exe"
    [ "$status" -ne 0 ]
}

@test "check_ioc con lista vuota non matcha" {
    IOC_LIST=()
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
