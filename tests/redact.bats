#!/usr/bin/env bats
#
# Test della modalità --redact.
#
# Due rischi speculari, e i test coprono entrambi:
#   - lasciar passare un segreto in una copia dichiarata condivisibile;
#   - oscurare quello che non è un segreto. Un hash SHA-256 di un reperto e un
#     hash NTLM sono entrambi stringhe esadecimali: cancellare il primo
#     spezzerebbe la catena di custodia.
# Più il vincolo non negoziabile: l'originale non si tocca.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    REPORT_BASE_DIR="$FIXTURE"
    REDACT=true; REDACT_DEFANG=false
    EXPORT_JSONL=false; CUSTODY=false
    GENERATED_REPORTS=()
    HTML="$FIXTURE/report.html"
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
    rm -rf "${TMPDIR:-/tmp}/fiuto_redact_$$"
}

_write() { printf '%s\n' "<html><head><title>T</title></head><body><main>" "$@" "</main></body></html>" > "$HTML"; }
_red() { echo "$FIXTURE/report.redacted.html"; }

# ------------------------------------------------------- non si tocca -------

@test "l'originale resta byte per byte identico" {
    _write "password = SuperSegreta123"
    BEFORE=$(sha256sum "$HTML" | cut -d' ' -f1)
    redact_report "$HTML"
    AFTER=$(sha256sum "$HTML" | cut -d' ' -f1)
    [ "$BEFORE" = "$AFTER" ]
}

@test "la copia oscurata nasce accanto all'originale" {
    _write "nulla di sensibile"
    redact_report "$HTML"
    [ -f "$(_red)" ]
}

@test "una copia oscurata non viene oscurata a sua volta" {
    # Altrimenti ogni sessione produrrebbe report.redacted.redacted.html
    _write "password = x"
    redact_report "$HTML"
    redact_report "$(_red)"
    [ ! -f "$FIXTURE/report.redacted.redacted.html" ]
}

# ------------------------------------------------------ cosa si oscura ------

@test "il valore di un campo etichettato sparisce, l'etichetta resta" {
    # Sapere CHE c'era una password è un dato dell'analisi; il valore no.
    _write "password = SuperSegreta123" "api_key: abcd1234efgh5678"
    redact_report "$HTML"
    R=$(_red)
    ! grep -q "SuperSegreta123" "$R"
    ! grep -q "abcd1234efgh5678" "$R"
    grep -qi "password" "$R"
    grep -q "REDACTED" "$R"
}

@test "gli hash NTLM di un riversamento SAM spariscono, utente e RID restano" {
    _write "Administrator:500:aad3b435b51404eeaad3b435b51404ee:31d6cfe0d16ae931b73c59d7e0c089c0:::"
    redact_report "$HTML"
    R=$(_red)
    ! grep -q "31d6cfe0d16ae931b73c59d7e0c089c0" "$R"
    ! grep -q "aad3b435b51404ee" "$R"
    grep -q "Administrator:500:" "$R"
}

@test "una chiave privata PEM viene rimossa per intero" {
    _write "-----BEGIN RSA PRIVATE KEY-----" "MIIEpAIBAAKCAQEA1234567890" "abcdefghijklmnop" "-----END RSA PRIVATE KEY-----"
    redact_report "$HTML"
    R=$(_red)
    ! grep -q "MIIEpAIBAAKCAQEA" "$R"
    ! grep -q "abcdefghijklmnop" "$R"
}

@test "la PSK di un profilo WLAN sparisce" {
    _write "&lt;keyMaterial&gt;PasswordDelWifi2024&lt;/keyMaterial&gt;"
    redact_report "$HTML"
    ! grep -q "PasswordDelWifi2024" "$(_red)"
}

@test "credenziali dalla forma inequivocabile spariscono senza etichetta" {
    _write "AKIAIOSFODNN7EXAMPLE" \
           "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.dBjftJeZ4CVPmB92K27uhbUJU1p1r3wa"
    redact_report "$HTML"
    R=$(_red)
    ! grep -q "AKIAIOSFODNN7EXAMPLE" "$R"
    ! grep -q "eyJhbGciOiJIUzI1NiJ9" "$R"
}

# --------------------------------------------- cosa NON si deve oscurare ----

@test "gli hash SHA-256 dei reperti restano intatti" {
    # Sono dati di integrità, non segreti: rimuoverli spezzerebbe la catena
    # di custodia proprio nel file che si condivide.
    H="e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    _write "SHA-256: $H" "<td>pagefile.sys</td><td>$H</td>"
    redact_report "$HTML"
    [ "$(grep -c "$H" "$(_red)")" -ge 2 ]
}

@test "percorsi, date e nomi utente non vengono toccati" {
    _write "<td>2026-03-05 10:00:00</td><td>C:\\Users\\mrossi\\payload.exe</td>"
    redact_report "$HTML"
    R=$(_red)
    grep -q "2026-03-05 10:00:00" "$R"
    grep -q "mrossi" "$R"
    grep -q "payload.exe" "$R"
}

# ---------------------------------------------------------- dichiarazione ---

@test "la copia dichiara in testa di non essere l'originale" {
    # Un file che sembra un report ma non lo è del tutto, senza dirlo, è
    # peggio di nessun file.
    _write "password = x"
    redact_report "$HTML"
    R=$(_red)
    grep -qi "REDACTED COPY\|COPIA OSCURATA" "$R"
    grep -qi "chain of custody\|catena di custodia" "$R"
    grep -q "REDACTED" <(head -20 "$R")
}

@test "il cartiglio elenca quante occorrenze ha oscurato e per quale regola" {
    _write "password = uno" "pwd = due"
    redact_report "$HTML"
    grep -qi "campo etichettato" "$(_red)"
}

@test "anche senza nulla da oscurare la copia lo dichiara invece di mancare" {
    # Se il file mancasse, chi condivide dovrebbe chiedersi perché.
    _write "<td>openssl</td><td>3.0.2</td>"
    redact_report "$HTML"
    R=$(_red)
    [ -f "$R" ]
    grep -qi "Nessuna regola\|No rule found" "$R"
}

# ---------------------------------------------------------------- defang ----

@test "con --defang URL e IP diventano inerti nella sola copia" {
    _write "GET http://evil.com/x da 203.0.113.9"
    REDACT_DEFANG=true
    redact_report "$HTML"
    R=$(_red)
    grep -q "hxxp://" "$R"
    grep -q "203\[.\]0\[.\]113\[.\]9" "$R"
    # L'originale resta navigabile e fedele.
    grep -q "http://evil.com/x" "$HTML"
}

@test "senza --defang gli URL restano come sono" {
    _write "GET http://evil.com/x"
    redact_report "$HTML"
    grep -q "http://evil.com/x" "$(_red)"
}

# ------------------------------------------------------------- aggancio -----

@test "register_report produce la copia solo con REDACT attivo" {
    _write "password = x"
    REDACT=false
    register_report "$HTML"
    [ ! -f "$(_red)" ]
    REDACT=true
    register_report "$HTML"
    [ -f "$(_red)" ]
}

@test "il conteggio di sessione tiene traccia delle copie prodotte" {
    _write "password = x"
    redact_report "$HTML"
    cp "$HTML" "$FIXTURE/altro.html"
    redact_report "$FIXTURE/altro.html"
    [ "$(redacted_count)" -eq 2 ]
}
