#!/usr/bin/env bats
#
# Test del motore IoC tipizzato.
#
# Il match a sottostringa sbagliava in due direzioni opposte, e i test coprono
# entrambe:
#   - falsi positivi: 10.0.0.5 dentro 110.0.0.55, evil.com dentro notevil.com;
#   - falsi negativi: gli indicatori arrivano defanged e cercati alla lettera
#     non corrispondono a niente.
# La seconda è la più insidiosa, perché si manifesta come "disco pulito".

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    IOC_LIST=(); IOC_TYPES=(); IOC_REGEX=""
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

_load() { printf '%s\n' "$@" > "$FIXTURE/ioc.txt"; load_ioc_file "$FIXTURE/ioc.txt" > /dev/null; }

# ------------------------------------------------------- tipizzazione -------

@test "gli indicatori vengono riconosciuti per tipo" {
    _load "10.0.0.5" "evil.com" "https://evil.com/a" \
          "d41d8cd98f00b204e9800998ecf8427e" \
          "da39a3ee5e6b4b0d3255bfef95601890afd80709" \
          "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855" \
          "tizio@evil.com" "payload.dll"
    T=$(printf '%s\n' "${IOC_TYPES[@]}" | cut -d'|' -f1 | tr '\n' ' ')
    [ "$T" = "ipv4 domain url md5 sha1 sha256 email literal " ]
}

# -------------------------------------------------- falsi positivi ----------

@test "un IP non corrisponde dentro un IP più lungo" {
    _load "10.0.0.5"
    run check_ioc "connessione a 10.0.0.5 riuscita"
    [ "$status" -eq 0 ]
    run check_ioc "connessione a 110.0.0.55 riuscita"
    [ "$status" -ne 0 ]
}

@test "un dominio non corrisponde come sottostringa di un altro" {
    _load "evil.com"
    run check_ioc "GET http://evil.com/x"
    [ "$status" -eq 0 ]
    run check_ioc "GET http://notevil.com.au/x"
    [ "$status" -ne 0 ]
    run check_ioc "GET http://evil.com.altro.net/x"
    [ "$status" -ne 0 ]
}

@test "un dominio corrisponde nei propri sottodomini" {
    # mail.evil.com È evil.com: escluderlo perderebbe il caso più comune.
    _load "evil.com"
    run check_ioc "smtp mail.evil.com"
    [ "$status" -eq 0 ]
}

@test "un hash non corrisponde dentro un hash più lungo" {
    _load "d41d8cd98f00b204e9800998ecf8427e"
    run check_ioc "md5=d41d8cd98f00b204e9800998ecf8427e ok"
    [ "$status" -eq 0 ]
    run check_ioc "blob d41d8cd98f00b204e9800998ecf8427eff00 ok"
    [ "$status" -ne 0 ]
}

@test "il match sugli hash ignora maiuscole e minuscole" {
    _load "D41D8CD98F00B204E9800998ECF8427E"
    run check_ioc "hash d41d8cd98f00b204e9800998ecf8427e"
    [ "$status" -eq 0 ]
}

# -------------------------------------------------- falsi negativi ----------

@test "gli indicatori defanged vengono normalizzati" {
    # È così che arrivano in una mail o in un bollettino CERT.
    _load "hxxp://evil[.]com/payload" "1[.]2[.]3[.]4" "tizio[at]evil[.]com"
    run check_ioc "download da http://evil.com/payload"
    [ "$status" -eq 0 ]
    run check_ioc "peer 1.2.3.4:443"
    [ "$status" -eq 0 ]
    run check_ioc "from tizio@evil.com"
    [ "$status" -eq 0 ]
}

@test "il valore normalizzato finisce anche in IOC_LIST" {
    # I moduli che fanno il proprio match in Python leggono IOC_LIST: se ci
    # finisse la forma defanged, non troverebbero mai nulla.
    _load "hxxp://evil[.]com/x"
    [ "${IOC_LIST[0]}" = "http://evil.com/x" ]
}

# ------------------------------------------------------------ regex --------

@test "un indicatore fra slash è trattato come regex" {
    _load "/inv[o0]ke-expr[e3]ssion/"
    run check_ioc "cmd: Inv0ke-Expr3ssion x"
    [ "$status" -eq 0 ]
    run check_ioc "cmd: get-childitem"
    [ "$status" -ne 0 ]
}

@test "i metacaratteri di un valore letterale non vengono interpretati" {
    # Un nome di file con la parentesi non deve diventare un gruppo di regex,
    # né far esplodere il match su tutto.
    _load 'fattura(1).exe'
    run check_ioc "eseguito fattura(1).exe"
    [ "$status" -eq 0 ]
    run check_ioc "eseguito fattura1.exe"
    [ "$status" -ne 0 ]
}

# ------------------------------------------------------------ import -------

@test "un bundle STIX 2.x viene importato con i tipi corretti" {
    cat > "$FIXTURE/stix.json" <<'EOF'
{"type": "bundle", "id": "bundle--1", "objects": [
 {"type": "indicator", "pattern": "[ipv4-addr:value = '203.0.113.9']", "pattern_type": "stix"},
 {"type": "indicator", "pattern": "[domain-name:value = 'evil.example']", "pattern_type": "stix"},
 {"type": "indicator", "pattern": "[file:hashes.'SHA-256' = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855']", "pattern_type": "stix"},
 {"type": "malware", "name": "non un indicatore"}
]}
EOF
    load_ioc_file "$FIXTURE/stix.json" > /dev/null
    [ "${#IOC_LIST[@]}" -eq 3 ]
    run check_ioc "peer 203.0.113.9 attivo"
    [ "$status" -eq 0 ]
    run check_ioc "host evil.example"
    [ "$status" -eq 0 ]
}

@test "un export MISP viene importato, oggetti annidati compresi" {
    cat > "$FIXTURE/misp.json" <<'EOF'
{"Event": {"info": "caso", "Attribute": [
   {"type": "ip-dst", "value": "198.51.100.7"},
   {"type": "domain", "value": "bad.example"}],
 "Object": [{"name": "file", "Attribute": [
   {"type": "md5", "value": "d41d8cd98f00b204e9800998ecf8427e"}]}]}}
EOF
    load_ioc_file "$FIXTURE/misp.json" > /dev/null
    [ "${#IOC_LIST[@]}" -eq 3 ]
    run check_ioc "md5 d41d8cd98f00b204e9800998ecf8427e"
    [ "$status" -eq 0 ]
}

@test "un JSON che non è né STIX né MISP non viene accettato in silenzio" {
    printf '{"qualcosa": [1,2,3]}\n' > "$FIXTURE/altro.json"
    run load_ioc_file "$FIXTURE/altro.json"
    [ "$status" -ne 0 ]
}

@test "i duplicati vengono contati una volta sola" {
    _load "evil.com" "EVIL.COM" "evil.com"
    [ "${#IOC_LIST[@]}" -eq 1 ]
}

@test "commenti e righe vuote non diventano indicatori" {
    # Un commento caricato come IoC letterale farebbe scattare match ovunque.
    _load "# indicatori del caso 42" "" "   " "evil.com"
    [ "${#IOC_LIST[@]}" -eq 1 ]
}

# ----------------------------------------------------------- utilità -------

@test "matched_ioc dice quale indicatore ha corrisposto" {
    _load "10.0.0.5" "evil.com"
    run matched_ioc "GET http://evil.com/x"
    [ "$status" -eq 0 ]
    [ "$output" = "domain|evil.com" ]
}

@test "defang_value rende inerte un URL" {
    # Un URL cliccabile in un documento di analisi è un rischio senza ragione.
    run defang_value "http://evil.com/x"
    [ "$output" = "hxxp://evil[.]com/x" ]
}

@test "senza indicatori caricati check_ioc non corrisponde a nulla" {
    run check_ioc "qualunque testo"
    [ "$status" -ne 0 ]
}
