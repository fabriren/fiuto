#!/usr/bin/env bats
#
# Test del compilatore di regole Sigma.
#
# La parte rischiosa di un motore di detection non è leggere i log: è decidere
# cosa una regola significa. Un modificatore interpretato male o una condizione
# valutata al contrario produce un falso negativo che nessuno vede. Per questo
# il compilatore sta in una libreria condivisa e qui viene esercitato con
# eventi sintetici, senza bisogno di un .evtx.
#
# Due categorie di test, ugualmente importanti:
#   - la regola scatta quando deve;
#   - la regola che il motore NON sa valutare finisce fra le scartate, con un
#     motivo. Una regola silenziosamente ignorata è indistinguibile da una
#     regola che non ha trovato nulla.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    RULES="$FIXTURE/rules"; mkdir -p "$RULES"
    SIGMA_RULES=""
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

_rule() { cat > "$RULES/$1.yml"; }

# _eval '<json evento>' — stampa "MATCH <titolo>" per ogni regola che scatta,
# e "REJECT <titolo> :: <motivo>" per ogni regola scartata.
_eval() {
    EV_JSON="$1" run_py_with_lib pylib_sigma "$RULES" << 'PYEOF'
import sys, os, json

# Il canale è finto: al compilatore serve solo sapere che esiste.
available = {'security.evtx': '/fake/Security.evtx',
             'microsoft-windows-sysmon%4operational.evtx': '/fake/Sysmon.evtx',
             'microsoft-windows-powershell%4operational.evtx': '/fake/PS.evtx'}
rules, rejected = load_sigma_rules(sys.argv[1], available)
ev = json.loads(os.environ['EV_JSON'])
for r in rules:
    if r['pred'](ev):
        print('MATCH %s' % r['title'])
for r in rejected:
    print('REJECT %s :: %s' % (r['rule'], r['reason']))
PYEOF
}

# ------------------------------------------------------------- selezioni ----

@test "una selezione con un solo campo scatta sul valore esatto" {
    _rule esatta <<'EOF'
title: esatta
logsource: {product: windows, service: security}
detection:
  selection:
    EventID: 4688
  condition: selection
EOF
    run _eval '{"EventID": "4688"}'
    [[ "$output" == *"MATCH esatta"* ]]
    run _eval '{"EventID": "4689"}'
    [[ "$output" != *"MATCH esatta"* ]]
}

@test "i campi di una stessa selezione sono in AND" {
    _rule andcampi <<'EOF'
title: andcampi
logsource: {product: windows, service: security}
detection:
  selection:
    EventID: 4688
    SubjectUserName: admin
  condition: selection
EOF
    run _eval '{"EventID": "4688", "SubjectUserName": "admin"}'
    [[ "$output" == *"MATCH andcampi"* ]]
    run _eval '{"EventID": "4688", "SubjectUserName": "guest"}'
    [[ "$output" != *"MATCH andcampi"* ]]
}

@test "una lista di valori è un OR" {
    _rule orvalori <<'EOF'
title: orvalori
logsource: {product: windows, service: security}
detection:
  selection:
    EventID:
      - 4624
      - 4625
  condition: selection
EOF
    run _eval '{"EventID": "4625"}'
    [[ "$output" == *"MATCH orvalori"* ]]
    run _eval '{"EventID": "4634"}'
    [[ "$output" != *"MATCH orvalori"* ]]
}

@test "una lista di mappe è un OR fra le mappe" {
    _rule ormappe <<'EOF'
title: ormappe
logsource: {category: process_creation}
detection:
  selection:
    - Image|endswith: '\powershell.exe'
    - Image|endswith: '\cmd.exe'
  condition: selection
EOF
    run _eval '{"Image": "C:\\Windows\\System32\\cmd.exe"}'
    [[ "$output" == *"MATCH ormappe"* ]]
    run _eval '{"Image": "C:\\Windows\\System32\\notepad.exe"}'
    [[ "$output" != *"MATCH ormappe"* ]]
}

# --------------------------------------------------------- modificatori -----

@test "contains, startswith ed endswith sono case-insensitive per default" {
    # Sigma è case-insensitive salvo |cased: sbagliarlo perderebbe i match su
    # percorsi scritti in maiuscolo, che su Windows sono la norma.
    _rule mods <<'EOF'
title: mods
logsource: {category: process_creation}
detection:
  sel_contains:
    CommandLine|contains: '-encodedcommand'
  condition: sel_contains
EOF
    run _eval '{"CommandLine": "powershell.exe -EncodedCommand SQBFAFgA"}'
    [[ "$output" == *"MATCH mods"* ]]
}

@test "cased rende il confronto sensibile alle maiuscole" {
    _rule cased <<'EOF'
title: cased
logsource: {category: process_creation}
detection:
  selection:
    CommandLine|contains|cased: 'EncodedCommand'
  condition: selection
EOF
    run _eval '{"CommandLine": "powershell -EncodedCommand X"}'
    [[ "$output" == *"MATCH cased"* ]]
    run _eval '{"CommandLine": "powershell -encodedcommand X"}'
    [[ "$output" != *"MATCH cased"* ]]
}

@test "contains|all richiede tutti i valori, non uno solo" {
    # Confondere |all con l'OR di default trasformerebbe una regola precisa in
    # una rumorosa, che è il modo più rapido per far ignorare le detection.
    _rule allmod <<'EOF'
title: allmod
logsource: {category: process_creation}
detection:
  selection:
    CommandLine|contains|all:
      - 'iex'
      - 'downloadstring'
  condition: selection
EOF
    run _eval '{"CommandLine": "iex (new-object net.webclient).downloadstring(1)"}'
    [[ "$output" == *"MATCH allmod"* ]]
    run _eval '{"CommandLine": "iex whoami"}'
    [[ "$output" != *"MATCH allmod"* ]]
}

@test "il modificatore re applica una regex" {
    _rule remod <<'EOF'
title: remod
logsource: {category: process_creation}
detection:
  selection:
    CommandLine|re: '-enc\s+[A-Za-z0-9+/=]{20,}'
  condition: selection
EOF
    run _eval '{"CommandLine": "powershell -enc SQBFAFgAIAAoAE4AZQB3AC0A"}'
    [[ "$output" == *"MATCH remod"* ]]
    run _eval '{"CommandLine": "powershell -enc breve"}'
    [[ "$output" != *"MATCH remod"* ]]
}

@test "null significa campo assente o vuoto" {
    _rule nullo <<'EOF'
title: nullo
logsource: {category: process_creation}
detection:
  selection:
    ParentImage: null
  condition: selection
EOF
    run _eval '{"Image": "x.exe"}'
    [[ "$output" == *"MATCH nullo"* ]]
    run _eval '{"Image": "x.exe", "ParentImage": "explorer.exe"}'
    [[ "$output" != *"MATCH nullo"* ]]
}

# ---------------------------------------------------------- condizioni ------

@test "and, or e not compongono le selezioni" {
    _rule cond <<'EOF'
title: cond
logsource: {category: process_creation}
detection:
  selection:
    Image|endswith: '\powershell.exe'
  filtro:
    ParentImage|endswith: '\explorer.exe'
  condition: selection and not filtro
EOF
    run _eval '{"Image": "c:\\ps\\powershell.exe", "ParentImage": "c:\\w\\winword.exe"}'
    [[ "$output" == *"MATCH cond"* ]]
    run _eval '{"Image": "c:\\ps\\powershell.exe", "ParentImage": "c:\\w\\explorer.exe"}'
    [[ "$output" != *"MATCH cond"* ]]
}

@test "1 of selection* è un OR fra le selezioni con quel prefisso" {
    _rule oneof <<'EOF'
title: oneof
logsource: {category: process_creation}
detection:
  selection_a:
    Image|endswith: '\mimikatz.exe'
  selection_b:
    CommandLine|contains: 'sekurlsa'
  condition: 1 of selection*
EOF
    run _eval '{"CommandLine": "x sekurlsa::logonpasswords"}'
    [[ "$output" == *"MATCH oneof"* ]]
    run _eval '{"CommandLine": "dir"}'
    [[ "$output" != *"MATCH oneof"* ]]
}

@test "all of them richiede che scattino tutte le selezioni" {
    _rule allof <<'EOF'
title: allof
logsource: {category: process_creation}
detection:
  sel_a:
    Image|endswith: '\rundll32.exe'
  sel_b:
    CommandLine|contains: 'javascript:'
  condition: all of them
EOF
    run _eval '{"Image": "c:\\w\\rundll32.exe", "CommandLine": "rundll32 javascript:x"}'
    [[ "$output" == *"MATCH allof"* ]]
    run _eval '{"Image": "c:\\w\\rundll32.exe", "CommandLine": "rundll32 shell32.dll"}'
    [[ "$output" != *"MATCH allof"* ]]
}

# ----------------------------------------------------- regole scartate ------

@test "una condizione con parentesi viene scartata con il motivo" {
    _rule parentesi <<'EOF'
title: parentesi
logsource: {category: process_creation}
detection:
  a: {EventID: 1}
  b: {EventID: 2}
  c: {EventID: 3}
  condition: (a or b) and c
EOF
    run _eval '{"EventID": "1"}'
    [[ "$output" == *"REJECT parentesi :: condizione con parentesi"* ]]
    [[ "$output" != *"MATCH parentesi"* ]]
}

@test "un'aggregazione viene scartata, non valutata a metà" {
    _rule aggreg <<'EOF'
title: aggreg
logsource: {product: windows, service: security}
detection:
  selection: {EventID: 4625}
  condition: selection | count() > 5
EOF
    run _eval '{"EventID": "4625"}'
    [[ "$output" == *"REJECT aggreg"* ]]
    [[ "$output" != *"MATCH aggreg"* ]]
}

@test "un modificatore non implementato scarta la regola" {
    # base64offset cambia completamente il significato del confronto: valutarlo
    # come uguaglianza semplice darebbe un "non scattata" privo di senso.
    _rule b64 <<'EOF'
title: b64
logsource: {category: process_creation}
detection:
  selection:
    CommandLine|base64offset|contains: 'whoami'
  condition: selection
EOF
    run _eval '{"CommandLine": "whoami"}'
    [[ "$output" == *"REJECT b64 :: modificatore |base64offset"* ]]
    [[ "$output" != *"MATCH b64"* ]]
}

@test "una logsource non mappabile scarta la regola invece di girare a tappeto" {
    _rule linuxrule <<'EOF'
title: linuxrule
logsource: {product: linux, service: auditd}
detection:
  selection: {type: EXECVE}
  condition: selection
EOF
    run _eval '{"type": "EXECVE"}'
    [[ "$output" == *"REJECT linuxrule :: logsource non mappata"* ]]
}

@test "un canale assente dal volume scarta la regola con il proprio motivo" {
    _rule wmirule <<'EOF'
title: wmirule
logsource: {product: windows, service: wmi}
detection:
  selection: {EventID: 5861}
  condition: selection
EOF
    # available nel test contiene solo Security, Sysmon e PowerShell.
    run _eval '{"EventID": "5861"}'
    [[ "$output" == *"REJECT wmirule :: canale non presente"* ]]
}

@test "uno YAML rotto non impedisce alle altre regole di caricarsi" {
    _rule buona <<'EOF'
title: buona
logsource: {product: windows, service: security}
detection:
  selection: {EventID: 4688}
  condition: selection
EOF
    printf 'title: rotta\n  detection: [\n' > "$RULES/rotta.yml"
    run _eval '{"EventID": "4688"}'
    [[ "$output" == *"MATCH buona"* ]]
    [[ "$output" == *"REJECT rotta.yml"* ]]
}

@test "una regola senza detection viene scartata" {
    printf 'title: vuota\nlogsource: {service: security}\n' > "$RULES/vuota.yml"
    run _eval '{"EventID": "1"}'
    [[ "$output" == *"REJECT vuota :: detection assente"* ]]
}

# ---------------------------------------------------------------- modulo ----

@test "senza --sigma la guardia salta il modulo con un motivo" {
    run _guard_sigma
    [ "$status" -ne 0 ]
    [ -n "$output" ]
}

@test "il modulo Sigma è registrato per Windows con la sua guardia" {
    printf '%s\n' "${MODULES_WIN[@]}" | grep -q "module_sigma|.*|_guard_sigma"
}

@test "Sigma non è nei registri Linux e macOS" {
    # Le regole mappate qui puntano a canali EVTX: offrirlo altrove
    # prometterebbe una copertura che non esiste.
    ! printf '%s\n' "${MODULES_LINUX[@]}" "${MODULES_MACOS[@]}" | grep -q "module_sigma"
}
