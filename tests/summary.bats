#!/usr/bin/env bats
#
# Test dell'executive summary.
#
# Il rischio di questa funzione non è mancare un riscontro: è produrne di
# falsi. Un riepilogo che segnala cose che non esistono fa perdere fiducia
# anche nelle voci giuste, e allora tanto vale non averlo. I test insistono
# quindi sui casi in cui NON deve scattare.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    REPORT_BASE_DIR="$FIXTURE"
    WIN_ROOT="/mnt/disk"; OS_TYPE="windows"; HOST_NAME="WS-01"
    BATCH_MODE=true              # niente prompt "apri il report?"
    GENERATED_REPORTS=()
    IOC_LIST=(); IOC_TYPES=(); IOC_REGEX=""
    TIME_SINCE=""; TIME_UNTIL=""
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

# _report <slug_con_timestamp> — legge l'HTML da stdin e lo registra
_report() {
    mkdir -p "$FIXTURE/$1"
    cat > "$FIXTURE/$1/report.html"
    GENERATED_REPORTS+=("$FIXTURE/$1/report.html")
}

_json() { python3 -c '
import json, sys
print(json.dumps(json.load(open(sys.argv[1]))))' "$FIXTURE/findings.json"; }

_q() { python3 -c '
import json, sys
d = json.load(open(sys.argv[1]))
print(eval(sys.argv[2], {"d": d}))' "$FIXTURE/findings.json" "$1"; }

# ------------------------------------------------------------- riscontri ---

@test "una riga di dato che corrisponde a una regola produce un riscontro" {
    _report ps_scriptblock_20260301_100000 <<'EOF'
<html><body><table>
<tr><td>2026-03-05 10:00:00</td><td>powershell.exe -enc SQBFAFgAIAAoAE4AZQB3AC0ATwBiAGoAZQBjAHQA</td></tr>
</table></body></html>
EOF
    generate_executive_summary
    [ -f "$FIXTURE/findings.json" ]
    [ "$(_q 'd["counts"]["ALTA"]')" -eq 1 ]
    [ "$(_q 'd["findings"][0]["mitre"]')" = "T1059.001" ]
}

@test "la prosa esplicativa del report non fa scattare le regole" {
    # I moduli scrivono cartigli tipo "questo artefatto è usato dal malware
    # per...". Un motore che leggesse l'intera pagina segnalerebbe sé stesso.
    _report ps_scriptblock_20260301_100000 <<'EOF'
<html><body>
<div class='card'><div style='padding:1rem'>Nota: qui compaiono spesso comandi
con -enc AAAAAAAAAAAAAAAAAAAAAAAAAAAA e invoke-expression usati dal malware.</div></div>
<table><tr><td>2026-03-05 10:00:00</td><td>Get-ChildItem C:\Users</td></tr></table>
</body></html>
EOF
    generate_executive_summary
    [ "$(_q 'len(d["findings"])')" -eq 0 ]
}

@test "un report senza nulla di notevole non produce riscontri" {
    _report linux_packages_20260301_100000 <<'EOF'
<html><body><table>
<tr><td>2026-03-05 10:00:00</td><td>openssl</td><td>3.0.2</td></tr>
<tr><td>2026-03-05 10:00:01</td><td>bash</td><td>5.1</td></tr>
</table></body></html>
EOF
    generate_executive_summary
    [ "$(_q 'len(d["findings"])')" -eq 0 ]
    [ "$(_q 'd["score"]')" -eq 0 ]
}

@test "il conteggio è per riga, non per cella" {
    # "3 riscontri" su una riga sola sarebbe una bugia sul volume del problema.
    _report usb_20260301_100000 <<'EOF'
<html><body><table>
<tr><td>2026-03-05 10:05:00</td><td>Kingston DataTraveler</td><td>USBSTOR</td></tr>
</table></body></html>
EOF
    generate_executive_summary
    [ "$(_q 'd["findings"][0]["count"]')" -eq 1 ]
}

@test "le righe marcate sensitive dal modulo diventano un riscontro" {
    _report linux_syslog_20260301_100000 <<'EOF'
<html><body><pre><span class="line sensitive">    1 Mar  5 10:09:00 srv sudo: root shell</span>
<span class="line">    2 Mar  5 10:10:00 srv cron: normale</span></pre></body></html>
EOF
    generate_executive_summary
    [ "$(_q 'd["findings"][0]["source"]')" = "modulo" ]
    [ "$(_q 'd["findings"][0]["count"]')" -eq 1 ]
}

@test "gli IoC caricati producono un riscontro ad alta severità" {
    # Si passa dal caricatore vero: il riepilogo usa la regex compilata dal
    # motore IoC, non l'elenco grezzo dei valori.
    printf 'evil.example\n' > "$FIXTURE/ioc.txt"
    load_ioc_file "$FIXTURE/ioc.txt" > /dev/null
    _report webcache_20260301_100000 <<'EOF'
<html><body><table>
<tr><td>2026-03-05 10:00:00</td><td>http://evil.example/payload</td></tr>
<tr><td>2026-03-05 10:01:00</td><td>https://www.google.com</td></tr>
</table></body></html>
EOF
    generate_executive_summary
    [ "$(_q '[f["source"] for f in d["findings"]].count("ioc")')" -eq 1 ]
    [ "$(_q '[f for f in d["findings"] if f["source"]=="ioc"][0]["count"]')" -eq 1 ]
}

@test "i riscontri sono ordinati per severità decrescente" {
    IOC_LIST=(); IOC_TYPES=(); IOC_REGEX=""
    _report usb_20260301_100000 <<'EOF'
<html><body><table><tr><td>2026-03-05 10:05:00</td><td>Kingston</td></tr></table></body></html>
EOF
    _report ifeo_20260301_100100 <<'EOF'
<html><body><table><tr><td>notepad.exe</td><td>Debugger</td><td>C:\x.exe</td></tr></table></body></html>
EOF
    generate_executive_summary
    [ "$(_q 'd["findings"][0]["severity"]')" = "ALTA" ]
    [ "$(_q 'd["findings"][-1]["severity"]')" = "BASSA" ]
}

# ------------------------------------------------------------ correlazione --

_exfil_reports() {
    _report usb_20260301_100000 <<'EOF'
<html><body><table><tr><td>2026-03-05 10:05:00</td><td>Kingston DataTraveler</td></tr></table></body></html>
EOF
    _report lnk_20260301_100100 <<'EOF'
<html><body><table><tr><td>2026-03-05 10:07:00</td><td>E:\progetto.xlsx</td></tr></table></body></html>
EOF
    _report usn_20260301_100200 <<'EOF'
<html><body><table><tr><td>2026-03-05 10:08:00</td><td>progetto.xlsx</td><td>DATA_OVERWRITE</td></tr></table></body></html>
EOF
}

@test "USB, LNK e USN nella stessa finestra producono la correlazione esfiltrazione" {
    _exfil_reports
    generate_executive_summary
    [ "$(_q 'len(d["correlations"])')" -eq 1 ]
    [ "$(_q 'sorted(d["correlations"][0]["modules"])')" = "['lnk', 'usb', 'usn']" ]
}

@test "gli stessi moduli a giorni di distanza non vengono correlati" {
    # La coincidenza temporale è tutto ciò che rende l'ipotesi sensata:
    # senza, resterebbe una correlazione inventata.
    _report usb_20260301_100000 <<'EOF'
<html><body><table><tr><td>2026-01-05 10:05:00</td><td>Kingston</td></tr></table></body></html>
EOF
    _report lnk_20260301_100100 <<'EOF'
<html><body><table><tr><td>2026-03-05 10:07:00</td><td>E:\progetto.xlsx</td></tr></table></body></html>
EOF
    _report usn_20260301_100200 <<'EOF'
<html><body><table><tr><td>2026-06-05 10:08:00</td><td>progetto.xlsx</td></tr></table></body></html>
EOF
    generate_executive_summary
    [ "$(_q 'len(d["correlations"])')" -eq 0 ]
}

@test "un solo modulo attivo non produce correlazioni" {
    _report usb_20260301_100000 <<'EOF'
<html><body><table>
<tr><td>2026-03-05 10:05:00</td><td>Kingston</td></tr>
<tr><td>2026-03-05 10:06:00</td><td>SanDisk</td></tr>
</table></body></html>
EOF
    generate_executive_summary
    [ "$(_q 'len(d["correlations"])')" -eq 0 ]
}

@test "la correlazione elenca anche gli altri moduli attivi nella finestra" {
    _exfil_reports
    _report linux_packages_20260301_100300 <<'EOF'
<html><body><table><tr><td>2026-03-05 10:06:30</td><td>openssl</td></tr></table></body></html>
EOF
    generate_executive_summary
    [ "$(_q '"linux_packages" in d["correlations"][0]["all_modules"]')" = "True" ]
    [ "$(_q '"linux_packages" in d["correlations"][0]["modules"]')" = "False" ]
}

# ---------------------------------------------------------------- output ----

@test "il punteggio dichiara la propria formula e il proprio limite" {
    # Un numero senza definizione è peggio di nessun numero.
    _exfil_reports
    generate_executive_summary
    [ "$(_q '"40" in d["score_formula"]')" = "True" ]
    [ "$(_q '"NON misura" in d["score_meaning"]')" = "True" ]
}

@test "il punteggio non supera 100" {
    printf 'evil.example\n' > "$FIXTURE/ioc.txt"
    load_ioc_file "$FIXTURE/ioc.txt" > /dev/null
    local i
    for i in 1 2 3 4 5 6 7 8; do
        _report "ifeo_2026030${i}_100000" <<'EOF'
<html><body><table><tr><td>notepad.exe</td><td>Debugger</td><td>http://evil.example/x.exe</td></tr></table></body></html>
EOF
    done
    generate_executive_summary
    [ "$(_q 'd["score"]')" -le 100 ]
}

@test "viene prodotta la pagina HTML oltre al JSON" {
    _exfil_reports
    generate_executive_summary
    [ -f "$FIXTURE/executive_summary.html" ]
    grep -q "Executive Summary" "$FIXTURE/executive_summary.html"
    # L'avvertenza sul significato dell'assenza di riscontri non è opzionale.
    grep -qi "not a clean bill of health\|attestato di pulizia" "$FIXTURE/executive_summary.html"
}

@test "il riepilogo non rilegge sé stesso né la dashboard" {
    _exfil_reports
    generate_executive_summary
    # Seconda passata: executive_summary.html è ora in GENERATED_REPORTS.
    generate_executive_summary
    [ "$(_q '"executive_summary" in " ".join(d["modules"])')" = "False" ]
}

# ------------------------------------------------------------- struttura ---

@test "ogni regola ha i sei campi attesi e una severità valida" {
    fail=0
    while IFS= read -r line; do
        n=$(awk -F'\t' '{print NF}' <<< "$line")
        [ "$n" -eq 6 ] || { echo "campi=$n: ${line:0:60}"; fail=1; continue; }
        sev=$(cut -f2 <<< "$line")
        case "$sev" in
            CRITICA|ALTA|MEDIA|BASSA) ;;
            *) echo "severità sconosciuta '$sev': ${line:0:60}"; fail=1 ;;
        esac
    done < <(finding_rules_tsv)
    [ "$fail" -eq 0 ]
}

@test "ogni pattern delle regole è una regex valida" {
    # Una regex rotta verrebbe scartata in silenzio e la regola non girerebbe.
    finding_rules_tsv | python3 -c '
import re, sys
bad = 0
for line in sys.stdin:
    f = line.rstrip("\n").split("\t")
    if len(f) != 6:
        continue
    for rx in (f"^({f[0]})$", f[3]):
        try:
            re.compile(rx)
        except re.error as exc:
            print(f"{rx}: {exc}")
            bad = 1
sys.exit(bad)'
}

@test "gli slug citati dalle regole corrispondono a moduli esistenti" {
    # Uno slug sbagliato non fa fallire nulla: la regola semplicemente non
    # scatta mai, e nessuno se ne accorge finché non serve.
    grep -ohE '(finish_report|prepare_report_dir) "[a-z0-9_]+"' "$REPO_ROOT/fiuto.sh" \
        | sed -E 's/.*"([a-z0-9_]+)"/\1/' | sort -u > "$FIXTURE/slugs"
    fail=0
    while IFS= read -r line; do
        IFS='|' read -ra SLUGS <<< "$(cut -f1 <<< "$line")"
        for s in "${SLUGS[@]}"; do
            grep -qx "$s" "$FIXTURE/slugs" || { echo "slug inesistente: $s"; fail=1; }
        done
    done < <(finding_rules_tsv)
    while IFS= read -r line; do
        # Non usare GROUPS: è una variabile speciale di bash e l'assegnazione
        # fallisce in silenzio, uccidendo il test invece di farlo fallire.
        IFS=',' read -ra GRP <<< "$(cut -f1 <<< "$line")"
        for g in "${GRP[@]}"; do
            IFS='|' read -ra SLUGS <<< "$g"
            for s in "${SLUGS[@]}"; do
                grep -qx "$s" "$FIXTURE/slugs" || { echo "slug inesistente (scenario): $s"; fail=1; }
            done
        done
    done < <(correlation_scenarios_tsv)
    [ "$fail" -eq 0 ]
}

@test "ogni scenario di correlazione richiede almeno due gruppi di moduli" {
    fail=0
    while IFS= read -r line; do
        n=$(awk -F'\t' '{print NF}' <<< "$line")
        [ "$n" -eq 3 ] || { echo "campi=$n: ${line:0:50}"; fail=1; continue; }
        g=$(cut -f1 <<< "$line" | tr ',' '\n' | wc -l)
        # Un solo gruppo non è una correlazione: è un riscontro travestito.
        [ "$g" -ge 2 ] || { echo "gruppi=$g: ${line:0:50}"; fail=1; }
    done < <(correlation_scenarios_tsv)
    [ "$fail" -eq 0 ]
}

@test "senza alcun report il riepilogo non viene generato" {
    GENERATED_REPORTS=()
    run generate_executive_summary
    [ ! -f "$FIXTURE/findings.json" ]
}

# -------------------------------------------------- dashboard di sessione ---

@test "la dashboard mette l'executive summary come prima scheda" {
    # Dopo --all e' il documento da cui si comincia a leggere, non uno dei
    # novanta report da cercare nell'elenco.
    SUMMARY_TABLE=("1|Alfa|SI|$FIXTURE/alfa/report.html")
    mkdir -p "$FIXTURE/alfa"; : > "$FIXTURE/alfa/report.html"
    : > "$FIXTURE/executive_summary.html"
    generate_full_dashboard > /dev/null 2>&1
    [ -f "$FIXTURE/index.html" ]
    nav=$(sed -n 's/.*<nav id="tabs">\(.*\)<\/nav>.*/\1/p' "$FIXTURE/index.html")
    [[ "$nav" == "<button class='tab tab-summary'"* ]]
    grep -q "data-src='executive_summary.html'" "$FIXTURE/index.html"
}

@test "la dashboard si apre già posizionata sul riepilogo" {
    SUMMARY_TABLE=("1|Alfa|SI|$FIXTURE/alfa/report.html")
    mkdir -p "$FIXTURE/alfa"; : > "$FIXTURE/alfa/report.html"
    : > "$FIXTURE/executive_summary.html"
    generate_full_dashboard > /dev/null 2>&1
    grep -q "querySelector('.tab-summary')" "$FIXTURE/index.html"
}

@test "senza riepilogo la dashboard non inventa una scheda" {
    # Puo' mancare: --all su un volume da cui nessun modulo estrae nulla.
    SUMMARY_TABLE=("1|Alfa|SI|$FIXTURE/alfa/report.html")
    mkdir -p "$FIXTURE/alfa"; : > "$FIXTURE/alfa/report.html"
    rm -f "$FIXTURE/executive_summary.html"
    generate_full_dashboard > /dev/null 2>&1
    # Il CSS della classe c'e' sempre; cio' che non deve esserci e' il BOTTONE.
    # grep -q su un pattern assente esce 1 e, sotto errexit, ucciderebbe il
    # test: si cattura l'esito invece di negarlo in linea.
    trovato=$(grep -c "button class='tab tab-summary'" "$FIXTURE/index.html" || true)
    [ "$trovato" -eq 0 ]
}

@test "in batch il riepilogo non chiede di aprirsi da solo" {
    # Chiederlo vorrebbe dire due domande di fila e due finestre aperte sullo
    # stesso contenuto, visto che poi si apre la dashboard.
    grep -q 'SUMMARY_NO_PROMPT=true generate_executive_summary' "$REPO_ROOT/src/lib/11-runner.sh"
    grep -q 'SUMMARY_NO_PROMPT:-false' "$REPO_ROOT/src/lib/18-summary.sh"
}
