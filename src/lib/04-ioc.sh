
# ================================================================
#  MOTORE IoC
#
#  Fino alla 2.2 il match era una sottostringa case-insensitive. Comodo, e
#  sbagliato in due direzioni opposte:
#
#  - FALSI POSITIVI. L'indirizzo 10.0.0.5 corrisponde dentro 110.0.0.55, dentro
#    un numero di build, dentro un offset esadecimale. Il dominio evil.com
#    corrisponde dentro notevil.com.au. In un report con centomila righe questo
#    non e' rumore fastidioso: e' rumore che nasconde i match veri.
#  - FALSI NEGATIVI. Gli IoC arrivano quasi sempre defanged — hxxp://,
#    1[.]2[.]3[.]4, evil[.]com — perche' e' cosi' che si scrivono in una mail o
#    in un bollettino. Cercati alla lettera non corrispondono a niente, e
#    l'analista conclude che il disco e' pulito.
#
#  Qui ogni indicatore viene TIPIZZATO al caricamento e confrontato con i
#  confini giusti per il suo tipo, dopo essere stato normalizzato. Si accettano
#  liste piatte, bundle STIX 2.x e export MISP.
#
#  IOC_LIST resta l'elenco dei valori normalizzati: i moduli che fanno il
#  proprio match in Python continuano a funzionare, e guadagnano comunque la
#  normalizzazione del defanging.
# ================================================================

declare -a IOC_TYPES=()     # "tipo|valore" per ogni indicatore caricato
IOC_REGEX=""                # ERE unica, minuscola, usata da check_ioc

# defang_value <stringa> — rende inerte un indicatore per la visualizzazione.
# Serve nei report: un URL cliccabile o copiabile per sbaglio in un browser e'
# un rischio che non ha ragione di esistere in un documento di analisi.
defang_value() {
    local V="$1"
    V="${V//http:/hxxp:}"; V="${V//https:/hxxps:}"
    V="${V//./[.]}"
    V="${V//@/[at]}"
    printf '%s' "$V"
}

# Estensioni che rendono un valore un nome di file e non un dominio. Senza
# questo elenco "payload.dll" verrebbe classificato dominio: sintatticamente lo
# e'. La riga di riepilogo per tipo stampata al caricamento serve proprio a far
# accorgere di una classificazione sbagliata; per i casi davvero ambigui
# (payload.com) esiste il prefisso esplicito "file:payload.com".
_IOC_FILE_EXT="|exe|dll|sys|bat|cmd|ps1|psm1|vbs|vbe|js|jse|wsf|scr|hta|jar|lnk|tmp|dat|bin|log|txt|ini|cfg|conf|zip|rar|7z|gz|tar|iso|img|doc|docx|xls|xlsx|ppt|pptx|pdf|rtf|one|php|asp|aspx|jsp|py|sh|pl|rb|dmp|pf|evtx|reg|crt|pem|key|"

# Riconosce il tipo di un indicatore gia' normalizzato.
_ioc_type_of() {
    local V="${1,,}"
    case "$V" in
        /*/) echo regex; return ;;
        http://*|https://*|ftp://*) echo url; return ;;
    esac
    if [[ "$V" =~ ^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}(/[0-9]{1,2})?$ ]]; then
        echo ipv4; return
    fi
    if [[ "$V" =~ ^[0-9a-f]{64}$ ]]; then echo sha256; return; fi
    if [[ "$V" =~ ^[0-9a-f]{40}$ ]]; then echo sha1; return; fi
    if [[ "$V" =~ ^[0-9a-f]{32}$ ]]; then echo md5; return; fi
    if [[ "$V" =~ ^[^@[:space:]]+@[^@[:space:]]+\.[a-z]{2,}$ ]]; then echo email; return; fi
    if [[ "$V" == *:*:* ]]; then echo ipv6; return; fi
    if [[ "$V" =~ ^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$ ]]; then
        [[ "$_IOC_FILE_EXT" == *"|${V##*.}|"* ]] && { echo literal; return; }
        echo domain; return
    fi
    # Nomi di file, chiavi di registro, nomi di mutex: nessun confine sensato
    # da imporre, si resta sulla sottostringa e lo si dichiara.
    echo literal
}

# Escapa i metacaratteri ERE di un valore letterale.
_ioc_quote() { printf '%s' "$1" | sed 's/[][\.^$*+?(){}|\\/]/\\&/g'; }

# ERE con i confini adatti al tipo. Il confronto avviene su testo minuscolo.
_ioc_regex_for() {
    local TYPE="$1" V="${2,,}" Q
    Q=$(_ioc_quote "$V")
    case "$TYPE" in
        regex)  printf '%s' "${V:1:${#V}-2}" ;;
        ipv4)   # niente cifre o punti attaccati: 10.0.0.5 non deve valere
                # dentro 110.0.0.55.
                printf '(^|[^0-9.])%s($|[^0-9.])' "$Q" ;;
        md5|sha1|sha256)
                printf '(^|[^0-9a-f])%s($|[^0-9a-f])' "$Q" ;;
        domain) # I sottodomini contano (mail.evil.com per evil.com), il
                # suffisso no: evil.com.altro.net non e' evil.com.
                printf '(^|[^a-z0-9._-])(([a-z0-9-]+\.)*)%s($|[^a-z0-9.-])' "$Q" ;;
        email)  printf '(^|[^a-z0-9._%%+-])%s($|[^a-z0-9.-])' "$Q" ;;
        *)      printf '%s' "$Q" ;;
    esac
}

# Normalizza un indicatore: toglie il defanging e gli spazi.
_ioc_normalize() {
    local V="$1"
    V="${V#"${V%%[![:space:]]*}"}"; V="${V%"${V##*[![:space:]]}"}"
    V="${V//hxxp/http}"; V="${V//hXXp/http}"; V="${V//HXXP/HTTP}"
    V="${V//\[.\]/.}"; V="${V//(.)/.}"; V="${V//\{.\}/.}"
    V="${V//\[:\]/:}"; V="${V//\[\/\]/\/}"
    V="${V//\[at\]/@}"; V="${V//\[@\]/@}"
    V="${V//\[dot\]/.}"
    printf '%s' "$V"
}

# Estrae type<TAB>value da un bundle STIX 2.x o da un export MISP.
_ioc_from_json() {
    "$PY3" - "$1" << 'PYEOF' 2>/dev/null
import sys, json, re

try:
    doc = json.load(open(sys.argv[1], encoding='utf-8', errors='replace'))
except Exception:
    sys.exit(1)

out = []

# STIX 2.x: il pattern e' un linguaggio a se'. Non lo si interpreta: se ne
# estraggono i confronti di uguaglianza, che coprono il caso reale degli
# indicatori distribuiti nei feed.
EQ = re.compile(r"([a-z0-9:._'\"-]+)\s*=\s*'([^']+)'", re.I)


def stix(obj):
    pat = obj.get('pattern') or ''
    for path, val in EQ.findall(pat):
        p = path.lower()
        if 'hashes' in p:
            out.append(('hash', val))
        elif 'ipv4' in p:
            out.append(('ipv4', val))
        elif 'ipv6' in p:
            out.append(('ipv6', val))
        elif 'domain-name' in p:
            out.append(('domain', val))
        elif 'url' in p:
            out.append(('url', val))
        elif 'email' in p:
            out.append(('email', val))
        elif 'file:name' in p or 'directory:path' in p:
            out.append(('literal', val))
        else:
            out.append(('literal', val))


def misp_attrs(container):
    for a in container.get('Attribute', []) or []:
        v = a.get('value')
        if v:
            out.append((a.get('type', 'literal'), v))
    for o in container.get('Object', []) or []:
        misp_attrs(o)


if isinstance(doc, dict) and doc.get('type') == 'bundle':
    for obj in doc.get('objects', []):
        if obj.get('type') == 'indicator':
            stix(obj)
elif isinstance(doc, dict) and ('Event' in doc or 'response' in doc):
    events = []
    if 'Event' in doc:
        events = [doc['Event']]
    else:
        for item in doc.get('response', []):
            if isinstance(item, dict) and 'Event' in item:
                events.append(item['Event'])
    for ev in events:
        misp_attrs(ev)
else:
    # JSON valido ma schema non riconosciuto. Codice diverso da "non e' JSON":
    # il chiamante deve fermarsi, non trattare il file come lista piatta —
    # finirebbe per caricare le graffe come indicatore letterale.
    sys.exit(2)

if not out:
    sys.exit(2)

for t, v in out:
    print('%s\t%s' % (t, v))
PYEOF
}

# Carica un file IoC: lista piatta (una entry per riga, # ignorate),
# bundle STIX 2.x o export MISP. Il formato viene riconosciuto dal contenuto.
load_ioc_file() {
    local IOCFILE="$1"
    if [[ ! -f "$IOCFILE" ]]; then
        warn "$(t ioc_not_found) $IOCFILE"
        return 1
    fi

    local SRC; SRC=$(mktemp); register_tmp "$SRC"
    local FORMAT="lista"
    if [[ "$(head -c 1 "$IOCFILE" 2>/dev/null)" == "{" ]]; then
        _ioc_from_json "$IOCFILE" > "$SRC"
        local RC=$?
        if [[ "$RC" -eq 2 ]]; then
            err "$(L "JSON riconosciuto ma non e' un bundle STIX 2.x ne' un export MISP:" \
                     "JSON recognised but it is neither a STIX 2.x bundle nor a MISP export:") $IOCFILE"
            info "$(L "Il file non viene caricato: interpretarlo come lista piatta trasformerebbe le righe JSON in indicatori letterali." \
                     "The file is not loaded: reading it as a flat list would turn JSON lines into literal indicators.")"
            return 1
        fi
        [[ "$RC" -eq 0 && -s "$SRC" ]] && FORMAT="STIX/MISP"
    fi
    if [[ "$FORMAT" == "lista" ]]; then
        # Lista piatta: il tipo si deduce dal valore.
        local LINE
        while IFS= read -r LINE || [[ -n "$LINE" ]]; do
            [[ -z "${LINE//[[:space:]]/}" || "$LINE" == \#* ]] && continue
            # Prefisso esplicito "tipo:valore" per i casi che l'euristica non
            # puo' risolvere (payload.com e' un nome di file o un dominio?).
            if [[ "$LINE" =~ ^(ipv4|ipv6|domain|url|email|md5|sha1|sha256|file|literal|regex):(.+)$ ]]; then
                printf '%s\t%s\n' "${BASH_REMATCH[1]}" "${BASH_REMATCH[2]}" >> "$SRC"
            else
                printf 'auto\t%s\n' "$LINE" >> "$SRC"
            fi
        done < "$IOCFILE"
    fi

    local NDUP=0
    local -A _seen=()
    local RAW_TYPE RAW_VAL VAL TYPE RE
    while IFS=$'\t' read -r RAW_TYPE RAW_VAL; do
        [[ -n "$RAW_VAL" ]] || continue
        VAL=$(_ioc_normalize "$RAW_VAL")
        [[ -n "$VAL" ]] || continue
        if [[ -n "${_seen[${VAL,,}]:-}" ]]; then NDUP=$((NDUP + 1)); continue; fi
        _seen[${VAL,,}]=1
        case "$RAW_TYPE" in
            auto|"")            TYPE=$(_ioc_type_of "$VAL") ;;
            file|filename)      TYPE=literal ;;
            ip-src|ip-dst|ip)   TYPE=$(_ioc_type_of "$VAL") ;;
            hash)               TYPE=$(_ioc_type_of "$VAL") ;;
            ipv4|ipv6|domain|url|email|md5|sha1|sha256|literal|regex) TYPE="$RAW_TYPE" ;;
            # MISP usa decine di tipi (hostname, uri, sha512, mutex...): quelli
            # che non conosciamo si fanno riconoscere dal valore, che e' piu'
            # affidabile di una mappatura incompleta.
            *)                  TYPE=$(_ioc_type_of "$VAL") ;;
        esac
        IOC_LIST+=("$VAL")
        IOC_TYPES+=("${TYPE}|${VAL}")
        RE=$(_ioc_regex_for "$TYPE" "$VAL")
        [[ -n "$RE" ]] && IOC_REGEX="${IOC_REGEX:+${IOC_REGEX}|}(${RE})"
    done < "$SRC"

    if [[ ${#IOC_LIST[@]} -eq 0 ]]; then
        warn "$(L "Nessun indicatore utilizzabile in" "No usable indicator in") $IOCFILE"
        return 1
    fi

    # Il conteggio per tipo dice subito se il file e' stato interpretato come
    # ci si aspettava: 40 "literal" al posto di 40 hash sono un errore di
    # formato, e senza questa riga si scoprirebbe solo dall'assenza di match.
    local BREAKDOWN; BREAKDOWN=$(printf '%s\n' "${IOC_TYPES[@]}" | cut -d'|' -f1 \
        | sort | uniq -c | awk '{printf "%s=%s ", $2, $1}')
    ok "$(t ioc_loaded) ${#IOC_LIST[@]} ($FORMAT) - ${BREAKDOWN}"
    [[ "$NDUP" -gt 0 ]] && info "$(L "Duplicati ignorati:" "Duplicates ignored:") $NDUP"
    log_msg "[IOC] $FORMAT: ${#IOC_LIST[@]} indicatori da $IOCFILE - $BREAKDOWN"
    return 0
}

# Vero se il testo contiene almeno un indicatore caricato.
#
# Puro bash, senza fork: viene chiamata riga per riga su report da decine di
# migliaia di righe, e un processo per riga renderebbe i moduli inutilizzabili.
check_ioc() {
    [[ -n "$IOC_REGEX" ]] || return 1
    local TEXT="${1,,}"
    [[ "$TEXT" =~ $IOC_REGEX ]]
}

# Il tipo e il valore dell'indicatore che corrisponde, per i report che vogliono
# dire QUALE indicatore ha fatto match e non solo che ce n'e' stato uno.
matched_ioc() {
    local TEXT="${1,,}" entry TYPE VAL RE
    for entry in "${IOC_TYPES[@]}"; do
        TYPE="${entry%%|*}"; VAL="${entry#*|}"
        RE=$(_ioc_regex_for "$TYPE" "$VAL")
        [[ -n "$RE" && "$TEXT" =~ $RE ]] && { printf '%s|%s' "$TYPE" "$VAL"; return 0; }
    done
    return 1
}
