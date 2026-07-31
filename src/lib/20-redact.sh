
# ================================================================
#  MODALITA' --redact
#
#  I report contengono materiale che serve all'analisi e non deve uscire dallo
#  studio: hash NTLM da SAM e NTDS, PSK Wi-Fi, token OAuth, chiavi private,
#  segreti LSA. Finche' l'unica opzione e' "mandare il report intero o non
#  mandarlo", quel materiale finisce in una mail a un cliente o in un allegato
#  di un ticket. --redact esiste per rendere praticabile la terza via.
#
#  DUE SCELTE DI FONDO.
#
#  1. Il report originale NON viene toccato. Accanto ad esso nasce un
#     report.redacted.html. Oscurare sul posto distruggerebbe evidenza per una
#     necessita' di comunicazione, che e' un baratto che non spetta al tool
#     fare. Il manifesto di custodia elenca entrambi, con ruoli distinti.
#
#  2. Si oscura per CONTESTO, non per forma. Un hash SHA-256 e un hash NTLM
#     sono entrambi stringhe esadecimali: il primo e' l'impronta di integrita'
#     di un reperto e cancellarlo distruggerebbe la catena di custodia, il
#     secondo e' una credenziale. A distinguerli e' l'etichetta accanto, non il
#     numero di caratteri. Per questo le regole guardano il campo, non il
#     valore, e i conteggi per regola sono stampati: una regola che non scatta
#     mai o che scatta ovunque si vede subito.
#
#  Ogni copia oscurata dichiara in testa cosa e' stata: un file che sembra un
#  report ma non lo e' del tutto, senza dirlo, e' peggio di nessun file.
# ================================================================

REDACT=false          # --redact: genera una copia oscurata di ogni report
REDACT_DEFANG=false   # --defang: nella copia oscurata rende inerti URL e IP

# Numero di copie oscurate prodotte nella sessione (per il riepilogo finale).
_redact_dir() {
    local D="${TMPDIR:-/tmp}/fiuto_redact_$$"
    [[ -d "$D" ]] || mkdir -p "$D" 2>/dev/null || return 1
    echo "$D"
}

redacted_count() {
    local D="${TMPDIR:-/tmp}/fiuto_redact_$$"
    [[ -f "${D}/done" ]] && wc -l < "${D}/done" || echo 0
}

# redact_report <report.html>
# Scrive <report>.redacted.html accanto all'originale. L'originale resta
# intatto. Non fallisce mai in modo rumoroso: se non c'e' niente da oscurare
# la copia viene prodotta ugualmente, cosi' chi condivide non deve chiedersi
# perche' per un modulo il file manchi.
redact_report() {
    local HTML="$1"
    [[ -n "$HTML" && -f "$HTML" ]] || return 0
    case "$HTML" in *.redacted.html) return 0 ;; esac
    local OUT="${HTML%.html}.redacted.html"

    local N
    N=$(FIUTO_DEFANG="$REDACT_DEFANG" FIUTO_LANG="${LANG:-en}" \
        "$PY3" - "$HTML" "$OUT" << 'PYEOF' 2>/dev/null
import sys, os, re, html

src, dst = sys.argv[1], sys.argv[2]
DEFANG = os.environ.get('FIUTO_DEFANG', 'false') == 'true'
IT = os.environ.get('FIUTO_LANG', 'en') == 'it'


def L(i, e):
    return i if IT else e


try:
    text = open(src, encoding='utf-8', errors='replace').read()
except Exception:
    sys.exit(1)

MARK = '[REDACTED]'
counts = {}


def sub(name, pattern, repl, flags=re.I):
    global text
    text, n = re.subn(pattern, repl, text, flags=flags)
    if n:
        counts[name] = counts.get(name, 0) + n


# 1) Chiavi private in formato PEM: il blocco intero, non solo l'intestazione.
sub('chiave privata PEM',
    r'-----BEGIN [A-Z ]*PRIVATE KEY-----.*?-----END [A-Z ]*PRIVATE KEY-----',
    '-----BEGIN PRIVATE KEY----- ' + MARK + ' -----END PRIVATE KEY-----',
    re.I | re.DOTALL)

# 2) Coppie etichetta/valore. L'etichetta resta: sapere CHE c'era una password
#    e' un dato dell'analisi, il suo valore no.
LABELS = (r'password|passwd|pwd|psk|pre[- ]?shared[- ]?key|keymaterial|'
          r'key\s*material|secret|client_secret|access_token|refresh_token|'
          r'bearer|token|api[_-]?key|apikey|aws_secret_access_key|'
          r'connection\s*string|nt\s*hash|ntlm\s*hash|lm\s*hash|hash\s*nt')
sub('campo etichettato',
    r'(\b(?:' + LABELS + r')\b\s*(?:[:=]|&gt;|&#61;)\s*)(?:&quot;|["\']?)([^\s<"\'&]{3,})',
    lambda m: m.group(1) + MARK)

# 3) Riversamenti SAM/NTDS: utente:RID:LM:NT:::. Utente e RID restano — sono
#    l'informazione forense — gli hash no.
sub('hash SAM/NTDS',
    r'([^\s:<>]{1,64}):(\d+):([0-9a-f]{32}):([0-9a-f]{32}):::',
    lambda m: '%s:%s:%s:%s:::' % (m.group(1), m.group(2), MARK, MARK))

# 4) Materiale WLAN, sia come tag sia come entita' HTML.
sub('WLAN keyMaterial',
    r'(&lt;keyMaterial&gt;|<keyMaterial>)([^<&]+)',
    lambda m: m.group(1) + MARK)

# 5) Credenziali con forma inequivocabile: il contesto non serve, la stringa
#    stessa non e' altro che una credenziale.
sub('AWS access key', r'\bAKIA[0-9A-Z]{16}\b', MARK)
sub('JWT', r'\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b', MARK)
sub('Slack/GitHub token', r'\b(?:xox[baprs]-[A-Za-z0-9-]{10,}|gh[pousr]_[A-Za-z0-9]{20,})\b', MARK)

# 6) Defanging facoltativo di URL e IP nella copia condivisibile.
if DEFANG:
    sub('URL defanged', r'\bhttps?://', lambda m: m.group(0).replace('http', 'hxxp'), 0)
    sub('IPv4 defanged',
        r'(?<![\w.])((?:\d{1,3}\.){3}\d{1,3})(?![\w.])',
        lambda m: m.group(1).replace('.', '[.]'), 0)

rows = ''.join('<tr><td class="mono">%s</td><td class="mono">%d</td></tr>'
               % (html.escape(k), v) for k, v in sorted(counts.items()))
total = sum(counts.values())

banner = (
    '<div class="card" style="margin:1rem 0;border-color:#f0883e">'
    '<div style="padding:1rem 1.5rem;font-size:.82rem;line-height:1.7">'
    '<b>' + L('COPIA OSCURATA - non e\' il report originale',
              'REDACTED COPY - this is not the original report') + '</b><br>'
    + L('Da questa copia sono stati rimossi i valori delle credenziali (hash NTLM, PSK, '
        'token, chiavi private). Il report originale, integro, sta nello stesso percorso '
        'senza il suffisso <code>.redacted</code>: e\' quello da conservare come reperto e '
        'da citare in una perizia.',
        'Credential values (NTLM hashes, PSKs, tokens, private keys) have been removed from '
        'this copy. The original, intact report sits in the same path without the '
        '<code>.redacted</code> suffix: that is the one to retain as evidence and to cite '
        'in a report.')
    + '<br><br>'
    + L('Gli hash SHA-256 dei reperti <b>non</b> sono stati toccati: sono dati di '
        'integrita\', non segreti, e rimuoverli spezzerebbe la catena di custodia.',
        'Evidence SHA-256 hashes were <b>not</b> touched: they are integrity data, not '
        'secrets, and removing them would break the chain of custody.'))
if rows:
    banner += ('<br><br><table><tr><th>' + L('Regola', 'Rule') + '</th><th>'
               + L('Occorrenze', 'Occurrences') + '</th></tr>' + rows + '</table>')
else:
    banner += ('<br><br>' + L('Nessuna regola ha trovato materiale da oscurare in questo report.',
                              'No rule found material to redact in this report.'))
banner += '</div></div>'

# Il cartiglio va subito dopo l'apertura di <main>: e' la prima cosa che si
# legge, prima di qualunque dato.
if '<main>' in text:
    text = text.replace('<main>', '<main>\n' + banner, 1)
else:
    text = banner + text
text = text.replace('<title>', '<title>[REDACTED] ', 1)

with open(dst, 'w', encoding='utf-8') as fh:
    fh.write(text)
print(total)
PYEOF
    )

    [[ -f "$OUT" ]] || return 0
    local D; D=$(_redact_dir) && printf '%s\t%s\n' "${N:-0}" "$OUT" >> "${D}/done" 2>/dev/null
    evidence_note "$OUT" "copia oscurata" 2>/dev/null || true
    log_msg "[REDACT] $OUT - ${N:-0} occorrenze oscurate"
    return 0
}

# Riepilogo di fine sessione. Va detto quante copie esistono e dove: una copia
# oscurata che nessuno sa di avere non serve a niente, e una che si confonde
# con l'originale e' un rischio.
redact_summary() {
    [[ "$REDACT" == "true" ]] || return 0
    local D="${TMPDIR:-/tmp}/fiuto_redact_$$"
    [[ -f "${D}/done" ]] || return 0
    local NC NT
    NC=$(wc -l < "${D}/done")
    NT=$(awk -F'\t' '{s+=$1} END{print s+0}' "${D}/done")
    echo ""
    ok "$(L "Copie oscurate prodotte:" "Redacted copies produced:") ${BOLD}${NC}${RESET}  ·  $(L "valori rimossi:" "values removed:") ${BOLD}${NT}"
    info "$(L "Sono i file *.redacted.html. Gli originali restano integri accanto a loro." \
             "They are the *.redacted.html files. The originals remain intact beside them.")"
}
