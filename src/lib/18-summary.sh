
# ================================================================
#  EXECUTIVE SUMMARY
#
#  Novanta report HTML sono un dump, non un'analisi. Chi apre la cartella al
#  termine di --all non sa da dove cominciare, e la differenza fra un tool DFIR
#  e un collector sta esattamente li'.
#
#  Il riepilogo risponde a due domande: *da dove comincio* e *cosa e' successo
#  insieme a cos'altro*. Non risponde a "la macchina e' compromessa": quella e'
#  una conclusione dell'analista, e il report lo dice a chiare lettere.
#
#  TRE SOSTRATI, in ordine di affidabilita' decrescente:
#
#  1. Le righe che i moduli hanno gia' marcato come notevoli (classe
#     'sensitive', applicata da render_pre_block sulle parole chiave che ogni
#     modulo passa). E' un giudizio espresso da chi conosce l'artefatto.
#  2. Gli IoC caricati con --ioc, se presenti.
#  3. Una tabella di regole esplicite, deliberatamente piccola. Le regole
#     girano SOLO sul dato (celle di tabella e righe dei log), mai sulla prosa
#     dei cartigli esplicativi: un motore che cercasse "malware" nell'intera
#     pagina scatterebbe sulle spiegazioni scritte da FIUTO stesso.
#
#  Piu' la correlazione temporale fra moduli diversi, che e' l'unica cosa che
#  nessun singolo report puo' mostrare.
# ================================================================

# Regole di detection. TSV per non litigare con le pipe dentro le regex.
#
#   slug_regex  severita'  MITRE  pattern  titolo§title  perche'§why
#
# severita': CRITICA | ALTA | MEDIA | BASSA
# Il pattern e' case-insensitive e viene applicato a ogni cella/riga di dato.
# Aggiungendo regole: preferire pattern che descrivono la FORMA del dato
# (un base64 lungo dopo -enc) a parole che potrebbero comparire in un nome di
# file legittimo. Una regola rumorosa in un executive summary costa piu' di
# una regola mancante: fa perdere fiducia in tutte le altre.
finding_rules_tsv() {
    cat << 'RULESEOF'
ps_history|ps_scriptblock|ps_transcripts	ALTA	T1059.001	(-enc(odedcommand)?\s+[A-Za-z0-9+/=]{24,})|frombase64string|downloadstring|iex\s*\(|invoke-expression	PowerShell offuscato o download-and-execute§Obfuscated or download-and-execute PowerShell	Comando codificato in base64 o che scarica ed esegue codice dalla rete: e' la forma piu' comune di esecuzione iniziale su Windows.§Base64-encoded command, or one that downloads and runs code from the network: the most common initial-execution shape on Windows.
ifeo	ALTA	T1546.012	debugger	Debugger IFEO impostato§IFEO debugger set	Un valore Debugger sotto Image File Execution Options fa partire un eseguibile arbitrario al posto del programma indicato. Ha usi legittimi rari e un uso malevolo frequente.§A Debugger value under Image File Execution Options launches an arbitrary executable in place of the named program. Legitimate uses are rare, malicious ones frequent.
wmi	ALTA	T1546.003	eventfilter|eventconsumer|__filtertoconsumerbinding	Sottoscrizione WMI permanente§Permanent WMI subscription	Le sottoscrizioni WMI sopravvivono al riavvio e non compaiono in nessuna delle cartelle di avvio che si guardano di solito.§WMI subscriptions survive reboot and appear in none of the startup locations usually inspected.
run_keys|services|scheduled_tasks	ALTA	T1547.001	((users\\[^\\]+\\appdata|windows\\temp|programdata|users\\public)\\[^\\]*\.(exe|dll|ps1|bat|vbs|js|scr))	Persistenza che punta a una cartella scrivibile dall'utente§Persistence pointing at a user-writable folder	Un meccanismo di avvio che esegue un binario in AppData, Temp, ProgramData o Public: il software installato normalmente non sta li'.§A startup mechanism running a binary in AppData, Temp, ProgramData or Public: normally installed software does not live there.
defender_quarantine	ALTA	T1204	.	Defender ha messo elementi in quarantena§Defender quarantined items	Su questa macchina l'antivirus ha rilevato e isolato qualcosa. E' un evento gia' accaduto, non un'ipotesi: da qui si ricava data, percorso e nome della minaccia.§On this machine the antivirus detected and isolated something. This already happened, it is not a hypothesis: date, path and threat name follow from here.
bits	MEDIA	T1197	https?://	Job BITS con destinazione remota§BITS job with a remote endpoint	BITS scarica in background con il processo di sistema: e' un canale di download che non compare nei log del browser.§BITS downloads in the background under a system process: a download channel that never shows up in browser logs.
xplat_esp_bootkit	CRITICA	T1542.003	anomal|unsigned|non firmat|sconosciut|unknown	Anomalia nella partizione EFI§Anomaly in the EFI System Partition	Il codice nella ESP gira prima del sistema operativo e di qualunque EDR, e sopravvive alla reinstallazione. Un'anomalia qui va verificata subito.§Code in the ESP runs before the OS and any EDR, and survives reinstallation. An anomaly here must be verified first.
lsa_secrets|sam|ntds|credential_manager|wlan|linux_cloud_credentials	BASSA	T1003	.	Materiale soggetto a restrizioni nel report§Restricted material in the report	Il report contiene hash di credenziali, chiavi o segreti. Non e' un indizio di compromissione: e' un vincolo su come il report va conservato e condiviso.§The report contains credential hashes, keys or secrets. This is not an indicator of compromise: it is a constraint on how the report may be stored and shared.
linux_pam	ALTA	T1556.003	pam_exec|pam_python|nullok|pam_permit	Configurazione PAM che permette un bypass§PAM configuration allowing a bypass	pam_exec e pam_python eseguono codice a ogni autenticazione; nullok e pam_permit accettano password vuote o qualsiasi credenziale.§pam_exec and pam_python run code on every authentication; nullok and pam_permit accept empty passwords or any credential.
linux_kernel_modules	ALTA	T1014	out-of-tree|unsigned|non firmato|taint	Modulo kernel non firmato o fuori albero§Unsigned or out-of-tree kernel module	Un LKM caricato fuori dai pacchetti della distribuzione puo' nascondere processi, file e connessioni a tutto cio' che gira in user space.§An LKM loaded outside distribution packages can hide processes, files and connections from everything running in user space.
linux_webserver	ALTA	T1505.003	(cmd|exec|shell|passthru|system)=|union\s+select|\.\./\.\./|<\?php	Richiesta compatibile con una webshell§Request consistent with a web shell	Parametri che passano comandi, path traversal o SQL nella query string. Vanno confrontati con la risposta HTTP: un 200 su una di queste e' molto peggio di un 404.§Parameters passing commands, path traversal or SQL in the query string. Compare with the HTTP status: a 200 on one of these is far worse than a 404.
linux_persistence|linux_suid_caps	ALTA	T1053.003	(/tmp/|/dev/shm/|/var/tmp/)	Persistenza o SUID in una directory temporanea§Persistence or SUID in a temporary directory	/tmp, /dev/shm e /var/tmp sono scrivibili da tutti: niente di legittimo installa li' un meccanismo di avvio o un binario SUID.§/tmp, /dev/shm and /var/tmp are world-writable: nothing legitimate installs a startup mechanism or a SUID binary there.
linux_ssh	MEDIA	T1098.004	ssh-(rsa|ed25519|dss)|ecdsa-sha2	Chiave in authorized_keys§Key in authorized_keys	Una chiave pubblica autorizzata e' un accesso permanente che non passa dalla password e non scade. Vanno verificate una per una con il proprietario dell'account.§An authorized public key is permanent access that bypasses the password and never expires. Each must be verified with the account owner.
macos_persistence|macos_loginitems	ALTA	T1543.001	(/tmp/|/users/shared/|/private/var/tmp/)	LaunchAgent o LoginItem in una directory scrivibile§LaunchAgent or LoginItem in a writable directory	Un elemento di avvio che punta a una cartella condivisa o temporanea invece che a /Applications o ai bundle di sistema.§A startup item pointing at a shared or temporary folder rather than /Applications or system bundles.
macos_tcc	MEDIA	T1123	kTCCServiceScreenCapture|kTCCServiceListenEvent|kTCCServiceAccessibility|kTCCServiceMicrophone	Permesso TCC ad alto impatto concesso§High-impact TCC permission granted	Accessibility, cattura schermo, tastiera e microfono: con questi permessi un'applicazione vede e registra tutto quello che fa l'utente.§Accessibility, screen capture, keystrokes and microphone: with these an application sees and records everything the user does.
xplat_sqlite_recovery	MEDIA	T1555	password|passwd|token|api[_-]?key|secret	Credenziale in un record SQLite cancellato§Credential in a deleted SQLite record	Una stringa che sembra una credenziale recuperata da spazio non allocato: era stata cancellata dall'applicazione ma e' rimasta nel file.§A credential-looking string recovered from unallocated space: the application deleted it but it stayed in the file.
yara	ALTA	T1204	.	Match di una regola YARA§YARA rule match	Una firma esterna ha riconosciuto un file sul volume. Il peso dipende da chi ha scritto la regola: la severita' qui e' uniforme perche' FIUTO non puo' giudicare la qualita' di una regola di terze parti.§An external signature recognised a file on the volume. Its weight depends on who wrote the rule: the severity here is uniform because FIUTO cannot judge the quality of a third-party rule.
usb|setupapi	BASSA	T1052.001	.	Supporti rimovibili collegati§Removable media connected	Da solo non significa nulla. Conta per la correlazione: un supporto collegato nella stessa finestra in cui compaiono LNK e attivita' USN e' il profilo di un'esfiltrazione.§On its own it means nothing. It matters for correlation: media connected in the same window as LNK and USN activity is the shape of an exfiltration.
RULESEOF
}

# Scenari di correlazione: insiemi di moduli che, se attivi nella stessa
# finestra, suggeriscono un'ipotesi. Sono IPOTESI DA VERIFICARE e il report le
# presenta come tali — la coincidenza temporale non e' una prova di nesso.
#
#   moduli_regex_separati_da_virgola	severita'	ipotesi§hypothesis
correlation_scenarios_tsv() {
    cat << 'SCENEOF'
usb|setupapi,lnk|jumplists|opensave|shellbags,usn|mft	ALTA	Supporto rimovibile collegato mentre venivano aperti file e il file system registrava modifiche: profilo compatibile con una copia di dati verso l'esterno.§Removable media connected while files were being opened and the file system recorded changes: consistent with copying data out.
ps_history|ps_scriptblock|ps_transcripts,browser|webcache|bits|cloud_sync	ALTA	Esecuzione PowerShell in prossimita' di attivita' di rete: possibile download ed esecuzione di codice.§PowerShell execution close to network activity: possible download and execution of code.
defender_quarantine,prefetch|amcache|bam|userassist	ALTA	Rilevamento antivirus in prossimita' di un'esecuzione registrata: il binario eseguito e quello messo in quarantena vanno confrontati.§Antivirus detection close to a recorded execution: the binary that ran and the one quarantined must be compared.
evtx|linux_logins|macos_logs,run_keys|services|scheduled_tasks|linux_persistence|macos_persistence	ALTA	Accesso al sistema seguito a breve da una modifica dei meccanismi di avvio: profilo di un impianto di persistenza dopo l'ingresso.§System access shortly followed by a change to startup mechanisms: the shape of persistence being installed after entry.
linux_ssh|linux_logins,linux_pam|linux_accounts	ALTA	Accesso remoto in prossimita' di modifiche ad autenticazione o utenze.§Remote access close to changes in authentication or accounts.
SCENEOF
}

# Finestra (minuti) entro cui due eventi di moduli diversi sono considerati
# correlati. 30 minuti e' un compromesso: piu' stretto perde le catene manuali
# di un operatore, piu' largo correla mezza giornata di lavoro normale.
CORRELATION_WINDOW_MIN=30

# Genera executive_summary.html e findings.json in REPORT_BASE_DIR.
generate_executive_summary() {
    [[ -n "$REPORT_BASE_DIR" && -d "$REPORT_BASE_DIR" ]] || return 0

    # I report da leggere sono quelli dei moduli: la dashboard e il riepilogo
    # stesso non sono fonti, sono aggregatori.
    local LIST; LIST=$(mktemp); register_tmp "$LIST"
    local R
    for R in "${GENERATED_REPORTS[@]:-}"; do
        [[ -f "$R" ]] || continue
        case "$(basename "$R")" in
            index.html|executive_summary.html) continue ;;
        esac
        printf '%s\n' "$R" >> "$LIST"
    done
    if [[ ! -s "$LIST" ]]; then
        warn "$(L "Nessun report da riepilogare." "No report to summarise.")"
        return 0
    fi

    local RULES; RULES=$(mktemp); register_tmp "$RULES"
    finding_rules_tsv > "$RULES"
    local SCEN; SCEN=$(mktemp); register_tmp "$SCEN"
    correlation_scenarios_tsv > "$SCEN"
    local IOCF; IOCF=$(mktemp); register_tmp "$IOCF"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCF"

    local JSON="${REPORT_BASE_DIR}/findings.json"
    local BODY; BODY=$(mktemp); register_tmp "$BODY"

    info "$(L "Correlazione dei report in corso..." "Correlating reports...")"
    FIUTO_LANG="${LANG:-en}" FIUTO_WINDOW="$CORRELATION_WINDOW_MIN" \
    run_py_with_lib pylib_timeline "$LIST" "$RULES" "$SCEN" "$IOCF" "$JSON" "$BODY" << 'PYEOF' 2>/dev/null
import sys, os, re, json, html, datetime

list_path, rules_path, scen_path, ioc_path, json_path, body_path = sys.argv[1:7]
IT = os.environ.get('FIUTO_LANG', 'en') == 'it'
WINDOW = int(os.environ.get('FIUTO_WINDOW', '30')) * 60


def L(it, en):
    return it if IT else en


def bilingual(s):
    return L(*s.split('§', 1)) if '§' in s else s


SEV_ORDER = {'CRITICA': 0, 'ALTA': 1, 'MEDIA': 2, 'BASSA': 3}
SEV_WEIGHT = {'CRITICA': 40, 'ALTA': 15, 'MEDIA': 5, 'BASSA': 1}
SEV_LABEL = {'CRITICA': L('CRITICA', 'CRITICAL'), 'ALTA': L('ALTA', 'HIGH'),
             'MEDIA': L('MEDIA', 'MEDIUM'), 'BASSA': L('BASSA', 'LOW')}
SEV_CSS = {'CRITICA': 'crit', 'ALTA': 'high', 'MEDIA': 'med', 'BASSA': 'low'}

MAX_EXAMPLES = 6          # esempi mostrati per finding
MAX_EXAMPLE_LEN = 300

reports = [p.strip() for p in open(list_path, encoding='utf-8') if p.strip()]
iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]

rules = []
for line in open(rules_path, encoding='utf-8'):
    f = line.rstrip('\n').split('\t')
    if len(f) != 6:
        continue
    slug_rx, sev, mitre, pat, title, why = f
    try:
        rules.append((re.compile(f'^({slug_rx})$'), sev, mitre,
                      re.compile(pat, re.I), title, why))
    except re.error:
        continue

scenarios = []
for line in open(scen_path, encoding='utf-8'):
    f = line.rstrip('\n').split('\t')
    if len(f) != 3:
        continue
    groups = [re.compile(f'^({g})$') for g in f[0].split(',')]
    scenarios.append((groups, f[1], f[2]))


def slug_of(path):
    return re.sub(r'_\d{8}_\d{6}$', '', os.path.basename(os.path.dirname(path)))


def rel(path):
    base = os.path.dirname(json_path)
    return os.path.relpath(path, base)


findings = []
events = []          # (epoch_ordinabile, slug, message)
per_module = {}

for path in reports:
    slug = slug_of(path)
    data = read_report_data(path)
    sens = read_sensitive_lines(path)
    evs = extract_events(path)
    for ev in evs:
        events.append((ev['datetime'], slug, ev['message']))
    per_module[slug] = {'report': rel(path), 'rows': len(data), 'events': len(evs)}

    # 1) righe marcate dal modulo stesso
    if sens:
        findings.append({
            'severity': 'MEDIA',
            'module': slug,
            'report': rel(path),
            'mitre': '',
            'title': L('Righe segnalate dal modulo', 'Rows flagged by the module'),
            'why': L('Il modulo ha marcato queste righe con le proprie parole chiave: '
                     'sono il punto in cui l\'autore dell\'artefatto guarderebbe per primo.',
                     'The module flagged these rows with its own keywords: they are where '
                     'someone who knows the artefact would look first.'),
            'count': len(sens),
            'source': 'modulo',
            'examples': [s[:MAX_EXAMPLE_LEN] for s in sens[:MAX_EXAMPLES]],
        })

    # 2) IoC caricati con --ioc
    if iocs:
        hits = [d for d in data if any(i in d.lower() for i in iocs)]
        if hits:
            findings.append({
                'severity': 'ALTA',
                'module': slug,
                'report': rel(path),
                'mitre': '',
                'title': L('Corrispondenza con gli IoC forniti', 'Match against the supplied IoCs'),
                'why': L('Il dato contiene uno degli indicatori passati con --ioc.',
                         'The data contains one of the indicators passed with --ioc.'),
                'count': len(hits),
                'source': 'ioc',
                'examples': [h[:MAX_EXAMPLE_LEN] for h in hits[:MAX_EXAMPLES]],
            })

    # 3) regole esplicite, solo sul dato
    for slug_rx, sev, mitre, pat, title, why in rules:
        if not slug_rx.match(slug):
            continue
        hits = [d for d in data if pat.search(d)]
        if not hits:
            continue
        findings.append({
            'severity': sev,
            'module': slug,
            'report': rel(path),
            'mitre': mitre,
            'title': bilingual(title),
            'why': bilingual(why),
            'count': len(hits),
            'source': 'regola',
            'examples': [h[:MAX_EXAMPLE_LEN] for h in hits[:MAX_EXAMPLES]],
        })

findings.sort(key=lambda f: (SEV_ORDER.get(f['severity'], 9), -f['count'], f['module']))

# --- correlazione temporale ------------------------------------------------
# Gli eventi sono ordinati per stringa ISO, che per questo formato coincide con
# l'ordine cronologico. La finestra scorre e si chiude quando il salto supera
# CORRELATION_WINDOW_MIN.
_EPOCH = datetime.datetime(1970, 1, 1)


def to_secs(iso):
    """Secondi assoluti. Un'aritmetica approssimata sui campi della data
    (giorni=31 al mese) sballerebbe i salti attorno ai cambi di mese e
    spezzerebbe cluster legittimi senza che nulla lo segnali."""
    try:
        return int((datetime.datetime.fromisoformat(iso[:19]) - _EPOCH).total_seconds())
    except Exception:
        return None


events = [(to_secs(dt), dt, slug, msg) for dt, slug, msg in events]
events = sorted((e for e in events if e[0] is not None), key=lambda e: e[0])

clusters = []
cur = []
for e in events:
    if cur and e[0] - cur[-1][0] > WINDOW:
        clusters.append(cur)
        cur = []
    cur.append(e)
if cur:
    clusters.append(cur)

correlations = []
for cl in clusters:
    mods = sorted({e[2] for e in cl})
    if len(mods) < 2:
        continue
    for groups, sev, hypo in scenarios:
        matched = []
        for g in groups:
            m = [x for x in mods if g.match(x)]
            if not m:
                break
            matched.append(m[0])
        else:
            correlations.append({
                'severity': sev,
                'from': cl[0][1],
                'to': cl[-1][1],
                'modules': matched,
                'all_modules': mods,
                'events': len(cl),
                'hypothesis': bilingual(hypo),
            })
            break

correlations.sort(key=lambda c: (SEV_ORDER.get(c['severity'], 9), c['from']))

# --- punteggio -------------------------------------------------------------
# Somma di pesi, tetto 100. E' una CODA DI LAVORO, non una misura di
# compromissione: due macchine con lo stesso punteggio possono avere storie
# opposte. Il report riporta la formula perche' un numero senza definizione e'
# peggio di nessun numero.
score = min(100, sum(SEV_WEIGHT.get(f['severity'], 0) for f in findings)
            + sum(SEV_WEIGHT.get(c['severity'], 0) for c in correlations))

counts = {s: sum(1 for f in findings if f['severity'] == s) for s in SEV_ORDER}

with open(json_path, 'w', encoding='utf-8') as fh:
    json.dump({
        'schema_version': 1,
        'score': score,
        'score_formula': 'CRITICA=40 ALTA=15 MEDIA=5 BASSA=1, somma limitata a 100',
        'score_meaning': ('Ordina la coda di lavoro. NON misura la compromissione: '
                          'e\' la somma delle severita\' dei riscontri, non un giudizio.'),
        'counts': counts,
        'modules': per_module,
        'findings': findings,
        'correlations': correlations,
        'correlation_window_minutes': WINDOW // 60,
    }, fh, indent=2, ensure_ascii=False)


# --- corpo HTML ------------------------------------------------------------
def esc(s):
    return html.escape(str(s))


out = []
out.append("<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>")
out.append('<b>' + esc(L('Come leggere questa pagina', 'How to read this page')) + '</b><br>')
out.append(esc(L(
    "Questa pagina ordina il lavoro, non lo conclude. Ogni voce dice da quale modulo viene, "
    "su quale dato ha fatto match e perche' quel dato conta: il giudizio su cosa sia successo "
    "resta all'analista, che deve aprire il report del modulo e guardare il contesto. "
    "L'assenza di riscontri non e' un attestato di pulizia — significa che le regole applicate, "
    "che sono poche e deliberatamente conservative, non hanno trovato nulla.",
    "This page orders the work, it does not conclude it. Every entry states which module it comes "
    "from, which data it matched and why that data matters: the judgement on what happened stays "
    "with the analyst, who must open the module report and look at the context. "
    "No findings is not a clean bill of health — it means the applied rules, which are few and "
    "deliberately conservative, found nothing.")))
out.append('</div></div>')

if correlations:
    out.append("<h2 style='font-family:var(--mono);font-size:.9rem;margin:1.4rem 0 .7rem;color:var(--accent4)'>"
               + esc(L('Correlazioni temporali fra moduli', 'Cross-module time correlations')) + '</h2>')
    out.append("<div class='card' style='margin-bottom:1rem'><div style='padding:.8rem 1.5rem;font-size:.78rem;color:var(--text-mid);line-height:1.6'>"
               + esc(L(
                   "Moduli diversi con attivita' nella stessa finestra di %d minuti. La coincidenza "
                   "temporale non e' un nesso di causa: sono ipotesi da verificare aprendo i report "
                   "citati." % (WINDOW // 60),
                   "Different modules active within the same %d-minute window. Temporal coincidence "
                   "is not causation: these are hypotheses to verify by opening the reports named."
                   % (WINDOW // 60)))
               + '</div></div>')
    out.append("<div class='cards'>")
    for c in correlations[:30]:
        out.append("<div class='card'><div class='card-header'>")
        out.append("<div class='badge sev-%s'>%s</div>" % (SEV_CSS[c['severity']], esc(SEV_LABEL[c['severity']])))
        out.append("<div class='user-info' style='margin-left:.8rem'><div class='uname'>%s → %s</div>"
                   % (esc(c['from'].replace('T', ' ')), esc(c['to'].replace('T', ' '))))
        out.append("<div class='upath'>%s · %d %s</div></div></div>"
                   % (esc(' + '.join(c['modules'])), c['events'], esc(L('eventi', 'events'))))
        out.append("<div style='padding:1rem 1.5rem;font-size:.82rem;line-height:1.7'>%s" % esc(c['hypothesis']))
        others = [m for m in c['all_modules'] if m not in c['modules']]
        if others:
            out.append("<div style='margin-top:.6rem;font-size:.72rem;color:var(--text-dim)'>%s %s</div>"
                       % (esc(L('Altri moduli attivi nella finestra:', 'Other modules active in the window:')),
                          esc(', '.join(others[:12]))))
        out.append('</div></div>')
    out.append('</div>')

out.append("<h2 style='font-family:var(--mono);font-size:.9rem;margin:1.6rem 0 .7rem;color:var(--accent)'>"
           + esc(L('Riscontri per severita\'', 'Findings by severity')) + '</h2>')

if not findings:
    out.append("<div class='card'><div style='padding:1.4rem 1.5rem;font-size:.85rem'>%s</div></div>"
               % esc(L('Nessuna regola ha prodotto riscontri. Vedi l\'avvertenza in testa: '
                       'non equivale a un esito negativo dell\'analisi.',
                       'No rule produced a finding. See the notice above: this is not equivalent '
                       'to a negative analysis result.')))
else:
    out.append("<div class='cards'>")
    for f in findings:
        out.append("<div class='card'><div class='card-header'>")
        out.append("<div class='badge sev-%s'>%s</div>" % (SEV_CSS[f['severity']], esc(SEV_LABEL[f['severity']])))
        out.append("<div class='user-info' style='margin-left:.8rem'><div class='uname'>%s</div>" % esc(f['title']))
        sub = f['module']
        if f['mitre']:
            sub += ' · MITRE ' + f['mitre']
        out.append("<div class='upath'>%s</div></div>" % esc(sub))
        out.append("<div style='margin-left:auto;font-family:var(--mono);font-size:1.1rem;font-weight:800'>%d</div>"
                   % f['count'])
        out.append('</div>')
        out.append("<div style='padding:.9rem 1.5rem;font-size:.8rem;line-height:1.7'>%s" % esc(f['why']))
        if f['examples']:
            out.append("<pre style='margin-top:.7rem;padding:.7rem;background:var(--bg);border:1px solid var(--border);"
                       "font-size:.7rem;overflow-x:auto;white-space:pre-wrap;word-break:break-all'>%s</pre>"
                       % esc('\n'.join(f['examples'])))
            if f['count'] > len(f['examples']):
                out.append("<div style='font-size:.7rem;color:var(--text-dim)'>%s</div>"
                           % esc(L('… e altri %d nel report del modulo' % (f['count'] - len(f['examples'])),
                                   '… and %d more in the module report' % (f['count'] - len(f['examples'])))))
        out.append("<div style='margin-top:.7rem;font-size:.72rem'><a href='%s' style='color:var(--accent)'>%s</a></div>"
                   % (esc(f['report']), esc(L('Apri il report del modulo', 'Open the module report'))))
        out.append('</div></div>')
    out.append('</div>')

with open(body_path, 'w', encoding='utf-8') as fh:
    fh.write('\n'.join(out))
PYEOF

    if [[ ! -s "$JSON" ]]; then
        warn "$(L "Riepilogo non generato." "Summary not generated.")"
        return 0
    fi

    local SCORE NCRIT NHIGH NMED NLOW NCORR
    read -r SCORE NCRIT NHIGH NMED NLOW NCORR < <("$PY3" -c '
import json, sys
d = json.load(open(sys.argv[1]))
c = d["counts"]
print(d["score"], c["CRITICA"], c["ALTA"], c["MEDIA"], c["BASSA"], len(d["correlations"]))' "$JSON" 2>/dev/null)

    local OUT="${REPORT_BASE_DIR}/executive_summary.html"
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local STATS
    STATS="$(stat_box "$(L "Priorità" "Priority")" "$SCORE" "$([[ "${SCORE:-0}" -ge 40 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Critici" "Critical")" "${NCRIT:-0}")"
    STATS+="$(stat_box "$(L "Alti" "High")" "${NHIGH:-0}")"
    STATS+="$(stat_box "$(L "Medi" "Medium")" "${NMED:-0}" "info")"
    STATS+="$(stat_box "$(L "Bassi" "Low")" "${NLOW:-0}" "info")"
    STATS+="$(stat_box "$(L "Correlazioni" "Correlations")" "${NCORR:-0}" "$([[ "${NCORR:-0}" -gt 0 ]] && echo warn || echo info)")"

    {
        html_header "Executive Summary"
        html_page_header "SUM" "Executive Summary" \
            "$(L "Riepilogo di sessione — priorità e correlazioni" "Session summary — priorities and correlations")" \
            "$SCAN" "$WIN_ROOT"
        printf "<div class='statsbar'>%s</div>\n" "$STATS"
        echo "<main>"
        pre_style_block
        _summary_style_block
        time_window_html
        cat "$BODY"
        printf "<div class='card' style='margin-top:1.2rem'><div style='padding:1rem 1.5rem;font-size:.75rem;color:var(--text-mid);line-height:1.7'><b>%s</b> %s</div></div>\n" \
            "$(L "Come è calcolata la priorità:" "How the priority is computed:")" \
            "$(L "somma dei pesi dei riscontri (critico 40, alto 15, medio 5, basso 1), limitata a 100. Ordina la coda di lavoro; non misura la compromissione. Due macchine con lo stesso valore possono avere storie opposte." \
                 "sum of finding weights (critical 40, high 15, medium 5, low 1), capped at 100. It orders the work queue; it does not measure compromise. Two machines with the same value can have opposite stories.")"
        echo "</main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$OUT"

    register_report "$OUT"
    echo ""
    ok "$(L "Executive summary:" "Executive summary:") ${BOLD}${OUT}"
    info "$(L "Riscontri:" "Findings:") ${BOLD}$(( ${NCRIT:-0} + ${NHIGH:-0} + ${NMED:-0} + ${NLOW:-0} ))${RESET}  ·  $(L "critici/alti:" "critical/high:") ${BOLD}$(( ${NCRIT:-0} + ${NHIGH:-0} ))${RESET}  ·  $(L "correlazioni:" "correlations:") ${BOLD}${NCORR:-0}"
    log_msg "[SUMMARY] score=$SCORE crit=${NCRIT:-0} high=${NHIGH:-0} corr=${NCORR:-0}"
    open_report_prompt "$OUT"
}

# Stili delle sole severita': il resto lo porta html_header.
_summary_style_block() {
    cat << 'CSSEOF'
<style>
  .badge.sev-crit { background:#ff7b72; color:#0d1117; }
  .badge.sev-high { background:#f0883e; color:#0d1117; }
  .badge.sev-med  { background:#58a6ff; color:#0d1117; }
  .badge.sev-low  { background:#1e2d3d; color:#c9d1d9; }
</style>
CSSEOF
}
