# FIUTO — Piano di implementazione

Documento di lavoro per portare FIUTO da 2.1 a 3.0. Le fasi 1-7 sono complete
e rilasciate nella 2.3; restano le voci elencate in fondo come debito noto. È pensato per essere
ripreso a distanza di tempo, anche da un'altra sessione o da un'altra persona:
ogni fase dichiara **cosa fare**, **dove**, **come verificarlo** e **perché**.

Stato aggiornato al: **2026-07-31** (versione 2.3, **tutte le fasi complete**).

---

## Come usare questo documento

1. Le fasi sono ordinate per dipendenza, non per importanza. La Fase 1 (split)
   è prerequisito pratico di tutto il resto: aggiungere ~25 moduli a un file
   già da 13k righe lo renderebbe ingestibile.
2. Ogni fase è un branch e una PR. Non accorpare fasi diverse: il valore del
   piano è poter fermarsi in mezzo senza lasciare l'albero incoerente.
3. **Prima di ogni commit**, la porta di qualità è sempre la stessa:

   ```bash
   bash -n fiuto.sh
   shellcheck -S warning -x fiuto.sh
   python3 tests/lint_embedded_python.py fiuto.sh
   bats tests/
   ```

   Tutti e quattro devono passare. Se una modifica richiede di abbassare la
   soglia, il problema è la modifica.

---

## Stato attuale (fatto in v2.1)

| Area | Stato |
|---|---|
| CI (bash -n, ShellCheck, bats, parser Python su 3.9 + 3.12) | ✅ |
| Suite di test — 258 test bats | ✅ |
| Replay transaction log registro (`.LOG1`/`.LOG2`) | ✅ |
| Export JSONL / schema Timesketch (`--jsonl`) | ✅ |
| macOS: FSEvents, Spotlight | ✅ |
| Linux: auditd, Container (Docker/Podman) | ✅ |
| Fase 1 — split `src/` + `build.sh` (file singolo generato) | ✅ |
| Fase 1 — registro moduli data-driven anche per Windows | ✅ |
| Fase 2 — 11 moduli Windows nuovi (40-50) | ✅ |
| Fase 3 — 5 moduli Linux nuovi (17-21) | ✅ |
| Fase 4 — 6 moduli macOS nuovi (14-19) | ✅ |
| Fase 5 — 2 moduli cross-OS (SQLite recovery, ESP/bootkit) | ✅ |
| Libreria Python condivisa SQLite (`src/lib/14-pylib-sqlite.sh`) | ✅ |
| Fase 6.1 — chain of custody (`evidence_manifest.json`) | ✅ |
| Fase 6.2 — finestra temporale `--since`/`--until` + fuso del volume | ✅ |
| Fase 6.3 — executive summary, scoring e correlazione cross-modulo | ✅ |
| Fase 6.4 — scansione YARA (`--yara`), ambito dichiarato | ✅ |
| Fase 6.5 — motore Sigma sugli EVTX (`--sigma`), sottoinsieme dichiarato | ✅ |
| Fase 6.6 — esecuzione parallela (`--jobs N`) con esito invariante | ✅ |
| Fase 6.7 — motore IoC tipizzato, defanging, import STIX/MISP | ✅ |
| Fase 6.8 — `--redact` / `--defang`, copie condivisibili | ✅ |
| Fase 6.9 — immagini E01/raw e volumi cifrati (`--image`) | ✅ |
| Fase 7 — immagine Docker con tutte le dipendenze + CI su GHCR | ✅ |
| Build macOS (`buildMac.sh` + `src/compat/macos.sh`) | ✅ |
| Libreria Python condivisa Sigma (`src/lib/19-pylib-sigma.sh`) | ✅ |
| Libreria Python condivisa timeline (`src/lib/17-pylib-timeline.sh`) | ✅ |
| Flag `defer` nel registro (numerazione stabile) | ✅ |
| Libreria Python condivisa LevelDB/Snappy (`src/lib/13-pylib-leveldb.sh`) | ✅ |

Bug corretti in v2.1, da non reintrodurre:

- **Modulo 38 (PAD Offline)** non compilava affatto (f-string con quoting
  annidato). Il linter dei parser incorporati esiste per questo.
- **`ci_find_file` duplicata** con semantiche incompatibili. Il test
  strutturale "nessuna funzione è definita due volte" presidia il caso.
- **`find | xargs` senza `-print0`** perdeva i file con spazi nel nome.

---

## Fase 1 — Split per sistema operativo

**Obiettivo:** da un `fiuto.sh` monolitico a un albero navigabile, senza
cambiare il comportamento osservabile.

### Struttura realizzata

```
build.sh                   genera fiuto.sh concatenando src/build.order
fiuto.sh                   GENERATO — non modificare a mano
src/
  header.sh                shebang, banner, set -uo pipefail
  build.order              ordine di concatenazione (unica fonte di verità)
  lib/
    00-core.sh             globali, colori, cleanup, rilevamento Python
    01-i18n.sh             L(), t(), selezione lingua
    02-ui.sh               banner, info/ok/warn/err, prompt
    03-util.sh             html_esc, sha256_file, log_msg
    04-ioc.sh              caricamento e match IoC
    05-fs.sh               ci_find_*, detect_os_type, home utenti, sqlite, plist
    06-hostinfo.sh         raccolta informazioni macchina
    07-registry.sh         get_hive, recover_hive (replay .LOG1/.LOG2), regipy
    08-report.sh           HTML: header/footer, card, finish_report
    09-export.sh           export JSONL (Timesketch)
    10-mount.sh            rilevamento volumi, set_win_root, dir dei report
    11-runner.sh           menu, dispatcher, esecuzione batch
    12-registries.sh       MODULES_LINUX / MODULES_MACOS
  modules/{win,linux,macos,xplat}/   un file per modulo
  main.sh                  main() + entrypoint
tests/
docs/
```

### Decisione presa

**Distribuzione a file singolo mantenuta.** `build.sh` concatena i sorgenti
elencati in `src/build.order` producendo `fiuto.sh`, che resta versionato: il
tool si copia su una workstation forense e funziona senza installare nulla.
La CI verifica con `./build.sh --check` che il file committato coincida con il
build, così le due forme non possono divergere in silenzio.

Il build è una **concatenazione pura**, senza trasformazioni: il file generato
resta leggibile e un diff fra due build dice esattamente cosa è cambiato.

### Fatto (1/2)

Struttura prodotta — 82 sorgenti:

```
src/header.sh              intestazione + set -uo pipefail
src/lib/00-core.sh .. 12-registries.sh    13 file di libreria
src/modules/win/           39 moduli
src/modules/linux/         15 moduli
src/modules/macos/         12 moduli
src/modules/xplat/          1 modulo
src/main.sh                main() + entrypoint
```

Lo split è stato eseguito meccanicamente, non a mano, con verifica di
equivalenza: il multiset delle righe prima e dopo è identico a meno delle 16
righe del banner "file generato". **Zero righe perse.**

Nota per chi rifà un'operazione simile: un parser riga-per-riga non basta,
perché gli heredoc contengono CSS e Python con `}` a colonna 0 e le funzioni
one-liner (`nome() { ...; }`) non hanno mai una graffa di chiusura a colonna 0.

### Fatto (2/2) — registro moduli unificato

Windows usa ora `MODULES_WIN`, come Linux e macOS. Rimossi tre dispatcher:
`print_menu` (menu stampato a mano), `run_module_by_number` e un **secondo**
`case` da 39 rami inline in `main()` per `--module`, che nessuno aveva
notato: i due erano ancora identici, ma erano duplicazione in attesa di
divergere. Netto: −204 righe.

Il formato del registro è stato esteso in modo retrocompatibile:

```
"funzione|Nome|COLORE|descrizione[|guardia]"
```

- **Bilinguismo**: `Nome` e `descrizione` accettano la forma
  `italiano§english`; senza separatore lo stesso testo vale per entrambe le
  lingue. Serviva per non perdere le etichette bilingui che il menu Windows
  aveva e che i registri Linux/macOS non supportavano. `reg_text` fa la
  selezione. Nota: `~` non è utilizzabile come separatore, una descrizione
  Linux inizia con `~/.local/share/Trash`.
- **Guardia** facoltativa: funzione che ritorna 0 se il modulo va eseguito,
  oppure stampa il motivo e ritorna non-zero per farlo saltare in batch.
  Serviva a preservare l'unico comportamento speciale del batch Windows: il
  modulo 38 (PAD Offline) va saltato se il volume non è un Domain Controller
  (`_guard_pad_offline`).

Attenzione per il futuro: le descrizioni contengono `$Recycle.Bin` e
`$UsnJrnl:$J`. Dentro le virgolette doppie delle voci di registro il dollaro
**va escapato** (`\$`), altrimenti bash espande una variabile inesistente e
la descrizione si svuota. C'è un test che lo presidia.

La numerazione è stata verificata identica a quella precedente confrontando
l'ordine del registro con entrambi i vecchi dispatcher: `--module N` e
`--modules 1,4,6-8` invocano esattamente le stesse funzioni di prima.

### Verifica in essere

I test strutturali coprono ora tutti e tre i registri: funzioni esistenti,
campi obbligatori, colori definiti, guardie esistenti, nomi univoci, dollari
letterali preservati, assenza dei dispatcher rimossi, `reg_text` nelle due
lingue, e coerenza fra numero di moduli e tabelle del README.

---

## Fase 2 — Moduli Windows

Un file per modulo in `src/modules/win/`. Priorità decrescente. ✅ = fatto in v2.2 (moduli 40-50). **Fase 2 completata.**

I nuovi moduli usano `finish_report` + `generic_card_html` + `_rows_to_table`
invece di comporre l'HTML a mano come i 39 storici: molto meno codice e un
solo punto da toccare quando cambia il layout dei report. Conviene seguire
questa strada anche per i restanti.

| # | Modulo | Fonte | Note di implementazione |
|---|---|---|---|
| 1 | ✅ **Volume Shadow Copies** | snapshot VSS | Enumerazione + rilancio dei moduli sullo snapshot: abilita l'analisi differenziale storica. Oggi VSS è solo *citato* nei suggerimenti. |
| 2 | ✅ **Outlook PST/OST** | `*.pst`, `*.ost` | `libpff`/`pypff`. Zero copertura email oggi. Header, allegati, item cancellati. |
| 3 | ✅ **Cloud sync / esfiltrazione** | OneDrive `*.odl` + `SyncEngineDatabase.db`, Dropbox, Google Drive | Vettore di esfiltrazione moderno, oggi invisibile. |
| 4 | ✅ **Chat desktop** | Teams/Slack/Discord LevelDB | **Riusa il decompressore Snappy già scritto per il modulo 39**, che pero' vive dentro l'heredoc Python di `module_ai_chat`: va prima estratto in un helper condiviso (es. `src/lib/13-leveldb.py.sh`). Fare il refactor contestualmente, non prima. |
| 5 | ✅ **LSA Secrets / DCC2** | hive `SECURITY` | Il modulo 20 copre solo SAM: mancano cached domain credentials e password dei service account. |
| 6 | ✅ **WebCacheV01.dat** | ESE Edge/IE | History/cookie/download non coperti dal modulo 17 (solo SQLite). |
| 7 | ✅ **Thumbcache / Thumbs.db** | `thumbcache_*.db` | Prova visiva di file **cancellati**. |
| 8 | ✅ **PowerShell transcripts** | `PowerShell_transcript.*.txt` | Il modulo 1 copre solo PSReadLine; i transcript hanno l'output completo. |
| 9 | ✅ **SetupAPI dev log** | `setupapi.dev.log` | Timestamp di *first install* USB, complementa il modulo 9. |
| 10 | ✅ **BITS jobs** | `qmgr.db` | Download/persistenza LOLBin (T1197). |
| 11 | ✅ **Windows Search index** | `Windows.edb` | Contenuto indicizzato di file poi cancellati. |

---

## Fase 3 — Moduli Linux ✅ (completata in v2.2)

Tutti e cinque implementati come moduli 17-21.


| Modulo | Fonte | Perché |
|---|---|---|
| **PAM** | `/etc/pam.d`, moduli `.so` non pacchettizzati | Backdoor PAM classica, non rilevata dal modulo Persistence. |
| **Rootkit / LKM** | `/lib/modules` vs pacchetti, `modprobe.d`, initramfs | Oggi è coperto solo `ld.so.preload`. |
| **Log web server** | nginx/apache access+error | Webshell e initial access. |
| **Credenziali cloud/dev** | `~/.aws/credentials`, `~/.kube/config`, `~/.docker/config.json`, `.netrc`, `.git-credentials` | Credential access. |
| **SUID/SGID + capabilities** | scan filesystem | Superficie di privilege escalation. |

---

## Fase 4 — Moduli macOS ✅ (completata in v2.2)

Tutti implementati come moduli 14-19. Sui unified log la valutazione di costo
prevista dalla roadmap si e' chiusa cosi': si decomprimono i chunk LZ4 (bv41) e
si estraggono le stringhe leggibili, senza ricostruire i messaggi — la
risoluzione dei riferimenti in .uuidtext e nel dyld_shared_cache resta fuori
portata per un tool bash+python senza dipendenze. Il limite e' dichiarato nel
report invece di essere lasciato intendere.


| Modulo | Fonte | Note |
|---|---|---|
| **Unified logs** | `.tracev3` | Oggi dichiarati fuori scope. Esistono parser Python offline; è il gap più citato nell'analisi macOS. Valutare il costo prima di impegnarsi. |
| **Cookies / Downloads** | `Cookies.binarycookies`, `Downloads.plist` | |
| **XProtect / Gatekeeper** | `xprotect_behavior_service.db` | |
| **Messages** | `chat.db` | |
| **Snapshot APFS** | elenco snapshot | Analogo VSS per macOS. |
| **Inventario app** | `/Applications` + firma/notarizzazione | |

---

## Fase 5 — Moduli cross-OS ✅ (completata in v2.2)

Entrambi implementati e registrati nei tre registri. Sul recupero SQLite la
scelta e' stata di estrarre **stringhe** dallo spazio non allocato, non di
ricostruire i record: servirebbero lo schema e l'interpretazione dei serial
type, e un record ricomposto male in una perizia e' peggio di nessun record.

Insidia trovata durante i test, da non reintrodurre: il pattern di estrazione
ammette i byte di continuazione UTF-8 per non spezzare gli accenti, ma in una
pagina SQLite l'header del record segue il payload, quindi il match finisce
quasi sempre con un byte di continuazione isolato. Con `decode(..., 'strict')`
l'intera stringa veniva scartata in silenzio — e sono proprio quelle
interessanti. Va usato `errors='ignore'` con pulizia dei caratteri di
controllo.


- **ESP / bootkit hunting** — hash e verifica dei bootloader nella partizione
  EFI, scan YARA (ESPecter, BlackLotus, Bootkitty), file non-EFI anomali.
  Persiste sotto ogni OS ed è invisibile a tutti i moduli attuali.
- **SQLite recovery** — recupero di record cancellati da freelist e aree non
  allocate. Da implementare come helper condiviso: `query_sqlite()` già copia
  il `-wal`, ma non recupera i cancellati. Un helper unico porta il beneficio a
  **tutti** i moduli SQLite (browser, TCC, KnowledgeC, quarantine, chat).

---

## Fase 6 — Motore

Ordinati per rapporto valore/costo.

1. **Chain of custody.** `evidence_manifest.json` con hash di ogni file
   toccato, versione del tool, timestamp UTC, comando eseguito, hash dei report
   finali, firma GPG opzionale. Oggi `sha256_file()` esiste ma è usata in pochi
   moduli. Per uso peritale è un requisito, non un extra.
2. ~~**Normalizzazione temporale.**~~ ✅ **Fatto (6.2)** — `src/lib/16-time.sh`.
   Fuso del volume rilevato da `/etc/timezone`, `/etc/localtime` e
   `SYSTEM\Control\TimeZoneInformation`, filtro globale `--since/--until`.

   **Decisione da non ribaltare senza motivo:** la voce chiedeva "ISO8601 UTC
   ovunque", cioè convertire tutte le date a UTC. Non è stato fatto, ed è
   deliberato: gli artefatti di uno stesso volume mescolano UTC (registro, log
   eventi) e ora locale (syslog, shell history), e non esiste un modo affidabile
   per sapere *dalla stringa* in quale dei due si è. Una conversione applicata
   alla cieca sposterebbe una parte degli eventi di ore — un errore inventato,
   molto peggio di una data dichiarata ambigua. Si rileva quindi il fuso e lo si
   dichiara (avvio, manifesto, campo `volume_timezone` nel JSONL), e il
   confronto usa le date così come compaiono. Convertire davvero è possibile
   solo artefatto per artefatto, sapendo cosa produce quel formato: è un lavoro
   per singolo modulo, non per il motore.

   Agganci universali usati: `_rows_to_table` (tabelle), `render_pre_block`
   (blocchi di log) ed `export_report_jsonl`. Filtrare solo il JSONL avrebbe
   prodotto due viste dello stesso modulo in contraddizione.

   Regole del filtro, presidiate dai test: le righe **prive** di data si
   tengono sempre (non sono valutabili), una riga con **più** date si tiene se
   *almeno una* cade nella finestra, i limiti sono inclusivi, e ogni blocco
   filtrato dichiara quante righe ha nascosto. Un limite malformato **ferma
   l'esecuzione**: ignorarlo produrrebbe un report che dichiara una copertura
   che non ha.
3. ~~**Executive summary con scoring.**~~ ✅ **Fatto (6.3)** —
   `src/lib/18-summary.sh`, `executive_summary.html` + `findings.json`.
   Generato a fine `--all` e dalla voce `[S]` del menu.

   **Tre sostrati, in ordine di affidabilità decrescente:** le righe che i
   moduli hanno già marcato `sensitive` (giudizio di chi conosce l'artefatto),
   le corrispondenze `--ioc`, e una tabella di regole esplicite in
   `finding_rules_tsv`. Le regole girano **solo sul dato** (`read_report_data`:
   righe di tabella e righe di log), mai sulla prosa dei cartigli: un motore
   che leggesse l'intera pagina scatterebbe sulle spiegazioni scritte da FIUTO
   stesso. C'è un test dedicato, ed è il primo che si rompe se qualcuno
   allarga l'estrazione.

   **Aggiungendo regole:** preferire pattern che descrivono la *forma* del dato
   (un base64 lungo dopo `-enc`) a parole che possono comparire in un nome di
   file legittimo. In un executive summary una regola rumorosa costa più di una
   regola mancante: fa perdere fiducia in tutte le altre. I test strutturali
   verificano campi, severità, validità delle regex e — soprattutto —
   **esistenza degli slug citati**: uno slug sbagliato non rompe niente, la
   regola semplicemente non scatta mai.

   **Correlazione:** cluster temporali di 30 minuti su tutti gli eventi, con
   scenari in `correlation_scenarios_tsv`. Sono ipotesi, e il report le
   presenta come tali. Ogni riscontro porta la sua tecnica MITRE: il punto 8
   copre gia' il mapping ATT&CK: ogni riscontro porta la sua tecnica.

   **Punteggio:** somma dei pesi (40/15/5/1) limitata a 100, con formula
   stampata nel report. Ordina la coda di lavoro, non misura la
   compromissione — e l'assenza di riscontri è dichiarata come "le poche regole
   applicate non hanno trovato nulla", non come esito negativo dell'analisi.

   **Nota di refactoring:** l'estrazione degli eventi dall'HTML viveva dentro
   `export_report_jsonl`. Serve identica al riepilogo, che deve poter correlare
   anche senza `--jsonl`, quindi è stata spostata in `pylib_timeline`
   (`src/lib/17-pylib-timeline.sh`). L'export ne è ora un consumatore: i 12
   test JSONL esistenti hanno presidiato il cambio.
4. ~~**Detection engine Sigma/YARA.**~~ ✅ **Fatto (6.4 YARA, 6.5 Sigma)** —
   `src/modules/xplat/04-yara.sh`, `src/modules/win/51-sigma.sh` con il
   compilatore in `src/lib/19-pylib-sigma.sh`.

   **Sull'ambito, che è il punto vero.** Un volume da un terabyte non si
   scansiona file per file, quindi il modulo copre un insieme limitato di
   posizioni (quelle scrivibili senza privilegi) e le **elenca nel report** con
   i conteggi, insieme ai file saltati e al perché. Senza quella contabilità un
   "nessun match" verrebbe letto come "il disco è pulito", che è esattamente il
   tipo di conclusione che il tool non deve indurre. Chi tocca questo modulo non
   tolga il cartiglio dell'ambito: c'è un test che lo pretende.

   Nessun ripiego se `yara-python` manca: una scansione YARA senza YARA non è
   una scansione, e dichiararlo è meglio di un report vuoto che sembra pulito.
   I symlink non vengono seguiti — su un volume montato porterebbero fuori
   dall'evidenza fino al file system della workstation.

   **Sigma — il sottoinsieme è dichiarato.** Il formato è vasto e un supporto
   parziale spacciato per completo sarebbe peggio di nessun supporto: una regola
   mai valutata comparirebbe come una regola non scattata, cioè un falso
   negativo invisibile. Implementato: selezioni (mappe, liste di valori, liste
   di mappe), modificatori `contains`/`startswith`/`endswith`/`re`/`all`/`cased`,
   condizioni `sel`, `a and/or b`, `a and not b`, `not a`, `1 of x*`,
   `all of x*`, `1 of them`, `all of them`, e `null` come campo assente.
   Scartato e **contato ed elencato nel report, raggruppato per motivo**: le
   condizioni con parentesi o aggregazioni, i modificatori
   base64/utf16/wide/cidr/gt/lt, e le logsource non mappabili.

   Una logsource non mappata **non** viene fatta girare a tappeto su tutti i
   canali come ripiego: la regola verrebbe valutata su campi che quel canale non
   ha, e il "non scattata" sarebbe privo di significato.

   Il compilatore sta in `src/lib/19-pylib-sigma.sh` e non dentro il modulo,
   perché è la parte rischiosa: lì è esercitabile dai test con eventi sintetici,
   senza bisogno di un `.evtx`. I 22 test coprono in parti uguali "la regola
   scatta quando deve" e "la regola non valutabile finisce fra le scartate".

   Aggiungendo supporto per altre costruzioni: prima il test che dimostra che
   oggi finisce fra le scartate, poi l'implementazione, poi il test che dimostra
   che scatta. L'ordine conta — è l'unico modo per sapere che il caso era
   davvero coperto dal ramo "scartate" e non silenziosamente ignorato.

   Sigma è registrato **solo per Windows**: le logsource mappate puntano a
   canali EVTX, offrirlo su Linux o macOS prometterebbe una copertura
   inesistente. C'è un test che lo presidia.
5. ~~**Parallelizzazione di `--all`.**~~ ✅ **Fatto (6.6)** — `--jobs N`,
   `run_batch_pool` in `src/lib/11-runner.sh`. Su un volume di prova: 15,5s
   sequenziale, 7,3s con `--jobs 4`, esito identico.

   **Opt-in, non default.** Su disco meccanico o volume di rete N lettori vanno
   più piano di uno: il collo di bottiglia è l'I/O, non la CPU, e qual è lo sa
   solo chi analizza.

   **Il criterio è che l'esito non dipenda da `--jobs`**: stessi report, stesso
   ordine nel riepilogo. Un risultato che cambia con il parallelismo non è
   affidabile, e in ambito peritale è peggio di un tool lento.

   Due trappole trovate implementando, entrambe con test di regressione:
   - i moduli `defer` vanno lanciati **dopo** aver raccolto gli esiti del pool,
     non solo dopo che il pool si è svuotato. La Master Timeline legge
     `GENERATED_REPORTS`: nella prima versione girava con l'array ancora vuoto e
     produceva una timeline vuota **senza segnalare nulla**;
   - `recover_hive` ora prende un lock (`mkdir`, atomico). Due moduli che
     chiedono lo stesso hive insieme facevano leggere al secondo la copia
     ricostruita mentre il primo la stava scrivendo. Un hive troncato non dà
     errore: dà risultati parziali.

   ESC per saltare un modulo non è disponibile in parallelo — richiede il
   controllo esclusivo del terminale — e viene dichiarato all'avvio.
6. ~~**IoC engine.**~~ ✅ **Fatto (6.7)** — `src/lib/04-ioc.sh` riscritta.

   Il match a sottostringa sbagliava in due direzioni: falsi positivi
   (`10.0.0.5` dentro `110.0.0.55`) e falsi negativi (gli IoC arrivano defanged
   e cercati alla lettera non trovano nulla — che si manifesta come "disco
   pulito"). Ora gli indicatori sono tipizzati al caricamento, l'input defanged
   viene normalizzato e ogni tipo ha i suoi confini.

   `check_ioc` resta **pura bash, senza fork**: viene chiamata riga per riga su
   report da decine di migliaia di righe, e un processo per riga renderebbe i
   moduli inutilizzabili. Il costo è una singola ERE compilata al caricamento —
   la stessa che viene passata all'executive summary, così i confini non
   possono divergere fra due implementazioni.

   `IOC_LIST` continua a contenere i valori normalizzati: i moduli che fanno il
   proprio match in Python non sono stati toccati e guadagnano comunque il
   defanging.

   Due decisioni da non ribaltare distrattamente:
   - un JSON valido ma di schema sconosciuto viene **rifiutato**, non letto come
     lista piatta: caricare le graffe come indicatore letterale le farebbe
     corrispondere ovunque;
   - `payload.dll` è sintatticamente un dominio. Un elenco di estensioni decide
     i casi ambigui e il prefisso `file:` li forza; il conteggio per tipo
     stampato al caricamento esiste per far accorgere di una classificazione
     sbagliata prima che si traduca in match mancati.

   Resta da fare il **defanging dell'output**: `defang_value` esiste ed è
   testata, ma non è ancora applicata ai valori mostrati nei report. Va fatta
   insieme a `--redact` (punto 7), che tocca gli stessi punti di rendering.
7. ~~**`--redact`.**~~ ✅ **Fatto (6.8)** — `src/lib/20-redact.sh`.
   `--redact` produce un `report.redacted.html` accanto a ogni report;
   `--defang` rende inerti anche URL e IP nella sola copia.

   Due scelte da non ribaltare:
   - **l'originale non si tocca.** Oscurare sul posto distruggerebbe evidenza
     per una necessità di comunicazione: un baratto che non spetta al tool.
   - **si oscura per contesto, non per forma.** Un SHA-256 e un hash NTLM sono
     entrambe stringhe esadecimali; il primo è integrità di un reperto, il
     secondo una credenziale. Le regole guardano l'etichetta accanto al valore,
     e c'è un test che pretende che gli SHA-256 sopravvivano.

   Il defanging dell'output, rimasto in sospeso dalla 6.7, è chiuso qui:
   `defang_value` per i valori singoli, `--defang` per la copia condivisibile.

   **Difetto preesistente trovato agganciando la modalità.** L'esecuzione batch
   ridefinisce `register_report` nel processo figlio per far risalire i percorsi
   (gli array bash non risalgono), e quella ridefinizione saltava gli effetti
   collaterali: **`--all --jsonl` non produceva alcun JSONL**, senza un errore
   né un avviso, da quando la funzione esiste. Gli effetti stanno ora in
   `_report_side_effects`, richiamata da tutte e tre le versioni di
   `register_report`, con due test di regressione. Lezione per il futuro: ogni
   volta che si aggiunge un effetto a `register_report`, va aggiunto lì.

   Aggiunto anche un `flock` sull'append alla timeline unica: con `--jobs` le
   righe JSON superano PIPE_BUF e si intreccerebbero, producendo JSON non
   parsabile proprio nel file destinato a un altro strumento.
8. ~~**Immagini senza mount manuale.**~~ ✅ **Fatto (6.9)** —
   `src/lib/21-image.sh`, flag `--image`, `--partition`, `--unlock`,
   `--list-partitions`.

   **L'invariante è la sola lettura per costruzione**: ogni anello riceve la
   propria opzione (`ewfmount` lo è per natura, `losetup --read-only`,
   `cryptsetup --readonly`, `dislocker -r`, `mount ro,noexec,nodev,nosuid`) e
   c'è un test che rilegge il sorgente per verificarlo. Chi tocca questo file
   non aggiunga un `mount` senza `ro`: è l'unica cosa che rende difendibile
   l'intera funzione.

   Decisioni:
   - **la cifratura si rileva prima dei privilegi.** "È BitLocker e manca la
     chiave" si sa da utente normale, ed evita di rilanciare con sudo per
     scoprirlo solo allora;
   - **su più partizioni senza `--partition` si rifiuta di scegliere.**
     Prenderne una a caso significa analizzare quella sbagliata e non
     accorgersene mai;
   - **FileVault 2 riconosciuto e rifiutato**, VMDK/VHDX/QCOW2 rifiutati con il
     comando di conversione. Leggerli come raw darebbe dati sbagliati senza
     errori;
   - `image_cleanup` **non propaga mai un errore**: gira dalla trap EXIT, spesso
     quando qualcosa è già andato storto, e un suo fallimento nasconderebbe la
     causa vera. (È anche il bug che i test hanno trovato: senza `|| true` la
     funzione usciva non-zero e uccideva il test invece di farlo fallire.)

   **Verificato per davvero**, non solo a fixture: E01 creata con `ewfacquire`,
   aperta con `ewfmount` ed elencata senza privilegi; LUKS vero creato con
   `cryptsetup luksFormat` e riconosciuto; catena completa raw → loop `-r` →
   `mount ro` eseguita dentro il container `--privileged`, con verifica che
   all'uscita non restino loop device orfani.

   Resta fuori: le immagini segmentate multi-file oltre il primo `.E01` (le
   gestisce `ewfmount` da sé, ma non è stato provato), e i container VeraCrypt.

---

## Fase 7 — Distribuzione ✅ **fatta**

`Dockerfile`, `requirements.txt`, `docker-entrypoint.sh`, job `docker` in CI.
Immagine di 328 MB su `python:3.12-slim-bookworm` con tutte e nove le librerie
che FIUTO cerca a runtime, più i binari (`file`, `strings`, `sqlite3`,
`ntfscat`, `ewfmount`, `flock`, `fls`, `icat`).

**Perché era la voce giusta da fare.** L'attrito vero non è il tool: sono
regipy, python-evtx, libesedb, libpff, yara-python, PyYAML. Finché mancano, i
moduli che le usano si degradano — dichiarandolo, ma degradandosi — e "nessun
match" diventa ambiguo. L'immagine è l'unico modo per garantire che quella
frase significhi davvero quello.

### Decisioni

- **Versioni fissate** in `requirements.txt`. Un parser che cambia fra due
  build produce report diversi sullo stesso disco, e in perizia la differenza
  va spiegata. Il pin si aggiorna deliberatamente, non per inerzia. C'è un test
  che rifiuta una riga senza `==`.
- **`build.sh --check` gira dentro il Dockerfile**, non solo in CI: un'immagine
  costruita a mano da un albero sporco conterrebbe altrimenti un `fiuto.sh`
  diverso dai suoi sorgenti.
- **`--deps`** stampa l'inventario dei parser presenti. Serve a poter
  dichiarare, in una relazione, con cosa è stata fatta l'analisi — e la CI lo
  usa per fallire se una libreria manca: un'immagine che si costruisce ma senza
  libesedb degraderebbe in silenzio proprio dove promette il contrario.
- **L'entrypoint verifica le due condizioni che rendono valida l'analisi** e
  che è facilissimo sbagliare scrivendo un `docker run` a mano: evidenza in
  `:ro` e `--user`. Nessuna delle due è bloccante — chi sa cosa fa deve poter
  procedere — ma nessuna resta implicita.
- **Pubblicazione solo dai tag `v*`.** Un `latest` che cambia a ogni push
  renderebbe irriproducibile un'analisi fatta la settimana prima.

### Verifica fatta

Build riuscito, `--deps` con nove librerie su nove, esecuzione `--all` su un
volume Linux reale montato `:ro`: report prodotti, di proprietà dell'utente
invocante grazie a `--user`, volume di evidenza non modificato. Con il mount
scrivibile e senza `--user` l'entrypoint emette entrambi gli avvisi.

### Resta da fare

- `docker buildx` multi-arch (arm64) — utile su Apple Silicon, non provato.
- Nessuno dei binari libyal per VSS (`vshadowinfo`) è pacchettizzato in Debian:
  il modulo 43 continua a dichiarare il comando invece di eseguirlo, immagine o
  no. Va compilato da sorgente se si vuole chiudere quel caso.

---

## Build macOS

`./buildMac.sh` produce `fiuto-macos.sh`, che gira SU un Mac come workstation
di analisi. **Non è un fork**: concatena gli stessi sorgenti della build Linux
più `src/compat/macos.sh`, lo strato che traduce gli strumenti GNU verso
l'userland BSD ridefinendoli come funzioni bash. Duplicare l'albero avrebbe
significato correggere ogni bug due volte.

Il compat sta in **fondo** all'ordine di build, non in testa: in bash i nomi di
funzione si risolvono alla chiamata, quindi metterlo alla fine basta perché le
ridefinizioni valgano ovunque — e metterlo in testa non basterebbe, perché
`_find_user_cwd`, `image_open` e `_find_windows_mounts` verrebbero
sovrascritte dai file di libreria che le definiscono.

La trappola da conoscere prima di toccare quel file: nella traduzione di `stat`
l'ordine conta. GNU `%s` (dimensione) diventa BSD `%z`, ma GNU `%z` (ctime)
diventa BSD `%Sc`. Tradurre `%s` per primo lo riconvertirebbe, e la chiamata
restituirebbe **una data dove il codice si aspetta dei byte, senza errore**.
C'è un test che inchioda l'ordine, con un userland BSD simulato.

**Non verificata su un Mac reale.** Da qui si è verificato che le due build
restano allineate, che le traduzioni producono i comandi BSD attesi e che la
suite passa. Il primo avvio sulla macchina di destinazione resta il collaudo.
Cose da controllare lì per prime: `hdiutil attach` sulle immagini reali, il
percorso di `brew` su Intel (`/usr/local/bin/bash` invece di
`/opt/homebrew/bin/bash`), e i moduli che invocano `sqlite3` e `plutil`.

Non supportato su macOS, per assenza degli strumenti: LUKS e BitLocker
(`cryptsetup`, `dislocker`). Il volume viene riconosciuto e dichiarato, non
montato male.

---

## Debito noto, da chiudere quando si passa di lì

- **Numerazione Linux:** l'aggiunta di auditd e Container ha spostato la Master
  Timeline da 14 a 16. È stato introdotto il flag `defer` proprio per non
  ripetere l'errore: i moduli nuovi si accodano e la Master Timeline mantiene
  il suo numero girando comunque per ultima con `--all`.
- **`--user` e `--silent` sono documentati nel README storico ma non esistono**
  nel parsing argomenti. In v2.1 sono stati tolti dal README. Vanno
  implementati (`--user` è utile davvero: filtra i moduli per-utente) oppure
  lasciati fuori consapevolmente.
- `datetime.utcfromtimestamp` è deprecata da Python 3.12: sostituire con
  `datetime.fromtimestamp(..., tz=timezone.utc)` quando si tocca un parser.
- Il menu Windows stampa le voci a mano: si disallinea dal dispatcher a ogni
  aggiunta. Si chiude con la Fase 1 punto 3.
