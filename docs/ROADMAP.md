# FIUTO — Piano di implementazione

Documento di lavoro per portare FIUTO da 2.1 a 3.0. È pensato per essere
ripreso a distanza di tempo, anche da un'altra sessione o da un'altra persona:
ogni fase dichiara **cosa fare**, **dove**, **come verificarlo** e **perché**.

Stato aggiornato al: **2026-07-30** (versione 2.2, Fasi 1 e 2 completate).

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
| Suite di test — 53 test bats | ✅ |
| Replay transaction log registro (`.LOG1`/`.LOG2`) | ✅ |
| Export JSONL / schema Timesketch (`--jsonl`) | ✅ |
| macOS: FSEvents, Spotlight | ✅ |
| Linux: auditd, Container (Docker/Podman) | ✅ |
| Fase 1 — split `src/` + `build.sh` (file singolo generato) | ✅ |
| Fase 1 — registro moduli data-driven anche per Windows | ✅ |
| Fase 2 — 11 moduli Windows nuovi (40-50) | ✅ |
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

## Fase 3 — Moduli Linux

| Modulo | Fonte | Perché |
|---|---|---|
| **PAM** | `/etc/pam.d`, moduli `.so` non pacchettizzati | Backdoor PAM classica, non rilevata dal modulo Persistence. |
| **Rootkit / LKM** | `/lib/modules` vs pacchetti, `modprobe.d`, initramfs | Oggi è coperto solo `ld.so.preload`. |
| **Log web server** | nginx/apache access+error | Webshell e initial access. |
| **Credenziali cloud/dev** | `~/.aws/credentials`, `~/.kube/config`, `~/.docker/config.json`, `.netrc`, `.git-credentials` | Credential access. |
| **SUID/SGID + capabilities** | scan filesystem | Superficie di privilege escalation. |

---

## Fase 4 — Moduli macOS

| Modulo | Fonte | Note |
|---|---|---|
| **Unified logs** | `.tracev3` | Oggi dichiarati fuori scope. Esistono parser Python offline; è il gap più citato nell'analisi macOS. Valutare il costo prima di impegnarsi. |
| **Cookies / Downloads** | `Cookies.binarycookies`, `Downloads.plist` | |
| **XProtect / Gatekeeper** | `xprotect_behavior_service.db` | |
| **Messages** | `chat.db` | |
| **Snapshot APFS** | elenco snapshot | Analogo VSS per macOS. |
| **Inventario app** | `/Applications` + firma/notarizzazione | |

---

## Fase 5 — Moduli cross-OS

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
2. **Normalizzazione temporale.** ISO8601 UTC ovunque, timezone del volume
   (`/etc/timezone`, `TimeZoneInformation`), filtro globale `--since/--until`.
   Riduce molto il rumore sui dischi grandi.
3. **Executive summary con scoring.** Findings ordinati per severità e
   correlazione cross-modulo (es. USB inserito + LNK + picco USN nella stessa
   finestra ⇒ candidato esfiltrazione). È ciò che distingue un tool DFIR da un
   dumper.
4. **Detection engine Sigma/YARA.** `--sigma <dir>` sugli EVTX (approccio
   Chainsaw/Hayabusa), `--yara <rules>` su file estratti, quarantena, ESP.
   Trasforma FIUTO da collector a triage.
5. **Parallelizzazione di `--all`.** Oggi strettamente sequenziale. Pool di job
   sui moduli indipendenti; unico vincolo la master timeline, che è già un
   aggregatore finale.
6. **IoC engine.** Oggi è match di sottostringa case-insensitive. Tipizzare
   (IP/dominio/hash/regex), import STIX/MISP, defanging in output.
7. **`--redact`.** I report contengono hash NTLM, PSK Wi-Fi, token: serve una
   modalità per condividerli senza i segreti.
8. **Mapping MITRE ATT&CK** per finding + export di un layer Navigator.
9. **Immagini senza mount manuale.** `ewfmount` per E01, `losetup` per raw/dd,
   volumi cifrati (BitLocker/`dislocker`, LUKS, FileVault).

---

## Fase 7 — Distribuzione

- **Immagine Docker** con tutte le dipendenze (regipy, impacket, libesedb,
  mft, snappy, libpff). L'attrito maggiore per chi usa FIUTO è l'installazione
  delle dipendenze: un'immagine ufficiale lo azzera.
- Pubblicazione su GHCR dalla CI, tag allineato alla versione.

---

## Debito noto, da chiudere quando si passa di lì

- **`--user` e `--silent` sono documentati nel README storico ma non esistono**
  nel parsing argomenti. In v2.1 sono stati tolti dal README. Vanno
  implementati (`--user` è utile davvero: filtra i moduli per-utente) oppure
  lasciati fuori consapevolmente.
- `datetime.utcfromtimestamp` è deprecata da Python 3.12: sostituire con
  `datetime.fromtimestamp(..., tz=timezone.utc)` quando si tocca un parser.
- Il menu Windows stampa le voci a mano: si disallinea dal dispatcher a ogni
  aggiunta. Si chiude con la Fase 1 punto 3.
