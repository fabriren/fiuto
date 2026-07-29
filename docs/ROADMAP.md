# FIUTO — Piano di implementazione

Documento di lavoro per portare FIUTO da 2.1 a 3.0. È pensato per essere
ripreso a distanza di tempo, anche da un'altra sessione o da un'altra persona:
ogni fase dichiara **cosa fare**, **dove**, **come verificarlo** e **perché**.

Stato aggiornato al: **2026-07-29** (versione 2.1).

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

### Struttura di destinazione

```
fiuto.sh                  # entrypoint: sourcing, parsing argomenti, menu
lib/
  core.sh                 # globali, colori, cleanup, log, i18n (L, t)
  fs.sh                   # ci_find_dir, ci_find_file, detect_os_type, homes
  report.sh               # html_header/footer, finish_report, register_report
  export.sh               # export_report_jsonl e futuri formati
  registry.sh             # get_hive, recover_hive, regipy_query
  ioc.sh                  # load_ioc_file, check_ioc
modules/
  win/*.sh                # un file per modulo (o per famiglia)
  linux/*.sh
  macos/*.sh
  xplat/*.sh              # master timeline, ESP/bootkit, SQLite recovery
tests/
docs/
```

### Vincoli non negoziabili

- **Distribuzione a file singolo.** Oggi FIUTO si usa copiando un solo file su
  una workstation forense; è una qualità reale, non un dettaglio. Serve quindi
  `build.sh` che concateni `lib/` + `modules/` in un `fiuto.sh` distribuibile,
  **e** la CI deve verificare che il file committato coincida con il build
  (`./build.sh --check`), altrimenti le due forme divergono silenziosamente.
- In alternativa: `fiuto.sh` fa `source` dei moduli quando li trova accanto a
  sé e usa la versione concatenata quando è solo. Più semplice da mantenere,
  ma va deciso **prima** di iniziare, non a metà.

### Passi

1. Estrarre prima `lib/` (basso rischio: nessuna logica forense).
2. Poi i moduli, **un OS alla volta**, eseguendo la suite a ogni passo.
3. Estendere il registro data-driven anche a Windows: oggi Linux e macOS usano
   `MODULES_LINUX`/`MODULES_MACOS`, mentre Windows ha `run_module_by_number`
   con un `case` di 39 rami e un menu stampato a mano. Unificare elimina la
   principale fonte di disallineamento fra menu, dispatcher e README.
4. Aggiungere `MODULES_WIN` e far derivare da lì menu, dispatcher e `--all`.

### Verifica

- Il test strutturale esistente ("ogni numero del menu Windows invoca una
  funzione esistente") va esteso al nuovo registro.
- Nuovo test: la numerazione dei moduli nei registri coincide con le tabelle
  del README (evita il disallineamento che si è già verificato).

---

## Fase 2 — Moduli Windows

Un file per modulo in `modules/win/`. Priorità decrescente.

| # | Modulo | Fonte | Note di implementazione |
|---|---|---|---|
| 1 | **Volume Shadow Copies** | snapshot VSS | Enumerazione + rilancio dei moduli sullo snapshot: abilita l'analisi differenziale storica. Oggi VSS è solo *citato* nei suggerimenti. |
| 2 | **Outlook PST/OST** | `*.pst`, `*.ost` | `libpff`/`pypff`. Zero copertura email oggi. Header, allegati, item cancellati. |
| 3 | **Cloud sync / esfiltrazione** | OneDrive `*.odl` + `SyncEngineDatabase.db`, Dropbox, Google Drive | Vettore di esfiltrazione moderno, oggi invisibile. |
| 4 | **Chat desktop** | Teams/Slack/Discord LevelDB | **Riusa il decompressore Snappy già scritto per il modulo 39**: costo marginale basso, resa alta. |
| 5 | **LSA Secrets / DCC2** | hive `SECURITY` | Il modulo 20 copre solo SAM: mancano cached domain credentials e password dei service account. |
| 6 | **WebCacheV01.dat** | ESE Edge/IE | History/cookie/download non coperti dal modulo 17 (solo SQLite). |
| 7 | **Thumbcache / Thumbs.db** | `thumbcache_*.db` | Prova visiva di file **cancellati**. |
| 8 | **PowerShell transcripts** | `PowerShell_transcript.*.txt` | Il modulo 1 copre solo PSReadLine; i transcript hanno l'output completo. |
| 9 | **SetupAPI dev log** | `setupapi.dev.log` | Timestamp di *first install* USB, complementa il modulo 9. |
| 10 | **BITS jobs** | `qmgr.db` | Download/persistenza LOLBin (T1197). |
| 11 | **Windows Search index** | `Windows.edb` | Contenuto indicizzato di file poi cancellati. |

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
