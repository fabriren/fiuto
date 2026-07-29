
# ================================================================
#  MODULO 38 — PAD Offline (Active Directory Analysis)
# ================================================================
module_pad_offline() {
    section_header "PAD Offline — Active Directory Analysis (NTDS.dit)" "$CYAN"
    check_win_root || return 1

    # ── Trova NTDS.dit ────────────────────────────────────────────
    local NTDS_DIR="" NTDS_PATH="" SYSTEM_HIVE=""
    NTDS_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/NTDS")
    if [[ -n "$NTDS_DIR" ]]; then
        NTDS_PATH=$(find "$NTDS_DIR" -maxdepth 1 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    fi
    if [[ -z "$NTDS_PATH" ]]; then
        NTDS_PATH=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f \
            -not -path "*/WinSxS/*" -not -path "*/WinSXS/*" \
            2>/dev/null | head -1)
    fi
    if [[ -z "$NTDS_PATH" || ! -f "$NTDS_PATH" ]]; then
        warn "$(L "ntds.dit non trovato sotto" "ntds.dit not found under") $WIN_ROOT"
        info "$(L "Percorso atteso: Windows/NTDS/ntds.dit" "Expected path: Windows/NTDS/ntds.dit")"
        info "$(L "Questo modulo richiede un Domain Controller (non una workstation)" "This module requires a Domain Controller (not a workstation)")"
        return 0
    fi
    info "NTDS.dit: $NTDS_PATH"

    # ── Trova SYSTEM hive (opzionale) ────────────────────────────
    local CFG_DIR
    CFG_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
    if [[ -n "$CFG_DIR" ]]; then
        SYSTEM_HIVE=$(find "$CFG_DIR" -maxdepth 1 -iname "SYSTEM" -type f 2>/dev/null | head -1)
    fi
    if [[ -n "$SYSTEM_HIVE" ]]; then
        info "SYSTEM hive: $SYSTEM_HIVE"
    else
        warn "$(L "SYSTEM hive non trovato — attributi cifrati (hash) non disponibili; tutti gli altri OK" "SYSTEM hive not found — encrypted attributes (hashes) unavailable; all others OK")"
        SYSTEM_HIVE=""
    fi

    # ── Avviso dirty database ─────────────────────────────────────
    warn "$(L "Se ntds.dit proviene da un DC attivo, alcune transazioni potrebbero essere nei log (.jrs)." "If ntds.dit comes from a live DC, some transactions may be in logs (.jrs).")"
    info "Per massima integrità usa una VSS snapshot o esegui 'esentutl /r edb /l <logdir>' prima."

    # ── Verifica dipendenza impacket ──────────────────────────────
    if ! "$PY3" -c "from impacket.ese import ESENT_DB" 2>/dev/null; then
        warn "impacket non trovato. $(L "Installalo con:" "Install it with:") ${PY3} -m pip install impacket"
        return 1
    fi

    info "$(L "Avvio analisi NTDS.dit (può richiedere qualche minuto su database grandi)..." "Starting NTDS.dit analysis (may take a few minutes on large databases)...")"
    separator

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "pad_offline")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local NTDS_HASH; NTDS_HASH=$(sha256_file "$NTDS_PATH")

    # Copia ntds.dit in /tmp (preferita: velocizza accesso ESE e abilita patch binaria).
    # Se fallisce (file bloccato da Windows su DC live, o rete lenta) si usa il path
    # originale — il recovery in-memory del dirty state non richiede write access al file.
    local TMP_PAD_DIR; TMP_PAD_DIR=$(mktemp -d /tmp/dfir_pad_XXXXXX)
    local TMP_NTDS; TMP_NTDS="${TMP_PAD_DIR}/ntds.dit"
    local NTDS_TO_USE="$NTDS_PATH"
    info "$(L "Tentativo copia ntds.dit in /tmp (timeout 180s)..." "Attempting to copy ntds.dit to /tmp (timeout 180s)...")"
    if portable_timeout 180 cp "$NTDS_PATH" "$TMP_NTDS" 2>/dev/null; then
        chmod 600 "$TMP_NTDS" 2>/dev/null || true
        NTDS_TO_USE="$TMP_NTDS"
        info "$(L "Copia completata — uso copia locale." "Copy completed — using local copy.")"
    else
        warn "$(L "Copia non riuscita (file bloccato da Windows o rete lenta) — uso path originale." "Copy failed (file locked by Windows or slow network) — using original path.")"
        info "$(L "Il dirty state verrà corretto in memoria (patch in-memory, senza write access)." "Dirty state will be corrected in memory (in-memory patch, no write access).")"
    fi

    TMP_NTDS_SIZE=$(wc -c < "$TMP_NTDS" 2>/dev/null | tr -d ' ' || echo "0")
    ORIG_NTDS_SIZE=$(wc -c < "$NTDS_TO_USE" 2>/dev/null | tr -d ' ' || echo "0")
    info "$(L "Dimensione NTDS.dit prima del parsing:" "NTDS.dit size before parsing:") ${TMP_NTDS_SIZE} bytes (originale: ${ORIG_NTDS_SIZE})"
    
    # ── Script Python: analisi completa + generazione HTML ───────
    local PYERR_FILE; PYERR_FILE="${TMP_PAD_DIR}/pyerr.log"
    "$PY3" - "$NTDS_TO_USE" "${SYSTEM_HIVE}" "$REPORT_HTML" "$SCAN" "$WIN_ROOT" "$NTDS_HASH" "$(dirname "$NTDS_PATH")" "$TMP_PAD_DIR" << 'PYEOF' 2>"$PYERR_FILE"
import sys, os, struct, datetime, traceback, html as html_mod

# Bilingual helper: reads LANG env var set by fiuto.sh
def L(it, en): return it if os.environ.get('LANG', 'en') == 'it' else en

ntds_path   = sys.argv[1]
sys_hive    = sys.argv[2] if len(sys.argv) > 2 else ''
report_path = sys.argv[3]
scan_date   = sys.argv[4]
source_root = sys.argv[5]
ntds_hash   = sys.argv[6]
ntds_log_dir= sys.argv[7] if len(sys.argv) > 7 else ''
tmp_work_dir= sys.argv[8] if len(sys.argv) > 8 else os.path.dirname(ntds_path)

# ── Copia log ESE in /tmp per permettere soft-recovery automatico ─
# I file .jrs devono essere nella stessa dir del ntds.dit aperto da impacket
# PRIMA che impacket lo apra, e il dirty state NON deve essere patchato
# altrimenti impacket non tenta il recovery.
_has_ese_logs = False
import shutil, glob as _glob
_log_search_dirs = []
if ntds_log_dir and os.path.isdir(ntds_log_dir):
    _log_search_dirs.append(ntds_log_dir)
# Cerca anche nella dir del file originale se diversa
_orig_ntds_dir = os.path.dirname(os.path.realpath(ntds_path))
if _orig_ntds_dir not in _log_search_dirs:
    _log_search_dirs.append(_orig_ntds_dir)

for _log_src_dir in _log_search_dirs:
    _log_files = (
        _glob.glob(os.path.join(_log_src_dir, 'edb*.log')) +
        _glob.glob(os.path.join(_log_src_dir, 'edb*.jrs')) +
        _glob.glob(os.path.join(_log_src_dir, 'edb.chk'))
    )
    for _lf in _log_files:
        try:
            _dst = os.path.join(tmp_work_dir, os.path.basename(_lf))
            if not os.path.exists(_dst):
                shutil.copy2(_lf, _dst)
                sys.stderr.write(f"INFO: copiato log ESE: {os.path.basename(_lf)}\n")
            _has_ese_logs = True
        except Exception as _le:
            sys.stderr.write(f"INFO: log copy skip {_lf}: {_le}\n")
    if _has_ese_logs:
        break

if _has_ese_logs:
    sys.stderr.write("INFO: log ESE trovati e copiati in tmp — impacket tenterà soft-recovery\n")
else:
    sys.stderr.write("INFO: no ESE log found — only dirty state patch will be applied\n")

try:
    from impacket.ese import ESENT_DB as _ESENT_DB_CLS
    ESENT_DB = _ESENT_DB_CLS   # alias usato in _open_esent_db
    _HAVE_IMPACKET = True
except ImportError:
    _HAVE_IMPACKET = False
    ESENT_DB = None

try:
    import pyesedb as _pyesedb
    _HAVE_PYESEDB = True
except ImportError:
    _HAVE_PYESEDB = False

if not _HAVE_IMPACKET and not _HAVE_PYESEDB:
    sys.stderr.write("FATAL: no ESE library available. Install impacket or libesedb-python.\n")
    sys.exit(1)

# ── Wrapper pyesedb compatibile con l'API impacket ───────────────
# Produce righe come dict {col_name: value}, compatibile col resto del codice.
class _PyesedbCursor:
    def __init__(self, table):
        self._table = table
        self._ncols = table.get_number_of_columns()
        self._nrec  = table.get_number_of_records()
        self._idx   = 0
        # Mappa indice colonna → nome
        self._colnames = [table.get_column(i).name for i in range(self._ncols)]
        # Mappa indice colonna → tipo (0=null,1=bool,2=uint8,3=int16,4=int32,5=currency,
        #  6=float,7=double,8=datetime,9=guid,10=uint16,11=int32,12=int64,13=float,
        #  14=double,15=datetime,16=binary,17=text,18=blob,19=text)
        self._coltypes = [table.get_column(i).type for i in range(self._ncols)]

    def _decode_val(self, record, idx):
        ct = self._coltypes[idx]
        try:
            raw = record.get_value_data(idx)
            if raw is None: return None
            # Interi
            if ct in (2, 3, 4, 10, 11):
                return record.get_value_data_as_integer(idx)
            if ct == 12:   # int64
                return record.get_value_data_as_integer(idx)
            if ct == 5:    # currency (int64)
                return record.get_value_data_as_integer(idx)
            if ct in (17, 19):  # testo Unicode (UTF-16-LE)
                return raw.decode('utf-16-le', errors='replace').rstrip('\x00')
            # Binario / GUID / tutto il resto → restituisce bytes
            return bytes(raw)
        except Exception:
            return None

    def next_row(self):
        if self._idx >= self._nrec:
            return None
        rec = self._table.get_record(self._idx)
        self._idx += 1
        row = {}
        for i, name in enumerate(self._colnames):
            row[name] = self._decode_val(rec, i)
        return row

class _PyesedbDB:
    def __init__(self, path):
        self._db = _pyesedb.file()
        self._db.open(path)
    def openTable(self, name):
        t = self._db.get_table_by_name(name)
        if t is None:
            raise KeyError(f"Tabella '{name}' non trovata")
        return _PyesedbCursor(t)
    def moveHead(self, cur):
        cur._idx = 0
    def getNextRow(self, cur):
        return cur.next_row()
    def closeTable(self, cur):
        pass

PRIV_GROUPS = {
    "Account Operators","Administrators","Backup Operators","Cert Publishers",
    "Cloneable Domain Controllers","DnsAdmins","DnsUpdateProxy","Domain Admins",
    "Domain Controllers","Debugger Users","Distributed COM Users","Enterprise Admins",
    "Enterprise Key Admins","Enterprise Read-only Domain Controllers",
    "Group Policy Creator Owners","Hyper-V Administrators",
    "Incoming Forest Trust Builders","Key Admins","Performance Log Users",
    "Print Operators","Remote Desktop Users","Remote Management Users",
    "Replicator","Schema Admins","Server Operators",
    "Storage Replica Administrators","System Managed Accounts Group"
}

UAC_DISABLED    = 0x00000002
UAC_WORKSTATION = 0x00001000
UAC_DC          = 0x00002000
UAC_NO_EXPIRE   = 0x00010000
UAC_SMARTCARD   = 0x00040000

SAMTYPE_USER     = 805306368
SAMTYPE_COMPUTER = 805306369
SAMTYPE_GROUP    = 268435456

FL_MAP = {0:'Windows 2000',1:'Windows Server 2003 Interim',2:'Windows Server 2003',
          3:'Windows Server 2008',4:'Windows Server 2008 R2',5:'Windows Server 2012',
          6:'Windows Server 2012 R2',7:'Windows Server 2016',10:'Windows Server 2025'}

WELLKNOWN_SIDS = {
    'S-1-1-0':'Everyone','S-1-5-11':'Authenticated Users','S-1-5-18':'SYSTEM',
    'S-1-3-0':'Creator Owner','S-1-5-9':'Enterprise Domain Controllers',
    'S-1-5-10':'Self','S-1-5-7':'Anonymous','S-1-5-15':'This Organization',
    'S-1-5-32-544':'BUILTIN\\Administrators','S-1-5-32-545':'BUILTIN\\Users',
    'S-1-5-32-546':'BUILTIN\\Guests','S-1-5-32-547':'BUILTIN\\Power Users',
    'S-1-5-32-548':'BUILTIN\\Account Operators',
    'S-1-5-32-549':'BUILTIN\\Server Operators',
    'S-1-5-32-550':'BUILTIN\\Print Operators',
    'S-1-5-32-551':'BUILTIN\\Backup Operators',
    'S-1-5-32-552':'BUILTIN\\Replicators',
}

# SID skip per ACL (non interessanti per IR)
ACL_SKIP_SIDS = {'S-1-5-18','S-1-3-0','S-1-5-10','S-1-5-9'}

DANGEROUS_RIGHTS = [
    (0x10000000,'GenericAll'),(0xF01FF,'FullControl'),
    (0x00040000,'WriteDACL'),(0x00080000,'WriteOwner'),
    (0x00000100,'ExtendedRight (DCSync/etc)'),
    (0x00000020,'WriteProperty'),(0x00000001,'CreateChild'),
    (0x00000002,'DeleteChild'),
]

# ── Helper functions ─────────────────────────────────────────────
def find_col(candidates, avail):
    for c in candidates:
        if c in avail:
            return c
    return None

def filetime_to_dt(ft):
    if not ft or ft <= 0 or ft >= 0x7FFFFFFFFFFFFFFF:
        return None
    try:
        return datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=ft//10)
    except:
        return None

def ft_str(ft):
    dt = filetime_to_dt(ft)
    return dt.strftime('%d/%m/%Y %H:%M') if dt else 'N/A'

def ft_days_ago(ft):
    dt = filetime_to_dt(ft)
    if dt is None: return None
    return (datetime.datetime.utcnow() - dt).days

def gt_str(s):
    if not s: return 'N/A'
    try:
        if isinstance(s, bytes): s = s.decode('utf-8','replace')
        return datetime.datetime.strptime(s[:14],'%Y%m%d%H%M%S').strftime('%d/%m/%Y %H:%M')
    except:
        return str(s)[:16]

def sid_to_str(b):
    if not b or len(b) < 8: return ''
    try:
        rev = b[0]; n = b[1]
        auth = int.from_bytes(b[2:8],'big')
        subs = struct.unpack_from(f'<{n}I', b, 8)
        return f'S-{rev}-{auth}-' + '-'.join(str(s) for s in subs)
    except:
        return ''

def decode_str(v):
    if v is None: return ''
    if isinstance(v, bytes): return v.decode('utf-8','replace')
    return str(v)

def h(s):
    return html_mod.escape(str(s))

# ── Apri NTDS.dit (con recovery per dirty state da DC live) ─────
# PROBLEMA: alcune versioni di impacket chiamano mountDB() direttamente nel
# costruttore __init__ di ESENT_DB. Se ntds.dit ha dirty state (3,4,5),
# il crash avviene PRIMA che si possa intercettare con openDatabase().
#
# SOLUZIONE: patch binaria PREVENTIVA dell'offset 52 (dbstate ESE/JET Blue)
# sul file copiato in /tmp. Solo dopo, istanziamo ESENT_DB.
# Il file è già una copia locale → scrivere su di esso è sicuro e legale.
def _patch_ese_state(path):
    """
    Scansiona l'header ESE cercando il magic 0x89ABCDEF e,
    se lo trova, cerca dbstate nei byte successivi e lo patcha a 2 (CleanShutdown).
    Stampa hex dump diagnostico dei primi 80 byte.
    Restituisce (patched: bool, state_found: int).
    """
    ESE_MAGIC = 0x89ABCDEF
    try:
        with open(path, 'r+b') as f:
            header = f.read(80)
        # Hex dump diagnostico (visibile nei log con Debug ESE)
        hexdump = ' '.join(f'{b:02X}' for b in header[:80])
        sys.stderr.write(f"INFO: ESE header[0:80] = {hexdump}\n")

        # Cerca il magic nei primi 64 byte (offset tipici: 0 o 4)
        magic_off = -1
        for off in range(0, 64, 4):
            if len(header) < off + 4:
                break
            val = struct.unpack_from('<I', header, off)[0]
            if val == ESE_MAGIC:
                magic_off = off
                sys.stderr.write(f"INFO: ESE magic trovato a offset {off}\n")
                break

        if magic_off == -1:
            sys.stderr.write("INFO: ESE magic NON trovato nei primi 64 byte — patch impossibile\n")
            return False, -1

        # dbstate si trova tipicamente a +28 o +48 bytes dal magic
        # Prova entrambi gli offset e cerca un valore plausibile (0-6)
        state_off = -1
        state_val = -1
        for delta in (28, 48, 52, 44, 32, 36, 40, 56, 24, 20):
            cand_off = magic_off + delta
            if cand_off + 4 > len(header):
                continue
            v = struct.unpack_from('<I', header, cand_off)[0]
            if 1 <= v <= 6:   # valori ESE dbstate validi
                state_off = cand_off
                state_val = v
                sys.stderr.write(f"INFO: dbstate candidato: offset {cand_off} = {v}\n")
                if v in (3, 4, 5):   # dirty → prendi subito questo
                    break

        if state_off == -1:
            sys.stderr.write("INFO: dbstate non trovato con offset standard — skip patch\n")
            return False, 0

        sys.stderr.write(f"INFO: ESE dbstate = {state_val} (2=Clean, 3=DirtyShutdown) @ offset {state_off}\n")
        if state_val in (3, 4, 5):
            with open(path, 'r+b') as f:
                f.seek(state_off)
                f.write(struct.pack('<I', 2))   # CleanShutdown
            sys.stderr.write(f"INFO: dirty state patchato a CleanShutdown (2) @ offset {state_off}\n")
            return True, state_val

        return False, state_val

    except OSError as e:
        sys.stderr.write(f"INFO: impossibile patchare {path}: {e}\n")
        return False, -1

def _open_esent_db(path, skip_patch=False):
    import traceback

    # Step 1: patch binaria preventiva PRIMA di istanziare ESENT_DB.
    # SKIP se i log ESE sono disponibili: impacket usa il dirty state per capire
    # che deve applicare il soft-recovery. Se patchiamo a CleanShutdown, non lo fa.
    if skip_patch:
        patched, state = False, -1
        sys.stderr.write("INFO: skip patch dirty state — log ESE disponibili per soft-recovery\n")
    else:
        patched, state = _patch_ese_state(path)
        if state == -1:
            sys.stderr.write("WARNING: patch ESE non possibile, tentativo apertura diretta\n")
        else:
            sys.stderr.write(f"INFO: state={state}, patched={patched}\n")

    _last_err = "unknown"
    _last_tb  = ""

    # Step 2: apertura standard (funziona con impacket che richiede openDatabase())
    try:
        db = ESENT_DB(path)
        db.openDatabase()
        return db
    except Exception as e:
        _last_err = str(e)
        _last_tb  = traceback.format_exc()

    # Step 3: alcune versioni di impacket chiamano mountDB() in __init__ e caricano
    # già le tabelle senza richiedere openDatabase(). Proviamo a usare db direttamente.
    try:
        db3 = ESENT_DB(path)
        # Verifichiamo che le tabelle siano state caricate (indica mountDB ok)
        _ = db3.openTable('datatable')
        sys.stderr.write("INFO: aperto senza openDatabase() (mountDB in __init__)\n")
        return db3
    except Exception as e3:
        _last_err = f"{_last_err} | step3: {e3}"
        _last_tb  = traceback.format_exc()

    sys.stderr.write(f"FATAL cannot open NTDS.dit: {_last_err}\n")
    sys.stderr.write(_last_tb)
    sys.exit(1)

# ── Apertura DB (pyesedb → impacket fallback) ────────────────────
db = None
_db_backend = 'none'

if _HAVE_PYESEDB:
    try:
        # pyesedb handles logs automatically if they are in the same dir
        db = _PyesedbDB(ntds_path)
        _db_backend = 'pyesedb'
        sys.stderr.write("INFO: backend=pyesedb\n")
    except Exception as e2:
        sys.stderr.write(f"INFO: pyesedb fallito ({e2}), provo impacket\n")

if db is None and _HAVE_IMPACKET:
    try:
        import impacket.ese as _ese_mod
        
        # Monkey patch
        _orig_pp = _ese_mod.ESENT_DB.parsePage
        _ENTRIES_ATTR = '_ESENT_DB__pageEntries'
        def _safe_parsePage(self, page):
            try:
                _orig_pp(self, page)
            except Exception as _e:
                sys.stderr.write(f"INFO: parsePage tolerant skip pg: {_e}\n")
                try: self.__dict__[_ENTRIES_ATTR] = []
                except Exception: pass
        _ese_mod.ESENT_DB.parsePage = _safe_parsePage

        _orig_pc = _ese_mod.ESENT_DB.parseCatalog
        def _safe_parseCatalog(self, pageNum):
            try: _orig_pc(self, pageNum)
            except Exception as _ep: sys.stderr.write(f"INFO: parseCatalog tolerant skip pageNum={pageNum}: {_ep}\n")
        _ese_mod.ESENT_DB.parseCatalog = _safe_parseCatalog

        _orig_gnr = _ese_mod.ESENT_DB.getNextRow
        def _safe_getNextRow(self, cursor):
            # Loop di recupero tag/pagine corrotti.
            # NOTA: _orig_gnr chiama self.getNextRow() (= questa funzione) per le
            # transizioni di pagina normali → non richiamare mai _orig_gnr nel ramo
            # di errore, altrimenti si crea double-traversal esponenziale.
            # Strategia:
            #   - successo  → restituiamo la riga
            #   - eccezione → _orig_gnr ha già fatto CurrentTag += 1 prima di fallire
            #       * se CurrentTag >= FirstAvailablePageTag: pagina esaurita,
            #         avanziamo manualmente a NextPageNumber (reset tag=0)
            #       * altrimenti: tag corrotto nel mezzo della pagina;
            #         il prossimo giro incrementerà di nuovo e salterà il tag rotto
            for _attempts in range(500000):
                try:
                    return _orig_gnr(self, cursor)
                except Exception:
                    try:
                        page = cursor['CurrentPageData']
                        fap  = page.record['FirstAvailablePageTag']
                        if cursor['CurrentTag'] >= fap:
                            # Pagina esaurita con errore: segui NextPageNumber
                            next_pg = page.record['NextPageNumber']
                            if not next_pg:
                                return None
                            cursor['CurrentPageData'] = self.getPage(next_pg)
                            cursor['CurrentTag'] = 0
                        # else: tag corrotto nel mezzo — il prossimo _orig_gnr salta
                    except Exception:
                        return None
            return None
        _ese_mod.ESENT_DB.getNextRow = _safe_getNextRow
        ESENT_DB = _ese_mod.ESENT_DB
        if not hasattr(ESENT_DB, 'closeTable'): ESENT_DB.closeTable = lambda self, cur: None
        if not hasattr(ESENT_DB, 'moveHead'): ESENT_DB.moveHead = lambda self, cur: None

        # skip_patch=False: patchiamo sempre a CleanShutdown per sicurezza nella lettura.
        # I log ESE sono già stati copiati in tmp e verranno usati da openDatabase().
        db_imp = _open_esent_db(ntds_path, skip_patch=False)
        db = db_imp
        _db_backend = 'impacket'
        sys.stderr.write(f"INFO: backend=impacket has_logs={_has_ese_logs}\n")
    except SystemExit:
        raise
    except Exception as e:
        sys.stderr.write(f"INFO: impacket fallito ({e})\n")

if db is None:
    sys.stderr.write("FATAL: unable to open NTDS.dit: no backend available.\n")
    sys.stderr.write("INFO: installa libesedb-python: pip install libesedb-python\n")
    sys.exit(1)

# ── Scoperta colonne ─────────────────────────────────────────────
# Con impacket, getNextRow restituisce solo le colonne con dati nella riga corrente
# e le chiavi possono essere bytes. Normalizziamo tutto a str e campionamo più righe.
def _norm_key(k):
    return k.decode('utf-8', 'replace') if isinstance(k, bytes) else str(k)

def _norm_row(r):
    """Normalizza le chiavi del dict da bytes a str (impacket può restituire bytes)."""
    if r is None: return None
    return {_norm_key(k): v for k, v in r.items()}

_KEY_COLS = {'ATTm590045', 'ATTj590014', 'ATTi590014', 'ATTi590126', 'ATTq589920'}
try:
    cur = db.openTable('datatable')
    avail = set()
    for _ in range(500):
        _r = _norm_row(db.getNextRow(cur))
        if _r is None: break
        avail.update(_r.keys())
        if _KEY_COLS.issubset(avail): break   # colonne chiave trovate, stop anticipato
    db.closeTable(cur)
    sys.stderr.write(f"INFO: schema discovery — {len(avail)} colonne trovate in datatable\n")
    sys.stderr.write(f"INFO: COL_SAM={'ATTm590045' in avail} COL_STYPE={'ATTj590014' in avail or 'ATTi590014' in avail} COL_UAC={'ATTi590126' in avail}\n")
except Exception as e:
    sys.stderr.write(f"FATAL cannot read datatable schema: {e}\n"); sys.exit(1)

COL_SAM   = find_col(['ATTm590045'], avail)
COL_NAME  = find_col(['ATTm590480'], avail)
COL_UAC   = find_col(['ATTi590126'], avail)
COL_PWD   = find_col(['ATTq589920'], avail)
COL_LOGON = find_col(['ATTq589876'], avail)
COL_CRE   = find_col(['ATTm131074'], avail)
COL_CHG   = find_col(['ATTl131075','ATTm131075'], avail)
COL_SID   = find_col(['ATTb590597'], avail)
COL_ADMC  = find_col(['ATTj590420'], avail)
COL_EXP   = find_col(['ATTq591520','ATTq589984'], avail)
COL_NTSEC = find_col(['ATTp131353','ATTb131353'], avail)
COL_OS    = find_col(['ATTm590474'], avail)
COL_DISP  = find_col(['ATTm131085','ATTm590042'], avail)
COL_GPC   = find_col(['ATTm590258'], avail)
COL_BEHAV = find_col(['ATTj590182','ATTl590182'], avail)
COL_DEL   = find_col(['ATTi590829'], avail)
COL_STYPE = find_col(['ATTj590014','ATTi590014'], avail)
COL_SIDH  = find_col(['ATTb589971'], avail)
COL_OPTF  = find_col(['ATTb591520','ATTm591520'], avail)

# ── Diagnosi catalogo corrotto ────────────────────────────────────
_catalog_corrupt = (not COL_SAM or not COL_STYPE or not COL_UAC)
if _catalog_corrupt:
    _sample = sorted(c for c in list(avail)[:20] if c.startswith('ATT'))
    sys.stderr.write(f"WARNING: catalogo ESE parzialmente corrotto — colonne chiave non trovate\n")
    sys.stderr.write(f"WARNING: COL_SAM={COL_SAM} COL_STYPE={COL_STYPE} COL_UAC={COL_UAC}\n")
    sys.stderr.write(f"WARNING: campione colonne ATT trovate: {_sample[:10]}\n")
    sys.stderr.write(f"CATALOG_CORRUPT:1\n")
    # Scrivi comunque un report parziale, ma avvisa l'utente
    print("\n  [!] NTDS.dit con catalogo ESE parzialmente corrotto.")
    print("      I log ESE (.jrs) sono necessari per il soft-recovery.")
    print("      Vedi sezione 'Recupero NTDS.dit' nel report HTML.")

# ── Scansione datatable ──────────────────────────────────────────
objects      = {}   # dnt -> obj dict
names_by_dnt = {}   # dnt -> name (RDN)
pdnt_by_dnt  = {}   # dnt -> parent dnt
_scan_count = 0

try:
    sys.stdout.write("  [i] Lettura datatable in corso (0 oggetti)...\r")
    sys.stdout.flush()
    cur = db.openTable('datatable')
    db.moveHead(cur)
    while True:
        row = _norm_row(db.getNextRow(cur))
        if row is None: break

        _scan_count += 1
        if _scan_count % 500 == 0:
            sys.stdout.write(f"  [i] Lettura datatable in corso ({_scan_count} oggetti)... \r")
            sys.stdout.flush()

        try:
            dnt  = row.get('DNT_col')
            pdnt = row.get('PDNT_col')
            if dnt is None: continue
            if COL_DEL and row.get(COL_DEL): continue   # skip deleted

            rdn = decode_str(row.get(COL_NAME)) if COL_NAME else ''
            if rdn: names_by_dnt[dnt] = rdn
            if pdnt: pdnt_by_dnt[dnt] = pdnt

            sam     = decode_str(row.get(COL_SAM))   if COL_SAM  else ''
            ntsec   = row.get(COL_NTSEC)              if COL_NTSEC else None
            behav   = row.get(COL_BEHAV)              if COL_BEHAV else None
            gpc     = decode_str(row.get(COL_GPC))   if COL_GPC  else ''
            optf    = row.get(COL_OPTF)               if COL_OPTF  else None

            if not sam and not ntsec and behav is None and not gpc and not optf:
                continue

            sid_raw = row.get(COL_SID) if COL_SID else None
            objects[dnt] = {
                'dnt':       dnt,
                'pdnt':      pdnt,
                'name':      rdn,
                'sam':       sam,
                'uac':       (row.get(COL_UAC)   if COL_UAC  else None) or 0,
                'pwdLastSet':(row.get(COL_PWD)   if COL_PWD  else None),
                'lastLogon': (row.get(COL_LOGON) if COL_LOGON else None),
                'whenCreated':decode_str(row.get(COL_CRE)) if COL_CRE else '',
                'whenChanged':(row.get(COL_CHG)  if COL_CHG  else None),
                'sid':       sid_to_str(sid_raw) if sid_raw else '',
                'adminCount':(row.get(COL_ADMC)  if COL_ADMC else None) or 0,
                'expires':   (row.get(COL_EXP)   if COL_EXP  else None),
                'ntsec':     ntsec,
                'os':        decode_str(row.get(COL_OS))   if COL_OS  else '',
                'disp':      decode_str(row.get(COL_DISP)) if COL_DISP else '',
                'gpc':       gpc,
                'behavior':  behav,
                'samType':   (row.get(COL_STYPE) if COL_STYPE else None),
                'hasSidHist':bool(row.get(COL_SIDH)) if COL_SIDH else False,
                'optf':      optf,
            }
        except:
            continue
    db.closeTable(cur)
except Exception as e:
    sys.stderr.write(f"ERROR datatable scan: {e}\n")

sys.stderr.write(f"INFO: scan_count={_scan_count} objects={len(objects)} names={len(names_by_dnt)}\n")

# ── Classifica oggetti ────────────────────────────────────────────
def obj_type(o):
    st = o.get('samType')
    if st == SAMTYPE_COMPUTER: return 'computer'
    if st in (SAMTYPE_GROUP, 268435457): return 'group'
    if st == SAMTYPE_USER: return 'user'
    uac = o.get('uac', 0) or 0
    if uac & UAC_DC:          return 'dc'
    if uac & UAC_WORKSTATION: return 'computer'
    sam = o.get('sam', '') or ''
    if sam:
        # Fallback quando samAccountType non disponibile (catalogo corrotto):
        # account computer terminano con '$', gli altri sono utenti o gruppi.
        # I nomi gruppo coincidono spesso con o['name'] — distinguiamo per assenza di '$'.
        if sam.endswith('$'): return 'computer'
        return 'user'   # assume utente se ha SAM e non è un computer
    return 'other'

# ── Mappa SID → nome (dai dati estratti) ─────────────────────────
sid_map = dict(WELLKNOWN_SIDS)
for o in objects.values():
    if o['sid'] and o['sam']:
        sid_map[o['sid']] = o['sam']

# ── Costruzione DN ────────────────────────────────────────────────
def build_dn(dnt):
    parts = []; seen = set(); cur = dnt
    for _ in range(25):
        if cur in seen: break
        seen.add(cur)
        n = names_by_dnt.get(cur,'')
        if n: parts.append(f'CN={n}')
        p = pdnt_by_dnt.get(cur)
        if p is None or p == cur: break
        cur = p
    return ','.join(parts)

# ── Lettura link_table (group membership) ────────────────────────
group_members = {}   # group_dnt -> set(member_dnt)
member_of     = {}   # user_dnt  -> set(group_dnt)

try:
    cur = db.openTable('link_table')
    first_lrow = _norm_row(db.getNextRow(cur))
    lcols = set(first_lrow.keys()) if first_lrow else set()
    db.closeTable(cur)

    LCOL_GRP  = find_col(['link_DNT'],     lcols)
    LCOL_MBR  = find_col(['backlink_DNT'], lcols)
    LCOL_DEL  = find_col(['link_deltime'], lcols)
    LCOL_BASE = find_col(['link_base'],    lcols)
    # Filtriamo per link_base=2 (forward link dell'attributo `member` AD standard).
    # Se il catalogo è corrotto e link_base non è mappato correttamente,
    # accettiamo qualsiasi valore di link_base noto per l'attributo `member`:
    # linkID=2 nel default Windows AD schema.
    # Se LCOL_BASE è None (colonna non trovata), non filtriamo.
    MEMBER_LINK_BASE = 2  # forward linkID del `member` attribute

    if LCOL_GRP and LCOL_MBR:
        sys.stdout.write("\n  [i] Lettura link_table in corso (0 link)...\r")
        sys.stdout.flush()
        _link_count = 0
        cur = db.openTable('link_table')
        db.moveHead(cur)
        while True:
            row = _norm_row(db.getNextRow(cur))
            if row is None: break

            _link_count += 1
            if _link_count % 500 == 0:
                sys.stdout.write(f"  [i] Lettura link_table in corso ({_link_count} link)... \r")
                sys.stdout.flush()

            try:
                deltime = row.get(LCOL_DEL) if LCOL_DEL else None
                if deltime and deltime != 0: continue
                if LCOL_BASE:
                    lb = row.get(LCOL_BASE)
                    if lb is not None and lb != MEMBER_LINK_BASE: continue
                gd  = row.get(LCOL_GRP)
                md  = row.get(LCOL_MBR)
                if gd is None or md is None: continue
                group_members.setdefault(gd, set()).add(md)
                member_of.setdefault(md, set()).add(gd)
            except:
                continue
        db.closeTable(cur)
except Exception as e:
    sys.stderr.write(f"WARNING link_table: {e}\n")

# ── Indice gruppi per nome ────────────────────────────────────────
groups_by_name = {}  # name -> dnt
for dnt, o in objects.items():
    if _catalog_corrupt:
        # COL_NAME (ATTm590480) spesso corrotto → o['name'] quasi sempre vuoto.
        # I gruppi AD hanno sAMAccountName uguale al loro nome (es. "Domain Admins").
        # Indexiamo sia per name che per sam così PRIV_GROUPS viene trovato.
        if o['name']:
            groups_by_name[o['name']] = dnt
        if o['sam']:
            groups_by_name[o['sam']] = dnt
    else:
        if obj_type(o) in ('group','other') and o['name']:
            groups_by_name[o['name']] = dnt

priv_group_dnts = {dnt for name, dnt in groups_by_name.items() if name in PRIV_GROUPS}

# ── Membership transitiva (BFS) ────────────────────────────────────
def transitive_groups(user_dnt):
    visited = set(); queue = list(member_of.get(user_dnt, []))
    while queue:
        g = queue.pop()
        if g in visited: continue
        visited.add(g)
        for pg in member_of.get(g, []):
            if pg not in visited: queue.append(pg)
    return visited

# ── Analisi utenti privilegiati ────────────────────────────────────
priv_users = []
all_users  = [(dnt, o) for dnt, o in objects.items() if obj_type(o) in ('user','dc')]

for dnt, o in all_users:
    if not o['sam']: continue
    all_grp_dnts = transitive_groups(dnt)
    matched = all_grp_dnts & priv_group_dnts
    if not matched: continue

    memberships = []
    for gd in matched:
        go = objects.get(gd, {})
        gname = go.get('name', str(gd))
        is_direct = dnt in group_members.get(gd, set())
        memberships.append(f"{h(gname)} ({'Direct' if is_direct else 'Nested'})")

    uac = o['uac'] or 0
    pwd_dt  = filetime_to_dt(o['pwdLastSet'])
    logon_dt= filetime_to_dt(o['lastLogon'])
    expires_dt = filetime_to_dt(o.get('expires'))
    
    pwd_days= (datetime.datetime.utcnow()-pwd_dt).days if pwd_dt else None
    logon_days= (datetime.datetime.utcnow()-logon_dt).days if logon_dt else None
    is_expired= (datetime.datetime.utcnow() > expires_dt) if expires_dt else False
    is_inactive90 = (logon_days is not None and logon_days > 90)

    priv_users.append({
        'sam':        o['sam'],
        'groups':     '; '.join(memberships),
        'pwdLastSet': ft_str(o['pwdLastSet']),
        'pwdDays':    pwd_days,
        'lastLogon':  ft_str(o['lastLogon']),
        'disabled':   bool(uac & UAC_DISABLED),
        'noExpire':   bool(uac & UAC_NO_EXPIRE),
        'adminCount': o['adminCount'],
        'sid':        o['sid'],
        'hasSidHist': o['hasSidHist'],
        'stale':      (pwd_days is not None and pwd_days > 1),
        'expired':    is_expired,
        'inactive90': is_inactive90,
    })

priv_users.sort(key=lambda u: u['sam'].lower())

# ── KRBTGT ─────────────────────────────────────────────────────────
krbtgt_pwd = 'N/A'; krbtgt_days = None
for dnt, o in objects.items():
    if o['sam'].lower() == 'krbtgt':
        krbtgt_pwd  = ft_str(o['pwdLastSet'])
        krbtgt_days = ft_days_ago(o['pwdLastSet'])
        break

# ── Functional level ───────────────────────────────────────────────
domain_fl = 'N/A'; domain_root_dnt = None
for dnt, o in objects.items():
    if o.get('behavior') is not None:
        fv = o['behavior']
        domain_fl = f"{fv} — {FL_MAP.get(fv, 'Unknown')}"
        domain_root_dnt = dnt
        break

# ── Recycle Bin ────────────────────────────────────────────────────
recycle_bin = 'Disabled'
for dnt, o in objects.items():
    if o['name'] == 'Recycle Bin Feature':
        recycle_bin = 'Enabled'
        break
    if o.get('optf') and o['name']:
        recycle_bin = 'Enabled'
        break

# ── GPO ────────────────────────────────────────────────────────────
gpos = []
for dnt, o in objects.items():
    if not o['gpc']: continue
    disp = o['disp'] or o['name'] or '(no name)'
    when_chg = ft_str(o['whenChanged']) if isinstance(o['whenChanged'], int) else gt_str(o.get('whenCreated',''))
    days_ago = ft_days_ago(o['whenChanged']) if isinstance(o['whenChanged'], int) else None
    gpos.append({'name': disp, 'changed': when_chg, 'path': o['gpc'], 'days': days_ago})
gpos.sort(key=lambda g: g['changed'], reverse=True)

# ── Computer ────────────────────────────────────────────────────────
computers = []
for dnt, o in objects.items():
    if obj_type(o) not in ('computer','dc'): continue
    if not o['name']: continue
    computers.append({'name': o['name'], 'os': o['os'] or 'N/A', 'dn': build_dn(dnt)})
computers.sort(key=lambda c: c['name'].lower())

# ── Parser Security Descriptor ─────────────────────────────────────
def parse_sd(sd_bytes):
    aces = []
    if not sd_bytes or len(sd_bytes) < 20: return aces
    try:
        off_dacl = struct.unpack_from('<I', sd_bytes, 16)[0]
        if off_dacl == 0 or off_dacl >= len(sd_bytes): return aces
        ace_count = struct.unpack_from('<H', sd_bytes, off_dacl + 4)[0]
        ace_off   = off_dacl + 8
        for _ in range(min(ace_count, 200)):
            if ace_off + 8 > len(sd_bytes): break
            ace_type  = sd_bytes[ace_off]
            ace_flags = sd_bytes[ace_off + 1]
            ace_size  = struct.unpack_from('<H', sd_bytes, ace_off + 2)[0]
            if ace_size < 8: break
            access_mask = struct.unpack_from('<I', sd_bytes, ace_off + 4)[0]
            # SID offset: standard ACE=8, object ACE (type 5/6)=8+16[+16]
            sid_off = ace_off + 8
            if ace_type in (5, 6):
                flags_oe = struct.unpack_from('<I', sd_bytes, ace_off + 8)[0]
                sid_off = ace_off + 8 + 4
                if flags_oe & 1: sid_off += 16
                if flags_oe & 2: sid_off += 16
            sid_str = sid_to_str(sd_bytes[sid_off:sid_off + 68])
            if sid_str and sid_str not in ACL_SKIP_SIDS and ace_type in (0, 5):
                flagged = [nm for mask, nm in DANGEROUS_RIGHTS if access_mask & mask]
                if flagged:
                    aces.append({
                        'sid':       sid_str,
                        'name':      sid_map.get(sid_str, sid_str),
                        'rights':    ', '.join(flagged),
                        'mask':      hex(access_mask),
                        'inherited': bool(ace_flags & 0x10),
                        'everyone':  sid_str in ('S-1-1-0','S-1-5-11','S-1-5-7'),
                    })
            ace_off += ace_size
    except Exception as e:
        pass
    return aces

# ── ACL domain root + DC OU ────────────────────────────────────────
root_aces = []; dc_ou_aces = []

if domain_root_dnt is not None:
    ntsec = objects[domain_root_dnt].get('ntsec')
    if ntsec: root_aces = parse_sd(ntsec)

dc_ou_dnt = None
for dnt, o in objects.items():
    if o['name'] == 'Domain Controllers':
        p = pdnt_by_dnt.get(dnt)
        if p == domain_root_dnt or True:
            dc_ou_dnt = dnt; break

if dc_ou_dnt is not None:
    ntsec = objects[dc_ou_dnt].get('ntsec')
    if ntsec: dc_ou_aces = parse_sd(ntsec)

# ── Statistiche ────────────────────────────────────────────────────
total_priv    = len(priv_users)
total_stale   = sum(1 for u in priv_users if u['stale'])
total_disabled= sum(1 for u in priv_users if u['disabled'])
total_expired = sum(1 for u in priv_users if u['expired'])
total_inactive= sum(1 for u in priv_users if u['inactive90'])
total_gpo     = len(gpos)
gpo_recent    = sum(1 for g in gpos if g['days'] is not None and g['days'] <= 30)
total_comp    = len(computers)

# ── CSS / HTML ────────────────────────────────────────────────────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@300;400;500;700&family=DM+Sans:wght@400;500;700&display=swap');
:root{--bg:#080b0f;--bg2:#0d1117;--bg3:#131920;--bg4:#1a2332;
  --border:#1e2d3d;--accent:#58a6ff;--accent2:#ff7b72;
  --accent3:#3fb950;--accent4:#f0883e;
  --text:#c9d1d9;--text-dim:#3d5166;--text-mid:#6e8898;
  --mono:'Fira Code',monospace;--sans:'DM Sans',sans-serif;}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:var(--text);font-family:var(--sans);min-height:100vh;}
body::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background-image:linear-gradient(rgba(88,166,255,.03) 1px,transparent 1px),
  linear-gradient(90deg,rgba(88,166,255,.03) 1px,transparent 1px);
  background-size:40px 40px;}
header{position:relative;z-index:1;background:var(--bg2);border-bottom:1px solid var(--border);
  padding:1.8rem 3rem;display:flex;align-items:center;gap:2rem;}
.hicon{width:3rem;height:3rem;background:linear-gradient(135deg,var(--accent),#1f6feb);
  display:flex;align-items:center;justify-content:center;font-family:var(--mono);
  font-size:1.1rem;color:#fff;font-weight:700;flex-shrink:0;
  clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%);}
.htxt h1{font-size:1.35rem;font-weight:700;color:#fff;}
.htxt h1 span{color:var(--accent);}
.htxt .sub{font-size:.72rem;color:var(--text-dim);font-family:var(--mono);margin-top:.25rem;}
.hmeta{margin-left:auto;font-family:var(--mono);font-size:.7rem;color:var(--text-dim);text-align:right;line-height:2;}
.hmeta .val{color:var(--accent);}
.statsbar{position:relative;z-index:1;background:var(--bg3);border-bottom:1px solid var(--border);
  padding:1rem 3rem;display:flex;gap:3rem;flex-wrap:wrap;}
.stat .label{font-size:.6rem;text-transform:uppercase;letter-spacing:.15em;color:var(--text-dim);font-family:var(--mono);}
.stat .value{font-size:1.8rem;font-weight:800;font-family:var(--mono);color:var(--accent2);}
.stat.ok .value{color:var(--accent3);}
.stat.info .value{color:var(--accent);}
main{position:relative;z-index:1;padding:2rem 3rem;max-width:1600px;margin:0 auto;}
.stitle{font-size:.65rem;text-transform:uppercase;letter-spacing:.2em;color:var(--text-dim);
  font-family:var(--mono);margin-bottom:1.5rem;padding-bottom:.5rem;border-bottom:1px solid var(--border);}
.stitle::before{content:'> ';color:var(--accent3);}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:6px;overflow:hidden;margin-bottom:2rem;}
table{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:.78rem;}
th{padding:.5rem 1rem;text-align:left;font-size:.6rem;text-transform:uppercase;
  letter-spacing:.15em;color:var(--text-dim);border-bottom:1px solid var(--border);
  background:rgba(0,0,0,.3);}
td{padding:.5rem 1rem;border-bottom:1px solid rgba(30,45,61,.5);vertical-align:top;}
tr:last-child td{border-bottom:none;}
tr:hover{background:rgba(88,166,255,.03);}
.mono{font-family:var(--mono);}
.dim{color:var(--text-dim);}
.ok{color:var(--accent3);}
.bad{color:var(--accent2);}
.warn{color:var(--accent4);}
.tag{display:inline-block;padding:.1rem .4rem;border-radius:2px;font-size:.65rem;margin:.1rem;font-family:var(--mono);}
.tag-bad{background:rgba(255,123,114,.15);color:var(--accent2);border:1px solid rgba(255,123,114,.3);}
.tag-ok{background:rgba(63,185,80,.15);color:var(--accent3);border:1px solid rgba(63,185,80,.3);}
.tag-info{background:rgba(88,166,255,.15);color:var(--accent);border:1px solid rgba(88,166,255,.3);}
.tag-warn{background:rgba(240,136,62,.15);color:var(--accent4);border:1px solid rgba(240,136,62,.3);}
.row-bad{background:rgba(255,123,114,.06);border-left:3px solid var(--accent2);}
.row-warn{background:rgba(240,136,62,.06);}
footer{position:relative;z-index:1;border-top:1px solid var(--border);padding:1rem 3rem;
  font-family:var(--mono);font-size:.62rem;color:var(--text-dim);margin-top:3rem;
  display:flex;justify-content:space-between;}
footer span{color:var(--accent);}
"""

def tag(label, cls='info'):
    return f"<span class='tag tag-{cls}'>{h(label)}</span>"

def uac_tags(u):
    out = ''
    if u['disabled']:  out += tag('DISABLED','bad')
    if u['expired']:   out += tag('EXPIRED','bad')
    if u['inactive90']:out += tag('INACTIVE >90d','warn')
    if u['noExpire']:  out += tag('PWD-NO-EXPIRE','warn')
    if u['hasSidHist']:out += tag('SID-HISTORY','bad')
    if u['stale']:     out += tag('PWD-STALE','warn')
    return out

# ── Sezione utenti privilegiati ───────────────────────────────────
rows_priv = ''
for u in priv_users:
    cls = 'row-bad' if (u['disabled'] or u['hasSidHist'] or u['expired']) else ('row-warn' if (u['stale'] or u['inactive90']) else '')
    rows_priv += f"""<tr class='{cls}'>
<td class='mono ok' style='white-space:nowrap'>{h(u['sam'])}</td>
<td style='font-size:.68rem'>{u['groups']}</td>
<td class='mono {"bad" if u["stale"] else "ok"}' style='white-space:nowrap'>{h(u['pwdLastSet'])}</td>
<td class='mono dim' style='white-space:nowrap'>{h(u['lastLogon'])}</td>
<td>{uac_tags(u)}</td>
<td class='mono dim' style='font-size:.65rem;word-break:break-all'>{h(u['sid'])}</td>
</tr>"""

# ── Sezione GPO ────────────────────────────────────────────────────
rows_gpo = ''
for g in gpos:
    cls = 'row-warn' if (g['days'] is not None and g['days'] <= 30) else ''
    badge = tag(f"{g['days']}d ago",'warn') if (g['days'] is not None and g['days'] <= 30) else ''
    rows_gpo += f"""<tr class='{cls}'>
<td class='mono'>{h(g['name'])}{badge}</td>
<td class='mono {"warn" if (g["days"] is not None and g["days"]<=30) else "dim"}' style='white-space:nowrap'>{h(g['changed'])}</td>
<td class='mono dim' style='font-size:.65rem;word-break:break-all'>{h(g['path'])}</td>
</tr>"""

# ── Sezione Computer ────────────────────────────────────────────────
rows_comp = ''
for c in computers:
    rows_comp += f"""<tr>
<td class='mono ok'>{h(c['name'])}</td>
<td class='mono dim'>{h(c['os'])}</td>
<td class='mono dim' style='font-size:.65rem'>{h(c['dn'])}</td>
</tr>"""

# ── Sezione ACL ────────────────────────────────────────────────────
def ace_rows(aces):
    if not aces:
        return f"<tr><td colspan='5' class='dim' style='padding:1rem;text-align:center'>{L('Nessun ACE pericoloso rilevato (o nTSecurityDescriptor non trovato)', 'No dangerous ACE found (or nTSecurityDescriptor not present)')}</td></tr>"
    out = ''
    for a in aces:
        cls = 'row-bad' if a['everyone'] else 'row-warn'
        ev  = tag('EVERYONE-LIKE','bad') if a['everyone'] else ''
        inh = tag('inherited','info') if a['inherited'] else ''
        out += f"""<tr class='{cls}'>
<td class='mono {"bad" if a["everyone"] else "warn"}'>{h(a['name'])}{ev}</td>
<td class='mono dim' style='font-size:.65rem;word-break:break-all'>{h(a['sid'])}</td>
<td class='mono bad'>{h(a['rights'])}</td>
<td class='mono dim'>{h(a['mask'])}</td>
<td>{inh}</td>
</tr>"""
    return out

# ── Functional level badge ─────────────────────────────────────────
fl_badge_cls = 'ok' if domain_fl != 'N/A' else 'dim'
krbtgt_cls = 'bad' if (krbtgt_days is not None and krbtgt_days > 180) else 'ok'
krbtgt_badge = tag(f'{krbtgt_days}d ago','bad') if (krbtgt_days and krbtgt_days > 180) else (tag(f'{krbtgt_days}d ago','ok') if krbtgt_days else '')

# ── Frammenti HTML pre-generati (evitano f-string annidati con lo stesso quoting) ──
_krbtgt_risk_html = (
    f'<span class="tag tag-bad">{L("GOLDEN TICKET RISK: rotazione &gt;180gg", "GOLDEN TICKET RISK: rotation &gt;180d")}</span>'
    if (krbtgt_days and krbtgt_days > 180) else ''
)
def _empty_row(cols, it, en):
    return f"<tr><td colspan='{cols}' class='dim' style='padding:1rem;text-align:center'>{L(it, en)}</td></tr>"

_no_priv_row = _empty_row(6, 'Nessun utente privilegiato trovato', 'No privileged users found')
_no_gpo_row  = _empty_row(3, 'Nessuna GPO trovata', 'No GPOs found')
_no_comp_row = _empty_row(3, 'Nessun computer trovato', 'No computers found')

# ── Blocco HTML per catalogo corrotto (pre-generato per evitare f-string annidati) ──
_catalog_corrupt_html = ""
if _catalog_corrupt:
    _objects_count = len(objects)
    _catalog_corrupt_html = f"""
<div class="stitle" style="color:var(--accent2)">&#9888; NTDS.dit &mdash; {L("Catalogo ESE corrotto: analisi privilegiata non disponibile", "Corrupt ESE catalogue: privileged analysis unavailable")}</div>
<div class="card" style="border-color:rgba(255,123,114,.5);background:rgba(255,123,114,.06)">
<div style="padding:1.4rem 1.6rem;font-family:var(--mono);font-size:.82rem;line-height:1.9">
<div style="color:var(--accent2);font-weight:700;font-size:.95rem;margin-bottom:.9rem">Stato: DirtyShutdown &mdash; Pagine ESE corrotte nel catalogo di sistema</div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem">
  <div style="background:rgba(0,0,0,.25);border-radius:6px;padding:.9rem 1rem">
    <div style="color:var(--accent3);font-weight:700;margin-bottom:.5rem">&#10003; Dati estratti con successo</div>
    <div style="color:var(--text-mid)">&#x2022; Password KRBTGT (pwdLastSet)</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Livello funzionale dominio", "Domain functional level")}</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Oggetti GPO e percorsi", "GPO objects and paths")}</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Elenco computer (parziale)", "Computer list (partial)")}</div>
    <div style="color:var(--text-mid)">&#x2022; {_objects_count} oggetti datatable recuperati</div>
  </div>
  <div style="background:rgba(0,0,0,.25);border-radius:6px;padding:.9rem 1rem">
    <div style="color:var(--accent2);font-weight:700;margin-bottom:.5rem">&#10007; Dati non disponibili</div>
    <div style="color:var(--text-mid)">&#x2022; sAMAccountName (colonna non mappata)</div>
    <div style="color:var(--text-mid)">&#x2022; samAccountType / userAccountControl</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Membership gruppi privilegiati", "Privileged group membership")}</div>
    <div style="color:var(--text-mid)">&#x2022; Link-table group membership</div>
    <div style="color:var(--text-mid)">&#x2022; {L("Hash delle password", "Password hashes")}</div>
  </div>
</div>
<div style="color:var(--text-dim);margin-bottom:1rem">
  <b style="color:var(--text)">Causa tecnica:</b> Le pagine ESE del catalogo di sistema (pagine 14, 19&ndash;22) sono corrotte.
  Il catalogo mappa i nomi delle colonne AD (es. <code style="color:var(--accent)">ATTm590045</code> &rarr; sAMAccountName) agli ID interni:
  senza di esso, impacket e libesedb non possono identificare correttamente i campi.
  Le entry di group membership nella link_table risultano su pagine non leggibili.
</div>
<div style="margin-top:.8rem;color:var(--text);font-weight:700;border-top:1px solid rgba(255,255,255,.08);padding-top:.8rem">Procedura di recupero</div>
<div style="margin:.6rem 0">
  <div style="color:var(--accent3);font-weight:600">Opzione 1 &mdash; esentutl su Windows <span style="color:var(--text-dim)">(recovery completo, consigliato)</span></div>
  <div style="color:var(--text-dim);margin:.2rem 0 .2rem 1rem">{L("Sul DC sorgente o una VM Windows con ntds.dit + log ESE nella stessa cartella:", "On the source DC or a Windows VM with ntds.dit + ESE log files in the same folder:")}</div>
  <div style="background:rgba(0,0,0,.3);border-radius:4px;padding:.4rem .8rem;margin:.2rem 0 .4rem 1rem;color:var(--accent)">esentutl /r edb /l "C:\\Windows\\NTDS" /s "C:\\Windows\\NTDS" /!32768</div>
  <div style="color:var(--text-dim);margin-left:1rem">Poi copia ntds.dit recuperato e rilancia il modulo 38.</div>
</div>
<div style="margin:.6rem 0">
  <div style="color:var(--accent3);font-weight:600">Opzione 2 &mdash; ntdsutil IFM <span style="color:var(--text-dim)">(se il DC &egrave; ancora operativo)</span></div>
  <div style="background:rgba(0,0,0,.3);border-radius:4px;padding:.4rem .8rem;margin:.2rem 0 .4rem 1rem;color:var(--accent)">ntdsutil "activate instance ntds" "ifm" "create full C:\\ifm" quit quit</div>
  <div style="color:var(--text-dim);margin-left:1rem">Il file <code>C:\\ifm\\Active Directory\\ntds.dit</code> sar&agrave; gi&agrave; in CleanShutdown.</div>
</div>
<div style="margin:.6rem 0">
  <div style="color:var(--accent3);font-weight:600">Opzione 3 &mdash; esentutl /p (hard repair) <span style="color:var(--text-dim)">(richiede Windows &mdash; usa solo se Opzione 1 fallisce)</span></div>
  <div style="color:var(--text-dim);margin:.2rem 0 .2rem 1rem">
    Ripara le pagine fisicamente danneggiate azzerandole. <b style="color:var(--accent2)">Possibile perdita dati parziale.</b> Esegui prima un backup.
  </div>
  <div style="background:rgba(0,0,0,.3);border-radius:4px;padding:.4rem .8rem;margin:.2rem 0 .4rem 1rem;color:var(--accent)">
    esentutl /p ntds.dit /!32768 /8 /o
  </div>
  <div style="color:var(--text-dim);margin-left:1rem">
    Dopo /p: ri-esegui <code>/r edb</code> per il soft-recovery, poi rilancia il modulo 38.<br>
    <b>Nota:</b> il soft-recovery (<code>impacket</code> o <code>esentutl /r</code>) non ripara pagine fisicamente corrotte &mdash; serve /p prima.
  </div>
</div>
<div style="margin:.6rem 0;padding:.6rem .8rem;background:rgba(255,123,114,.08);border-radius:4px;border-left:3px solid rgba(255,123,114,.5)">
  <b style="color:var(--accent2)">Perch&eacute; il soft-recovery non basta:</b>
  <span style="color:var(--text-dim)"> Il replay dei log ESE corregge solo lo stato <i>DirtyShutdown</i> (transazioni non committed).
  Le pagine di catalogo 14, 19&ndash;22 di questo NTDS.dit sono fisicamente danneggiate (byte corrotti nel payload ESE):
  {L("nessun log pu&ograve; ripristinare dati che non esistono pi&ugrave; sul disco.", "no log files can restore data that no longer exists on disk.")}</span>
</div>
</div></div>"""

# ── HTML output ────────────────────────────────────────────────────
html_out = f"""<!DOCTYPE html>
<html lang="{L("it", "en")}"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>PAD Offline — DFIR Report</title>
<style>{CSS}</style></head><body>
<header>
  <div class="hicon">AD</div>
  <div class="htxt">
    <h1>PAD Offline — <span>{L("Panoramica Active Directory", "Active Directory Overview")}</span></h1>
    <div class="sub">NTDS.dit offline forensic analysis</div>
  </div>
  <div class="hmeta">
    <div>{L("Scansione:", "Scan:")} <span class="val">{h(scan_date)}</span></div>
    <div>{L("Sorgente:", "Source:")} <span class="val">{h(source_root)}</span></div>
    <div style="font-size:.6rem">SHA256: <span class="val" style="word-break:break-all">{h(ntds_hash)}</span></div>
  </div>
</header>
<div class="statsbar">
  <div class="stat"><div class="label">{L("Utenti Privilegiati", "Privileged Users")}</div><div class="value">{total_priv}</div></div>
  <div class="stat"><div class="label">{L("Gruppi Monitorati", "Monitored Groups")}</div><div class="value info" style="color:var(--accent)">{len(PRIV_GROUPS)}</div></div>
  <div class="stat"><div class="label">{L("KRBTGT (giorni)", "KRBTGT (days)")}</div><div class="value {krbtgt_cls}">{krbtgt_days if krbtgt_days is not None else 'N/A'}</div></div>
  <div class="stat"><div class="label">{L("Password Stale (&gt;24h)", "Stale Passwords (&gt;24h)")}</div><div class="value">{total_stale}</div></div>
  <div class="stat"><div class="label">{L("GPO Modificate (30gg)", "Modified GPOs (30d)")}</div><div class="value {"" if gpo_recent==0 else ""}">{gpo_recent}/{total_gpo}</div></div>
  <div class="stat ok"><div class="label">{L("Computer nel Dominio", "Domain Computers")}</div><div class="value">{total_comp}</div></div>
</div>
<main>

<div class="stitle">{L("Configurazione Dominio", "Domain Configuration")}</div>
<div class="card"><table>
<thead><tr><th>{L("Parametro", "Parameter")}</th><th>{L("Valore", "Value")}</th><th>Note</th></tr></thead>
<tbody>
<tr><td class="mono dim">Domain Functional Level</td>
    <td class="mono {fl_badge_cls}">{h(domain_fl)}</td><td></td></tr>
<tr><td class="mono dim">KRBTGT pwdLastSet</td>
    <td class="mono {krbtgt_cls}">{h(krbtgt_pwd)}</td>
    <td>{krbtgt_badge}{_krbtgt_risk_html}</td></tr>
<tr><td class="mono dim">AD Recycle Bin</td>
    <td class="mono {"ok" if recycle_bin=="Enabled" else "bad"}">{h(recycle_bin)}</td>
    <td>{"" if recycle_bin=="Enabled" else "<span class='tag tag-warn'>Oggetti eliminati non recuperabili</span>"}</td></tr>
<tr><td class="mono dim">{L("Utenti Privilegiati", "Privileged Users")} Disabilitati</td>
    <td class="mono {"warn" if total_disabled>0 else "ok"}">{total_disabled}</td><td></td></tr>
</tbody></table></div>

{_catalog_corrupt_html}
<div class="stitle">{L("Utenti Privilegiati", "Privileged Users")} — {total_priv} account ({len(PRIV_GROUPS)} gruppi monitorati)</div>
<div class="card"><table>
<thead><tr><th>sAMAccountName</th><th>{L("Gruppi (Direct/Nested)", "Groups (Direct/Nested)")}</th><th>pwdLastSet</th><th>lastLogon</th><th>Flags</th><th>SID</th></tr></thead>
<tbody>{rows_priv if rows_priv else _no_priv_row}</tbody>
</table></div>

<div class="stitle">{L("ACL Domain Root — ACE con diritti pericolosi", "ACL Domain Root — Dangerous ACE rights")}</div>
<div class="card"><table>
<thead><tr><th>Trustee</th><th>SID</th><th>{L("Diritti", "Rights")}</th><th>Access Mask</th><th>Flags</th></tr></thead>
<tbody>{ace_rows(root_aces)}</tbody>
</table></div>

<div class="stitle">{L("ACL OU Domain Controllers — ACE con diritti pericolosi", "ACL OU Domain Controllers — Dangerous ACE rights")}</div>
<div class="card"><table>
<thead><tr><th>Trustee</th><th>SID</th><th>{L("Diritti", "Rights")}</th><th>Access Mask</th><th>Flags</th></tr></thead>
<tbody>{ace_rows(dc_ou_aces)}</tbody>
</table></div>

<div class="stitle">Group Policy Objects — {total_gpo} {L("GPO totali", "total GPOs")}</div>
<div class="card"><table>
<thead><tr><th>Display Name</th><th>{L("Ultima Modifica", "Last Modified")}</th><th>Path SYSVOL</th></tr></thead>
<tbody>{rows_gpo if rows_gpo else _no_gpo_row}</tbody>
</table></div>

<div class="stitle">{L("Computer nel Dominio", "Domain Computers")} — {total_comp} oggetti</div>
<div class="card"><table>
<thead><tr><th>Computer Name</th><th>Operating System</th><th>Distinguished Name</th></tr></thead>
<tbody>{rows_comp if rows_comp else _no_comp_row}</tbody>
</table></div>

</main>
<footer>
  <span>PAD Offline</span>
  <span>fiuto.sh — DFIR Toolkit</span>
</footer>
</body></html>"""

try:
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html_out)
except Exception as e:
    sys.stderr.write(f"ERROR writing report: {e}\n"); sys.exit(1)

# Output stats per bash (su stderr per non interferire)
sys.stderr.write(f"STATS:{total_priv}:{total_stale}:{krbtgt_days or 'N/A'}:{gpo_recent}:{total_comp}\n")
PYEOF

    if [[ -f "$REPORT_HTML" ]]; then
        # Copia il pyerr.log nel report dir per debug postumo
        if [[ -f "$PYERR_FILE" ]]; then
            cp "$PYERR_FILE" "$(dirname "$REPORT_HTML")/impacket_debug.log" 2>/dev/null || true
        fi

        # Avviso catalogo corrotto
        if grep -q '^CATALOG_CORRUPT:1' "$PYERR_FILE" 2>/dev/null; then
            echo ""
            warn "$(L "NTDS.dit con catalogo ESE parzialmente corrotto — dati utente non disponibili." "NTDS.dit with partially corrupted ESE catalog — user data unavailable.")"
            info "$(L "Le colonne sAMAccountName/samAccountType sono su pagine di catalogo non leggibili." "Columns sAMAccountName/samAccountType are on unreadable catalog pages.")"
            echo ""
            echo -e "  ${BOLD}${YELLOW}Recupero necessario — scegli una delle opzioni:${RESET}"
            echo -e "  ${CYAN}1) Soft-recovery Linux${RESET} (richiede i log .jrs nella stessa dir di ntds.dit)"
            echo -e "     Copia ntds.dit + edb*.jrs + edb.chk in una cartella, poi rilancia."
            echo ""
            echo -e "  ${CYAN}2) esentutl su Windows:${RESET}"
            echo -e "     ${DIM}esentutl /r edb /l C:\\Windows\\NTDS /s C:\\Windows\\NTDS${RESET}"
            echo ""
            echo -e "  ${CYAN}3) IFM snapshot (DC attivo):${RESET}"
            echo -e "     ${DIM}ntdsutil \"activate instance ntds\" \"ifm\" \"create full C:\\ifm\" quit quit${RESET}"
            echo ""
            info "$(L "Il report HTML contiene la guida completa al recupero." "The HTML report contains the complete recovery guide.")"
        fi

        info "$(L "Report generato:" "Report generated:") ${BOLD}$REPORT_HTML${RESET}"
        open_report_prompt "$REPORT_HTML"
    else
        local _PYERR _PYINFO
        _PYERR=$(grep -v '^STATS:' "$PYERR_FILE" 2>/dev/null | grep -iv '^INFO:' | head -10 || true)
        _PYINFO=$(grep -i '^INFO:' "$PYERR_FILE" 2>/dev/null | head -10 || true)
        [[ -n "$_PYINFO" ]] && info "$(L "Debug ESE:" "ESE debug:") $_PYINFO"
        [[ -n "$_PYERR" ]] && warn "$(L "Errore Python:" "Python error:") $_PYERR"

        # Diagnosi specifica per errori noti
        if echo "${_PYERR}" | grep -qiE 'Errno 35|EAGAIN|temporarily unavailable|Resource temporarily'; then
            err "$(L "ntds.dit è bloccato dal servizio Active Directory (Errno 35 / EAGAIN)." "ntds.dit is locked by Active Directory service (Errno 35 / EAGAIN).")"
            echo ""
            echo -e "  ${YELLOW}$(L "Il client SMB rispetta il lock esclusivo tenuto da lsass/ntds.exe." "The SMB client respects the exclusive lock held by lsass/ntds.exe.")${RESET}"
            echo -e "  ${YELLOW}$(L "Il file non è apribile nemmeno in lettura finché il DC è attivo." "The file cannot be opened even for reading while the DC is active.")${RESET}"
            echo ""
            echo -e "  ${BOLD}$(L "Soluzioni per DC live:" "Solutions for live DC:")${RESET}"
            echo -e "  ${CYAN}$(L "1) VSS snapshot (consigliato):" "1) VSS snapshot (recommended):")${RESET}"
            echo -e "     $(L "Sul DC (PowerShell admin):" "On the DC (PowerShell admin):")"
            echo -e "     ${DIM}ntdsutil \"activate instance ntds\" \"ifm\" \"create full C:\\\\ifm_snapshot\" quit quit${RESET}"
            echo -e "     $(L "Poi monta/copia la cartella C:\\\\ifm_snapshot e rilancia su quel path." "Then mount/copy the C:\\\\ifm_snapshot folder and relaunch on that path.")"
            echo ""
            echo -e "  ${CYAN}$(L "2) impacket secretsdump remoto (solo hash NTLM, senza analisi ACL/GPO):" "2) remote impacket secretsdump (NTLM hashes only, no ACL/GPO analysis):")${RESET}"
            echo -e "     ${DIM}impacket-secretsdump -just-dc-ntlm DOMINIO/utente:password@IP_DC${RESET}"
            echo ""
            echo -e "  ${CYAN}$(L "3) Se hai già un'immagine disco (VHDX/E01) del DC:" "3) If you already have a disk image (VHDX/E01) of the DC:")${RESET}"
            echo -e "     $(L "Monta offline con" "Mount offline with") ${BOLD}./mount_image.sh${RESET} $(L "e rilancia il modulo su quel mount point." "and relaunch the module on that mount point.")"
        elif echo "${_PYERR}" | grep -qiE 'dirty|state|DirtyShutdown|Unknown state'; then
            err "$(L "Generazione report fallita — database ESE in stato dirty non recuperabile." "Report generation failed — ESE database in unrecoverable dirty state.")"
            info "$(L "Suggerimento: copia il file e prova 'esentutl /r edb /l .' in un ambiente Windows" "Hint: copy the file and try 'esentutl /r edb /l .' in a Windows environment")"
        else
            err "$(L "Generazione report fallita. Verifica che ntds.dit sia leggibile e non corrotto." "Report generation failed. Verify ntds.dit is readable and not corrupted.")"
            info "$(L "Suggerimento: copia il file e prova 'esentutl /r edb' in un ambiente Windows" "Hint: copy the file and try 'esentutl /r edb' in a Windows environment")"
        fi
    fi
    rm -rf "$TMP_PAD_DIR" 2>/dev/null || true
}
