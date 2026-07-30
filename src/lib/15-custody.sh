
# ================================================================
#  CATENA DI CUSTODIA
#
#  Un report forense vale quanto la tracciabilita' di cio' su cui si basa. Fino
#  alla 2.1 FIUTO calcolava qualche hash sparso in pochi moduli: non c'era modo
#  di rispondere a "quali file sono stati letti, in che stato erano, e i report
#  allegati sono quelli prodotti allora?".
#
#  Il manifesto risponde a tutte e tre. E' scritto a fine sessione in
#  evidence_manifest.json e contiene: identita' del tool e comando eseguito,
#  operatore e host di analisi, finestra temporale, volume analizzato, elenco
#  dei file di evidenza con dimensione, data e SHA256, ed elenco dei report
#  generati con il loro SHA256.
#
#  Aggancio: invece di modificare 90 moduli, si annotano i file nelle funzioni
#  che TUTTI attraversano per accedere a un'evidenza — ci_find_file,
#  query_sqlite, read_plist, recover_hive. La copertura e' quindi automatica
#  anche per i moduli futuri.
# ================================================================

CUSTODY=true              # --no-custody per disattivare
CUSTODY_HASH=true         # --no-hash: annota i file senza calcolarne l'hash
CUSTODY_HASH_LIMIT_MB=1024   # oltre questa soglia niente hash (--hash-limit)

# Directory di sessione. Come per il replay degli hive, il percorso e'
# deterministico e non tenuto in una variabile: le annotazioni arrivano da
# subshell (le funzioni sono usate in command substitution).
_custody_dir() {
    local D="${TMPDIR:-/tmp}/fiuto_custody_$$"
    [[ -d "$D" ]] || { mkdir -p "$D" 2>/dev/null || return 1; chmod 700 "$D" 2>/dev/null || true; }
    echo "$D"
}

# evidence_note <percorso> [ruolo]
# Annota un file come evidenza consultata. Non calcola nulla: l'hash arriva
# alla fine, una volta sola per file, per non pagarlo a ogni accesso.
evidence_note() {
    [[ "$CUSTODY" == "true" ]] || return 0
    local P="$1" ROLE="${2:-letto}"
    [[ -n "$P" && -f "$P" ]] || return 0
    # I file di lavoro di FIUTO non sono evidenza. L'esclusione deve pero'
    # essere precisa: scartare tutto cio' che sta sotto /tmp escluderebbe le
    # evidenze quando il volume e' montato li', cosa del tutto legittima.
    case "$P" in
        */fiuto_custody_*/*|*/fiuto_hives_*/*) return 0 ;;
    esac
    local D; D=$(_custody_dir) || return 0
    printf '%s\t%s\n' "$ROLE" "$P" >> "${D}/evidence.tsv" 2>/dev/null || true
}

# Quante evidenze sono state annotate finora.
custody_count() {
    local D="${TMPDIR:-/tmp}/fiuto_custody_$$"
    [[ -f "${D}/evidence.tsv" ]] && cut -f2 "${D}/evidence.tsv" | sort -u | wc -l || echo 0
}

# Scrive evidence_manifest.json in REPORT_BASE_DIR.
# Va chiamata a fine sessione, PRIMA della pulizia dei file temporanei.
write_evidence_manifest() {
    [[ "$CUSTODY" == "true" ]] || return 0
    [[ -n "$REPORT_BASE_DIR" && -d "$REPORT_BASE_DIR" ]] || return 0
    local D="${TMPDIR:-/tmp}/fiuto_custody_$$"
    local EV="${D}/evidence.tsv"
    [[ -s "$EV" ]] || return 0

    local OUT="${REPORT_BASE_DIR}/evidence_manifest.json"
    # Lo script legge il contesto da os.environ: va esportato esplicitamente.
    export FIUTO_VERSION PY3_VERSION REPORT_BASE_DIR WIN_ROOT OS_TYPE HOST_NAME
    export CUSTODY_HASH CUSTODY_HASH_LIMIT_MB CUSTODY_START_UTC
    export CUSTODY_CMDLINE CUSTODY_OPERATOR CUSTODY_HOST
    local REPORTS; REPORTS=$(mktemp)
    printf '%s\n' "${GENERATED_REPORTS[@]:-}" > "$REPORTS"
    local REPLAY; REPLAY=$(mktemp)
    hive_replay_report > "$REPLAY" 2>/dev/null || true

    "$PY3" - "$EV" "$REPORTS" "$REPLAY" "$OUT" << 'PYEOF' 2>/dev/null
import sys, os, json, hashlib, datetime

ev_path, rep_path, replay_path, out_path = sys.argv[1:5]

env = os.environ
do_hash = env.get('CUSTODY_HASH', 'true') == 'true'
try:
    limit = int(env.get('CUSTODY_HASH_LIMIT_MB', '1024')) * 1024 * 1024
except ValueError:
    limit = 1024 * 1024 * 1024


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def describe(path, roles):
    rec = {"path": path, "roles": sorted(roles)}
    try:
        st = os.stat(path)
        rec["size_bytes"] = st.st_size
        rec["mtime_utc"] = datetime.datetime.utcfromtimestamp(st.st_mtime).strftime('%Y-%m-%dT%H:%M:%SZ')
    except Exception as exc:
        rec["error"] = str(exc)
        return rec
    if not do_hash:
        rec["sha256"] = None
        rec["sha256_omitted"] = "hashing disabilitato (--no-hash)"
    elif rec["size_bytes"] > limit:
        # Su un pagefile o un $MFT da diversi GB l'hash costa quanto tutto il
        # resto dell'analisi: si dichiara perche' manca invece di ometterlo.
        rec["sha256"] = None
        rec["sha256_omitted"] = f"file oltre la soglia di {limit // (1024*1024)} MB"
    else:
        try:
            rec["sha256"] = sha256(path)
        except Exception as exc:
            rec["sha256"] = None
            rec["sha256_omitted"] = f"lettura fallita: {exc}"
    return rec


# --- evidenze -------------------------------------------------------------
items = {}
with open(ev_path, encoding='utf-8', errors='replace') as fh:
    for line in fh:
        parts = line.rstrip('\n').split('\t', 1)
        if len(parts) != 2:
            continue
        role, path = parts
        items.setdefault(path, set()).add(role)

evidence = [describe(p, r) for p, r in sorted(items.items())]

# --- report generati ------------------------------------------------------
reports = []
seen = set()
with open(rep_path, encoding='utf-8', errors='replace') as fh:
    for line in fh:
        p = line.strip()
        if not p or p in seen or not os.path.isfile(p):
            continue
        seen.add(p)
        rec = {"path": p}
        try:
            rec["size_bytes"] = os.path.getsize(p)
            rec["sha256"] = sha256(p)
        except Exception as exc:
            rec["error"] = str(exc)
        reports.append(rec)

# --- replay dei transaction log ------------------------------------------
replay = []
try:
    with open(replay_path, encoding='utf-8', errors='replace') as fh:
        for line in fh:
            f = line.rstrip('\n').split('\t')
            if len(f) >= 3:
                replay.append({"status": f[0], "hive": f[1], "detail": f[2]})
except Exception:
    pass

manifest = {
    "manifest_version": 1,
    "tool": {
        "name": "FIUTO",
        "version": env.get('FIUTO_VERSION', ''),
        "python": env.get('PY3_VERSION', ''),
    },
    "session": {
        "started_utc": env.get('CUSTODY_START_UTC', ''),
        "completed_utc": datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
        "command": env.get('CUSTODY_CMDLINE', ''),
        "operator": env.get('CUSTODY_OPERATOR', ''),
        "analysis_host": env.get('CUSTODY_HOST', ''),
        "report_directory": env.get('REPORT_BASE_DIR', ''),
    },
    "subject": {
        "volume_root": env.get('WIN_ROOT', ''),
        "detected_os": env.get('OS_TYPE', ''),
        "hostname_from_artefacts": env.get('HOST_NAME', ''),
    },
    "integrity_policy": {
        "algorithm": "SHA-256",
        "hashing_enabled": do_hash,
        "size_limit_bytes": limit,
        "note": ("Gli hash sono calcolati sui file COSI' COME LETTI dal volume montato. "
                 "FIUTO non modifica il volume: monta sempre in sola lettura e verifica "
                 "questi valori contro l'immagine di acquisizione."),
    },
    "registry_transaction_log_replay": replay,
    "evidence_files": evidence,
    "generated_reports": reports,
    "counts": {
        "evidence_files": len(evidence),
        "evidence_hashed": sum(1 for e in evidence if e.get("sha256")),
        "generated_reports": len(reports),
    },
}

with open(out_path, 'w', encoding='utf-8') as fh:
    json.dump(manifest, fh, indent=2, ensure_ascii=False)
    fh.write('\n')
PYEOF

    rm -f "$REPORTS" "$REPLAY"
    if [[ -s "$OUT" ]]; then
        local NEV NRE
        NEV=$("$PY3" -c "import json,sys; print(json.load(open(sys.argv[1]))['counts']['evidence_files'])" "$OUT" 2>/dev/null || echo "?")
        NRE=$("$PY3" -c "import json,sys; print(json.load(open(sys.argv[1]))['counts']['generated_reports'])" "$OUT" 2>/dev/null || echo "?")
        echo ""
        ok "$(L "Manifesto catena di custodia:" "Chain-of-custody manifest:") ${BOLD}${OUT}"
        info "$(L "Evidenze tracciate:" "Evidence files tracked:") ${BOLD}${NEV}${RESET}  ·  $(L "report:" "reports:") ${BOLD}${NRE}"
    fi
}
