
# Controlla se regipy è disponibile tramite il Python rilevato all'avvio ($PY3)
check_regipy() {
    "$PY3" -c "import regipy" 2>/dev/null
}

# ----------------------------------------------------------------
#  Replay dei transaction log del registro (.LOG1 / .LOG2)
#
#  Windows non scrive immediatamente nell'hive primario: accoda le modifiche
#  nei transaction log e le consolida solo a uno smontaggio pulito. Un hive
#  acquisito da una macchina spenta a caldo, da un'immagine o da uno snapshot
#  e' quindi quasi sempre "dirty": le scritture piu' recenti — tipicamente
#  proprio quelle dell'attaccante — esistono SOLO nei .LOG1/.LOG2.
#
#  Parsare l'hive grezzo le perde senza emettere alcun errore: il modulo
#  produce un report che sembra completo ma non lo e'. Qui i log vengono
#  riapplicati su una COPIA in area temporanea; il volume di evidenza resta
#  intatto e in sola lettura.
#
#  Nota: regipy scriverebbe di default l'hive ricostruito accanto
#  all'originale (<hive>.restored), cioe' SUL volume di evidenza. Il percorso
#  di destinazione viene quindi sempre passato esplicitamente.
# ----------------------------------------------------------------

# Area temporanea per gli hive ricostruiti.
#
# recover_hive viene invocata quasi sempre dentro una command substitution,
# quindi gira in una subshell: variabili globali e array modificati li' non
# risalgono al processo padre. Il percorso deve percio' essere DETERMINISTICO
# e non memorizzato in una variabile, altrimenti ogni chiamata creerebbe una
# nuova directory — rifacendo il replay ogni volta e, peggio, lasciando in
# giro copie di hive (che contengono credenziali) mai ripulite.
#
# $$ resta il PID della shell principale anche nelle subshell, quindi
# identifica la sessione in modo stabile.
_hive_replay_tmpdir() {
    local D="${TMPDIR:-/tmp}/fiuto_hives_$$"
    if [[ ! -d "$D" ]]; then
        mkdir -p "$D" 2>/dev/null || return 1
        chmod 700 "$D" 2>/dev/null || true   # gli hive ricostruiti contengono credenziali
    fi
    echo "$D"
}

# Percorso deterministico dell'hive ricostruito a partire dall'originale.
# Fa anche da cache: se il file esiste gia', il replay e' gia' stato fatto.
_hive_recovered_path() {
    local ORIG="$1" TMPD
    TMPD=$(_hive_replay_tmpdir) || return 1
    printf '%s/%s_%s.recovered' "$TMPD" \
        "$(printf '%s' "$ORIG" | sha256sum | cut -c1-16)" "$(basename "$ORIG")"
}

# Annota l'esito del replay su file: gli array in subshell non sopravvivono.
_hive_replay_note() {
    local TMPD; TMPD=$(_hive_replay_tmpdir) || return 0
    printf '%s\t%s\t%s\n' "$1" "$2" "$3" >> "${TMPD}/replay.log" 2>/dev/null || true
}

# Righe di esito del replay raccolte finora (una per hive, formato TSV).
hive_replay_report() {
    local TMPD="${TMPDIR:-/tmp}/fiuto_hives_$$"
    [[ -f "${TMPD}/replay.log" ]] && sort -u "${TMPD}/replay.log"
}

# recover_hive <path_hive>
# Stampa il percorso dell'hive da parsare: la copia ricostruita se erano
# presenti transaction log applicabili, altrimenti l'originale.
# Non fallisce mai: in caso di problema ripiega sull'originale e lo annota.
recover_hive() {
    local ORIG="$1"
    [[ -n "$ORIG" && -f "$ORIG" ]] || { echo ""; return; }
    evidence_note "$ORIG" "hive di registro"
    [[ "$HIVE_REPLAY" == "true" ]] || { echo "$ORIG"; return; }

    local OUT; OUT=$(_hive_recovered_path "$ORIG") || { echo "$ORIG"; return; }

    # Cache su disco: sopravvive alle subshell delle command substitution.
    [[ -s "$OUT" ]] && { echo "$OUT"; return; }
    # Marcatore "gia' valutato, si usa l'originale" (hive pulito o replay fallito):
    # evita di ritentare a ogni modulo.
    [[ -f "${OUT}.skip" ]] && { echo "$ORIG"; return; }

    # Con --jobs piu' moduli possono chiedere lo stesso hive nello stesso
    # istante. Senza lock il secondo leggerebbe la copia ricostruita mentre il
    # primo la sta ancora scrivendo: un hive troncato non da' errore, da'
    # risultati parziali — che e' peggio. mkdir e' atomico anche su NFS.
    local LOCK="${OUT}.lock" _held=false _tries=0
    while true; do
        if mkdir "$LOCK" 2>/dev/null; then _held=true; break; fi
        _tries=$((_tries + 1))
        # Un lock orfano (processo ucciso) non deve bloccare la sessione: dopo
        # due minuti si procede comunque, nel caso peggiore rifacendo il lavoro.
        [[ $_tries -gt 240 ]] && break
        sleep 0.5
        # Nel frattempo puo' aver finito qualcun altro.
        [[ -s "$OUT" ]] && { echo "$OUT"; return; }
        [[ -f "${OUT}.skip" ]] && { echo "$ORIG"; return; }
    done
    # Rilascia il lock solo chi lo detiene davvero: dopo un timeout il lock e'
    # di un altro processo, e rimuoverlo aprirebbe la corsa che il lock evita.
    _unlock() { [[ "$_held" == "true" ]] && rmdir "$LOCK" 2>/dev/null; return 0; }

    local DIR BASE LOG1 LOG2
    DIR=$(dirname "$ORIG")
    BASE=$(basename "$ORIG")
    LOG1=$(find "$DIR" -maxdepth 1 -iname "${BASE}.LOG1" -type f 2>/dev/null | head -1)
    LOG2=$(find "$DIR" -maxdepth 1 -iname "${BASE}.LOG2" -type f 2>/dev/null | head -1)

    # Nessun log, o log vuoti: l'hive e' gia' consolidato.
    if [[ ( -z "$LOG1" || ! -s "$LOG1" ) && ( -z "$LOG2" || ! -s "$LOG2" ) ]]; then
        : > "${OUT}.skip"
        _hive_replay_note "clean" "$BASE" "$(L "nessun transaction log da applicare" "no transaction log to apply")"
        _unlock; echo "$ORIG"
        return
    fi

    if ! check_regipy; then
        # L'avviso va emesso una sola volta per sessione: il marcatore sta su
        # disco perche' una variabile non sopravvivrebbe alla subshell.
        local WARNED; WARNED="$(_hive_replay_tmpdir)/.regipy_warned"
        if [[ ! -f "$WARNED" ]]; then
            : > "$WARNED"
            warn "$(L "regipy assente: i transaction log del registro non verranno applicati (dati recenti potenzialmente mancanti)." \
                     "regipy missing: registry transaction logs will not be applied (recent data potentially missing).")" >&2
        fi
        : > "${OUT}.skip"
        _hive_replay_note "skipped" "$BASE" "$(L "regipy non disponibile" "regipy unavailable")"
        _unlock; echo "$ORIG"
        return
    fi

    local RES
    RES=$("$PY3" - "$ORIG" "$OUT" "${LOG1:-}" "${LOG2:-}" << 'PYEOF' 2>/dev/null
import sys, os

hive_path, out_path = sys.argv[1], sys.argv[2]
log1 = sys.argv[3] if len(sys.argv) > 3 and sys.argv[3] else None
log2 = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] else None

# Se esiste solo il secondario, va passato come primario: regipy richiede
# comunque un primary_log_path valido.
if not log1 and log2:
    log1, log2 = log2, None

try:
    from regipy.recovery import apply_transaction_logs
    restored, dirty = apply_transaction_logs(
        hive_path, log1, secondary_log_path=log2, restored_hive_path=out_path
    )
    if restored and os.path.isfile(restored) and os.path.getsize(restored) > 0:
        print(f"OK\t{dirty}")
    else:
        print("FAIL\tnessun hive ricostruito prodotto")
except Exception as exc:
    print(f"FAIL\t{exc}")
PYEOF
    )

    local STATUS DETAIL
    STATUS="${RES%%	*}"
    DETAIL="${RES#*	}"

    if [[ "$STATUS" == "OK" && -s "$OUT" ]]; then
        _hive_replay_note "recovered" "$BASE" "${DETAIL} $(L "pagine dirty riapplicate" "dirty pages replayed")"
        info "$(L "Transaction log applicati a" "Transaction logs applied to") ${BOLD}${BASE}${RESET} - ${DETAIL} $(L "pagine dirty" "dirty pages")" >&2
        log_msg "[HIVE] replay OK: $ORIG -> $OUT (${DETAIL} dirty pages)"
        _unlock; echo "$OUT"
        return
    fi

    # Fallback non distruttivo: si continua sull'hive originale.
    rm -f "$OUT" 2>/dev/null
    : > "${OUT}.skip"
    _hive_replay_note "failed" "$BASE" "$DETAIL"
    warn "$(L "Replay dei transaction log fallito per" "Transaction log replay failed for") ${BASE}: ${DETAIL}" >&2
    log_msg "[HIVE] replay FAILED: $ORIG - $DETAIL"
    _unlock; echo "$ORIG"
}

# Torna il percorso di un hive di sistema, con i transaction log gia' applicati.
get_hive() {
    local NAME="$1"   # SOFTWARE, SYSTEM, SECURITY, SAM
    local FOUND=""
    local PATH1="$WIN_ROOT/Windows/System32/config/${NAME}"
    if [[ -f "$PATH1" ]]; then
        FOUND="$PATH1"
    else
        local PATH2
        PATH2=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config")
        [[ -n "$PATH2" ]] && FOUND=$(ci_find_file "$PATH2" "$NAME")
    fi
    [[ -n "$FOUND" ]] && recover_hive "$FOUND"
}

# Torna l'hive di un utente (NTUSER.DAT, UsrClass.dat) con i log applicati.
# $1 = directory da cui cercare · $2 = nome hive
get_user_hive() {
    local FOUND; FOUND=$(ci_find_file "$1" "$2")
    [[ -n "$FOUND" ]] && recover_hive "$FOUND"
}

# Avvia un parser regipy su un hive e una chiave
regipy_query() {
    local HIVE="$1"
    local KEY="$2"
    "$PY3" - "$HIVE" "$KEY" << 'PYEOF'
import sys
try:
    from regipy.registry import RegistryHive
    hive = RegistryHive(sys.argv[1])
    key = hive.get_key(sys.argv[2])
    for v in key.get_values():
        print(f"{v.name}\t{v.value}")
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
PYEOF
}
