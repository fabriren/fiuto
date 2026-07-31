#!/usr/bin/env bash
#
# Entrypoint dell'immagine FIUTO.
#
# Non si limita a lanciare il tool: verifica le due condizioni che rendono
# valida un'analisi dentro un container, e che sono facilissime da sbagliare
# scrivendo un docker run a mano.
#
#   1. L'EVIDENZA DEVE ESSERE IN SOLA LETTURA. Dimenticare :ro non produce
#      alcun errore: FIUTO non scrive sul volume, ma un mount scrivibile
#      significa che qualunque altro processo del container potrebbe, e in
#      perizia la domanda "il reperto poteva essere modificato?" ha una sola
#      risposta accettabile.
#   2. I REPORT DEVONO RESTARE DELL'UTENTE. Senza --user il container gira da
#      root e i file finiscono di root sull'host: l'analista non li puo'
#      archiviare senza sudo, e il manifesto di custodia registra un operatore
#      che non e' chi ha eseguito l'analisi.
#
# Nessuna delle due e' bloccante — chi sa cosa sta facendo deve poter
# procedere — ma nessuna delle due resta implicita.

set -uo pipefail

_y() { printf '\033[1;33m  [!] %s\033[0m\n' "$1" >&2; }
_c() { printf '\033[0;36m  [i] %s\033[0m\n' "$1" >&2; }

# --deps: inventario delle dipendenze presenti nell'immagine. Serve a poter
# dichiarare, in una relazione, con quali parser e' stata fatta l'analisi.
if [[ "${1:-}" == "--deps" ]]; then
    python3 - << 'PYEOF'
import importlib
import platform

MODS = [
    ('regipy', 'hive di registro, replay transaction log'),
    ('Evtx', 'log eventi .evtx, motore Sigma'),
    ('pyesedb', 'ESE: WebCacheV01, Windows.edb, SRUM'),
    ('pypff', 'archivi PST/OST'),
    ('yara', 'regole YARA'),
    ('yaml', 'regole Sigma'),
    ('impacket', 'SAM, NTDS, segreti LSA'),
    ('mft', 'MFT'),
    ('snappy', 'LevelDB delle app Electron'),
]
print("FIUTO - dipendenze dell'immagine")
print('Python %s\n' % platform.python_version())
for name, what in MODS:
    try:
        m = importlib.import_module(name)
        ver = getattr(m, '__version__', '') or ''
    except Exception as exc:
        print('  MANCANTE  %-10s %-42s (%s)' % (name, what, str(exc)[:40]))
        continue
    print('  presente  %-10s %-42s %s' % (name, what, ver))
PYEOF
    echo ""
    echo "  Binari di sistema:"
    for b in file strings sqlite3 ntfscat ewfmount flock fls icat; do
        if command -v "$b" > /dev/null 2>&1; then
            printf '    presente  %s\n' "$b"
        else
            printf '    MANCANTE  %s\n' "$b"
        fi
    done
    exit 0
fi

# --- verifica 1: l'evidenza e' montata in sola lettura ---------------------
if [[ -d /evidence ]]; then
    if [[ -w /evidence ]]; then
        _y "/evidence e' montato in SCRITTURA."
        _y "Rimontalo con :ro   ->   -v /mnt/disk:/evidence:ro"
        _y "FIUTO non scrive sul volume, ma un reperto scrivibile non e' un reperto."
    else
        _c "/evidence in sola lettura: corretto."
    fi
fi

# --- verifica 2: i report non finiscono di root ----------------------------
if [[ "$(id -u)" -eq 0 ]]; then
    _y "Il container gira come root: i report sull'host risulteranno di root."
    _y "Aggiungi --user \"\$(id -u):\$(id -g)\" al docker run."
elif [[ -d /report && ! -w /report ]]; then
    _y "/report non e' scrivibile dall'utente $(id -u): i report non verranno salvati."
    _y "Monta una directory scrivibile   ->   -v \"\$PWD/report\":/report"
fi

# HOME puo' non esistere per un --user arbitrario e alcune librerie ci scrivono
# cache: senza, fallirebbero con un errore che non c'entra niente con l'analisi.
export HOME="${HOME:-/tmp}"
[[ -w "$HOME" ]] || export HOME=/tmp

exec /opt/fiuto/fiuto.sh "$@"
