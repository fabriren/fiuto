#!/opt/homebrew/bin/bash
# ================================================================
#  fiuto.sh (macOS)  –  v2.3
#  Toolkit DFIR unificato per analisi offline di dischi Windows/Linux/macOS
#
#  Uso:
#    ./fiuto.sh                        # menu interattivo
#    ./fiuto.sh /mnt/disk              # specifica la root (Windows/Linux/macOS)
#    ./fiuto.sh /mnt/disk --all        # esegui tutto
#    ./fiuto.sh /mnt/disk --module 3   # modulo specifico
#    ./fiuto.sh --image disco.E01 --partition 2 --all   # da immagine forense
# ================================================================
#
#  ATTENZIONE — QUESTO FILE E' GENERATO. NON MODIFICARLO A MANO.
#
#  Il sorgente sta in src/, diviso per sistema operativo:
#    src/lib/                funzioni condivise (core, fs, report, registro...)
#    src/modules/win/        moduli Windows
#    src/modules/linux/      moduli Linux
#    src/modules/macos/      moduli macOS
#    src/modules/xplat/      moduli cross-OS
#
#  QUESTA E' LA BUILD PER macOS: gira SU un Mac come workstation di analisi.
#  Si genera con ./buildMac.sh, che concatena gli stessi sorgenti della build
#  Linux piu' src/compat/macos.sh, lo strato che traduce gli strumenti GNU
#  verso l'userland BSD. I sorgenti sono gli stessi byte per byte: la
#  differenza fra le due piattaforme sta tutta in quel file.
#
#  Richiede bash 4+ (macOS ne spedisce ancora la 3.2):
#    brew install bash
#
#  Dopo aver modificato un sorgente rigenera questo file con:
#    ./buildMac.sh
#
#  La CI verifica con ./build.sh --check che i due siano allineati: una
#  modifica fatta qui e non nei sorgenti verrebbe persa al build successivo.
# ================================================================

set -uo pipefail

# Versione del tool. Unica fonte: la usano il banner e il manifesto di catena
# di custodia, che non devono poter divergere.
FIUTO_VERSION="2.3"

# -- Cleanup globale file temporanei (signal-safe) ----------------
# La trap EXIT viene invocata su ritorno normale, SIGINT (Ctrl+C),
# SIGTERM, SIGHUP e errori.
# SIGKILL (kill -9) non e' intercettabile per design del kernel.
declare -a _GLOBAL_TMP_FILES=()
