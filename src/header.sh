#!/usr/bin/env bash
# ================================================================
#  fiuto.sh  –  v2.3
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
#  Dopo aver modificato un sorgente rigenera questo file con:
#    ./build.sh
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
