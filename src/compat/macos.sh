
# ================================================================
#  STRATO DI COMPATIBILITA' macOS
#
#  Questo file esiste in UNA sola build, quella prodotta da ./buildMac.sh, e
#  viene inserito subito dopo l'intestazione: tutto cio' che sta sotto vede le
#  definizioni che seguono.
#
#  PERCHE' UNO STRATO E NON UN FORK. I 117 sorgenti sotto src/ restano
#  identici byte per byte fra le due build. Duplicare l'albero per macOS
#  avrebbe significato correggere ogni bug due volte, e le due copie sarebbero
#  divergute al primo modulo nuovo. Qui la differenza fra le due piattaforme e'
#  raccolta in un punto solo, e si legge tutta insieme.
#
#  COME FUNZIONA. I moduli invocano gli strumenti GNU per nome (`stat -c %y`,
#  `sha256sum`, `numfmt`). Su macOS quegli strumenti non esistono o hanno
#  un'altra sintassi. Invece di modificare 52 punti di chiamata, qui si
#  definiscono funzioni bash con lo stesso nome che traducono verso l'userland
#  BSD. Una funzione ha la precedenza sull'eseguibile, quindi la sostituzione
#  e' trasparente; `command stat` resta la via per raggiungere il binario vero.
#
#  COSA NON PUO' FARE. Le funzionalita' che dipendono dal kernel Linux non si
#  traducono: i loop device, i namespace, /proc. Dove serve, il comportamento
#  viene ridefinito per intero (vedi _find_user_cwd e image_open) oppure il
#  modulo dichiara di non poter procedere, come gia' fa quando manca una
#  libreria. Nessuna funzione finge di funzionare.
# ================================================================

# --- bash: macOS ne spedisce ancora la 3.2 del 2007 ------------------------
# FIUTO usa array associativi, ${var,,}, mapfile, declare -n e wait -n, tutte
# cose che la 3.2 non ha. Fallire subito con l'istruzione per rimediare e'
# meglio di un errore di sintassi a meta' di un modulo.
if [[ -z "${BASH_VERSINFO[0]:-}" || "${BASH_VERSINFO[0]}" -lt 4 ]]; then
    echo "FIUTO richiede bash 4 o superiore; questo e' bash ${BASH_VERSION:-sconosciuto}." >&2
    echo "macOS spedisce ancora la 3.2. Installa una versione recente:" >&2
    echo "    brew install bash" >&2
    echo "poi rilancia con /opt/homebrew/bin/bash fiuto.sh (o ./fiuto.sh, che la usa gia')." >&2
    exit 1
fi

# --- stat ------------------------------------------------------------------
# GNU: stat -c FORMATO FILE      BSD: stat -f FORMATO [-t FORMATO_DATA] FILE
#
# L'ordine delle sostituzioni non e' arbitrario: GNU %s (dimensione) diventa
# BSD %z, ma GNU %z (ctime) diventa BSD %Sc. Tradurre %s per primo produrrebbe
# un %z che la regola successiva convertirebbe una seconda volta, restituendo
# la data al posto della dimensione. %z si traduce quindi PRIMA di %s.
stat() {
    local fmt="" ; local -a rest=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -c)          fmt="${2:-}"; shift 2 ;;
            -c*)         fmt="${1#-c}"; shift ;;
            --format=*)  fmt="${1#--format=}"; shift ;;
            --printf=*)  fmt="${1#--printf=}"; shift ;;
            *)           rest+=("$1"); shift ;;
        esac
    done
    if [[ -z "$fmt" ]]; then
        command stat "${rest[@]}"
        return $?
    fi
    local bsd
    bsd=$(printf '%s' "$fmt" | sed -e 's/%y/%Sm/g' \
                                   -e 's/%z/%Sc/g' \
                                   -e 's/%x/%Sa/g' \
                                   -e 's/%Y/%m/g'  \
                                   -e 's/%W/%B/g'  \
                                   -e 's/%X/%a/g'  \
                                   -e 's/%a/%Lp/g' \
                                   -e 's/%U/%Su/g' \
                                   -e 's/%G/%Sg/g' \
                                   -e 's/%F/%HT/g' \
                                   -e 's/%n/%N/g'  \
                                   -e 's/%s/%z/g')
    command stat -f "$bsd" -t '%Y-%m-%d %H:%M:%S' "${rest[@]}"
}

# --- hash ------------------------------------------------------------------
sha256sum() {
    # shasum stampa "hash  file", lo stesso formato di sha256sum: i chiamanti
    # che fanno awk '{print $1}' continuano a funzionare senza modifiche.
    shasum -a 256 "$@"
}
md5sum() { command md5 -r "$@"; }

# --- numfmt ----------------------------------------------------------------
# Non esiste su macOS. Serve solo nella forma --to=iec per rendere leggibile
# una dimensione in byte, quindi si implementa quella e basta: emulare l'intero
# numfmt sarebbe codice non esercitato da nessuno.
numfmt() {
    local iec=false n
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --to=iec|--to=iec-i) iec=true; shift ;;
            --to=*|--suffix=*|--format=*) shift ;;
            *) n="$1"; shift ;;
        esac
    done
    [[ -n "${n:-}" ]] || return 1
    if ! $iec; then printf '%s\n' "$n"; return 0; fi
    awk -v b="$n" 'BEGIN{
        split("K M G T P", u, " ")
        if (b < 1024) { printf "%d\n", b; exit }
        for (i = 1; i <= 5; i++) { b /= 1024; if (b < 1024) break }
        printf "%.1f%s\n", b, u[i]
    }'
}

# --- tac -------------------------------------------------------------------
tac() { tail -r "$@"; }

# --- realpath / readlink ---------------------------------------------------
# BSD non ha realpath -m (risolvi anche se il percorso non esiste). Si emula:
# si risolve la parte esistente e si riattacca il resto.
realpath() {
    local missing_ok=false ; local -a rest=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -m|--canonicalize-missing) missing_ok=true; shift ;;
            -e|-P|-L|--*) shift ;;
            *) rest+=("$1"); shift ;;
        esac
    done
    local p="${rest[0]:-}"
    [[ -n "$p" ]] || return 1
    [[ "$p" != /* ]] && p="$PWD/$p"
    if ! $missing_ok && [[ ! -e "$p" ]]; then return 1; fi
    local dir="$p" tail_part=""
    while [[ ! -d "$dir" && "$dir" != "/" ]]; do
        tail_part="$(basename "$dir")${tail_part:+/$tail_part}"
        dir="$(dirname "$dir")"
    done
    local abs; abs=$(cd "$dir" 2>/dev/null && pwd -P) || abs="$dir"
    printf '%s\n' "${abs%/}${tail_part:+/$tail_part}"
}

readlink() {
    local canon=false ; local -a rest=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -e|-f|-m) canon=true; shift ;;
            -n) shift ;;
            *) rest+=("$1"); shift ;;
        esac
    done
    if $canon; then
        realpath -m "${rest[0]:-}"
    else
        command readlink "${rest[@]}"
    fi
}

# --- flock -----------------------------------------------------------------
# Assente su macOS. Serve a un solo punto (l'append concorrente alla timeline
# unica con --jobs) e si emula con mkdir, che e' atomico anche qui. Senza
# attesa infinita: dopo il tetto si procede comunque, perche' perdere
# l'ordinamento di qualche riga e' meno grave che bloccare l'analisi.
flock() {
    local lock="" ; local -a cmd=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -c) cmd=("$2"); shift 2 ;;
            -*) shift ;;
            *) [[ -z "$lock" ]] && lock="$1" || cmd+=("$1"); shift ;;
        esac
    done
    [[ -n "$lock" ]] || return 1
    local d="${lock}.d" i=0
    while ! mkdir "$d" 2>/dev/null; do
        i=$((i + 1)); [[ $i -gt 200 ]] && break
        sleep 0.05
    done
    if [[ ${#cmd[@]} -gt 0 ]]; then bash -c "${cmd[0]}"; fi
    local rc=$?
    rmdir "$d" 2>/dev/null || true
    return $rc
}

# --- timeout ---------------------------------------------------------------
# coreutils di Homebrew lo installa come gtimeout. Se manca, si esegue il
# comando senza limite: meglio un modulo lento di un modulo che non parte.
if ! command -v timeout > /dev/null 2>&1; then
    timeout() {
        if command -v gtimeout > /dev/null 2>&1; then
            gtimeout "$@"
        else
            shift   # scarta la durata
            "$@"
        fi
    }
fi

# --- directory di invocazione ----------------------------------------------
# L'originale risale l'albero dei processi leggendo /proc, che su macOS non
# esiste. Qui si usa il PWD, che copre il caso reale (l'utente lancia lo script
# dalla directory in cui vuole i report) senza fingere di fare di piu'.
_find_user_cwd() { pwd; }
INVOCATION_DIR="$(pwd)"

# --- rilevamento dei volumi montati ----------------------------------------
# /proc/mounts non esiste: su macOS l'elenco lo da' mount(8), e i volumi
# montati stanno sotto /Volumes.
_macos_mount_points() {
    /sbin/mount | awk '{print $3}' | grep -v '^/$' 2>/dev/null
    [[ -d /Volumes ]] && find /Volumes -maxdepth 1 -mindepth 1 2>/dev/null
}

# --- montaggio di immagini --------------------------------------------------
# La catena Linux (losetup + cryptsetup + mount) non ha equivalente diretto.
# Su macOS il montaggio in sola lettura si fa con hdiutil, che gestisce da se'
# la tabella delle partizioni. La funzione sostituisce interamente quella
# generica: viene definita DOPO src/lib/21-image.sh nell'ordine di build.
_macos_image_open() {
    local IMG="$IMAGE_PATH"
    [[ -f "$IMG" ]] || { err "$(L "Immagine non trovata:" "Image not found:") $IMG" >&2; return 1; }
    local TYPE; TYPE=$(image_type "$IMG")
    info "$(L "Immagine:" "Image:") ${BOLD}$(basename "$IMG")${RESET} — $(L "formato" "format"): ${BOLD}${TYPE}" >&2

    case "$TYPE" in
        vmdk|vhdx|qcow2)
            err "$(L "Formato $TYPE non supportato." "Format $TYPE not supported.")" >&2
            info "  qemu-img convert -f $TYPE -O raw \"$(basename "$IMG")\" immagine.raw" >&2
            return 1 ;;
    esac

    local RAW="$IMG"
    if [[ "$TYPE" == ewf || "$TYPE" == ewf2 ]]; then
        command -v ewfmount > /dev/null 2>&1 || {
            err "$(L "ewfmount non disponibile: brew install libewf" "ewfmount unavailable: brew install libewf")" >&2
            return 1; }
        local D; D=$(_image_state) || return 1
        local EWFDIR="${D}/ewf"; mkdir -p "$EWFDIR"
        ewfmount "$IMG" "$EWFDIR" 2>/dev/null || {
            err "$(L "ewfmount non riuscito." "ewfmount failed.")" >&2; return 1; }
        _image_push fuse "$EWFDIR"
        RAW="${EWFDIR}/ewf1"
    fi

    if $IMAGE_LIST_ONLY; then
        image_show_partitions "$RAW" >&2 || \
            info "$(L "Nessuna tabella leggibile da mmls; hdiutil montera' comunque i volumi che riconosce." \
                     "No table readable by mmls; hdiutil will still mount the volumes it recognises.")" >&2
        return 2
    fi

    # -readonly e -nomount: si attacca senza montare, poi si monta a mano solo
    # cio' che serve. hdiutil non scrive mai sull'immagine con -readonly.
    local OUT; OUT=$(hdiutil attach -readonly -nomount -plist "$RAW" 2>/dev/null) || {
        err "$(L "hdiutil non e' riuscito ad attaccare l'immagine." "hdiutil could not attach the image.")" >&2
        return 1; }
    local DEV
    DEV=$(printf '%s' "$OUT" | grep -oE '/dev/disk[0-9]+' | head -1)
    [[ -n "$DEV" ]] || { err "$(L "Nessun device esposto da hdiutil." "hdiutil exposed no device.")" >&2; return 1; }
    _image_push hdiutil "$DEV"
    ok "$(L "Immagine attaccata in sola lettura:" "Image attached read-only:") ${BOLD}${DEV}" >&2

    local D; D=$(_image_state) || return 1
    local MNT="${D}/mnt"; mkdir -p "$MNT"
    # Si monta la partizione indicata, o la prima che il sistema riesce a
    # montare. Su un disco multi-partizione senza --partition NON si sceglie a
    # caso: si elencano e ci si ferma, come nella versione Linux.
    local TARGET="${DEV}s${IMAGE_PARTITION:-}"
    if [[ -z "${IMAGE_PARTITION:-}" ]]; then
        local -a SLICES=()
        local s
        while IFS= read -r s; do [[ -n "$s" ]] && SLICES+=("$s"); done < <(diskutil list -plist "$DEV" 2>/dev/null | grep -oE "disk[0-9]+s[0-9]+" | sort -u)
        if [[ ${#SLICES[@]} -gt 1 ]]; then
            err "$(L "L'immagine contiene ${#SLICES[@]} partizioni: indica quale con --partition." \
                     "The image holds ${#SLICES[@]} partitions: pick one with --partition.")" >&2
            printf '  %s\n' "${SLICES[@]}" >&2
            return 1
        fi
        TARGET="/dev/${SLICES[0]:-}"
        [[ "$TARGET" == "/dev/" ]] && TARGET="$DEV"
    fi

    if ! /sbin/mount -o rdonly,noexec,nosuid -t apfs "$TARGET" "$MNT" 2>/dev/null \
       && ! /sbin/mount -o rdonly,noexec,nosuid "$TARGET" "$MNT" 2>/dev/null; then
        err "$(L "Mount fallito. Il volume potrebbe essere cifrato (FileVault) o di un tipo non supportato dal sistema." \
                 "Mount failed. The volume may be encrypted (FileVault) or of a type the system does not support.")" >&2
        return 1
    fi
    _image_push mount "$MNT"
    evidence_note "$IMG" "immagine forense montata" 2>/dev/null || true
    ok "$(L "Immagine montata in sola lettura:" "Image mounted read-only:") ${BOLD}${MNT}" >&2
    echo "$MNT"
}

# --- smontaggio ------------------------------------------------------------
# La pila di image_cleanup non conosce il tipo "hdiutil": lo si aggiunge
# ridefinendo la funzione, che resta per il resto identica all'originale.
image_cleanup() {
    local D="${TMPDIR:-/tmp}/fiuto_image_$$"
    local S="${D}/stack"
    [[ -f "$S" ]] || return 0
    local KIND TARGET
    while IFS=$'\t' read -r KIND TARGET; do
        case "$KIND" in
            mount)   umount "$TARGET" 2>/dev/null || diskutil unmount force "$TARGET" 2>/dev/null || true ;;
            hdiutil) hdiutil detach "$TARGET" -force 2>/dev/null || true ;;
            fuse)    umount "$TARGET" 2>/dev/null || true ;;
        esac
    done < <(tail -r "$S" 2>/dev/null)
    rm -f "$S"
    return 0
}

# image_open generica -> versione hdiutil. La sostituzione avviene qui e non
# dentro 21-image.sh perche' quel file resta identico fra le due build.
# Nessun argomento: come l'originale, legge IMAGE_PATH e le altre globali.
image_open() { _macos_image_open; }

# --- volumi montati ---------------------------------------------------------
# L'originale legge /proc/mounts con tre campi (device, mount point, fstype).
# Su macOS la stessa informazione la da' mount(8) in un formato diverso, quindi
# si ridefinisce la funzione che li raccoglie mantenendone il contratto: stampa
# un punto di mount per riga, gia' filtrato.
_find_windows_mounts() {
    local -a CANDIDATES=()
    local MNT
    while IFS= read -r MNT; do
        [[ -z "$MNT" || "$MNT" == "/" ]] && continue
        # Volumi di sistema e snapshot APFS del sistema operativo in uso: non
        # sono reperti, sono la macchina su cui stiamo lavorando.
        case "$MNT" in
            /System/*|/private/*|/dev/*|/net/*|/home/*) continue ;;
        esac
        if [[ -d "$MNT/Windows" || -d "$MNT/Users" ]] \
           || [[ -f "$MNT/etc/os-release" || -f "$MNT/etc/passwd" ]] \
           || [[ -d "$MNT/System/Library/CoreServices" ]]; then
            CANDIDATES+=("$MNT")
        fi
    done < <(_macos_mount_points)
    local -A SEEN=()
    for MNT in "${CANDIDATES[@]}"; do
        [[ -n "${SEEN[$MNT]+x}" ]] && continue
        SEEN["$MNT"]=1; echo "$MNT"
    done
}

# debug_mounts stampa /proc/mounts: su macOS si mostra l'equivalente reale.
debug_mounts() {
    clear 2>/dev/null || true
    print_banner
    echo -e "  ${DIM}── mount(8) + /Volumes ─────────────────────────────${RESET}"
    echo ""
    /sbin/mount | sed 's/^/  /'
    echo ""
    echo -e "  ${DIM}── candidati riconosciuti da FIUTO ─────────────────${RESET}"
    echo ""
    _find_windows_mounts | sed 's/^/  /'
    echo ""
    pause_key
}
