
# ================================================================
#  IMMAGINI FORENSI E VOLUMI CIFRATI
#
#  Fino alla 2.2 FIUTO pretendeva un volume gia' montato. Il montaggio a mano
#  di una E01 e' una catena di quattro comandi in cui e' facile sbagliare
#  proprio il pezzo che conta: dimenticare il -r su losetup, o il ro su mount,
#  significa aver montato il reperto in scrittura senza accorgersene.
#
#  --image chiude quella catena. Ogni anello e' in SOLA LETTURA per costruzione
#  e non per convenzione: ewfmount lo e' per natura, losetup riceve -r,
#  cryptsetup --readonly, dislocker -r, mount -o ro. Non esiste un percorso di
#  codice che monti in scrittura.
#
#  COSA NON FA, dichiarato invece che scoperto a meta' analisi:
#  - non indovina chiavi. BitLocker e LUKS si aprono solo con --unlock;
#  - non tocca FileVault 2: richiede il keybag di macOS e strumenti che su
#    Linux non esistono in forma affidabile. Il volume viene riconosciuto e la
#    cosa viene detta, non aggirata;
#  - VMDK, VHDX e QCOW2 sono riconosciuti e rifiutati con l'indicazione del
#    comando per convertirli: fingere di supportarli produrrebbe letture
#    sbagliate su formati sparsi o a snapshot.
#
#  Il montaggio richiede privilegi (losetup, mount, cryptsetup). L'elenco delle
#  partizioni no: `--list-partitions` funziona da utente normale, cosi' si puo'
#  capire cosa c'e' in un'immagine prima di decidere se serve sudo.
# ================================================================

IMAGE_PATH=""              # --image: immagine da montare
IMAGE_PARTITION=""         # --partition N: quale partizione (numero di mmls)
IMAGE_UNLOCK=""            # --unlock: password/recovery key o file che la contiene
IMAGE_LIST_ONLY=false      # --list-partitions: elenca ed esci

# Pila di smontaggio. Su file e non in un array: il cleanup gira dalla trap
# EXIT, e deve funzionare anche se qualcosa e' stato aperto da una subshell.
_image_state() {
    local D="${TMPDIR:-/tmp}/fiuto_image_$$"
    [[ -d "$D" ]] || mkdir -p "$D" 2>/dev/null || return 1
    echo "$D"
}

# _image_push <tipo> <bersaglio> — tipo: fuse | crypt | mount | loop
_image_push() {
    local D; D=$(_image_state) || return 0
    printf '%s\t%s\n' "$1" "$2" >> "${D}/stack"
}

# Smonta tutto in ordine inverso. Idempotente: puo' essere chiamata due volte.
image_cleanup() {
    local D="${TMPDIR:-/tmp}/fiuto_image_$$"
    local S="${D}/stack"
    [[ -f "$S" ]] || return 0
    local KIND TARGET
    # tac e non un ciclo al contrario: l'ordine di chiusura e' l'inverso esatto
    # di quello di apertura, e sbagliarlo lascia loop device orfani.
    while IFS=$'\t' read -r KIND TARGET; do
        # Ogni ramo termina con true: una funzione di pulizia non deve MAI
        # propagare un errore. Gira dalla trap EXIT, spesso quando qualcosa e'
        # gia' andato storto, e un suo fallimento nasconderebbe la causa vera.
        case "$KIND" in
            mount) umount "$TARGET" 2>/dev/null || umount -l "$TARGET" 2>/dev/null || true ;;
            crypt) cryptsetup close "$TARGET" 2>/dev/null || true ;;
            loop)  losetup -d "$TARGET" 2>/dev/null || true ;;
            fuse)  fusermount -u "$TARGET" 2>/dev/null || umount "$TARGET" 2>/dev/null || true ;;
        esac
    done < <(tac "$S" 2>/dev/null)
    rm -f "$S"
    return 0
}

# Tipo dell'immagine dal contenuto, non dall'estensione: un .dd rinominato .img
# resta un raw, e una E01 chiamata .001 resta una E01.
image_type() {
    local F="$1"
    [[ -f "$F" ]] || { echo "assente"; return 1; }
    local MAGIC; MAGIC=$(head -c 8 "$F" 2>/dev/null | od -An -tx1 | tr -d ' \n')
    case "$MAGIC" in
        45564609*) echo "ewf" ; return ;;   # "EVF\x09" — EnCase E01
        45564632*) echo "ewf2"; return ;;   # "EVF2"    — Ex01
    esac
    case "$MAGIC" in
        4b444d56*) echo "vmdk" ; return ;;  # "KDMV"
        76686478*) echo "vhdx" ; return ;;  # "vhdx"
        514649fb*) echo "qcow2"; return ;;  # "QFI\xfb"
    esac
    echo "raw"
}

# Cifratura del contenuto a un dato offset. Riconoscere e dichiarare vale piu'
# che tentare: un volume BitLocker montato come raw non da' errore, da' dati
# illeggibili che sembrano corruzione.
image_crypto_at() {
    local DEV="$1" OFF="${2:-0}"
    local HDR; HDR=$(dd if="$DEV" bs=1 skip="$OFF" count=16 2>/dev/null | od -An -c | tr -s ' ')
    case "$HDR" in
        *'- F V E - F S -'*) echo bitlocker; return ;;
        *'M S W I N 4 . 1'*) : ;;
    esac
    local L; L=$(dd if="$DEV" bs=1 skip="$OFF" count=6 2>/dev/null | od -An -tx1 | tr -d ' \n')
    [[ "$L" == 4c554b53babe* ]] && { echo luks; return; }
    # APFS cifrato: il container si riconosce, la chiave sta nel keybag.
    # Il confronto passa dall'esadecimale: i byte grezzi in una command
    # substitution fanno emettere a bash un avviso sui null, e su una
    # partizione vuota (tutti zeri) l'avviso comparirebbe a ogni riga.
    local A; A=$(dd if="$DEV" bs=1 skip=$((OFF + 32)) count=4 2>/dev/null | od -An -tx1 | tr -d ' \n')
    [[ "$A" == "4e585342" ]] && { echo apfs; return; }
    echo none
}

# Elenco delle partizioni via mmls. Ritorna righe "slot<TAB>offset_byte<TAB>settori<TAB>descrizione".
image_partitions() {
    local DEV="$1"
    command -v mmls > /dev/null 2>&1 || return 1
    # L'output di mmls passa da un file e non da una pipe: lo stdin del
    # processo Python e' gia' occupato dall'heredoc con il programma.
    local TMP; TMP=$(mktemp) || return 1
    mmls -M "$DEV" > "$TMP" 2>/dev/null
    "$PY3" - "$TMP" << 'PYEOF'
import sys, re

sector = 512
rows = []
for line in open(sys.argv[1], encoding='utf-8', errors='replace'):
    m = re.search(r'Units are in (\d+)-byte sectors', line)
    if m:
        sector = int(m.group(1))
        continue
    # "  002:  000:000   0000002048   0000206847   0000204800   NTFS / exFAT"
    m = re.match(r'\s*(\d+):\s+\S+\s+(\d+)\s+(\d+)\s+(\d+)\s+(.*?)\s*$', line)
    if not m:
        continue
    slot, start, _end, length, desc = m.groups()
    # Le voci meta (tabella delle partizioni, spazio non allocato) non si
    # montano: elencarle come montabili farebbe scegliere quella sbagliata.
    if re.search(r'unallocated|meta|extended|primary table', desc, re.I):
        continue
    rows.append('%s\t%d\t%s\t%s' % (slot, int(start) * sector, length, desc))
print('\n'.join(rows))
PYEOF
    rm -f "$TMP"
}

# Stampa a console le partizioni trovate. Funziona senza privilegi.
image_show_partitions() {
    local DEV="$1"
    local ROWS; ROWS=$(image_partitions "$DEV")
    if [[ -z "$ROWS" ]]; then
        warn "$(L "Nessuna tabella delle partizioni: l'immagine sembra contenere un solo filesystem." \
                 "No partition table: the image seems to hold a single filesystem.")"
        return 1
    fi
    echo ""
    printf "  ${BOLD}%-5s %-14s %-12s %-8s %s${RESET}\n" \
        "$(L "SLOT" "SLOT")" "$(L "OFFSET" "OFFSET")" "$(L "SETTORI" "SECTORS")" \
        "$(L "CIFRAT." "CRYPTO")" "$(L "DESCRIZIONE" "DESCRIPTION")"
    local SLOT OFF LEN DESC
    while IFS=$'\t' read -r SLOT OFF LEN DESC; do
        local C; C=$(image_crypto_at "$DEV" "$OFF")
        local COL="$RESET"; [[ "$C" != "none" ]] && COL="$YELLOW"
        printf "  ${CYAN}%-5s${RESET} %-14s %-12s ${COL}%-8s${RESET} %s\n" \
            "$SLOT" "$OFF" "$LEN" "$([[ "$C" == "none" ]] && echo "-" || echo "$C")" "$DESC"
    done <<< "$ROWS"
    echo ""
    return 0
}

# Legge la chiave da --unlock: valore diretto o percorso di un file.
_image_key() {
    [[ -n "$IMAGE_UNLOCK" ]] || return 1
    if [[ -f "$IMAGE_UNLOCK" ]]; then
        head -1 "$IMAGE_UNLOCK"
    else
        printf '%s' "$IMAGE_UNLOCK"
    fi
}

# image_open — apre IMAGE_PATH e stampa il punto di mount su stdout.
# Ogni messaggio va su stderr: lo stdout e' il percorso, e sporcarlo
# significherebbe restituire una root inesistente.
image_open() {
    local IMG="$IMAGE_PATH"
    [[ -f "$IMG" ]] || { err "$(L "Immagine non trovata:" "Image not found:") $IMG" >&2; return 1; }

    local TYPE; TYPE=$(image_type "$IMG")
    info "$(L "Immagine:" "Image:") ${BOLD}$(basename "$IMG")${RESET} — $(L "formato" "format"): ${BOLD}${TYPE}" >&2

    case "$TYPE" in
        vmdk|vhdx|qcow2)
            err "$(L "Formato $TYPE non supportato." "Format $TYPE not supported.")" >&2
            info "$(L "Sono formati sparsi o a snapshot: leggerli come raw darebbe dati sbagliati senza errori. Converti prima:" \
                     "These are sparse or snapshot formats: reading them as raw would give wrong data with no error. Convert first:")" >&2
            info "  qemu-img convert -f $TYPE -O raw \"$(basename "$IMG")\" immagine.raw" >&2
            return 1 ;;
    esac

    # --- E01/Ex01: ewfmount espone un raw. E' FUSE, non serve root. ---------
    local RAW="$IMG"
    if [[ "$TYPE" == ewf || "$TYPE" == ewf2 ]]; then
        command -v ewfmount > /dev/null 2>&1 || {
            err "$(L "ewfmount non disponibile (pacchetto ewf-tools)." "ewfmount unavailable (ewf-tools package).")" >&2
            return 1; }
        local D; D=$(_image_state) || return 1
        local EWFDIR="${D}/ewf"; mkdir -p "$EWFDIR"
        if ! ewfmount "$IMG" "$EWFDIR" 2>/dev/null; then
            err "$(L "ewfmount non riuscito. Se l'immagine e' segmentata, indica il primo segmento (.E01)." \
                     "ewfmount failed. If the image is segmented, point at the first segment (.E01).")" >&2
            return 1
        fi
        _image_push fuse "$EWFDIR"
        RAW="${EWFDIR}/ewf1"
        [[ -e "$RAW" ]] || { err "$(L "ewfmount non ha esposto ewf1." "ewfmount did not expose ewf1.")" >&2; return 1; }
        ok "$(L "E01 aperta in sola lettura." "E01 opened read-only.")" >&2
    fi

    # --- scelta della partizione -------------------------------------------
    local ROWS; ROWS=$(image_partitions "$RAW")
    local OFFSET=0 PDESC="$(L "immagine intera" "whole image")"
    if [[ -n "$ROWS" ]]; then
        if $IMAGE_LIST_ONLY; then
            image_show_partitions "$RAW" >&2
            return 2
        fi
        local NPART; NPART=$(wc -l <<< "$ROWS")
        if [[ -n "$IMAGE_PARTITION" ]]; then
            local LINE; LINE=$(awk -F'\t' -v s="$IMAGE_PARTITION" '$1==s || $1+0==s+0' <<< "$ROWS" | head -1)
            [[ -n "$LINE" ]] || {
                err "$(L "Partizione inesistente:" "No such partition:") $IMAGE_PARTITION" >&2
                image_show_partitions "$RAW" >&2
                return 1; }
            OFFSET=$(cut -f2 <<< "$LINE"); PDESC=$(cut -f4 <<< "$LINE")
        elif [[ "$NPART" -eq 1 ]]; then
            OFFSET=$(cut -f2 <<< "$ROWS"); PDESC=$(cut -f4 <<< "$ROWS")
        else
            # Sceglierne una a caso su un disco multi-partizione significa
            # analizzare la partizione sbagliata e non accorgersene.
            err "$(L "L'immagine contiene $NPART partizioni: indica quale con --partition." \
                     "The image holds $NPART partitions: pick one with --partition.")" >&2
            image_show_partitions "$RAW" >&2
            return 1
        fi
    elif $IMAGE_LIST_ONLY; then
        image_show_partitions "$RAW" >&2
        return 2
    fi
    info "$(L "Partizione:" "Partition:") ${BOLD}${PDESC}${RESET} (offset ${OFFSET})" >&2

    # --- cifratura: si rileva PRIMA di chiedere i privilegi ----------------
    # "Questa partizione e' BitLocker e non mi hai dato la chiave" si sa da
    # utente normale, ed e' un'informazione piu' utile di "serve root": evita
    # di rilanciare con sudo per scoprire solo allora che manca la chiave.
    local CRYPTO; CRYPTO=$(image_crypto_at "$RAW" "$OFFSET")
    if [[ "$CRYPTO" == "apfs" ]]; then
        err "$(L "Container APFS cifrato (FileVault 2)." "Encrypted APFS container (FileVault 2).")" >&2
        info "$(L "Non supportato: la chiave sta nel keybag di macOS e su Linux non esistono strumenti affidabili per estrarla. Sblocca il volume su un Mac e ri-acquisiscilo, oppure analizzalo montato da macOS." \
                 "Not supported: the key lives in the macOS keybag and Linux has no reliable tooling for it. Unlock the volume on a Mac and re-acquire it, or analyse it mounted from macOS.")" >&2
        return 1
    fi
    if [[ "$CRYPTO" != "none" ]] && ! _image_key > /dev/null; then
        err "$(L "Partizione ${CRYPTO}: serve la chiave con --unlock <chiave|file>." \
                 "Partition is ${CRYPTO}: a key is required — --unlock <key|file>.")" >&2
        info "$(L "FIUTO non tenta di indovinarla." "FIUTO does not attempt to guess it.")" >&2
        return 1
    fi

    # --- da qui servono i privilegi ----------------------------------------
    if [[ "$(id -u)" -ne 0 ]]; then
        err "$(L "Il montaggio richiede privilegi di root (losetup, mount, cryptsetup)." \
                 "Mounting requires root privileges (losetup, mount, cryptsetup).")" >&2
        info "$(L "Con --list-partitions l'elenco funziona anche senza." \
                 "With --list-partitions the listing works without them.")" >&2
        return 1
    fi

    local D; D=$(_image_state) || return 1
    local MNT="${D}/mnt"; mkdir -p "$MNT"
    local SRC=""

    case "$CRYPTO" in
        none)
            SRC=$(losetup --find --show --read-only --offset "$OFFSET" "$RAW" 2>/dev/null) || {
                err "$(L "losetup non riuscito." "losetup failed.")" >&2; return 1; }
            _image_push loop "$SRC"
            ;;
        luks)
            warn "$(L "Partizione LUKS." "LUKS partition.")" >&2
            local KEY; KEY=$(_image_key) || {
                err "$(L "Serve la passphrase: --unlock <passphrase|file>." \
                         "Passphrase required: --unlock <passphrase|file>.")" >&2; return 1; }
            local LOOP; LOOP=$(losetup --find --show --read-only --offset "$OFFSET" "$RAW" 2>/dev/null) || {
                err "$(L "losetup non riuscito." "losetup failed.")" >&2; return 1; }
            _image_push loop "$LOOP"
            local NAME="fiuto_luks_$$"
            if ! printf '%s' "$KEY" | cryptsetup open --readonly --key-file=- "$LOOP" "$NAME" 2>/dev/null; then
                err "$(L "Apertura LUKS fallita: passphrase errata o header danneggiato." \
                         "LUKS unlock failed: wrong passphrase or damaged header.")" >&2
                return 1
            fi
            _image_push crypt "$NAME"
            SRC="/dev/mapper/${NAME}"
            ok "$(L "Volume LUKS aperto in sola lettura." "LUKS volume opened read-only.")" >&2
            ;;
        bitlocker)
            warn "$(L "Partizione BitLocker." "BitLocker partition.")" >&2
            command -v dislocker > /dev/null 2>&1 || {
                err "$(L "dislocker non disponibile." "dislocker unavailable.")" >&2; return 1; }
            local KEY; KEY=$(_image_key) || {
                err "$(L "Serve la chiave: --unlock <password|recovery key|file>." \
                         "Key required: --unlock <password|recovery key|file>.")" >&2; return 1; }
            local DIS="${D}/dislocker"; mkdir -p "$DIS"
            # Una recovery key BitLocker sono 48 cifre in 8 gruppi: la si passa
            # con -p, una password utente con -u. Distinguerle a occhio evita
            # un fallimento che sembrerebbe "chiave sbagliata".
            local FLAG="-u"
            [[ "$KEY" =~ ^[0-9]{6}(-[0-9]{6}){7}$ ]] && FLAG="-p"
            if ! dislocker -r -V "$RAW" -O "$OFFSET" "${FLAG}${KEY}" -- "$DIS" 2>/dev/null; then
                err "$(L "Apertura BitLocker fallita: chiave errata o volume non supportato." \
                         "BitLocker unlock failed: wrong key or unsupported volume.")" >&2
                return 1
            fi
            _image_push fuse "$DIS"
            SRC="${DIS}/dislocker-file"
            ok "$(L "Volume BitLocker aperto in sola lettura." "BitLocker volume opened read-only.")" >&2
            ;;
    esac

    # --- mount, sempre e solo in sola lettura ------------------------------
    # noexec/nodev/nosuid: il contenuto e' materiale ostile per definizione, e
    # non c'e' ragione perche' il kernel lo consideri eseguibile.
    local OPTS="ro,noexec,nodev,nosuid"
    if mount -o "$OPTS" "$SRC" "$MNT" 2>/dev/null; then
        :
    elif mount -t ntfs-3g -o "ro,noexec,nodev,nosuid,streams_interface=windows" "$SRC" "$MNT" 2>/dev/null; then
        :
    else
        err "$(L "Mount fallito. Filesystem non riconosciuto o servono moduli kernel assenti." \
                 "Mount failed. Unrecognised filesystem, or missing kernel modules.")" >&2
        return 1
    fi
    _image_push mount "$MNT"

    # Il montaggio e' un fatto della catena di custodia quanto il file.
    evidence_note "$IMG" "immagine forense montata" 2>/dev/null || true
    log_msg "[IMAGE] $IMG ($TYPE) offset=$OFFSET crypto=$CRYPTO -> $MNT"
    ok "$(L "Immagine montata in sola lettura:" "Image mounted read-only:") ${BOLD}${MNT}" >&2
    echo "$MNT"
}
