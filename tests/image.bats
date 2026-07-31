#!/usr/bin/env bats
#
# Test del montaggio di immagini forensi e volumi cifrati.
#
# Il montaggio vero richiede root e qui non c'è: si esercita tutto quello che
# gira da utente normale — riconoscimento del formato, lettura della tabella
# delle partizioni, rilevamento della cifratura, pila di smontaggio — più i
# rifiuti, che sono la parte che conta di più.
#
# Il rischio di questa funzione non è tecnico: è montare la partizione
# sbagliata, o montare in scrittura, senza che nulla lo dica.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    IMAGE_PATH=""; IMAGE_PARTITION=""; IMAGE_UNLOCK=""; IMAGE_LIST_ONLY=false
}

teardown() {
    image_cleanup 2>/dev/null || true
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
    rm -rf "${TMPDIR:-/tmp}/fiuto_image_$$"
}

_need() { command -v "$1" > /dev/null 2>&1 || skip "$1 non installato"; }

# Disco raw con tabella DOS e due partizioni.
_disk() {
    _need sfdisk
    truncate -s 64M "$FIXTURE/disco.raw"
    sfdisk --quiet "$FIXTURE/disco.raw" > /dev/null 2>&1 <<'EOF'
label: dos
start=2048, size=40960, type=07
start=43008, size=40960, type=83
EOF
    echo "$FIXTURE/disco.raw"
}

# ------------------------------------------------------ riconoscimento -----

@test "il formato viene dedotto dal contenuto, non dall'estensione" {
    # Un .dd rinominato .img resta un raw, una E01 chiamata .001 resta una E01.
    printf 'EVF\x09\x0d\x0a\xff\x00' > "$FIXTURE/finta.img"
    [ "$(image_type "$FIXTURE/finta.img")" = "ewf" ]
    printf 'EVF2\x0d\x0a\x81\x00' > "$FIXTURE/finta2.dd"
    [ "$(image_type "$FIXTURE/finta2.dd")" = "ewf2" ]
    head -c 512 /dev/zero > "$FIXTURE/vuota.E01"
    [ "$(image_type "$FIXTURE/vuota.E01")" = "raw" ]
}

@test "i formati virtuali vengono riconosciuti e rifiutati" {
    # Leggerli come raw darebbe dati sbagliati senza alcun errore: sono
    # formati sparsi o a snapshot.
    printf 'KDMV\x01\x00\x00\x00' > "$FIXTURE/d.vmdk"
    [ "$(image_type "$FIXTURE/d.vmdk")" = "vmdk" ]
    IMAGE_PATH="$FIXTURE/d.vmdk"
    run image_open
    [ "$status" -ne 0 ]
    [[ "$output" == *"qemu-img convert"* ]]
}

# ------------------------------------------------------- partizioni --------

@test "la tabella delle partizioni viene letta con gli offset in byte" {
    _need mmls
    D=$(_disk)
    ROWS=$(image_partitions "$D")
    [ "$(wc -l <<< "$ROWS")" -eq 2 ]
    # 2048 settori da 512 byte = 1048576
    [ "$(head -1 <<< "$ROWS" | cut -f2)" -eq 1048576 ]
}

@test "le voci non montabili non compaiono fra le partizioni" {
    # mmls elenca anche lo spazio non allocato e la tabella stessa: proporle
    # come montabili farebbe scegliere quella sbagliata.
    _need mmls
    D=$(_disk)
    ! image_partitions "$D" | grep -qi "unallocated"
}

@test "su un disco multi-partizione senza --partition si rifiuta di scegliere" {
    # Sceglierne una a caso significa analizzare la partizione sbagliata e non
    # accorgersene mai.
    _need mmls
    IMAGE_PATH=$(_disk)
    run image_open
    [ "$status" -ne 0 ]
    [[ "$output" == *"--partition"* ]]
}

@test "una partizione inesistente è un errore, non un ripiego sulla prima" {
    _need mmls
    IMAGE_PATH=$(_disk)
    IMAGE_PARTITION=99
    run image_open
    [ "$status" -ne 0 ]
    [[ "$output" == *"99"* ]]
}

# -------------------------------------------------------- cifratura -------

@test "un volume LUKS viene riconosciuto" {
    _need cryptsetup
    truncate -s 32M "$FIXTURE/luks.raw"
    printf 'passwordprova' | cryptsetup luksFormat --batch-mode --key-file=- "$FIXTURE/luks.raw" 2>/dev/null || skip "luksFormat non riuscito"
    [ "$(image_crypto_at "$FIXTURE/luks.raw" 0)" = "luks" ]
}

@test "un volume BitLocker viene riconosciuto dalla firma -FVE-FS-" {
    truncate -s 1M "$FIXTURE/bl.raw"
    printf '\xeb\x58\x90-FVE-FS-' | dd of="$FIXTURE/bl.raw" bs=1 conv=notrunc status=none
    [ "$(image_crypto_at "$FIXTURE/bl.raw" 0)" = "bitlocker" ]
}

@test "una partizione in chiaro non viene scambiata per cifrata" {
    truncate -s 1M "$FIXTURE/plain.raw"
    [ "$(image_crypto_at "$FIXTURE/plain.raw" 0)" = "none" ]
}

@test "senza --unlock un volume cifrato non viene montato" {
    # Nessun tentativo di indovinare: si dice cosa manca.
    truncate -s 1M "$FIXTURE/bl.raw"
    printf '\xeb\x58\x90-FVE-FS-' | dd of="$FIXTURE/bl.raw" bs=1 conv=notrunc status=none
    IMAGE_PATH="$FIXTURE/bl.raw"
    run image_open
    [ "$status" -ne 0 ]
    [[ "$output" == *"unlock"* ]]
}

@test "FileVault viene dichiarato non supportato invece di essere tentato" {
    truncate -s 1M "$FIXTURE/apfs.raw"
    printf 'NXSB' | dd of="$FIXTURE/apfs.raw" bs=1 seek=32 conv=notrunc status=none
    [ "$(image_crypto_at "$FIXTURE/apfs.raw" 0)" = "apfs" ]
}

# ----------------------------------------------------------- E01 -----------

@test "una E01 viene aperta e le sue partizioni elencate senza root" {
    # È il motivo per cui --list-partitions esiste: capire cosa c'è
    # nell'immagine prima di decidere se serve sudo.
    _need ewfacquire
    _need ewfmount
    _need mmls
    D=$(_disk)
    ( cd "$FIXTURE" && ewfacquire -q -u -t e01 -f encase6 -c deflate:none -S 0 -C 1 \
        -D t -e t -E t -m removable -M logical -N test "$D" ) > /dev/null 2>&1 \
        || skip "ewfacquire non riuscito"
    [ "$(image_type "$FIXTURE/e01.E01")" = "ewf" ]
    IMAGE_PATH="$FIXTURE/e01.E01"; IMAGE_LIST_ONLY=true
    run image_open
    # 2 = elenco stampato, niente da montare
    [ "$status" -eq 2 ]
    [[ "$output" == *"NTFS"* ]]
}

# ------------------------------------------------------- smontaggio -------

@test "la pila viene smontata in ordine inverso a quello di apertura" {
    # Sbagliare l'ordine lascia loop device orfani e mount occupati.
    _image_push loop /dev/loop-finto
    _image_push crypt nome-finto
    _image_push mount /mnt/finto
    S="${TMPDIR:-/tmp}/fiuto_image_$$/stack"
    [ "$(wc -l < "$S")" -eq 3 ]
    [ "$(head -1 "$S" | cut -f1)" = "loop" ]
    [ "$(tail -1 "$S" | cut -f1)" = "mount" ]
    image_cleanup
    [ ! -f "$S" ]
}

@test "image_cleanup può essere chiamata due volte senza errori" {
    # Gira dalla trap EXIT, che può scattare dopo una chiusura esplicita.
    _image_push loop /dev/loop-finto
    image_cleanup
    run image_cleanup
    [ "$status" -eq 0 ]
}

# ------------------------------------------------- sola lettura -----------

@test "nessun percorso di codice monta in scrittura" {
    # È l'invariante dell'intera funzione: ogni anello della catena riceve la
    # sua opzione di sola lettura, e non per convenzione.
    src="$REPO_ROOT/src/lib/21-image.sh"
    grep -q 'losetup --find --show --read-only' "$src"
    grep -q 'cryptsetup open --readonly' "$src"
    grep -q 'dislocker -r' "$src"
    grep -q 'OPTS="ro,noexec,nodev,nosuid"' "$src"
    # Nessun mount senza ro fra le opzioni.
    ! grep -E '^\s*mount ' "$src" | grep -v 'ro,' | grep -q 'mount '
}

@test "senza privilegi il montaggio si ferma dicendo perché" {
    [ "$(id -u)" -eq 0 ] && skip "il test presuppone un utente non privilegiato"
    _need mmls
    IMAGE_PATH=$(_disk)
    IMAGE_PARTITION=002
    run image_open
    [ "$status" -ne 0 ]
    [[ "$output" == *"root"* ]]
    [[ "$output" == *"list-partitions"* ]]
}
