
# ================================================================
#  CROSS-OS — EFI System Partition e bootkit
#
#  La ESP e' una piccola partizione FAT che il firmware legge all'accensione
#  per trovare il bootloader. Il codice che sta li' viene eseguito PRIMA del
#  sistema operativo, del kernel e di qualunque agente EDR: e' la posizione di
#  persistenza piu' ambita, e sopravvive alla reinstallazione del sistema e
#  alla formattazione della partizione di sistema.
#
#  I bootkit noti (ESPecter, BlackLotus, Bootkitty, CosmicStrand) agiscono
#  sostituendo o affiancando i loader legittimi. Il modulo inventaria la ESP,
#  calcola gli hash di tutto cio' che vi trova e segnala le anomalie
#  strutturali: file non-EFI, loader in posizioni inattese, date isolate.
#
#  Vale per tutti e tre i sistemi: la ESP e' condivisa fra gli OS installati.
# ================================================================
module_xplat_esp_bootkit() {
    section_header "EFI System Partition — bootkit" "$RED"
    check_target_root || return 1

    # La ESP puo' essere il volume stesso (montata a parte) oppure trovarsi
    # sotto /boot/efi, /efi o /Volumes/EFI del volume in analisi.
    local -a ESPS=()
    local D
    for D in "EFI" "boot/efi/EFI" "efi/EFI" "Volumes/EFI/EFI" "boot/EFI"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ESPS+=("$R")
    done

    if [[ ${#ESPS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna EFI System Partition raggiungibile da questo volume." \
                 "No EFI System Partition reachable from this volume.")"
        info "$(L "La ESP e' una partizione separata: se non e' montata, va montata a parte (di norma la prima partizione FAT32 del disco) e analizzata indicandola come root." \
                 "The ESP is a separate partition: if not mounted, mount it separately (usually the disk's first FAT32 partition) and analyse it as the root.")"
        return 0
    fi

    local ESP; ESP="${ESPS[0]}"
    info "ESP: ${BOLD}${ESP}"

    # Loader legittimi attesi. Un nome fuori da questo elenco non e' di per se'
    # malevolo (molte distribuzioni ne aggiungono di propri), ma va giustificato.
    local KNOWN="bootx64.efi|bootia32.efi|bootaa64.efi|bootmgfw.efi|bootmgr.efi|memtest.efi|shimx64.efi|shimaa64.efi|shim.efi|grubx64.efi|grubaa64.efi|mmx64.efi|fbx64.efi|mokmanager.efi|systemd-bootx64.efi|fwupdx64.efi|kernel.efi|vmlinuz.efi|refind_x64.efi|BOOT.EFI|apfs.efi|boot.efi|firmware.scap|immutablekernel"

    # Nomi ricorrenti nei bootkit documentati e nei loro payload.
    local BADNAMES="grubx64_real|bootmgfw_original|bootmgfw\.efi\.bak|winload\.efi|bootkit|especter|blacklotus|bootlicker|cosmicstrand|\.sys$|\.dll$|\.ps1$|\.bat$|\.vbs$|\.exe$"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    local NFILE=0 NEFI=0 NNONEFI=0 NUNKNOWN=0 NSUSP=0
    local F
    while IFS= read -r F; do
        [[ -f "$F" ]] || continue
        NFILE=$((NFILE + 1))
        local BASE REL SZ MT HASH FLAGS
        BASE=$(basename "$F")
        REL="${F#$ESP}"
        SZ=$(stat -c %s "$F" 2>/dev/null || echo 0)
        MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
        HASH=$(sha256_file "$F")
        FLAGS=""

        # Un binario EFI e' un PE: inizia con MZ. Un file nella ESP che non lo
        # e' — e non e' un file di configurazione noto — non ha motivo di stare li'.
        local MAGIC; MAGIC=$(head -c 2 "$F" 2>/dev/null)
        if [[ "$MAGIC" == "MZ" ]]; then
            NEFI=$((NEFI + 1))
        else
            case "${BASE,,}" in
                *.efi) FLAGS+="EFI_SENZA_HEADER_PE;"; NSUSP=$((NSUSP + 1)) ;;
                *.cfg|*.conf|*.ini|*.json|*.txt|bcd|*.crt|*.cer|*.esl|*.auth|*.scap|*.dat) : ;;
                *) FLAGS+="NON_EFI;"; NNONEFI=$((NNONEFI + 1)) ;;
            esac
        fi
        if ! printf '%s' "${BASE,,}" | grep -qiE "^(${KNOWN})$"; then
            case "${BASE,,}" in
                *.efi) FLAGS+="LOADER_NON_STANDARD;"; NUNKNOWN=$((NUNKNOWN + 1)) ;;
            esac
        fi
        if printf '%s' "${BASE,,}" | grep -qiE "$BADNAMES"; then
            FLAGS+="NOME_SOSPETTO;"
            NSUSP=$((NSUSP + 1))
        fi
        printf '%s\t%s\t%s\t%s\t%s\n' "$MT" "${REL#/}" "$SZ" "${FLAGS%;}" "$HASH" >> "$OUT"
    done < <(find "$ESP" -type f 2>/dev/null | sort)

    separator
    ok "$(L "File nella ESP:" "Files in the ESP:") ${BOLD}$NFILE"
    info "$(L "Binari EFI (header PE):" "EFI binaries (PE header):") ${BOLD}$NEFI"
    [[ "$NUNKNOWN" -gt 0 ]] && warn "$(L "Loader non standard:" "Non-standard loaders:") ${BOLD}$NUNKNOWN"
    [[ "$NNONEFI" -gt 0 ]] && warn "$(L "File non-EFI nella ESP:" "Non-EFI files in the ESP:") ${BOLD}$NNONEFI"
    [[ "$NSUSP" -gt 0 ]] && warn "$(L "Segnalazioni forti:" "Strong flags:") ${BOLD}$NSUSP"
    if [[ "$NUNKNOWN" -gt 0 || "$NNONEFI" -gt 0 || "$NSUSP" -gt 0 ]]; then
        echo ""
        awk -F'\t' '$4!=""{printf "      [%s] %s\n", $4, $2}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi
    [[ "$NFILE" -eq 0 ]] && { warn "$(L "ESP vuota o illeggibile." "ESP empty or unreadable.")"; return 0; }

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$4!=""' "$OUT"; awk -F'\t' '$4==""' "$OUT"; } )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Percorso nella ESP" "Path in ESP")" \
        "$(L "Byte" "Bytes")" "$(L "Segnalazioni" "Flags")" "SHA256")

    local NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' la ESP conta piu' di quanto sembri" "Why the ESP matters more than it looks")</b><br>"
    NOTE+="$(L "Il codice nella ESP viene eseguito prima del sistema operativo, del kernel e di qualunque agente EDR. Una persistenza qui sopravvive alla reinstallazione del sistema e alla formattazione della partizione di sistema: se un incidente si ripresenta dopo un ripristino completo, questo e' il primo posto da guardare." \
        "Code in the ESP runs before the operating system, the kernel and any EDR agent. Persistence here survives OS reinstallation and formatting of the system partition: if an incident recurs after a full rebuild, this is the first place to look.")<br><br>"
    NOTE+="<b>$(L "Come usare gli hash" "How to use the hashes")</b><br>"
    NOTE+="$(L "Gli SHA256 in tabella vanno confrontati con quelli dei loader legittimi della distribuzione o della versione di Windows installata, e cercati nelle basi di reputazione. Un bootx64.efi con un hash che non corrisponde a nessuna build ufficiale e' il reperto." \
        "The SHA256 values in the table should be compared against the legitimate loaders of the installed distribution or Windows build, and looked up in reputation databases. A bootx64.efi whose hash matches no official build is the finding.")<br><br>"
    NOTE+="<b>$(L "Le segnalazioni" "The flags")</b><br>"
    NOTE+="$(L "LOADER_NON_STANDARD significa 'nome non nell'elenco dei loader noti': molte distribuzioni ne aggiungono di legittimi, quindi va giustificato, non temuto. NON_EFI e EFI_SENZA_HEADER_PE sono piu' pesanti: un file nella ESP che non e' un binario EFI ne' una configurazione nota non ha una ragione ovvia per stare li'." \
        "NON_STANDARD_LOADER means 'name not in the known-loader list': many distributions add legitimate ones, so it needs justifying, not fearing. NON_EFI and EFI_WITHOUT_PE_HEADER are heavier: a file in the ESP that is neither an EFI binary nor a known configuration has no obvious reason to be there.")<br><br>"
    NOTE+="$(L "Per un controllo mirato sulle famiglie note serve YARA con regole aggiornate: questo modulo non ne include, perche' regole ferme al momento del rilascio darebbero una falsa sensazione di copertura." \
        "For targeted checks against known families use YARA with current rules: this module ships none, because rules frozen at release time would give a false sense of coverage.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "File" "Files")" "$NFILE")"
    STATS+="$(stat_box "$(L "Binari EFI" "EFI binaries")" "$NEFI" "info")"
    STATS+="$(stat_box "$(L "Non standard" "Non-standard")" "$NUNKNOWN" "$([[ "$NUNKNOWN" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Segnalati" "Flagged")" "$((NNONEFI + NSUSP))" "$([[ $((NNONEFI + NSUSP)) -gt 0 ]] && echo warn || echo info)")"
    finish_report "xplat_esp_bootkit" "EFI System Partition" "ESP" "$(L "bootkit e persistenza pre-boot" "bootkits and pre-boot persistence")" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Contenuto della ESP" "ESP contents")" "$ESP" "$NFILE" "$TABLE" "⏻")</div>"
}
