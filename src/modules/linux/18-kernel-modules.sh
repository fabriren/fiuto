
# ================================================================
#  LINUX 18 — Moduli kernel e rootkit
#
#  Un LKM rootkit gira in kernel space: puo' nascondere processi, file,
#  connessioni e se stesso agli strumenti in user space. Nessun modulo che
#  legge /proc o esegue comandi sulla macchina viva puo' vederlo — l'analisi
#  offline del filesystem, invece, si'.
#
#  Il modulo Persistence copre solo ld.so.preload. Qui si guardano i punti da
#  cui un modulo viene caricato all'avvio e i .ko presenti sul disco che non
#  appartengono a nessun pacchetto.
# ================================================================
module_linux_kernel_modules() {
    section_header "Linux - $(L "Moduli kernel e rootkit" "Kernel modules and rootkits")" "$RED"
    check_target_root || return 1

    local BODY="" ROWS="" NCFG=0 NSUSP=0
    local F D

    # --- Punti di caricamento automatico ---------------------------------
    local KW="^[a-zA-Z0-9_-]+"
    for D in "etc/modules-load.d" "etc/modprobe.d" "usr/lib/modules-load.d" "lib/modules-load.d"; do
        local DIR; DIR=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$DIR" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue
            NCFG=$((NCFG + 1))
            BODY+=$(file_card_html "$F" "install |^blacklist|/tmp/|/dev/shm|insmod|modprobe" "KO")
        done < <(find "$DIR" -maxdepth 1 -type f 2>/dev/null | sort)
    done
    F=$(ci_find_file "$WIN_ROOT" "etc/modules")
    if [[ -s "$F" ]]; then
        NCFG=$((NCFG + 1))
        BODY+=$(file_card_html "$F" "$KW" "KO")
    fi

    # "install <modulo> <comando>" in modprobe.d esegue un comando arbitrario
    # al caricamento del modulo: e' una persistenza poco nota e molto efficace.
    local INSTROWS=""
    for D in "etc/modprobe.d" "usr/lib/modprobe.d" "lib/modprobe.d"; do
        local DIR; DIR=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$DIR" ]] && continue
        while IFS= read -r LN; do
            [[ -n "$LN" ]] && { INSTROWS+="${LN}
"; NSUSP=$((NSUSP + 1)); }
        done < <(grep -rhnE '^[[:space:]]*install[[:space:]]+\S+[[:space:]]+\S' "$DIR" 2>/dev/null | grep -vE '/bin/(true|false)[[:space:]]*$' | head -50)
    done

    # --- Moduli .ko sul disco --------------------------------------------
    local NKO=0 NORPH=0 KOROWS=""
    local MODDIR; MODDIR=$(ci_find_dir "$WIN_ROOT" "lib/modules")
    [[ -z "$MODDIR" ]] && MODDIR=$(ci_find_dir "$WIN_ROOT" "usr/lib/modules")
    if [[ -n "$MODDIR" ]]; then
        local DPKGINFO; DPKGINFO=$(ci_find_dir "$WIN_ROOT" "var/lib/dpkg/info")
        # Un rootkit installato a mano finisce quasi sempre in extra/ o misc/,
        # o direttamente nella radice della versione del kernel.
        while IFS= read -r F; do
            NKO=$((NKO + 1))
            local MT SZ OWN REL
            MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
            SZ=$(stat -c %s "$F" 2>/dev/null)
            REL="${F#$WIN_ROOT}"
            OWN="?"
            if [[ -n "$DPKGINFO" ]]; then
                if grep -rqsF "$REL" "$DPKGINFO"/*.list 2>/dev/null; then
                    OWN="$(L "pacchetto" "package")"
                else
                    OWN="$(L "NON PACCHETTIZZATO" "NOT PACKAGED")"
                    NORPH=$((NORPH + 1))
                    KOROWS+="${MT}	$(basename "$F")	${SZ}	${OWN}	${REL}
"
                fi
            fi
        done < <(find "$MODDIR" -type f \( -name '*.ko' -o -name '*.ko.xz' -o -name '*.ko.zst' -o -name '*.ko.gz' \) 2>/dev/null)
    fi

    # --- initramfs --------------------------------------------------------
    local NINIT=0 INITROWS=""
    local BOOT; BOOT=$(ci_find_dir "$WIN_ROOT" "boot")
    if [[ -n "$BOOT" ]]; then
        while IFS= read -r F; do
            NINIT=$((NINIT + 1))
            INITROWS+="$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)	$(basename "$F")	$(stat -c %s "$F" 2>/dev/null)
"
        done < <(find "$BOOT" -maxdepth 1 -type f \( -iname 'initrd*' -o -iname 'initramfs*' \) 2>/dev/null | sort)
    fi

    separator
    info "$(L "Configurazioni di caricamento:" "Load configurations:") ${BOLD}$NCFG"
    [[ "$NSUSP" -gt 0 ]] && warn "$(L "Direttive 'install' con comando:" "'install' directives with a command:") ${BOLD}$NSUSP"
    info "$(L "Moduli .ko sul disco:" "Kernel modules on disk:") ${BOLD}$NKO"
    if [[ "$NORPH" -gt 0 ]]; then
        warn "$(L "Moduli non appartenenti ad alcun pacchetto:" "Modules owned by no package:") ${BOLD}$NORPH"
        printf '%s' "$KOROWS" | head -15 | while IFS=$'\t' read -r MT NM SZ OWN P; do
            echo -e "      ${RED}${MT}  ${NM}  ${P}${RESET}"
        done
    fi
    info "initramfs: ${BOLD}$NINIT"

    if [[ $NCFG -eq 0 && $NKO -eq 0 && $NINIT -eq 0 ]]; then
        warn "$(L "Nessun artefatto relativo ai moduli kernel trovato." "No kernel-module artefact found.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Perche' l'analisi offline conta qui" "Why offline analysis matters here")</b><br>"
    NOTE+="$(L "Un rootkit LKM gira in kernel space e puo' nascondersi a qualunque strumento eseguito sulla macchina compromessa, incluso lsmod. Sul filesystem montato da fuori non ha modo di mentire: i .ko ci sono o non ci sono." \
        "An LKM rootkit runs in kernel space and can hide from any tool executed on the compromised machine, lsmod included. On a filesystem mounted from outside it has no way to lie: the .ko files are either there or not.")<br><br>"
    NOTE+="<b>$(L "La direttiva install" "The install directive")</b><br>"
    NOTE+="$(L "In modprobe.d, 'install &lt;modulo&gt; &lt;comando&gt;' esegue quel comando quando il modulo viene caricato. E' una persistenza poco nota: non compare in cron, systemd o autostart, e scatta al primo uso di un dispositivo qualsiasi." \
        "In modprobe.d, 'install &lt;module&gt; &lt;command&gt;' runs that command when the module is loaded. It is a little-known persistence: it appears in no cron, systemd or autostart list, and triggers on the first use of any device.")"
    NOTE+="</div></div>"

    local CARDS=""
    [[ -n "$INSTROWS" ]] && CARDS+="<div class='cards'>$(generic_card_html "$(L "Direttive install" "install directives")" "modprobe.d" "$NSUSP" "$(_rows_to_table "$INSTROWS" "$(L "Riga" "Line")")" "⚑")</div>"
    [[ -n "$KOROWS" ]] && CARDS+="<div class='cards'>$(generic_card_html "$(L "Moduli non pacchettizzati" "Modules not packaged")" "$MODDIR" "$NORPH" "$(_rows_to_table "$KOROWS" "$(L "Ultima modifica" "Last modified")" "$(L "Modulo" "Module")" "$(L "Byte" "Bytes")" "$(L "Origine" "Origin")" "$(L "Percorso" "Path")")" "◧")</div>"
    [[ -n "$INITROWS" ]] && CARDS+="<div class='cards'>$(generic_card_html "initramfs" "$BOOT" "$NINIT" "$(_rows_to_table "$INITROWS" "$(L "Ultima modifica" "Last modified")" "$(L "File" "File")" "$(L "Byte" "Bytes")")" "▤")</div>"

    local STATS
    STATS="$(stat_box "$(L "Moduli .ko" "Modules .ko")" "$NKO")"
    STATS+="$(stat_box "$(L "Non pacchettizzati" "Not packaged")" "$NORPH" "$([[ "$NORPH" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "install" "$NSUSP" "$([[ "$NSUSP" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "initramfs" "$NINIT" "info")"
    finish_report "linux_kernel_modules" "Linux Kernel Modules" "KRN" "lib/modules · modprobe.d · initramfs" "$STATS" \
        "${NOTE}${CARDS}<div class='cards'>${BODY}</div>"
}
