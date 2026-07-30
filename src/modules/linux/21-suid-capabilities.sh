
# ================================================================
#  LINUX 21 — SUID/SGID, capabilities e file scrivibili da tutti
#
#  E' la superficie di privilege escalation del sistema. Conta in due momenti
#  diversi dell'indagine:
#
#  - come CAUSA: un binario SUID inatteso (o un /bin/bash con il bit SUID)
#    spiega come l'attaccante e' passato da utente a root;
#  - come EFFETTO: molti attaccanti lasciano dietro di se' un SUID come
#    backdoor di riserva, piu' discreta di un utente aggiuntivo.
#
#  L'analisi offline e' l'unico modo affidabile di farla: su una macchina
#  compromessa find puo' essere sostituito o un rootkit puo' nascondere i file.
# ================================================================
module_linux_suid_caps() {
    section_header "Linux — SUID/SGID $(L "e capabilities" "and capabilities")" "$ORANGE"
    check_target_root || return 1

    info "$(L "Scansione del filesystem in corso..." "Scanning the filesystem...")"

    # Binari SUID/SGID considerati normali su una distribuzione: servono a
    # separare il rumore dal segnale, NON a dichiarare sicuro cio' che vi
    # corrisponde (un /usr/bin/passwd modificato resta in questa lista).
    local EXPECTED="/(passwd|chsh|chfn|newgrp|gpasswd|su|sudo|mount|umount|ping|ping6|fusermount|fusermount3|pkexec|crontab|at|ssh-agent|unix_chkpwd|expiry|chage|wall|write|dotlockfile|dbus-daemon-launch-helper|polkit-agent-helper-1|sg|staprun|Xorg|snap-confine|utempter|screen|mount\.nfs|pam_timestamp_check|cockpit-session|vmware-user-suid-wrapper)$"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    # -xdev: resta sul volume montato, non segue mount annidati o /proc.
    find "$WIN_ROOT" -xdev \( -perm -4000 -o -perm -2000 \) -type f -printf '%M\t%u\t%g\t%s\t%TY-%Tm-%Td %TH:%TM:%TS\t%p\n' 2>/dev/null \
        | sed 's/\.[0-9]*\t/\t/' > "$OUT" || true

    # NB: `grep -c` stampa 0 ma esce con stato 1 quando non trova nulla, quindi
    # un `|| echo 0` produrrebbe due zeri. Si usa wc -l.
    local NSUID NUNEXP
    NSUID=$(wc -l < "$OUT" 2>/dev/null || echo 0)
    local UNEXP; UNEXP=$(mktemp); register_tmp "$UNEXP"
    grep -vE "$EXPECTED" "$OUT" > "$UNEXP" 2>/dev/null || true
    NUNEXP=$(wc -l < "$UNEXP" 2>/dev/null || echo 0)

    # File con capabilities: alternativa moderna al SUID, spesso trascurata.
    # getcap non funziona su un volume montato senza supporto xattr, quindi si
    # tenta e si dichiara l'esito invece di far finta di aver guardato.
    local CAPOUT; CAPOUT=$(mktemp); register_tmp "$CAPOUT"
    local CAPS_OK=false NCAP=0
    if command -v getcap > /dev/null 2>&1; then
        CAPS_OK=true
        getcap -r "$WIN_ROOT" 2>/dev/null | head -500 > "$CAPOUT" || true
        NCAP=$(wc -l < "$CAPOUT" 2>/dev/null || echo 0)
    fi

    # Directory e file scrivibili da tutti fuori dalle aree temporanee attese.
    local WWOUT; WWOUT=$(mktemp); register_tmp "$WWOUT"
    find "$WIN_ROOT" -xdev -perm -0002 ! -type l \
         ! -path "*/tmp/*" ! -path "*/var/tmp/*" ! -path "*/dev/shm/*" ! -path "*/proc/*" \
         -printf '%M\t%u\t%s\t%TY-%Tm-%Td %TH:%TM\t%p\n' 2>/dev/null | head -400 > "$WWOUT" || true
    local NWW; NWW=$(wc -l < "$WWOUT" 2>/dev/null || echo 0)

    separator
    ok "$(L "Binari SUID/SGID:" "SUID/SGID binaries:") ${BOLD}$NSUID"
    if [[ "$NUNEXP" -gt 0 ]]; then
        warn "$(L "Fuori dall'elenco atteso:" "Outside the expected set:") ${BOLD}$NUNEXP"
        awk -F'\t' '{printf "      %s  %-8s %s\n", $1, $2, $6}' "$UNEXP" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    fi
    if $CAPS_OK; then
        info "$(L "File con capabilities:" "Files with capabilities:") ${BOLD}$NCAP"
    else
        info "$(L "getcap non disponibile: capabilities non verificate." "getcap unavailable: capabilities not checked.")"
    fi
    info "$(L "Scrivibili da tutti (fuori da tmp):" "World-writable (outside tmp):") ${BOLD}$NWW"

    if [[ "$NSUID" -eq 0 && "$NCAP" -eq 0 && "$NWW" -eq 0 ]]; then
        warn "$(L "Nessun risultato: il volume potrebbe essere montato senza supporto ai permessi." \
                 "No result: the volume may be mounted without permission support.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Come leggere l'elenco atteso" "How to read the expected set")</b><br>"
    BODY+="$(L "I binari SUID di sistema (passwd, sudo, mount...) sono separati dagli altri solo per ridurre il rumore. Non sono dichiarati sicuri: un /usr/bin/passwd sostituito resta nell'elenco atteso, e va confrontato con l'hash del pacchetto. Il segnale forte e' un SUID FUORI dall'elenco — soprattutto una shell, un interprete o un binario in /tmp, /home o /var." \
        "System SUID binaries (passwd, sudo, mount...) are separated from the rest only to cut noise. They are not declared safe: a replaced /usr/bin/passwd still sits in the expected set and must be checked against the package hash. The strong signal is a SUID OUTSIDE the set — especially a shell, an interpreter, or a binary under /tmp, /home or /var.")<br><br>"
    BODY+="$(L "Le capabilities sono l'alternativa moderna al SUID e vengono spesso dimenticate in fase di audit: CAP_SETUID, CAP_SYS_ADMIN e CAP_DAC_OVERRIDE su un binario arbitrario equivalgono di fatto a root." \
        "Capabilities are the modern alternative to SUID and are often overlooked during audits: CAP_SETUID, CAP_SYS_ADMIN and CAP_DAC_OVERRIDE on an arbitrary binary are effectively root.")"
    if ! $CAPS_OK; then
        BODY+="<br><br><b>$(L "Capabilities non verificate" "Capabilities not checked")</b><br>"
        BODY+="$(L "getcap non era disponibile sull'host di analisi: questa parte del controllo non e' stata eseguita e va rifatta a parte." \
            "getcap was unavailable on the analysis host: this part of the check was not performed and must be repeated separately.")"
    fi
    BODY+="</div></div>"

    if [[ "$NUNEXP" -gt 0 ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "SUID/SGID fuori dall'elenco atteso" "SUID/SGID outside the expected set")" "$WIN_ROOT" "$NUNEXP" \
            "$(_rows_to_table "$(cat "$UNEXP")" "$(L "Permessi" "Mode")" "$(L "Utente" "Owner")" "$(L "Gruppo" "Group")" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")" "$(L "Percorso" "Path")")" "⚑")</div>"
    fi
    BODY+="<div class='cards'>$(generic_card_html "$(L "Tutti i SUID/SGID" "All SUID/SGID")" "$WIN_ROOT" "$NSUID" \
        "$(_rows_to_table "$(head -2000 "$OUT")" "$(L "Permessi" "Mode")" "$(L "Utente" "Owner")" "$(L "Gruppo" "Group")" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")" "$(L "Percorso" "Path")")" "◧")</div>"
    if [[ "$NCAP" -gt 0 ]]; then
        BODY+="<div class='cards'>$(generic_card_html "Capabilities" "getcap -r" "$NCAP" \
            "$(_rows_to_table "$(sed 's/ /\t/' "$CAPOUT")" "$(L "Percorso" "Path")" "Capabilities")" "⚙")</div>"
    fi
    if [[ "$NWW" -gt 0 ]]; then
        BODY+="<div class='cards'>$(generic_card_html "$(L "Scrivibili da tutti" "World-writable")" "$(L "fuori da tmp" "outside tmp")" "$NWW" \
            "$(_rows_to_table "$(cat "$WWOUT")" "$(L "Permessi" "Mode")" "$(L "Utente" "Owner")" "$(L "Byte" "Bytes")" "$(L "Ultima modifica" "Last modified")" "$(L "Percorso" "Path")")" "◔")</div>"
    fi

    local STATS
    STATS="$(stat_box "SUID/SGID" "$NSUID")"
    STATS+="$(stat_box "$(L "Inattesi" "Unexpected")" "$NUNEXP" "$([[ "$NUNEXP" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "Capabilities" "$($CAPS_OK && echo "$NCAP" || echo "n/d")" "info")"
    STATS+="$(stat_box "$(L "Scrivibili da tutti" "World-writable")" "$NWW" "$([[ "$NWW" -gt 0 ]] && echo warn || echo info)")"
    finish_report "linux_suid_caps" "Linux SUID & Capabilities" "SUID" "$(L "superficie di privilege escalation" "privilege escalation surface")" "$STATS" "$BODY"
}
