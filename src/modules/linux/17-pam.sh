
# ================================================================
#  LINUX 17 — PAM (Pluggable Authentication Modules)
#
#  PAM decide chi puo' autenticarsi e come. Una backdoor PAM e' fra le
#  persistenze piu' efficaci su Linux: aggiungendo una riga a /etc/pam.d/sshd
#  o sostituendo pam_unix.so con una versione modificata, l'attaccante ottiene
#  accesso con una password universale che nessun log segnala come anomala —
#  l'autenticazione risulta semplicemente riuscita.
#
#  Il modulo Persistence copre cron, systemd e autostart, ma non tocca PAM.
#  Qui si guardano tre cose: righe di configurazione sospette, moduli .so che
#  non appartengono a nessun pacchetto, e moduli con data di modifica
#  disallineata rispetto ai vicini.
# ================================================================
module_linux_pam() {
    section_header "Linux - PAM" "$RED"
    check_target_root || return 1

    local PAMD; PAMD=$(ci_find_dir "$WIN_ROOT" "etc/pam.d")
    local PAMCONF; PAMCONF=$(ci_find_file "$WIN_ROOT" "etc/pam.conf")
    if [[ -z "$PAMD" && -z "$PAMCONF" ]]; then
        warn "$(L "Configurazione PAM non trovata." "PAM configuration not found.")"
        return 0
    fi

    # Righe che meritano attenzione in una configurazione PAM.
    #  pam_exec      esegue un comando esterno a ogni autenticazione
    #  pam_permit    concede l'accesso incondizionatamente
    #  nullok        accetta password vuote
    #  pam_python /
    #  percorsi non standard  moduli caricati da fuori la directory di sistema
    local KW="pam_exec|pam_permit|nullok|pam_python|pam_script|/tmp/|/dev/shm|/home/|pam_succeed_if.*quiet.*uid|debug"

    local BODY="" NFILE=0 NSUSP=0 SUSPROWS=""
    local F
    if [[ -n "$PAMD" ]]; then
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue
            NFILE=$((NFILE + 1))
            local N; N=$(grep -cE "$KW" "$F" 2>/dev/null || echo 0)
            if [[ "$N" -gt 0 ]]; then
                NSUSP=$((NSUSP + 1))
                ok "$(basename "$F") - ${BOLD}${N}${RESET} $(L "righe da verificare" "lines to check")"
                local LN
                while IFS= read -r LN; do
                    SUSPROWS+="$(basename "$F")	${LN}
"
                done < <(grep -nE "$KW" "$F" 2>/dev/null | head -20)
            fi
            BODY+=$(file_card_html "$F" "$KW" "PAM")
        done < <(find "$PAMD" -maxdepth 1 -type f 2>/dev/null | sort)
    fi
    if [[ -n "$PAMCONF" && -s "$PAMCONF" ]]; then
        NFILE=$((NFILE + 1))
        BODY+=$(file_card_html "$PAMCONF" "$KW" "PAM")
    fi

    # Moduli PAM sul disco: un .so non pacchettizzato o con mtime fuori linea
    # rispetto ai vicini e' il segnale piu' concreto di manomissione.
    local SOROWS="" NSO=0 NORPH=0
    local D
    for D in "lib/x86_64-linux-gnu/security" "lib64/security" "lib/security" \
             "usr/lib/x86_64-linux-gnu/security" "usr/lib64/security" "usr/lib/security"; do
        local SD; SD=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -z "$SD" ]] && continue
        while IFS= read -r F; do
            [[ -s "$F" ]] || continue
            NSO=$((NSO + 1))
            local MT SZ OWN
            MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
            SZ=$(stat -c %s "$F" 2>/dev/null)
            # Appartenenza a un pacchetto: si verifica offline nei database
            # di dpkg (list dei file) e rpm non e' interrogabile senza il DB,
            # quindi qui si usa solo dpkg quando presente.
            OWN="?"
            local DPKGINFO; DPKGINFO=$(ci_find_dir "$WIN_ROOT" "var/lib/dpkg/info")
            if [[ -n "$DPKGINFO" ]]; then
                local REL="${F#$WIN_ROOT}"
                if grep -rqsF "$REL" "$DPKGINFO"/*.list 2>/dev/null; then
                    OWN="$(L "pacchetto" "package")"
                else
                    OWN="$(L "NON PACCHETTIZZATO" "NOT PACKAGED")"
                    NORPH=$((NORPH + 1))
                fi
            fi
            SOROWS+="${MT}	$(basename "$F")	${SZ}	${OWN}	${F}
"
        done < <(find "$SD" -maxdepth 1 -type f -name '*.so' 2>/dev/null | sort)
    done

    separator
    info "$(L "File di configurazione PAM:" "PAM configuration files:") ${BOLD}$NFILE"
    [[ $NSUSP -gt 0 ]] && warn "$(L "File con righe da verificare:" "Files with lines to check:") ${BOLD}$NSUSP"
    info "$(L "Moduli PAM (.so):" "PAM modules (.so):") ${BOLD}$NSO"
    if [[ "$NORPH" -gt 0 ]]; then
        warn "$(L "Moduli non appartenenti ad alcun pacchetto:" "Modules not owned by any package:") ${BOLD}$NORPH"
    fi
    [[ $NFILE -eq 0 && $NSO -eq 0 ]] && return 0

    ask_yn "Generare report HTML?" || return 0

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Come leggere le segnalazioni" "How to read the flags")</b><br>"
    NOTE+="$(L "Le parole chiave evidenziate non sono di per se' malevole: pam_exec e nullok hanno usi legittimi. Vanno lette come 'questa riga cambia chi puo' entrare e come, quindi va giustificata'. Il segnale piu' forte e' invece un modulo .so che non appartiene a nessun pacchetto, o la cui data di modifica e' isolata rispetto agli altri moduli della stessa directory." \
        "The highlighted keywords are not malicious in themselves: pam_exec and nullok have legitimate uses. Read them as 'this line changes who can log in and how, so it needs justifying'. The stronger signal is a .so module owned by no package, or whose modification date stands alone among the other modules in the same directory.")<br><br>"
    NOTE+="$(L "Una backdoor PAM non lascia tracce nei log di autenticazione: l'accesso risulta semplicemente riuscito." \
        "A PAM backdoor leaves no trace in authentication logs: the login simply succeeds.")"
    NOTE+="</div></div>"

    local CARDS=""
    if [[ -n "$SUSPROWS" ]]; then
        CARDS+="<div class='cards'>$(generic_card_html "$(L "Righe da verificare" "Lines to check")" "/etc/pam.d" "$NSUSP" \
            "$(_rows_to_table "$SUSPROWS" "$(L "File" "File")" "$(L "Riga" "Line")")" "⚑")</div>"
    fi
    if [[ -n "$SOROWS" ]]; then
        CARDS+="<div class='cards'>$(generic_card_html "$(L "Moduli PAM installati" "Installed PAM modules")" "security/*.so" "$NSO" \
            "$(_rows_to_table "$SOROWS" "$(L "Ultima modifica" "Last modified")" "$(L "Modulo" "Module")" "$(L "Byte" "Bytes")" "$(L "Origine" "Origin")" "$(L "Percorso" "Path")")" "◧")</div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Config" "Config")" "$NFILE")"
    STATS+="$(stat_box "$(L "Da verificare" "To check")" "$NSUSP" "$([[ $NSUSP -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Moduli .so" "Modules .so")" "$NSO" "info")"
    STATS+="$(stat_box "$(L "Non pacchettizzati" "Not packaged")" "$NORPH" "$([[ $NORPH -gt 0 ]] && echo warn || echo info)")"
    finish_report "linux_pam" "Linux PAM" "PAM" "/etc/pam.d · security/*.so" "$STATS" \
        "${NOTE}${CARDS}<div class='cards'>${BODY}</div>"
}
