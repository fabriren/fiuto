
# ================================================================
#  LINUX 20 — Credenziali cloud e di sviluppo
#
#  Su una workstation di sviluppo o su un server di build, il bottino non e'
#  la password dell'utente: sono le chiavi che danno accesso all'infrastruttura.
#  Un file ~/.aws/credentials o un ~/.kube/config permettono di passare dal
#  singolo host compromesso all'intero ambiente cloud, e non sono coperti da
#  nessun altro modulo.
#
#  Il modulo rileva la PRESENZA e i metadati di questi file, e mostra gli
#  identificativi non segreti (nome del profilo, access key ID, cluster,
#  registry). NON stampa mai il materiale segreto: la sua esistenza e' il dato
#  che serve all'analista, il valore in chiaro sarebbe solo un rischio in piu'
#  dentro il report.
# ================================================================
module_linux_cloud_credentials() {
    section_header "Linux — $(L "Credenziali cloud e sviluppo" "Cloud and development credentials")" "$RED"
    check_target_root || return 1

    local ROWS="" NFILE=0 NSECRET=0
    local HOME_DIR

    # spec: percorso relativo alla home | etichetta | tipo
    local -a SPECS=(
        ".aws/credentials|AWS|secret"
        ".aws/config|AWS|config"
        ".kube/config|Kubernetes|secret"
        ".docker/config.json|Docker Registry|secret"
        ".config/gcloud/credentials.db|Google Cloud|secret"
        ".config/gcloud/application_default_credentials.json|Google Cloud|secret"
        ".azure/accessTokens.json|Azure|secret"
        ".azure/azureProfile.json|Azure|config"
        ".netrc|netrc|secret"
        ".git-credentials|Git|secret"
        ".npmrc|npm|secret"
        ".pypirc|PyPI|secret"
        ".config/rclone/rclone.conf|rclone|secret"
        ".s3cfg|S3|secret"
        ".terraformrc|Terraform|secret"
        ".config/gh/hosts.yml|GitHub CLI|secret"
        ".ssh/id_rsa|SSH|key"
        ".ssh/id_ed25519|SSH|key"
        ".ssh/id_ecdsa|SSH|key"
    )

    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local SPEC REL LABEL KIND
        for SPEC in "${SPECS[@]}"; do
            IFS='|' read -r REL LABEL KIND <<< "$SPEC"
            local F; F=$(ci_find_file "$HOME_DIR" "$REL")
            [[ -s "$F" ]] || continue
            NFILE=$((NFILE + 1))
            [[ "$KIND" != "config" ]] && NSECRET=$((NSECRET + 1))

            local MT PERM SZ IDENT
            MT=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
            PERM=$(stat -c %a "$F" 2>/dev/null)
            SZ=$(stat -c %s "$F" 2>/dev/null)

            # Identificativi NON segreti, utili a capire a cosa da' accesso.
            IDENT=""
            case "$LABEL" in
                AWS)
                    IDENT=$(grep -aoE '^\[[^]]+\]|aws_access_key_id[[:space:]]*=[[:space:]]*[A-Z0-9]{16,}' "$F" 2>/dev/null \
                            | sed 's/.*=[[:space:]]*//' | tr '\n' ' ' | cut -c1-160) ;;
                Kubernetes)
                    IDENT=$(grep -aoE 'server:[[:space:]]*\S+|name:[[:space:]]*\S+' "$F" 2>/dev/null | head -6 | tr '\n' ' ' | cut -c1-160) ;;
                "Docker Registry")
                    IDENT=$(grep -aoE '"[a-z0-9.-]+\.[a-z]{2,}(:[0-9]+)?"' "$F" 2>/dev/null | sort -u | tr '\n' ' ' | cut -c1-160) ;;
                Git|netrc)
                    # Solo host e utente: la password viene esclusa.
                    IDENT=$(grep -aoE 'https://[^:/@]+@[^/]+|machine[[:space:]]+\S+|login[[:space:]]+\S+' "$F" 2>/dev/null \
                            | sed -E 's#https://([^:@]+):[^@]*@#https://\1@#' | head -6 | tr '\n' ' ' | cut -c1-160) ;;
                SSH)
                    IDENT=$(head -1 "$F" 2>/dev/null | grep -aoE 'BEGIN [A-Z ]+PRIVATE KEY' || echo "$(L "chiave privata" "private key")")
                    grep -aq 'ENCRYPTED' "$F" 2>/dev/null && IDENT+=" ($(L "protetta da passphrase" "passphrase protected"))" \
                        || IDENT+=" ($(L "SENZA passphrase" "NO passphrase"))" ;;
                *)
                    IDENT=$(grep -aoE '^\[[^]]+\]|^[a-z_]+[[:space:]]*=' "$F" 2>/dev/null | head -5 | tr '\n' ' ' | cut -c1-120) ;;
            esac

            local WARNP=""
            # Un file di credenziali leggibile da altri utenti e' un problema
            # a prescindere dall'incidente in corso.
            [[ -n "$PERM" && "${PERM: -2}" != "00" ]] && WARNP="$(L "permessi larghi" "loose permissions")"

            ROWS+="${MT}	${U}	${LABEL}	${REL}	${PERM} ${WARNP}	${SZ}	${IDENT}
"
            if [[ -n "$WARNP" ]]; then
                warn "$U — $REL ($PERM) — $WARNP"
            else
                ok "$U — ${BOLD}${LABEL}${RESET} ${DIM}${REL}${RESET}"
            fi
        done
    done < <(get_target_user_homes)

    separator
    if [[ "$NFILE" -eq 0 ]]; then
        info "$(L "Nessun file di credenziali cloud o di sviluppo trovato." "No cloud or development credential file found.")"
        return 0
    fi
    ok "$(L "File di credenziali trovati:" "Credential files found:") ${BOLD}$NFILE"
    warn "$(L "Di cui contenenti materiale segreto:" "Of which holding secret material:") ${BOLD}$NSECRET"

    ask_yn "Generare report HTML?" || return 0

    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Utente" "User")" "$(L "Servizio" "Service")" \
        "$(L "File" "File")" "$(L "Permessi" "Permissions")" "$(L "Byte" "Bytes")" "$(L "Identificativi" "Identifiers")")

    local NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Cosa mostra e cosa no" "What is shown and what is not")</b><br>"
    NOTE+="$(L "La tabella riporta identificativi non segreti — profilo, access key ID, endpoint del cluster, registry, host — perche' servono a capire a cosa quelle credenziali davano accesso. Il materiale segreto (secret key, token, password, contenuto delle chiavi private) NON viene mai stampato: e' la sua esistenza a essere il dato utile, riprodurlo qui aggiungerebbe solo un altro posto da cui puo' trapelare." \
        "The table shows non-secret identifiers — profile, access key ID, cluster endpoint, registry, host — because they establish what those credentials granted access to. Secret material (secret keys, tokens, passwords, private key contents) is never printed: its existence is the useful fact, reproducing it here would only add one more place it can leak from.")<br><br>"
    NOTE+="<b>$(L "Priorita' di risposta" "Response priority")</b><br>"
    NOTE+="$(L "Ogni credenziale elencata va considerata compromessa e ruotata. Le chiavi SSH senza passphrase e i file con permessi larghi vanno per primi." \
        "Every credential listed must be treated as compromised and rotated. SSH keys without a passphrase and files with loose permissions come first.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "File" "Files")" "$NFILE")"
    STATS+="$(stat_box "$(L "Con segreti" "With secrets")" "$NSECRET" "warn")"
    finish_report "linux_cloud_credentials" "Linux Cloud & Dev Credentials" "KEY" "~/.aws · ~/.kube · ~/.docker · ~/.ssh" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Credenziali rilevate" "Credentials found")" "$(L "home utenti" "user homes")" "$NFILE" "$TABLE" "⚿")</div>"
}
