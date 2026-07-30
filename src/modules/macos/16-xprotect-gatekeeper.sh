
# ================================================================
#  macOS 16 — XProtect, Gatekeeper e policy di sistema
#
#  Sono le difese native di macOS. In un'indagine servono a rispondere a due
#  domande diverse:
#
#  - le difese hanno visto qualcosa? (XProtect Remediator, versione delle
#    firme al momento dei fatti)
#  - le difese sono state indebolite? Un'app approvata a mano in SystemPolicy,
#    un kext di terze parti autorizzato in KextPolicy o Gatekeeper disattivato
#    sono azioni deliberate che spesso precedono l'esecuzione del malware.
#
#  La versione delle firme XProtect e' rilevante di per se': se e' molto
#  anteriore alla data dei fatti, l'assenza di rilevamenti non significa nulla.
# ================================================================
module_macos_xprotect() {
    section_header "macOS — XProtect / Gatekeeper" "$RED"
    check_target_root || return 1

    local BODY="" ROWS="" NFIND=0

    # ---------- Versione delle firme XProtect ----------
    local XPV="" XPD=""
    local META
    for META in "Library/Apple/System/Library/CoreServices/XProtect.bundle/Contents/Info.plist" \
                "System/Library/CoreServices/XProtect.bundle/Contents/Info.plist"; do
        local F; F=$(ci_find_file "$WIN_ROOT" "$META")
        [[ -s "$F" ]] || continue
        XPV=$(read_plist "$F" 2>/dev/null | grep -aoE 'CFBundleShortVersionString[^0-9]*([0-9]+)' | grep -oE '[0-9]+$' | head -1)
        XPD=$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
        [[ -n "$XPV" ]] && break
    done
    if [[ -n "$XPV" ]]; then
        ok "XProtect $(L "versione firme:" "signature version:") ${BOLD}${XPV}${RESET} ${DIM}(${XPD})${RESET}"
        ROWS+="XProtect	$(L "versione firme" "signature version")	${XPV}	${XPD}
"
    else
        warn "$(L "Versione delle firme XProtect non determinabile." "XProtect signature version not determinable.")"
    fi

    # ---------- Rilevamenti di XProtect Remediator ----------
    local XPDB
    for XPDB in "private/var/protected/xprotect/XPdb" "var/protected/xprotect/XPdb"; do
        local F; F=$(ci_find_file "$WIN_ROOT" "$XPDB")
        [[ -s "$F" ]] || continue
        local R; R=$(query_sqlite "$F" "SELECT * FROM sqlite_master WHERE type='table'")
        [[ -z "$R" || "$R" == ERROR* ]] && continue
        info "XPdb: ${DIM}${F}${RESET}"
        ROWS+="XProtect Remediator	$(L "database presente" "database present")	$(stat -c %s "$F" 2>/dev/null) B	$(stat -c %y "$F" 2>/dev/null | cut -d. -f1)
"
    done

    # ---------- SystemPolicy: eseguibili approvati a mano ----------
    local NAPPROVED=0 APPROWS=""
    local SP; SP=$(ci_find_file "$WIN_ROOT" "private/var/db/SystemPolicy")
    [[ -z "$SP" ]] && SP=$(ci_find_file "$WIN_ROOT" "var/db/SystemPolicy")
    if [[ -s "$SP" ]]; then
        local R
        R=$(query_sqlite "$SP" "SELECT datetime(ctime,'unixepoch'), type, COALESCE(requirement,''), COALESCE(remarks,'') FROM authority WHERE allow=1 AND (remarks IS NOT NULL AND remarks <> '') ORDER BY ctime DESC LIMIT 500")
        if [[ -n "$R" && "$R" != ERROR* ]]; then
            NAPPROVED=$(printf '%s\n' "$R" | grep -c . || true)
            APPROWS="$R"
            [[ "$NAPPROVED" -gt 0 ]] && warn "$(L "Autorizzazioni Gatekeeper con annotazione:" "Gatekeeper authorities with remarks:") ${BOLD}$NAPPROVED"
        fi
    fi

    # ---------- KextPolicy: estensioni kernel di terze parti approvate ----------
    local NKEXT=0 KEXTROWS=""
    local KP; KP=$(ci_find_file "$WIN_ROOT" "private/var/db/SystemPolicyConfiguration/KextPolicy")
    [[ -z "$KP" ]] && KP=$(ci_find_file "$WIN_ROOT" "var/db/SystemPolicyConfiguration/KextPolicy")
    if [[ -s "$KP" ]]; then
        local R
        R=$(query_sqlite "$KP" "SELECT team_id, bundle_id, allowed, COALESCE(developer_name,'') FROM kext_policy ORDER BY bundle_id")
        if [[ -n "$R" && "$R" != ERROR* ]]; then
            NKEXT=$(printf '%s\n' "$R" | grep -c . || true)
            KEXTROWS="$R"
            [[ "$NKEXT" -gt 0 ]] && warn "$(L "Estensioni kernel di terze parti registrate:" "Third-party kernel extensions registered:") ${BOLD}$NKEXT"
        fi
    fi

    # ---------- Stato di Gatekeeper ----------
    local GK; GK=$(ci_find_file "$WIN_ROOT" "private/var/db/.LastGKReject")
    [[ -n "$GK" ]] && ROWS+="Gatekeeper	$(L "ultimo rifiuto registrato" "last recorded rejection")	-	$(stat -c %y "$GK" 2>/dev/null | cut -d. -f1)
"
    local SIP; SIP=$(ci_find_file "$WIN_ROOT" "private/var/db/.AppleSetupDone")
    [[ -n "$SIP" ]] && ROWS+="$(L "Setup completato" "Setup completed")	.AppleSetupDone	-	$(stat -c %y "$SIP" 2>/dev/null | cut -d. -f1)
"

    NFIND=$(( NAPPROVED + NKEXT ))
    separator
    if [[ -z "$ROWS" && "$NFIND" -eq 0 ]]; then
        warn "$(L "Nessun artefatto XProtect/Gatekeeper leggibile." "No readable XProtect/Gatekeeper artefact.")"
        return 0
    fi
    ask_yn "Generare report HTML?" || return 0

    BODY="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "La versione delle firme cambia il significato dell'assenza di rilevamenti" "The signature version changes what 'no detections' means")</b><br>"
    BODY+="$(L "XProtect rileva solo cio' che le sue firme conoscono. Se la versione qui riportata e' molto anteriore alla data dei fatti, il fatto che non abbia segnalato nulla non e' un'assoluzione: e' semplicemente un dato privo di valore probatorio." \
        "XProtect only detects what its signatures know. If the version reported here long predates the events, the absence of detections is not an acquittal: it is simply a fact with no evidential weight.")<br><br>"
    BODY+="<b>$(L "Autorizzazioni concesse a mano" "Manually granted authorisations")</b><br>"
    BODY+="$(L "Una voce in SystemPolicy con annotazione indica un'app che qualcuno ha approvato esplicitamente superando Gatekeeper; una in KextPolicy indica un'estensione kernel di terze parti autorizzata. Entrambe richiedono un'azione deliberata dell'utente e spesso precedono immediatamente l'esecuzione di codice indesiderato." \
        "An entry in SystemPolicy with remarks indicates an app someone explicitly approved past Gatekeeper; one in KextPolicy indicates an authorised third-party kernel extension. Both require a deliberate user action and often immediately precede the execution of unwanted code.")"
    BODY+="</div></div>"

    [[ -n "$ROWS" ]] && BODY+="<div class='cards'>$(generic_card_html "$(L "Stato delle difese" "Defence status")" "$WIN_ROOT" "-" \
        "$(_rows_to_table "$ROWS" "$(L "Componente" "Component")" "$(L "Dato" "Item")" "$(L "Valore" "Value")" "$(L "Data" "Date")")" "🛡")</div>"
    [[ -n "$APPROWS" ]] && BODY+="<div class='cards'>$(generic_card_html "$(L "Autorizzazioni Gatekeeper" "Gatekeeper authorities")" "SystemPolicy" "$NAPPROVED" \
        "$(_rows_to_table "$APPROWS" "$(L "Data" "Date")" "$(L "Tipo" "Type")" "Requirement" "$(L "Annotazione" "Remarks")")" "⚑")</div>"
    [[ -n "$KEXTROWS" ]] && BODY+="<div class='cards'>$(generic_card_html "$(L "Estensioni kernel" "Kernel extensions")" "KextPolicy" "$NKEXT" \
        "$(_rows_to_table "$KEXTROWS" "Team ID" "Bundle ID" "$(L "Consentita" "Allowed")" "$(L "Sviluppatore" "Developer")")" "⚙")</div>"

    local STATS
    STATS="$(stat_box "$(L "Firme XProtect" "XProtect signatures")" "${XPV:-n/d}" "info")"
    STATS+="$(stat_box "$(L "Autorizzazioni" "Authorities")" "$NAPPROVED" "$([[ "$NAPPROVED" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "kext" "$NKEXT" "$([[ "$NKEXT" -gt 0 ]] && echo warn || echo info)")"
    finish_report "macos_xprotect" "macOS XProtect / Gatekeeper" "XPR" "XProtect · SystemPolicy · KextPolicy" "$STATS" "$BODY"
}
