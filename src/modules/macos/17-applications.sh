
# ================================================================
#  macOS 17 — Inventario applicazioni
#
#  Serve a rispondere a "cosa era installato e da dove veniva". Su macOS un
#  bundle .app e' una directory: puo' essere copiato ovunque ed eseguito senza
#  installazione, quindi le applicazioni fuori da /Applications meritano
#  attenzione — soprattutto quelle in ~/Downloads, /tmp o /Users/Shared.
#
#  Nota sulla firma: verificarla richiede `codesign`, disponibile solo su un
#  host macOS. Da Linux si puo' stabilire se il bundle CONTIENE una firma
#  (_CodeSignature) e se e' notarizzato in modo evidente, non se la firma sia
#  valida. Il report distingue le due cose invece di lasciarlo intendere.
# ================================================================
module_macos_applications() {
    section_header "macOS — $(L "Inventario applicazioni" "Application inventory")" "$GREEN"
    check_target_root || return 1

    local -a ROOTS=()
    local D
    for D in "Applications" "Applications/Utilities" "System/Applications" "Library/Application Support"; do
        local R; R=$(ci_find_dir "$WIN_ROOT" "$D")
        [[ -n "$R" ]] && ROOTS+=("$R")
    done
    local HOME_DIR
    while IFS= read -r HOME_DIR; do
        ROOTS+=("$HOME_DIR")
    done < <(get_macos_user_homes)

    if [[ ${#ROOTS[@]} -eq 0 ]]; then
        warn "$(L "Nessuna directory di applicazioni trovata." "No application directory found.")"
        return 0
    fi

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    info "$(L "Enumerazione dei bundle..." "Enumerating bundles...")"

    local R APP
    for R in "${ROOTS[@]}"; do
        while IFS= read -r APP; do
            [[ -d "$APP" ]] || continue
            local NAME PLIST BID VER SIGNED NOTAR MT LOC
            NAME=$(basename "$APP")
            MT=$(stat -c %y "$APP" 2>/dev/null | cut -d. -f1)
            PLIST="$APP/Contents/Info.plist"
            BID=""; VER=""
            if [[ -s "$PLIST" ]]; then
                local TXT; TXT=$(read_plist "$PLIST" 2>/dev/null | head -400)
                BID=$(printf '%s' "$TXT" | grep -aoE 'CFBundleIdentifier[^A-Za-z0-9]*[A-Za-z0-9.-]+' | grep -oE '[A-Za-z0-9.-]+$' | head -1)
                VER=$(printf '%s' "$TXT" | grep -aoE 'CFBundleShortVersionString[^0-9]*[0-9][0-9A-Za-z._-]*' | grep -oE '[0-9][0-9A-Za-z._-]*$' | head -1)
            fi
            # Presenza della firma, NON sua validita'.
            if [[ -d "$APP/Contents/_CodeSignature" ]]; then
                SIGNED="$(L "firma presente" "signature present")"
            else
                SIGNED="$(L "NESSUNA FIRMA" "NO SIGNATURE")"
            fi
            # La ricevuta App Store distingue l'origine ufficiale.
            if [[ -f "$APP/Contents/_MASReceipt/receipt" ]]; then
                NOTAR="App Store"
            else
                NOTAR="-"
            fi
            # Posizione: fuori dalle directory di sistema e' il dato che conta.
            # NB: */Applications/* copre gia' anche */System/Applications/*.
            case "$APP" in
                */Applications/*) LOC="$(L "standard" "standard")" ;;
                *)                LOC="$(L "FUORI DA /Applications" "OUTSIDE /Applications")" ;;
            esac
            printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
                "$MT" "$NAME" "${BID:-?}" "${VER:-?}" "$SIGNED" "$NOTAR" "$LOC" "$APP" >> "$OUT"
        done < <(find "$R" -maxdepth 4 -type d -name '*.app' -prune 2>/dev/null)
    done

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun bundle applicativo trovato." "No application bundle found.")"
        return 0
    fi
    local NUNSIGNED NOUTSIDE
    NUNSIGNED=$(awk -F'\t' '$5 ~ /NESSUNA|NO SIGNATURE/' "$OUT" | wc -l)
    NOUTSIDE=$(awk -F'\t' '$7 ~ /FUORI|OUTSIDE/' "$OUT" | wc -l)

    ok "$(L "Applicazioni trovate:" "Applications found:") ${BOLD}$TOTAL"
    [[ "$NUNSIGNED" -gt 0 ]] && warn "$(L "Senza firma nel bundle:" "Without a signature in the bundle:") ${BOLD}$NUNSIGNED"
    if [[ "$NOUTSIDE" -gt 0 ]]; then
        warn "$(L "Fuori dalle directory standard:" "Outside standard directories:") ${BOLD}$NOUTSIDE"
        awk -F'\t' '$7 ~ /FUORI|OUTSIDE/{printf "      %s  %s\n", $1, $8}' "$OUT" | head -15 | while IFS= read -r LN; do
            echo -e "      ${YELLOW}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$( { awk -F'\t' '$7 ~ /FUORI|OUTSIDE/' "$OUT"; awk -F'\t' '$5 ~ /NESSUNA|NO SIGNATURE/ && $7 !~ /FUORI|OUTSIDE/' "$OUT"; awk -F'\t' '$5 !~ /NESSUNA|NO SIGNATURE/ && $7 !~ /FUORI|OUTSIDE/' "$OUT"; } )
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "$(L "Ultima modifica" "Last modified")" "$(L "Applicazione" "Application")" "Bundle ID" \
        "$(L "Versione" "Version")" "$(L "Firma" "Signature")" "$(L "Origine" "Origin")" \
        "$(L "Posizione" "Location")" "$(L "Percorso" "Path")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Cosa dice davvero la colonna Firma" "What the Signature column actually says")</b><br>"
    NOTE+="$(L "Indica se il bundle contiene una directory _CodeSignature, non se la firma sia valida: la verifica richiede codesign, disponibile solo su un host macOS. Un'app puo' avere una firma presente ma non valida, revocata o di uno sviluppatore qualsiasi. Per la verifica: codesign -dv --verbose=4 /percorso/App.app e spctl -a -vv /percorso/App.app." \
        "It indicates whether the bundle contains a _CodeSignature directory, not whether the signature is valid: verification requires codesign, available only on a macOS host. An app may carry a present but invalid or revoked signature, or one from any developer. To verify: codesign -dv --verbose=4 /path/App.app and spctl -a -vv /path/App.app.")<br><br>"
    NOTE+="$(L "Un bundle .app e' una directory: puo' essere copiato ed eseguito da qualunque posizione senza installazione. Le applicazioni fuori dalle directory standard sono elencate per prime." \
        "An .app bundle is a directory: it can be copied and run from anywhere without installation. Applications outside the standard directories are listed first.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Applicazioni" "Applications")" "$TOTAL")"
    STATS+="$(stat_box "$(L "Senza firma" "Unsigned")" "$NUNSIGNED" "$([[ "$NUNSIGNED" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Fuori standard" "Non-standard")" "$NOUTSIDE" "$([[ "$NOUTSIDE" -gt 0 ]] && echo warn || echo info)")"
    finish_report "macos_applications" "macOS Applications" "APP" "/Applications · home utenti" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Applicazioni installate" "Installed applications")" "$WIN_ROOT" "$TOTAL" "$TABLE" "▣")</div>"
}
