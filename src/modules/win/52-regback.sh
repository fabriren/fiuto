
# ================================================================
#  MODULO 52 — RegBack (backup periodico degli hive di registro)
#
#  %windir%\System32\config\RegBack contiene copie di SOFTWARE, SYSTEM, SAM,
#  SECURITY e DEFAULT, rinfrescate ogni ~10 giorni dal task pianificato
#  RegIdleBackup. Dalla 1803 Microsoft lo disabilita sui client (file a 0
#  byte), resta attivo di default sui server.
#
#  Il valore per il DFIR: e' una fotografia del registro precedente
#  all'ultimo rinfresco. Confrontarla con l'hive vivo isola cosa e' cambiato
#  in quella finestra — NON "da prima dell'incidente a oggi": se l'attaccante
#  e' rimasto piu' a lungo dell'intervallo di backup, o se RegIdleBackup e'
#  girato DOPO la compromissione, la copia contiene gia' le sue modifiche.
#  E' un bonus temporalmente delimitato, non una baseline pulita garantita —
#  lo stesso principio di onesta' di MODULO 43 (VSS) vale qui.
#
#  Ambito scelto deliberatamente: policy Defender (il caso che ha motivato
#  il modulo: un DisableAntiSpyware comparso senza un impianto di persistenza
#  visibile altrove), Run/RunOnce, Winlogon su SOFTWARE; il flag
#  EnablePeriodicBackup su SYSTEM (letto solo dal vivo: la sua assenza o il
#  suo azzeramento e' un indicatore anti-forense indipendente dal diff); il
#  solo INSIEME DEI NOMI account su SAM, mai gli hash. Un differ generico
#  sull'intero hive produrrebbe centinaia di righe di rumore (timestamp,
#  contatori, GUID che cambiano ad ogni boot) e seppellirebbe il segnale.
# ================================================================
module_regback() {
    section_header "RegBack" "$CYAN"
    check_win_root || return 1

    if ! check_regipy; then
        warn "$(L "regipy non trovato per" "regipy not found for") ${PY3} (${PY3_VERSION})"
        warn "$(L "Installalo con: pip install regipy  oppure" "Install it with: pip install regipy  or")  ${PY3} -m pip install regipy"
        return 1
    fi

    local REGBACK_DIR
    REGBACK_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/config/RegBack")

    local -a HIVE_NAMES=(SOFTWARE SYSTEM SAM SECURITY DEFAULT)
    local -A LIVE_PATH REGBACK_PATH REGBACK_OK
    local INV_ROWS="" USABLE=0 NEWEST_EPOCH=0

    if [[ -z "$REGBACK_DIR" ]]; then
        warn "$(L "Cartella RegBack non trovata: nessuna baseline di registro disponibile." \
                 "RegBack folder not found: no registry baseline available.")"
    fi

    local H
    for H in "${HIVE_NAMES[@]}"; do
        LIVE_PATH[$H]=$(get_hive "$H")
        REGBACK_PATH[$H]=""
        REGBACK_OK[$H]=false

        local RB_RAW SIZE=0 MTIME="-"
        if [[ -n "$REGBACK_DIR" ]]; then
            RB_RAW=$(ci_find_file "$REGBACK_DIR" "$H")
            if [[ -n "$RB_RAW" ]]; then
                SIZE=$(stat -c %s "$RB_RAW" 2>/dev/null || echo 0)
                if [[ "$SIZE" -gt 0 ]]; then
                    REGBACK_PATH[$H]=$(recover_hive "$RB_RAW")
                    REGBACK_OK[$H]=true
                    USABLE=$((USABLE + 1))
                    MTIME=$(stat -c '%y' "$RB_RAW" 2>/dev/null | cut -d. -f1)
                    local EPOCH; EPOCH=$(stat -c %Y "$RB_RAW" 2>/dev/null || echo 0)
                    [[ "$EPOCH" -gt "$NEWEST_EPOCH" ]] && NEWEST_EPOCH=$EPOCH
                else
                    MTIME=$(stat -c '%y' "$RB_RAW" 2>/dev/null | cut -d. -f1)
                fi
            fi
        fi

        local STATUS
        if [[ -z "${RB_RAW:-}" ]]; then
            STATUS=$(L "assente" "missing")
        elif [[ "$SIZE" -eq 0 ]]; then
            STATUS=$(L "disabilitato (0 byte, default client dalla 1803)" "disabled (0 bytes, client default since 1803)")
        else
            STATUS=$(L "disponibile" "available")
        fi
        INV_ROWS+="${H}	$([[ -n "${LIVE_PATH[$H]}" ]] && echo "$(L "si" "yes")" || echo no)	$(numfmt --to=iec "$SIZE" 2>/dev/null || echo "$SIZE B")	${MTIME}	${STATUS}
"
        unset RB_RAW
    done

    separator
    if [[ "$USABLE" -eq 0 ]]; then
        warn "$(L "Nessun hive RegBack utilizzabile: nessun confronto possibile su questo volume." \
                 "No usable RegBack hive: no comparison possible on this volume.")"
    else
        ok "$(L "Hive RegBack utilizzabili:" "Usable RegBack hives:") ${BOLD}${USABLE}${RESET}/${#HIVE_NAMES[@]}"
        local AGE_DAYS=$(( ( $(date +%s) - NEWEST_EPOCH ) / 86400 ))
        if [[ "$AGE_DAYS" -gt 15 ]]; then
            warn "$(L "Backup piu' recente di $AGE_DAYS giorni fa: oltre il ciclo di ~10 giorni di RegIdleBackup. Verifica che il task sia ancora attivo (modulo Scheduled Tasks)." \
                     "Most recent backup is $AGE_DAYS days old: past RegIdleBackup's ~10-day cycle. Check the task is still active (Scheduled Tasks module).")"
        else
            info "$(L "Backup piu' recente:" "Most recent backup:") ${AGE_DAYS} $(L "giorni fa" "days ago")"
        fi
    fi

    # --- EnablePeriodicBackup, solo dal vivo -----------------------------
    # 0 = backup disattivato a monte: nessuna copia futura, indipendentemente
    # da quanto sopra. E' un controllo separato dal diff, non un suo prerequisito.
    local PERIODIC_STATUS="?"
    if [[ -n "${LIVE_PATH[SYSTEM]}" ]]; then
        PERIODIC_STATUS=$("$PY3" - "${LIVE_PATH[SYSTEM]}" << 'PYEOF' 2>/dev/null
import sys
from regipy.registry import RegistryHive
try:
    hive = RegistryHive(sys.argv[1])

    # \Select\Current indica il ControlSet davvero attivo: un altro
    # ControlSet puo' aprirsi comunque (Configuration Manager e' una chiave
    # standard quasi ovunque) senza essere quello che Windows usa. Se
    # EnablePeriodicBackup e' stato azzerato solo sul set attivo, fermarsi al
    # primo che si apre - anziche' al primo che e' davvero quello giusto -
    # leggerebbe il set sbagliato e mancherebbe l'indicatore.
    candidates = []
    try:
        sel = hive.get_key('Select')
        current = next(v.value for v in sel.get_values() if v.name.lower() == 'current')
        candidates.append(f'ControlSet{int(current):03d}')
    except Exception:
        pass
    for cs in ['ControlSet001', 'ControlSet002', 'CurrentControlSet']:
        if cs not in candidates:
            candidates.append(cs)

    for cs in candidates:
        try:
            key = hive.get_key(f'{cs}\\Control\\Session Manager\\Configuration Manager')
        except Exception:
            continue
        for v in key.get_values():
            if v.name.lower() == 'enableperiodicbackup':
                print(v.value)
                sys.exit(0)
    print('default')
except Exception:
    print('?')
PYEOF
)
    fi
    case "$PERIODIC_STATUS" in
        0) warn "$(L "EnablePeriodicBackup = 0 sull'hive vivo: RegIdleBackup non produrra' nuovi backup finche' resta cosi'. Possibile indicatore anti-forense." \
                     "EnablePeriodicBackup = 0 on the live hive: RegIdleBackup will not produce new backups while this stands. Possible anti-forensic indicator.")" ;;
        1) info "$(L "EnablePeriodicBackup = 1 (riattivato esplicitamente)." "EnablePeriodicBackup = 1 (explicitly re-enabled).")" ;;
        default) info "$(L "EnablePeriodicBackup non impostato (comportamento di default della piattaforma)." "EnablePeriodicBackup not set (platform default behaviour).")" ;;
        *) warn "$(L "Impossibile verificare EnablePeriodicBackup sull'hive SYSTEM vivo." "Could not check EnablePeriodicBackup on the live SYSTEM hive.")" ;;
    esac

    # --- Diff SOFTWARE: policy Defender, Run/RunOnce, Winlogon ----------
    local -a DIFF_ROWS=()  # campi separati da \x1f: Hive, Chiave, Nome, Live, RegBack, Verdetto
    local HIVE_ERR=false

    if [[ "${REGBACK_OK[SOFTWARE]}" == "true" && -n "${LIVE_PATH[SOFTWARE]}" ]]; then
        local -a SW_KEYS=(
            "Policies\\Microsoft\\Windows Defender"
            "Policies\\Microsoft\\Windows Defender\\Real-Time Protection"
            "Microsoft\\Windows\\CurrentVersion\\Run"
            "Microsoft\\Windows\\CurrentVersion\\RunOnce"
            "Microsoft\\Windows NT\\CurrentVersion\\Winlogon"
        )
        local KEY
        for KEY in "${SW_KEYS[@]}"; do
            local OUT
            OUT=$(regipy_diff_key "${LIVE_PATH[SOFTWARE]}" "${REGBACK_PATH[SOFTWARE]}" "$KEY" 2>/dev/null)
            [[ -z "$OUT" ]] && continue
            if [[ "$OUT" == HIVE_ERROR* ]]; then
                HIVE_ERR=true
                continue
            fi
            while IFS=$'\x1f' read -r VNAME VLIVE VRB VERDICT; do
                [[ -z "$VNAME" ]] && continue
                [[ "$VERDICT" == "UNCHANGED" ]] && continue
                DIFF_ROWS+=("SOFTWARE"$'\x1f'"${KEY}"$'\x1f'"${VNAME}"$'\x1f'"${VLIVE}"$'\x1f'"${VRB}"$'\x1f'"${VERDICT}")
            done <<< "$OUT"
        done
    fi

    # --- Diff SAM: solo nomi account, nessun hash ------------------------
    if [[ "${REGBACK_OK[SAM]}" == "true" && -n "${LIVE_PATH[SAM]}" ]]; then
        local OUT
        OUT=$(regipy_diff_subkeys "${LIVE_PATH[SAM]}" "${REGBACK_PATH[SAM]}" \
            "SAM\\Domains\\Account\\Users\\Names" 2>/dev/null)
        if [[ "$OUT" == HIVE_ERROR* ]]; then
            HIVE_ERR=true
        elif [[ -n "$OUT" ]]; then
            while IFS=$'\x1f' read -r NAME VERDICT; do
                [[ -z "$NAME" ]] && continue
                local PRESENT; PRESENT=$(L "presente" "present")
                DIFF_ROWS+=("SAM"$'\x1f'"Domains\\Account\\Users\\Names"$'\x1f'"${NAME}"$'\x1f'"$([[ "$VERDICT" == "NEW" ]] && echo "$PRESENT" || echo -)"$'\x1f'"$([[ "$VERDICT" == "REMOVED" ]] && echo "$PRESENT" || echo -)"$'\x1f'"${VERDICT}")
            done <<< "$OUT"
        fi
    fi

    $HIVE_ERR && warn "$(L "Uno o piu' hive non si sono aperti correttamente: il confronto per quelle chiavi e' stato saltato, non trattato come \"nessuna differenza\"." \
                           "One or more hives failed to open: comparison for those keys was skipped, not treated as \"no difference\".")"

    local NDIFF=${#DIFF_ROWS[@]}
    echo ""
    if [[ "$NDIFF" -eq 0 ]]; then
        [[ "$USABLE" -gt 0 ]] && ok "$(L "Nessuna differenza sulle chiavi controllate." "No difference on the checked keys.")"
    else
        warn "$(L "Differenze trovate:" "Differences found:") ${BOLD}${NDIFF}${RESET}"
    fi

    ask_yn "$(L "Generare report HTML?" "Generate HTML report?")" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Cosa dice e cosa non dice un confronto con RegBack" "What a RegBack comparison does and does not say")</b><br>"
    BODY+="$(L "Una differenza indica un cambiamento avvenuto nella finestra fra l'ultimo rinfresco di RegIdleBackup e l'acquisizione, non necessariamente \"dopo l'incidente\": se la permanenza supera l'intervallo di backup, o se il task e' girato dopo la compromissione, la copia contiene gia' le modifiche dell'attaccante. Trattalo come un bonus temporalmente delimitato, non come una baseline pulita garantita." \
        "A difference marks a change within the window between RegIdleBackup's last refresh and acquisition, not necessarily \"after the incident\": if dwell time exceeded the backup interval, or the task ran after the compromise, the copy already contains the attacker's changes. Treat it as a time-bounded bonus, not a guaranteed clean baseline.")<br><br>"
    BODY+="$(L "Nota su Defender: a seconda di versione e Tamper Protection, il valore potrebbe essere ignorato dal motore. Resta comunque un indicatore forte: qualcuno ha tentato la disattivazione." \
        "Note on Defender: depending on platform version and Tamper Protection, the value may be ignored by the engine. It remains a strong indicator either way: someone attempted the disabling.")"
    BODY+="</div></div>"

    local INV_TABLE
    INV_TABLE=$(_rows_to_table "$INV_ROWS" \
        "Hive" "$(L "Vivo trovato" "Live found")" "$(L "Dim. RegBack" "RegBack size")" "$(L "Data RegBack" "RegBack date")" "$(L "Stato" "Status")")
    BODY+="<div class='stitle'>$(L "Inventario hive RegBack" "RegBack hive inventory")</div>"
    BODY+="<div class='cards'>$(generic_card_html "RegBack" "${REGBACK_DIR:-Windows/System32/config/RegBack}" "$USABLE/${#HIVE_NAMES[@]}" "$INV_TABLE" "◫")</div>"

    if [[ "$NDIFF" -gt 0 ]]; then
        local DROWS=""
        local E
        for E in "${DIFF_ROWS[@]}"; do
            IFS=$'\x1f' read -r DHIVE DKEY DNAME DLIVE DRB DVERDICT <<< "$E"
            local VLABEL="$DVERDICT" CSS="class='mono'"
            case "$DVERDICT" in
                NEW)     VLABEL=$(L "NUOVO" "NEW"); CSS="class='mono bad'" ;;
                CHANGED) VLABEL=$(L "MODIFICATO" "CHANGED"); CSS="class='mono bad'" ;;
                REMOVED) VLABEL=$(L "RIMOSSO" "REMOVED"); CSS="class='mono mid'" ;;
            esac
            DROWS+="<tr><td class='mono dim'>$(html_esc "$DHIVE")</td><td class='mono dim' style='word-break:break-all'>$(html_esc "$DKEY")</td><td class='mono'>$(html_esc "$DNAME")</td><td class='mono'>$(html_esc "$DLIVE")</td><td class='mono dim'>$(html_esc "$DRB")</td><td ${CSS}>$(html_esc "$VLABEL")</td></tr>"
        done
        BODY+="<div class='stitle' style='margin-top:1.5rem'>$(L "Differenze rispetto al backup" "Differences from the backup")</div>"
        BODY+="<div class='card'><table>
          <thead><tr><th>Hive</th><th>$(L "Chiave" "Key")</th><th>$(L "Nome" "Name")</th><th>$(L "Vivo" "Live")</th><th>RegBack</th><th>$(L "Verdetto" "Verdict")</th></tr></thead>
          <tbody>${DROWS}</tbody>
        </table></div>"
    fi

    local STATS
    STATS="$(stat_box "$(L "Hive RegBack" "RegBack hives")" "${USABLE}/${#HIVE_NAMES[@]}" "$([[ "$USABLE" -eq 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Differenze" "Differences")" "$NDIFF" "$([[ "$NDIFF" -gt 0 ]] && echo warn || echo info)")"
    finish_report "regback" "RegBack" "RB" "Windows/System32/config/RegBack" "$STATS" "$BODY"
}
