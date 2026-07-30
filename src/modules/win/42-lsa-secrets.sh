
# ================================================================
#  MODULO 42 — LSA Secrets & Cached Domain Credentials (DCC2)
#
#  Il modulo 20 estrae gli hash degli account LOCALI dall'hive SAM. L'hive
#  SECURITY contiene altro, e spesso di piu' valore:
#
#  - LSA Secrets: password IN CHIARO degli account di servizio, delle
#    scheduled task configurate con credenziali, del DefaultPassword di
#    autologon, delle VPN e dei trust di dominio.
#  - Cached Domain Credentials (DCC2/MSCACHE v2): hash degli ultimi utenti di
#    dominio autenticatisi sulla macchina — inclusi amministratori di dominio
#    che si sono collegati a una workstation compromessa.
#
#  Su una workstation membro di dominio questo modulo e' spesso il punto in cui
#  si capisce fin dove l'attaccante poteva arrivare.
# ================================================================
module_lsa_secrets() {
    section_header "LSA Secrets & Cached Domain Credentials" "$RED"
    check_win_root || return 1

    local SEC_HIVE SYS_HIVE
    SEC_HIVE=$(get_hive SECURITY)
    SYS_HIVE=$(get_hive SYSTEM)

    if [[ -z "$SEC_HIVE" ]]; then
        warn "$(L "Hive SECURITY non trovato." "SECURITY hive not found.")"
        return 0
    fi
    if [[ -z "$SYS_HIVE" ]]; then
        warn "$(L "Hive SYSTEM non trovato: serve per la bootkey." "SYSTEM hive not found: required for the boot key.")"
        return 0
    fi
    info "SECURITY: ${DIM}${SEC_HIVE}${RESET}"
    info "SYSTEM:   ${DIM}${SYS_HIVE}${RESET}"

    if ! "$PY3" -c "import impacket" 2>/dev/null; then
        warn "$(L "impacket non disponibile: LSA Secrets e DCC2 non estraibili." \
                 "impacket unavailable: LSA Secrets and DCC2 cannot be extracted.")"
        info "$(L "Installalo con:" "Install it with:") ${PY3} -m pip install impacket"
        return 0
    fi

    info "$(L "Estrazione in corso (impacket)..." "Extracting (impacket)...")"
    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$SYS_HIVE" "$SEC_HIVE" > "$OUT" << 'PYEOF' 2>/dev/null
import sys, io, os, contextlib

system_hive, security_hive = sys.argv[1], sys.argv[2]

def L(it, en):
    return it if os.environ.get('LANG', 'en') == 'it' else en

try:
    from impacket.examples.secretsdump import LocalOperations, LSASecrets
except Exception as exc:
    print(f"ERROR\t{L('impacket non importabile', 'impacket not importable')}: {exc}")
    sys.exit(0)

try:
    ops = LocalOperations(system_hive)
    boot_key = ops.getBootKey()
except Exception as exc:
    print(f"ERROR\t{L('bootkey non ricavabile dall hive SYSTEM', 'boot key not derivable from SYSTEM hive')}: {exc}")
    sys.exit(0)

# impacket scrive i risultati su stdout: li si cattura per classificarli
# invece di lasciarli finire grezzi nel report.
buf = io.StringIO()
try:
    lsa = LSASecrets(security_hive, boot_key, None, isRemote=False, perSecretCallback=lambda *a: None)
    with contextlib.redirect_stdout(buf):
        try:
            lsa.dumpCachedHashes()
        except Exception:
            pass
        try:
            lsa.dumpSecrets()
        except Exception:
            pass
except Exception as exc:
    print(f"ERROR\t{L('estrazione fallita', 'extraction failed')}: {exc}")
    sys.exit(0)

captured = buf.getvalue().splitlines()

# I DCC2 hanno forma  utente/DOMINIO:$DCC2$10240#utente#hash
# Gli LSA secret sono blocchi "NOME\n(valore)".
for line in captured:
    line = line.rstrip()
    if not line.strip():
        continue
    low = line.lower()
    if '$dcc2$' in low:
        kind = 'DCC2'
    elif line.startswith('[*]') or line.startswith('[-]'):
        continue
    else:
        kind = 'LSA'
    print(f"{kind}\t{line}")
PYEOF

    if grep -q '^ERROR' "$OUT" 2>/dev/null; then
        warn "$(sed -n 's/^ERROR\t//p' "$OUT" | head -1)"
        return 0
    fi

    local NDCC NLSA TOTAL
    NDCC=$(awk -F'\t' '$1=="DCC2"' "$OUT" | wc -l)
    NLSA=$(awk -F'\t' '$1=="LSA"' "$OUT" | wc -l)
    TOTAL=$((NDCC + NLSA))
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun segreto estratto (hive vuoto o non decifrabile)." "No secret extracted (empty or undecryptable hive).")"
        return 0
    fi

    ok "$(L "Cached domain credentials (DCC2):" "Cached domain credentials (DCC2):") ${BOLD}$NDCC"
    ok "$(L "LSA secrets:" "LSA secrets:") ${BOLD}$NLSA"
    [[ "$NDCC" -gt 0 ]] && warn "$(L "I DCC2 sono craccabili offline: hashcat -m 2100" "DCC2 are crackable offline: hashcat -m 2100")"

    ask_yn "Generare report HTML?" || return 0

    local BODY=""
    BODY+="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    BODY+="<b>$(L "Materiale altamente sensibile" "Highly sensitive material")</b><br>"
    BODY+="$(L "Questo report puo' contenere password in chiaro di account di servizio e hash di credenziali di dominio. Trattalo come materiale riservato del caso: non allegarlo a ticket, non condividerlo su canali non cifrati." \
        "This report may contain cleartext service-account passwords and domain credential hashes. Treat it as restricted case material: do not attach it to tickets or share it over unencrypted channels.")<br><br>"
    BODY+="$(L "Cracking offline:" "Offline cracking:") <code>hashcat -m 2100 dcc2.txt wordlist.txt</code>"
    BODY+="</div></div>"

    if [[ "$NDCC" -gt 0 ]]; then
        local R; R=$(awk -F'\t' '$1=="DCC2"{print $2}' "$OUT")
        BODY+="<div class='cards'>$(generic_card_html "Cached Domain Credentials (DCC2)" "$SEC_HIVE" "$NDCC" "$(_rows_to_table "$R" "$(L "Voce" "Entry")")" "⚿")</div>"
    fi
    if [[ "$NLSA" -gt 0 ]]; then
        local R2; R2=$(awk -F'\t' '$1=="LSA"{print $2}' "$OUT")
        BODY+="<div class='cards'>$(generic_card_html "LSA Secrets" "$SEC_HIVE" "$NLSA" "$(_rows_to_table "$R2" "$(L "Voce" "Entry")")" "⚿")</div>"
    fi

    local STATS
    STATS="$(stat_box "DCC2" "$NDCC" "$([[ "$NDCC" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "LSA Secrets" "$NLSA" "$([[ "$NLSA" -gt 0 ]] && echo warn || echo info)")"
    finish_report "lsa_secrets" "LSA Secrets & DCC2" "LSA" "Windows/System32/config/SECURITY" "$STATS" "$BODY"
}
