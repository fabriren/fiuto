
# ================================================================
#  MODULO 48 — Chat Desktop (Slack / Teams / Discord)
#
#  Le app di messaggistica aziendale sono Electron e conservano i messaggi in
#  LevelDB, esattamente come ChatGPT Desktop (modulo 39). In un'indagine
#  contano per due ragioni:
#
#  - social engineering interno: il messaggio che ha convinto la vittima ad
#    aprire l'allegato o ad autorizzare un pagamento spesso arriva da qui,
#    non dalla posta;
#  - esfiltrazione: file e credenziali condivisi in chat privata non passano
#    dal gateway di posta e non lasciano traccia negli artefatti USB.
#
#  Il carving riusa la libreria condivisa pylib_leveldb: le stringhe
#  recuperate sono frammenti di conversazione, non un export strutturato —
#  mancano interlocutori e timestamp, e il report lo dichiara.
# ================================================================
module_chat_desktop() {
    section_header "Chat Desktop — Slack / Teams / Discord" "$MAGENTA"
    check_win_root || return 1

    local MANIFEST; MANIFEST=$(mktemp); register_tmp "$MANIFEST"
    local HOME_DIR NAPP=0
    local -a FOUND_APPS=()

    while IFS= read -r HOME_DIR; do
        local U; U=$(basename "$HOME_DIR")
        local ROAM LOCAL
        ROAM=$(ci_find_dir "$HOME_DIR" "AppData/Roaming")
        LOCAL=$(ci_find_dir "$HOME_DIR" "AppData/Local")

        # app|percorso relativo alla radice indicata
        local SPEC APP REL BASE DIR N
        for SPEC in "Slack|Slack|roam" "Discord|discord|roam" "Teams|Microsoft/Teams|roam" \
                    "Teams (new)|Packages|local"; do
            IFS='|' read -r APP REL BASE <<< "$SPEC"
            local ROOT; [[ "$BASE" == "roam" ]] && ROOT="$ROAM" || ROOT="$LOCAL"
            [[ -z "$ROOT" ]] && continue
            local APPDIR; APPDIR=$(ci_find_dir "$ROOT" "$REL")
            [[ -z "$APPDIR" ]] && continue

            # Teams "new" sta dentro Packages/MSTeams_*: si restringe la ricerca.
            if [[ "$APP" == "Teams (new)" ]]; then
                APPDIR=$(find "$ROOT" -maxdepth 1 -type d -iname 'MSTeams_*' 2>/dev/null | head -1)
                [[ -z "$APPDIR" ]] && continue
            fi

            N=0
            while IFS= read -r F; do
                [[ -s "$F" ]] || continue
                printf '%s\t%s\t%s\n' "$APP" "$U" "$F" >> "$MANIFEST"
                N=$((N + 1))
            done < <(find "$APPDIR" -maxdepth 8 -type f \( -iname '*.ldb' -o -iname '*.log' \) \
                          \( -ipath '*Local Storage*' -o -ipath '*IndexedDB*' -o -ipath '*leveldb*' \) 2>/dev/null)
            if [[ "$N" -gt 0 ]]; then
                NAPP=$((NAPP + 1))
                FOUND_APPS+=("$APP ($U): $N")
                ok "$APP — $U: ${BOLD}${N}${RESET} $(L "file LevelDB" "LevelDB files")"
            fi
        done
    done < <(get_user_homes)

    if [[ ! -s "$MANIFEST" ]]; then
        warn "$(L "Nessuna app di chat desktop con dati LevelDB trovata." "No desktop chat app with LevelDB data found.")"
        return 0
    fi
    local NFILES; NFILES=$(wc -l < "$MANIFEST")
    info "$(L "File da analizzare:" "Files to analyse:") ${BOLD}$NFILES"
    info "$(L "Carving in corso (puo' richiedere tempo)..." "Carving (may take a while)...")"

    local IOCTMP; IOCTMP=$(mktemp); register_tmp "$IOCTMP"
    printf '%s\n' "${IOC_LIST[@]:-}" > "$IOCTMP"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    run_py_with_lib pylib_leveldb "$MANIFEST" "$IOCTMP" "$OUT" << 'PYEOF' 2>/dev/null
import sys, os

manifest, ioc_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]

try:
    iocs = [l.strip().lower() for l in open(ioc_path, encoding='utf-8', errors='replace') if l.strip()]
except Exception:
    iocs = []

# Termini che, dentro una chat aziendale, meritano di essere portati in cima.
SENSITIVE = ('password', 'passwd', 'credenziali', 'credential', 'token', 'api key',
             'apikey', 'secret', 'iban', 'bonifico', 'wire transfer', 'urgente',
             'urgent', 'riservato', 'confidential', 'vpn', 'rdp', '2fa', 'otp',
             'codice di verifica', 'verification code', 'invoice', 'fattura')

rows = []
seen = set()
for line in open(manifest, encoding='utf-8', errors='replace'):
    parts = line.rstrip('\n').split('\t')
    if len(parts) != 3:
        continue
    app, user, path = parts
    try:
        items = carve_leveldb(path, cap=1200)
    except Exception:
        continue
    src = os.path.basename(path)
    for _, role, text in items:
        if role in ('error', 'info'):
            continue
        t = text.strip()
        if not t:
            continue
        key = (app, user, t.lower()[:90])
        if key in seen:
            continue
        seen.add(key)
        low = t.lower()
        flags = []
        if any(s in low for s in SENSITIVE):
            flags.append('SENSIBILE')
        if iocs and any(i in low for i in iocs):
            flags.append('IOC')
        rows.append((app, user, role, t[:1500], ';'.join(flags), src))

# Prima i messaggi segnalati.
rows.sort(key=lambda r: (r[4] == '', r[0], r[1]))
with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write('\t'.join(x.replace('\t', ' ').replace('\n', ' ') for x in r) + '\n')
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun frammento di conversazione recuperato." "No conversation fragment recovered.")"
        return 0
    fi
    local NFLAG; NFLAG=$(awk -F'\t' '$5!=""' "$OUT" | wc -l)

    ok "$(L "Frammenti recuperati:" "Fragments recovered:") ${BOLD}$TOTAL"
    if [[ "$NFLAG" -gt 0 ]]; then
        warn "$(L "Frammenti segnalati (termini sensibili o IoC):" "Flagged fragments (sensitive terms or IoC):") ${BOLD}$NFLAG"
        awk -F'\t' '$5!=""{printf "      [%s] %s — %s\n", $5, $1, substr($4,1,80)}' "$OUT" | head -15 | while IFS= read -r LN; do
            echo -e "      ${MAGENTA}${LN}${RESET}"
        done
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6}' "$OUT" | head -20000)
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "App" "$(L "Utente" "User")" "$(L "Ruolo" "Role")" "$(L "Testo" "Text")" \
        "$(L "Segnalazioni" "Flags")" "$(L "Origine" "Source")")

    local NOTE="<div class='card' style='margin-bottom:1rem'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
    NOTE+="<b>$(L "Cosa sono questi frammenti" "What these fragments are")</b><br>"
    NOTE+="$(L "Sono stringhe carvate dai database LevelDB delle app: NON sono un export strutturato della conversazione. Mancano interlocutore, canale e timestamp, l'ordine non e' quello cronologico e i messaggi possono risultare troncati. Servono a stabilire che una conversazione e' avvenuta e cosa conteneva, non a ricostruirne il filo." \
        "These are strings carved from the apps' LevelDB stores: they are NOT a structured conversation export. Counterpart, channel and timestamp are missing, the order is not chronological and messages may be truncated. They establish that a conversation happened and what it contained, not its thread.")<br><br>"
    NOTE+="$(L "Il ruolo indicato e' dedotto dalla presenza di markdown nel testo: e' un'euristica, non un dato del formato." \
        "The role shown is inferred from markdown present in the text: it is a heuristic, not a field of the format.")"
    NOTE+="</div></div>"

    local STATS
    STATS="$(stat_box "$(L "Frammenti" "Fragments")" "$TOTAL")"
    STATS+="$(stat_box "App" "$NAPP" "info")"
    STATS+="$(stat_box "$(L "Segnalati" "Flagged")" "$NFLAG" "$([[ "$NFLAG" -gt 0 ]] && echo warn || echo info)")"
    finish_report "chat_desktop" "Chat Desktop" "CHT" "Slack · Teams · Discord (LevelDB)" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Frammenti di conversazione" "Conversation fragments")" "$(L "segnalati in testa" "flagged first")" "$TOTAL" "$TABLE" "💬")</div>"
}
