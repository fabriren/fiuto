
# ================================================================
#  TEMPO — timezone del volume e finestra di analisi
#
#  Due problemi distinti, spesso confusi.
#
#  1) IN CHE FUSO SONO LE DATE CHE LEGGO. Gli artefatti non concordano: gli
#     hive di registro e i log Windows portano quasi sempre UTC, syslog e le
#     shell history portano l'ora locale della macchina, i database SQLite dei
#     browser dipendono dal browser. Convertire tutto a UTC richiederebbe di
#     sapere, artefatto per artefatto, quale delle due si sta leggendo: una
#     conversione applicata alla cieca produrrebbe date sbagliate di qualche
#     ora, che e' molto peggio di date dichiarate ambigue. Qui quindi NON si
#     converte nulla: si rileva il fuso del volume e lo si dichiara, cosi'
#     l'analista sa cosa significa un'ora locale trovata in un report.
#
#  2) QUANTE DATE DEVO GUARDARE. Su un disco da un terabyte un report puo'
#     contenere centinaia di migliaia di righe che coprono anni, mentre
#     l'incidente sta in tre giorni. --since/--until tagliano tutto cio' che
#     cade fuori dalla finestra.
#
#  Il filtro confronta le date COSI' COME COMPAIONO nell'artefatto, senza
#  riportarle a un fuso comune: e' l'unico confronto che non introduce errori
#  inventati. Su una finestra di giorni la differenza e' irrilevante; se la
#  finestra e' di ore, va allargata di un margine pari all'offset del volume.
#  Il report lo dichiara esplicitamente.
# ================================================================

TIME_SINCE=""          # limite inferiore normalizzato (YYYY-MM-DDTHH:MM:SS)
TIME_UNTIL=""          # limite superiore normalizzato
VOLUME_TZ=""           # fuso del volume analizzato, se rilevato
VOLUME_TZ_SOURCE=""    # da dove e' stato letto (serve a poterlo contestare)

# Directory di stato: come per custody e hive, il percorso e' deterministico
# perche' il conteggio delle righe scartate viene incrementato da subshell.
_time_dir() {
    local D="${TMPDIR:-/tmp}/fiuto_time_$$"
    [[ -d "$D" ]] || mkdir -p "$D" 2>/dev/null || return 1
    echo "$D"
}

# --- normalizzazione dei limiti ---------------------------------------------

# parse_time_bound <stringa> [end]
# Accetta: YYYY-MM-DD, YYYY-MM-DD HH:MM[:SS], la stessa con 'T' e/o 'Z' finale,
# e le forme relative -7d / -36h / -90m rispetto all'istante di avvio.
# Con "end" una data senza orario diventa fine giornata invece che inizio,
# cosi' --until 2026-03-01 include il 1 marzo per intero: e' quello che
# chiunque si aspetta scrivendolo.
parse_time_bound() {
    local RAW="$1" KIND="${2:-start}"
    [[ -n "$RAW" ]] || return 1

    # Forme relative: -7d, -36h, -90m
    if [[ "$RAW" =~ ^-([0-9]+)([dhm])$ ]]; then
        local N="${BASH_REMATCH[1]}" U="${BASH_REMATCH[2]}" SPEC BSD
        # Le unita' di date(1) BSD non coincidono con le nostre: 'm' li' sono
        # MESI, i minuti sono 'M'. Usare la stessa lettera darebbe una finestra
        # sbagliata di ordini di grandezza, in silenzio.
        case "$U" in
            d) SPEC="$N days ago";    BSD="-${N}d" ;;
            h) SPEC="$N hours ago";   BSD="-${N}H" ;;
            m) SPEC="$N minutes ago"; BSD="-${N}M" ;;
        esac
        date -d "$SPEC" "+%Y-%m-%dT%H:%M:%S" 2>/dev/null && return 0
        # date(1) BSD (macOS): non ha -d.
        date -v"$BSD" "+%Y-%m-%dT%H:%M:%S" 2>/dev/null && return 0
        return 1
    fi

    local S="${RAW%Z}"
    S="${S//T/ }"
    if [[ "$S" =~ ^([0-9]{4}-[0-9]{2}-[0-9]{2})$ ]]; then
        [[ "$KIND" == "end" ]] && { echo "${S}T23:59:59"; return 0; }
        echo "${S}T00:00:00"; return 0
    fi
    if [[ "$S" =~ ^([0-9]{4}-[0-9]{2}-[0-9]{2})\ ([0-9]{2}:[0-9]{2})$ ]]; then
        echo "${BASH_REMATCH[1]}T${BASH_REMATCH[2]}:$([[ "$KIND" == "end" ]] && echo 59 || echo 00)"
        return 0
    fi
    if [[ "$S" =~ ^([0-9]{4}-[0-9]{2}-[0-9]{2})\ ([0-9]{2}:[0-9]{2}:[0-9]{2})$ ]]; then
        echo "${BASH_REMATCH[1]}T${BASH_REMATCH[2]}"; return 0
    fi
    return 1
}

# Vero se e' stata impostata almeno una delle due estremita'.
time_window_active() { [[ -n "$TIME_SINCE" || -n "$TIME_UNTIL" ]]; }

# Descrizione leggibile della finestra, per console e report.
time_window_label() {
    time_window_active || return 0
    if [[ -n "$TIME_SINCE" && -n "$TIME_UNTIL" ]]; then
        echo "${TIME_SINCE/T/ } → ${TIME_UNTIL/T/ }"
    elif [[ -n "$TIME_SINCE" ]]; then
        echo "$(L "dal" "from") ${TIME_SINCE/T/ }"
    else
        echo "$(L "fino al" "until") ${TIME_UNTIL/T/ }"
    fi
}

# --- rilevamento del fuso del volume ----------------------------------------

# Legge il fuso dal volume analizzato. Non converte niente: serve a dichiarare
# cosa significano le ore locali che compaiono nei report.
detect_volume_timezone() {
    VOLUME_TZ=""; VOLUME_TZ_SOURCE=""
    [[ -n "$WIN_ROOT" ]] || return 0

    case "$OS_TYPE" in
        linux)
            local TZF; TZF=$(ci_find_file "$WIN_ROOT/etc" "timezone" 2>/dev/null)
            if [[ -n "$TZF" && -s "$TZF" ]]; then
                VOLUME_TZ=$(head -1 "$TZF" | tr -d '[:space:]')
                VOLUME_TZ_SOURCE="/etc/timezone"
            fi
            # Su systemd /etc/timezone puo' mancare: il fuso e' allora nel
            # target del symlink /etc/localtime.
            if [[ -z "$VOLUME_TZ" && -L "$WIN_ROOT/etc/localtime" ]]; then
                local T; T=$(readlink "$WIN_ROOT/etc/localtime" 2>/dev/null)
                VOLUME_TZ="${T##*/zoneinfo/}"
                VOLUME_TZ_SOURCE="/etc/localtime"
            fi
            ;;
        macos)
            if [[ -L "$WIN_ROOT/etc/localtime" || -L "$WIN_ROOT/private/etc/localtime" ]]; then
                local T
                T=$(readlink "$WIN_ROOT/etc/localtime" 2>/dev/null \
                    || readlink "$WIN_ROOT/private/etc/localtime" 2>/dev/null)
                VOLUME_TZ="${T##*/zoneinfo/}"
                VOLUME_TZ_SOURCE="/etc/localtime"
            fi
            ;;
        windows)
            local HIVE; HIVE=$(get_hive SYSTEM 2>/dev/null)
            [[ -n "$HIVE" && -f "$HIVE" ]] || return 0
            local OUT
            OUT=$("$PY3" - "$HIVE" << 'PYEOF' 2>/dev/null
import sys
try:
    from regipy.registry import RegistryHive
except Exception:
    sys.exit(0)

try:
    hive = RegistryHive(sys.argv[1])
except Exception:
    sys.exit(0)

# Il ControlSet corrente non e' sempre il 001: su una macchina che ha avuto un
# avvio fallito la differenza c'e' davvero.
sel = 1
try:
    for v in hive.get_key('\\Select').get_values():
        if v['name'].lower() == 'current':
            sel = int(v['value'])
except Exception:
    pass

name = bias = std_bias = None
for cs in (f'ControlSet{sel:03d}', 'ControlSet001', 'CurrentControlSet'):
    try:
        vals = hive.get_key(f'\\{cs}\\Control\\TimeZoneInformation').get_values()
    except Exception:
        continue
    for v in vals:
        n = v['name'].lower()
        if n == 'timezonekeyname' and v['value']:
            name = str(v['value']).rstrip('\x00').strip()
        elif n == 'bias':
            bias = v['value']
        elif n == 'standardbias':
            std_bias = v['value']
    if name or bias is not None:
        break

if not name and bias is None:
    sys.exit(0)

# Bias e' in minuti da sottrarre all'ora locale per ottenere UTC: il segno e'
# quindi invertito rispetto all'offset che si scrive di solito (UTC+1 -> -60).
off = ''
if isinstance(bias, int):
    total = -(bias + (std_bias if isinstance(std_bias, int) else 0))
    sign = '+' if total >= 0 else '-'
    total = abs(total)
    off = f'UTC{sign}{total // 60:02d}:{total % 60:02d}'
print('\t'.join([name or '', off]))
PYEOF
            )
            local TZNAME TZOFF
            IFS=$'\t' read -r TZNAME TZOFF <<< "$OUT"
            if [[ -n "${TZNAME:-}" || -n "${TZOFF:-}" ]]; then
                VOLUME_TZ="${TZNAME:-?}${TZOFF:+ (${TZOFF})}"
                VOLUME_TZ_SOURCE="SYSTEM\\Control\\TimeZoneInformation"
            fi
            ;;
    esac
    [[ -n "$VOLUME_TZ" ]] && log_msg "[TZ] volume: $VOLUME_TZ (${VOLUME_TZ_SOURCE})"
    return 0
}

# --- filtro ------------------------------------------------------------------

# Somma le righe scartate dal filtro, per poterlo dichiarare a fine sessione.
time_filtered_add() {
    local N="${1:-0}"
    [[ "$N" -gt 0 ]] 2>/dev/null || return 0
    local D; D=$(_time_dir) || return 0
    echo "$N" >> "${D}/dropped" 2>/dev/null || true
}

time_filtered_total() {
    local D="${TMPDIR:-/tmp}/fiuto_time_$$"
    [[ -f "${D}/dropped" ]] || { echo 0; return 0; }
    awk '{s+=$1} END{print s+0}' "${D}/dropped" 2>/dev/null || echo 0
}

# Nota HTML da inserire nei report quando la finestra e' attiva. Un report
# filtrato che non dichiara di esserlo e' una trappola: chi lo legge conclude
# che prima di quella data non e' successo niente.
time_window_html() {
    time_window_active || return 0
    printf "<div class='card' style='margin-bottom:1rem'><div style='padding:.9rem 1.5rem;font-size:.8rem;line-height:1.7'><b>%s</b> %s<br>%s%s</div></div>" \
        "$(L "Report filtrato per data:" "Report filtered by date:")" \
        "$(html_esc "$(time_window_label)")" \
        "$(L "Le righe con una data fuori dalla finestra non compaiono. Le righe prive di data sono state mantenute: non erano valutabili." \
             "Rows carrying a date outside the window are not shown. Rows without a date were kept: they could not be evaluated.")" \
        "$([[ -n "$VOLUME_TZ" ]] && printf " %s" "$(L "Il confronto usa le date come compaiono nell'artefatto, senza riportarle a un fuso comune; il volume risulta configurato su" "The comparison uses dates as they appear in the artefact, without normalising them to a common zone; the volume is configured for") $(html_esc "$VOLUME_TZ").")"
}
