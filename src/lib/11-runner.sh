
# Esegue un modulo in modalità batch e registra il risultato in SUMMARY_TABLE.
# Durante l'esecuzione mostra il suggerimento ESC; premendo ESC il modulo viene
# interrotto e si passa al successivo.
run_batch_module() {
    local mod_num="$1"
    local mod_func="$2"
    local mod_name="$3"
    local total_mods="${4:-38}"
    local _ESC_HINT; _ESC_HINT="$(L "[ESC: salta modulo]" "[ESC: skip module]")"

    echo -ne "  ${CYAN}[*]${RESET} [${mod_num}/${total_mods}] $(L "Esecuzione modulo" "Running module") $mod_num ($mod_name)...  ${DIM}${_ESC_HINT}${RESET}\r"
    log_msg "[BATCH] Modulo $mod_num: $mod_name"

    # I report generati dal subshell vengono scritti su un file temp
    # perché gli array bash non si propagano al processo padre.
    local _REP_TMP; _REP_TMP=$(mktemp)

    (
        # Override locale: scrive il path nel file temp invece dell'array
        register_report() { [[ -n "${1:-}" && -f "$1" ]] && echo "$1" >> "$_REP_TMP"; }
        $mod_func >/dev/null 2>&1
    ) &
    local MOD_PID=$!

    # Monitoraggio tasto ESC (solo se /dev/tty è disponibile)
    local SKIPPED=0
    local _OLD_STTY=""
    if [[ -c /dev/tty ]]; then
        _OLD_STTY=$(stty -g </dev/tty 2>/dev/null) || true
        stty -echo -icanon min 0 time 0 </dev/tty 2>/dev/null || true
        while kill -0 "$MOD_PID" 2>/dev/null; do
            local KEY=""
            IFS= read -r -s -t 0.2 -N 1 KEY </dev/tty 2>/dev/null || true
            if [[ "$KEY" == $'\033' ]]; then
                kill "$MOD_PID" 2>/dev/null
                SKIPPED=1
                break
            fi
        done
        [[ -n "$_OLD_STTY" ]] && stty "$_OLD_STTY" </dev/tty 2>/dev/null || true
    fi

    wait "$MOD_PID" 2>/dev/null

    if [[ $SKIPPED -eq 1 ]]; then
        printf '\r\033[K'
        echo -e "  ${YELLOW}[⏭]${RESET} [${mod_num}/${total_mods}] $mod_name — $(L "annullato (ESC)" "cancelled (ESC)")"
        SUMMARY_TABLE+=("$mod_num|$mod_name|SKIP|$(L "annullato" "cancelled")")
        rm -f "$_REP_TMP"
        return
    fi

    # Importa i report generati nel subshell
    if [[ -s "$_REP_TMP" ]]; then
        while IFS= read -r _rep; do
            [[ -n "$_rep" && -f "$_rep" ]] && GENERATED_REPORTS+=("$_rep")
        done < "$_REP_TMP"
        local rep_path="${GENERATED_REPORTS[-1]}"
        printf '\r\033[K'
        echo -e "  ${GREEN}[✓]${RESET} [${mod_num}/${total_mods}] $mod_name — report: ${DIM}${rep_path}${RESET}"
        SUMMARY_TABLE+=("$mod_num|$mod_name|SI|$rep_path")
    else
        printf '\r\033[K'
        echo -e "  ${DIM}[i] [${mod_num}/${total_mods}] $mod_name — $(L "nessun risultato" "no results")${RESET}"
        SUMMARY_TABLE+=("$mod_num|$mod_name|NO|-")
    fi
    rm -f "$_REP_TMP"
}

# ================================================================
#  DASHBOARD "FULL" — indice navigabile con tab + iframe centrale.
#  Generata al termine di "esegui TUTTI i moduli" (Windows/Linux/macOS).
#  Costruita interamente da SUMMARY_TABLE (righe "num|nome|SI/NONE/SKIP|path").
# ================================================================
generate_full_dashboard() {
    [[ -z "$REPORT_BASE_DIR" ]] && return 0
    [[ ${#SUMMARY_TABLE[@]} -eq 0 ]] && return 0
    local OSL; OSL=$(os_label)
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")
    local DASH="${REPORT_BASE_DIR}/index.html"
    local TABS="" COUNT_OK=0 COUNT_TOTAL=0

    for row in "${SUMMARY_TABLE[@]}"; do
        IFS='|' read -r mnum mname msy mpath <<< "$row"
        [[ -z "$mnum" ]] && continue
        COUNT_TOTAL=$((COUNT_TOTAL + 1))
        local NUM2; NUM2=$(printf '%02d' "$mnum" 2>/dev/null || echo "$mnum")
        local NAME_ESC; NAME_ESC=$(html_esc "$mname")
        if [[ "$msy" == "SI" ]]; then
            COUNT_OK=$((COUNT_OK + 1))
            local rel="${mpath#$REPORT_BASE_DIR/}"
            TABS+="<button class='tab' data-src='$(html_esc "$rel")'><span class='tn'>${NUM2}</span><span class='tl'>${NAME_ESC}</span><span class='dot ok'></span></button>"
        else
            local CLS="none" LBL
            [[ "$msy" == "SKIP" ]] && { CLS="skip"; LBL="skip"; } || LBL="—"
            TABS+="<button class='tab disabled' disabled title='$([ "$msy" = "SKIP" ] && echo "$(L "saltato" "skipped")" || echo "$(L "nessuna evidenza" "no findings")")'><span class='tn'>${NUM2}</span><span class='tl'>${NAME_ESC}</span><span class='dot ${CLS}'></span></button>"
        fi
    done

    local HOST_DISP="${HOST_NAME:-N/A}"
    # Icona "naso di cane che fiuta" (SVG inline, bianco su sfondo blu)
    local NOSE_SVG='<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M19 11c2-2 5-2 7 0" fill="none" stroke="#dbeafe" stroke-width="2.4" stroke-linecap="round" opacity=".85"/><path d="M38 11c2-2 5-2 7 0" fill="none" stroke="#dbeafe" stroke-width="2.4" stroke-linecap="round" opacity=".85"/><path fill="#fff" d="M32 50C17 39 10 32 10 25c0-6 6-9 12-7 4 1 7 4 10 7 3-3 6-6 10-7 6-2 12 1 12 7 0 7-7 14-22 25Z"/><ellipse cx="23" cy="28" rx="3" ry="4.3" fill="#1f6feb"/><ellipse cx="41" cy="28" rx="3" ry="4.3" fill="#1f6feb"/><path d="M32 33v10" stroke="#1f6feb" stroke-width="2.6" stroke-linecap="round"/></svg>'
    {
        cat << HTMLEOF
<!DOCTYPE html>
<html lang="$(L "it" "en")">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FIUTO — Full Report (${OSL})</title>
<style>
  :root{ --bg:#080b0f; --bg2:#0d1117; --bg3:#131920; --bg4:#1a2332; --border:#1e2d3d;
    --accent:#58a6ff; --accent2:#ff7b72; --accent3:#3fb950; --accent4:#f0883e;
    --text:#c9d1d9; --text-dim:#3d5166; --text-mid:#6e8898;
    --mono:'Fira Code',ui-monospace,monospace; --sans:'DM Sans',system-ui,sans-serif; }
  *{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%}
  body{background:var(--bg);color:var(--text);font-family:var(--sans);display:flex;flex-direction:column;height:100vh;overflow:hidden}
  header{background:var(--bg2);border-bottom:1px solid var(--border);padding:1rem 1.6rem;display:flex;align-items:center;gap:1.4rem;flex-shrink:0}
  .hicon{width:2.6rem;height:2.6rem;background:linear-gradient(135deg,var(--accent),#1f6feb);display:flex;align-items:center;justify-content:center;flex-shrink:0;clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%)}
  .hicon svg{width:62%;height:62%}
  .htxt h1{font-size:1.1rem;color:#fff;font-weight:700}
  .htxt .sub{font-size:.66rem;color:var(--text-dim);font-family:var(--mono);margin-top:.2rem}
  .hstats{margin-left:auto;display:flex;gap:1.6rem;font-family:var(--mono)}
  .hstats .s .v{font-size:1.3rem;font-weight:800}
  .hstats .s .l{font-size:.55rem;text-transform:uppercase;letter-spacing:.12em;color:var(--text-dim)}
  .s.ok .v{color:var(--accent3)} .s.tot .v{color:var(--accent)}
  #tabs{display:flex;flex-wrap:wrap;gap:.15rem .2rem;background:var(--bg2);border-bottom:1px solid var(--border);padding:.4rem .8rem;flex-shrink:0;max-height:45vh;overflow-y:auto;align-content:flex-start}
  #tabs::-webkit-scrollbar{width:6px}
  #tabs::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
  .tab{display:flex;align-items:center;gap:.4rem;white-space:nowrap;background:var(--bg3);border:1px solid var(--border);border-radius:5px;color:var(--text-mid);font-family:var(--mono);font-size:.72rem;padding:.4rem .6rem;cursor:pointer;transition:.15s}
  .tab:hover:not(.disabled){color:var(--text);border-color:var(--accent);background:rgba(88,166,255,.08)}
  .tab .tn{color:var(--text-dim);font-size:.64rem}
  .tab.active{color:#fff;border-color:var(--accent4);background:rgba(240,136,62,.12)}
  .tab.active .tn{color:var(--accent4)}
  .tab.disabled{opacity:.4;cursor:not-allowed}
  .dot{width:.45rem;height:.45rem;border-radius:50%;flex-shrink:0}
  .dot.ok{background:var(--accent3)} .dot.none{background:var(--border)} .dot.skip{background:var(--accent4)}
  main{flex:1;position:relative;background:var(--bg)}
  iframe{width:100%;height:100%;border:0;background:var(--bg);display:none}
  iframe.show{display:block}
  #placeholder{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1rem;text-align:center;padding:2rem}
  #placeholder .big{font-family:var(--mono);font-size:1.1rem;color:var(--text-mid)}
  #placeholder .small{font-size:.8rem;color:var(--text-dim);max-width:34rem;line-height:1.6}
  #placeholder .logo{width:4rem;height:4rem;background:linear-gradient(135deg,var(--accent),#1f6feb);display:flex;align-items:center;justify-content:center;clip-path:polygon(0 15%,15% 0,85% 0,100% 15%,100% 85%,85% 100%,15% 100%,0 85%)}
  #placeholder .logo svg{width:62%;height:62%}
</style>
</head>
<body>
<header>
  <div class="hicon">${NOSE_SVG}</div>
  <div class="htxt">
    <h1>FIUTO — Full Report</h1>
    <div class="sub">${OSL} · ${HOST_DISP} · ${SCAN}</div>
  </div>
  <div class="hstats">
    <div class="s ok"><div class="v">${COUNT_OK}</div><div class="l">$(L "con evidenze" "with findings")</div></div>
    <div class="s tot"><div class="v">${COUNT_TOTAL}</div><div class="l">$(L "moduli" "modules")</div></div>
  </div>
</header>
<nav id="tabs">${TABS}</nav>
<main>
  <iframe id="viewer" title="report"></iframe>
  <div id="placeholder">
    <div class="logo">${NOSE_SVG}</div>
    <div class="big">$(L "Seleziona un modulo dalle tab in alto" "Select a module from the tabs above")</div>
    <div class="small">$(L "Il report verrà caricato qui al centro. Puoi passare da un modulo all'altro senza aprire file separati." "The report will load here in the center. Switch between modules without opening separate files.")</div>
  </div>
</main>
<script>
(function(){
  var tabs=document.querySelectorAll('.tab:not(.disabled)'),
      viewer=document.getElementById('viewer'),
      ph=document.getElementById('placeholder');
  tabs.forEach(function(t){
    t.addEventListener('click',function(){
      document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});
      t.classList.add('active');
      var src=t.getAttribute('data-src');
      if(src){ viewer.src=src; viewer.classList.add('show'); ph.style.display='none'; }
    });
  });
})();
</script>
</body></html>
HTMLEOF
    } > "$DASH"

    register_report "$DASH"
    echo ""
    ok "$(L "Dashboard FULL generata:" "FULL dashboard generated:") ${BOLD}$DASH"
    open_report_prompt "$DASH"
}

# Estrae dalla forma bilingue "italiano§english" la variante per la lingua
# corrente. Senza separatore il testo vale per entrambe.
reg_text() {
    local _s="$1"
    [[ "$_s" != *§* ]] && { printf '%s' "$_s"; return; }
    if [[ "${LANG:-en}" == "it" ]]; then
        printf '%s' "${_s%%§*}"
    else
        printf '%s' "${_s##*§}"
    fi
}

active_registry_name() {
    case "$OS_TYPE" in
        windows) echo "MODULES_WIN" ;;
        linux)   echo "MODULES_LINUX" ;;
        macos)   echo "MODULES_MACOS" ;;
        *)       echo "" ;;
    esac
}

# Renderizza il menu a partire da un registro (equivalente dinamico di print_menu)
render_menu_from_registry() {
    local -n _REG="$1"
    local _NOT_SET _WRITABLE _READONLY _NOT_CREATED _PARENT_RO _DIAG _RUN_ALL _QUIT _CHOICE_LABEL _REPORTS_LABEL
    _NOT_SET="$(L "non impostata" "not set")"
    _WRITABLE="$(L "scrivibile" "writable")"
    _READONLY="$(L "SOLA LETTURA" "READ ONLY")"
    _NOT_CREATED="$(L "OK (non ancora creata)" "OK (not yet created)")"
    _PARENT_RO="$(L "PARENT NON SCRIVIBILE" "PARENT NOT WRITABLE")"
    _DIAG="$(L "Diagnostica volumi montati" "Diagnose mounted volumes")"
    _RUN_ALL="$(L "Esegui TUTTI i moduli" "Run ALL modules")"
    _QUIT="$(L "Esci" "Quit")"
    _CHOICE_LABEL="$(L "Scelta" "Choice")"
    _REPORTS_LABEL="$(L "Report generati" "Generated reports")"
    local _OSL; _OSL=$(os_label)
    local _TITLE; _TITLE="$(L "SELEZIONA UN MODULO" "SELECT A MODULE")"

    echo -e "  ${CYAN}${BOLD}╔══════════════════════════════════════════════════╗${RESET}"
    printf  "  ${CYAN}${BOLD}║   F I U T O  —  %-8s —  %-18s║${RESET}\n" "$_OSL" "$_TITLE"
    echo -e "  ${CYAN}${BOLD}╚══════════════════════════════════════════════════╝${RESET}"
    echo ""
    if [[ -n "$REPORT_BASE_DIR" ]]; then
        local _RW_LABEL _RW_COLOR
        if [[ -d "$REPORT_BASE_DIR" ]]; then
            if [[ -w "$REPORT_BASE_DIR" ]]; then _RW_LABEL="$_WRITABLE"; _RW_COLOR="$GREEN"
            else _RW_LABEL="$_READONLY"; _RW_COLOR="$RED"; fi
        else
            local _RD_PARENT; _RD_PARENT=$(dirname "$REPORT_BASE_DIR")
            if [[ -w "$_RD_PARENT" ]]; then _RW_LABEL="$_NOT_CREATED"; _RW_COLOR="$GREEN"
            else _RW_LABEL="$_PARENT_RO"; _RW_COLOR="$RED"; fi
        fi
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${DIM}${REPORT_BASE_DIR}${RESET}  ${_RW_COLOR}[${_RW_LABEL}]${RESET}"
    else
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${RED}$(L "non impostata — premi [P] per configurare" "not set — press [P] to configure")${RESET}"
    fi
    echo -e "  ${WHITE}[R]${RESET}  ${BOLD}$(L "Imposta root da analizzare" "Set analysis root")${RESET}  ${DIM}${WIN_ROOT:-($_NOT_SET)} [${_OSL}]${RESET}"
    echo -e "  ${YELLOW}[D]${RESET}  ${BOLD}$(L "Debug mount attivi" "Debug active mounts")${RESET}  ${DIM}${_DIAG}${RESET}"
    echo ""
    local _i=1 _entry _f _name _color _desc
    for _entry in "${_REG[@]}"; do
        IFS='|' read -r _f _name _color _desc <<< "$_entry"
        local _C="${!_color:-$RESET}"
        printf "  ${_C}[%2d]${RESET} %-28s ${DIM}%s${RESET}\n" \
            "$_i" "$(reg_text "$_name")" "$(reg_text "$_desc")"
        _i=$((_i + 1))
    done
    echo ""
    echo -e "  ${WHITE}${BOLD}[0]${RESET}  ${BOLD}${_RUN_ALL}${RESET}"
    echo ""
    if [[ ${#GENERATED_REPORTS[@]} -gt 0 ]]; then
        echo -e "  ${DIM}── ${_REPORTS_LABEL} (${#GENERATED_REPORTS[@]}) ──────────────────────────${RESET}"
        for _R in "${GENERATED_REPORTS[@]}"; do
            echo -e "  ${CYAN}↳${RESET} ${DIM}${_R}${RESET}"
        done
        echo ""
    fi
    echo -e "  ${RED}[Q]  ${_QUIT}${RESET}"
    echo ""
    echo -ne "  ${YELLOW}${_CHOICE_LABEL}:${RESET} "
}

# Esegue il modulo n-esimo (1-based) di un registro
dispatch_from_registry() {
    local _RNAME="$1" _N="$2"
    local -n _REG="$_RNAME"
    if ! [[ "$_N" =~ ^[0-9]+$ ]] || (( _N < 1 || _N > ${#_REG[@]} )); then
        err "$(L "Modulo sconosciuto:" "Unknown module:") $_N"
        return 1
    fi
    local _entry="${_REG[$((_N - 1))]}"
    local _f="${_entry%%|*}"
    "$_f"
}

# Esegue TUTTI i moduli di un registro in modalità batch (equivalente di run_all_modules)
run_all_from_registry() {
    local _RNAME="$1"
    local -n _REG="$_RNAME"
    clear
    print_banner
    info "$(t batch_running)"
    echo ""
    if [[ -z "$REPORT_BASE_DIR" ]]; then
        REPORT_BASE_DIR="${INVOCATION_DIR}/fiuto_reports_$(date +%Y%m%d_%H%M%S)"
        LOG_FILE="${REPORT_BASE_DIR}/fiuto_session_$(date +%Y%m%d_%H%M%S).log"
    fi
    info "$(t batch_report_dir) ${BOLD}$REPORT_BASE_DIR${RESET}"
    log_msg "$(t batch_started)$WIN_ROOT ==="
    sleep 1
    BATCH_MODE=true
    SUMMARY_TABLE=()
    echo ""
    local _total=${#_REG[@]} _i=1 _entry _f _name _color _desc _guard
    for _entry in "${_REG[@]}"; do
        IFS='|' read -r _f _name _color _desc _guard <<< "$_entry"
        local _label; _label=$(reg_text "$_name")
        # Guardia facoltativa: se fallisce il modulo viene saltato con motivo.
        if [[ -n "${_guard:-}" ]] && declare -F "$_guard" > /dev/null; then
            local _reason
            if ! _reason=$("$_guard"); then
                echo -e "  ${DIM}[i] [$_i/$_total] ${_label} — $(L "saltato" "skipped") (${_reason})${RESET}"
                SUMMARY_TABLE+=("$_i|$_label|SKIP|$_reason")
                _i=$((_i + 1))
                continue
            fi
        fi
        run_batch_module "$_i" "$_f" "$_label" "$_total"
        _i=$((_i + 1))
    done
    BATCH_MODE=false
    echo ""
    section_header "$(L "Riepilogo Scansione Globale" "Global Scan Summary")" "$GREEN"
    local _hdr_mod _hdr_name _hdr_evid _hdr_file _lbl_found _lbl_none _lbl_skip
    _hdr_mod="$(L "MOD" "MOD")"; _hdr_name="$(L "NOME MODULO" "MODULE NAME")"
    _hdr_evid="$(L "EVIDENZE" "FINDINGS")"; _hdr_file="$(L "FILE GENERATI" "GENERATED FILES")"
    _lbl_found="$(L "TROVATE" "FOUND")"; _lbl_none="$(L "NESSUNA" "NONE")"; _lbl_skip="$(L "SALTATO" "SKIPPED")"
    printf "  ${BOLD}%-4s %-32s %-12s %s${RESET}\n" "$_hdr_mod" "$_hdr_name" "$_hdr_evid" "$_hdr_file"
    echo "  ─────────────────────────────────────────────────────────────────────────────────────────"
    for row in "${SUMMARY_TABLE[@]}"; do
        IFS='|' read -r mnum mname msy mpath <<< "$row"
        if [[ "$msy" == "SI" ]]; then
            local rel_path="${mpath#$REPORT_BASE_DIR/}"
            printf "  ${CYAN}%02d${RESET}   ${BOLD}%-32s${RESET} ${GREEN}%-12s${RESET} ${DIM}%s${RESET}\n" "$mnum" "$mname" "$_lbl_found" "$rel_path"
        elif [[ "$msy" == "SKIP" ]]; then
            printf "  ${CYAN}%02d${RESET}   %-32s ${YELLOW}%-12s${RESET} ${DIM}%s${RESET}\n" "$mnum" "$mname" "$_lbl_skip" "$mpath"
        else
            printf "  ${CYAN}%02d${RESET}   %-32s ${DIM}%-12s${RESET} ${DIM}-${RESET}\n" "$mnum" "$mname" "$_lbl_none"
        fi
    done
    echo ""
    ok "$(L "Report salvati integralmente in:" "All reports saved in:") ${BOLD}$REPORT_BASE_DIR"
    generate_full_dashboard
}


# ================================================================
#  HELPER FUNCTIONS PER MODALITÀ NON INTERATTIVA
# ================================================================

# Espande una stringa tipo "1,3,5-8,12" in una lista di numeri
expand_module_list() {
    local spec="$1"
    local result=()
    IFS=',' read -ra PARTS <<< "$spec"
    for P in "${PARTS[@]}"; do
        if [[ "$P" =~ ^([0-9]+)-([0-9]+)$ ]]; then
            local from="${BASH_REMATCH[1]}" to="${BASH_REMATCH[2]}"
            for (( n=from; n<=to; n++ )); do result+=("$n"); done
        elif [[ "$P" =~ ^[0-9]+$ ]]; then
            result+=("$P")
        fi
    done
    printf '%s\n' "${result[@]}" | sort -n -u
}

