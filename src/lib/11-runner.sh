
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

run_all_modules() {
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
    run_batch_module 1 module_ps_history "PowerShell History" 39
    run_batch_module 2 module_notepad_tabstate "Notepad TabState" 39
    run_batch_module 3 module_ifeo "IFEO" 39
    run_batch_module 4 module_bam "BAM" 39
    run_batch_module 5 module_run_keys "Run Keys" 39
    run_batch_module 6 module_prefetch "Prefetch" 39
    run_batch_module 7 module_scheduled_tasks "Scheduled Tasks" 39
    run_batch_module 8 module_usb "USB Devices" 39
    run_batch_module 9 module_lnk "LNK Files" 39
    run_batch_module 10 module_rdp_cache "RDP Cache" 39
    run_batch_module 11 module_services "Services" 39
    run_batch_module 12 module_evtx "EVTX" 39
    run_batch_module 13 module_amcache "Amcache" 39
    run_batch_module 14 module_recycle_bin "Recycle Bin" 39
    run_batch_module 15 module_wmi "WMI" 39
    run_batch_module 16 module_srum "SRUM" 39
    run_batch_module 17 module_browser "Browser History" 39
    run_batch_module 18 module_userassist "UserAssist" 39
    run_batch_module 19 module_shellbags "Shellbags" 39
    run_batch_module 20 module_sam "SAM" 39
    run_batch_module 21 module_mft "MFT" 39
    run_batch_module 22 module_opensave "Open/Save MRU" 39
    run_batch_module 23 module_usn "USN Journal" 39
    run_batch_module 24 module_ntds "NTDS.dit" 39
    run_batch_module 25 module_hiberfil "Hibernation / Pagefile" 39
    run_batch_module 26 module_wer_files "WER Files" 39
    run_batch_module 27 module_credential_manager "Credential Manager" 39
    run_batch_module 28 module_wlan "WLAN Profiles" 39
    run_batch_module 29 module_appx "AppX / UWP" 39
    run_batch_module 30 module_browser_extra "Browser Logins/Downloads" 39
    run_batch_module 31 module_clipboard "Clipboard History" 39
    run_batch_module 32 module_office_mru "Office MRU" 39
    run_batch_module 33 module_defender_quarantine "Defender Quarantine" 39
    run_batch_module 34 module_ps_scriptblock "PS ScriptBlock Log" 39
    run_batch_module 35 module_jumplists "JumpLists" 39
    run_batch_module 36 module_network_artifacts "Network Artifacts" 39
    run_batch_module 37 module_master_timeline "Master Timeline" 39
    # Modulo 38: eseguito solo se il disco è un Domain Controller (ntds.dit presente)
    local _ntds_check
    _ntds_check=$(find "$WIN_ROOT" -maxdepth 8 -iname "ntds.dit" -type f 2>/dev/null | head -1)
    if [[ -n "$_ntds_check" ]]; then
        run_batch_module 38 module_pad_offline "PAD Offline AD" 39
    else
        echo -e "  ${DIM}[i] [38/39] PAD Offline AD — $(L "saltato (non è un Domain Controller)" "skipped (not a Domain Controller)")${RESET}"
        SUMMARY_TABLE+=("38|PAD Offline AD|SKIP|$(L "non è un DC" "not a DC")")
    fi
    run_batch_module 39 module_ai_chat "AI Chat History" 39

    BATCH_MODE=false

    echo ""
    section_header "$(L "Riepilogo Scansione Globale" "Global Scan Summary")" "$GREEN"
    local _hdr_mod;    _hdr_mod="$(L    "MOD" "MOD")"
    local _hdr_name;   _hdr_name="$(L  "NOME MODULO" "MODULE NAME")"
    local _hdr_evid;   _hdr_evid="$(L  "EVIDENZE" "FINDINGS")"
    local _hdr_file;   _hdr_file="$(L  "FILE GENERATI" "GENERATED FILES")"
    local _lbl_found;  _lbl_found="$(L "TROVATE" "FOUND")"
    local _lbl_none;   _lbl_none="$(L  "NESSUNA" "NONE")"
    local _lbl_skip;   _lbl_skip="$(L  "SALTATO" "SKIPPED")"
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
active_registry_name() {
    case "$OS_TYPE" in
        linux) echo "MODULES_LINUX" ;;
        macos) echo "MODULES_MACOS" ;;
        *)     echo "" ;;
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
        printf "  ${_C}[%2d]${RESET} %-26s ${DIM}%s${RESET}\n" "$_i" "$_name" "$_desc"
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
    local _total=${#_REG[@]} _i=1 _entry _f _name _rest
    for _entry in "${_REG[@]}"; do
        IFS='|' read -r _f _name _rest <<< "$_entry"
        run_batch_module "$_i" "$_f" "$_name" "$_total"
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

print_menu() {
    local _MENU_TITLE _SELECT_MODULE _NOT_SET _WRITABLE _READONLY _NOT_CREATED _PARENT_RO
    local _REPORT_DIR_LABEL _WIN_ROOT_LABEL _DEBUG_LABEL _RUN_ALL _QUIT _CHOICE_LABEL
    local _REPORTS_LABEL
    _MENU_TITLE="$(L "SELEZIONA UN MODULO" "SELECT A MODULE    ")"
    _SELECT_MODULE="$(L "Seleziona" "Select")"
    _NOT_SET="$(L "non impostata" "not set")"
    _WRITABLE="$(L "scrivibile" "writable")"
    _READONLY="$(L "SOLA LETTURA" "READ ONLY")"
    _NOT_CREATED="$(L "OK (non ancora creata)" "OK (not yet created)")"
    _PARENT_RO="$(L "PARENT NON SCRIVIBILE" "PARENT NOT WRITABLE")"
    _REPORT_DIR_LABEL="$(L "Imposta dir report" "Set report dir")"
    _WIN_ROOT_LABEL="$(L "Imposta root Windows" "Set Windows root    ")"
    _DEBUG_LABEL="$(L "Debug mount attivi " "Debug active mounts")"
    _DIAG="$(L "Diagnostica volumi montati" "Diagnose mounted volumes")"
    _RUN_ALL="$(L "Esegui TUTTI i moduli" "Run ALL modules")"
    _QUIT="$(L "Esci" "Quit")"
    _CHOICE_LABEL="$(L "Scelta" "Choice")"
    _REPORTS_LABEL="$(L "Report generati" "Generated reports")"

    echo -e "  ${CYAN}${BOLD}╔══════════════════════════════════════════════════╗${RESET}"
    echo -e "  ${CYAN}${BOLD}║           F I U T O  —  ${_MENU_TITLE}      ║ ${RESET}"
    echo -e "  ${CYAN}${BOLD}╚══════════════════════════════════════════════════╝${RESET}"
    echo ""
    if [[ -n "$REPORT_BASE_DIR" ]]; then
        local _RW_LABEL _RW_COLOR
        if [[ -d "$REPORT_BASE_DIR" ]]; then
            if [[ -w "$REPORT_BASE_DIR" ]]; then
                _RW_LABEL="$_WRITABLE"; _RW_COLOR="$GREEN"
            else
                _RW_LABEL="$_READONLY"; _RW_COLOR="$RED"
            fi
        else
            local _RD_PARENT; _RD_PARENT=$(dirname "$REPORT_BASE_DIR")
            if [[ -w "$_RD_PARENT" ]]; then
                _RW_LABEL="$_NOT_CREATED"; _RW_COLOR="$GREEN"
            else
                _RW_LABEL="$_PARENT_RO"; _RW_COLOR="$RED"
            fi
        fi
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${DIM}${REPORT_BASE_DIR}${RESET}  ${_RW_COLOR}[${_RW_LABEL}]${RESET}"
    else
        local _CONF_MSG="$(L "non impostata — premi [P] per configurare" "not set — press [P] to configure")"
        echo -e "  ${WHITE}[P]${RESET}  ${BOLD}Report dir:${RESET} ${RED}${_CONF_MSG}${RESET}"
    fi
    echo -e "  ${WHITE}[R]${RESET}  ${BOLD}${_WIN_ROOT_LABEL}${RESET}          ${DIM}${WIN_ROOT:-($_NOT_SET)}${RESET}"
    echo -e "  ${YELLOW}[D]${RESET}  ${BOLD}${_DEBUG_LABEL}${RESET}           ${DIM}${_DIAG}${RESET}"
    echo ""
    echo -e "  ${MAGENTA}[1]${RESET}  PowerShell History            ${DIM}PSReadLine *_history.txt${RESET}"
    echo -e "  ${MAGENTA}[2]${RESET}  Notepad TabState              ${DIM}$(L "Tab rimasti aperti (.bin)" "Open tabs (.bin)")${RESET}"
    echo -e "  ${RED}[3]${RESET}  IFEO Hijacking                ${DIM}Image File Execution Options${RESET}"
    echo -e "  ${BLUE}[4]${RESET}  BAM                           ${DIM}Background Activity Moderator${RESET}"
    echo -e "  ${ORANGE}[5]${RESET}  Run Keys & $(L "Persistenza" "Persistence")        ${DIM}$(L "Autorun nel registro" "Autorun in registry")${RESET}"
    echo -e "  ${GREEN}[6]${RESET}  Prefetch                      ${DIM}$(L "Eseguibili tracciati" "Tracked executables") (*.pf)${RESET}"
    echo -e "  ${YELLOW}[7]${RESET}  Scheduled Tasks               ${DIM}$(L "Task pianificati (XML)" "Scheduled tasks (XML)")${RESET}"
    echo -e "  ${BLUE}[8]${RESET}  USB Devices                   ${DIM}$(L "Dispositivi rimovibili (USBSTOR)" "Removable devices (USBSTOR)")${RESET}"
    echo -e "  ${GREEN}[9]${RESET}  LNK & JumpList                ${DIM}$(L "File recenti e target path" "Recent files and target path")${RESET}"
    echo -e "  ${CYAN}[10]${RESET} RDP Cache                     ${DIM}Terminal Server Client Cache${RESET}"
  echo -e "  ${RED}[11]${RESET} Services                      ${DIM}$(L "Servizi Windows (SYSTEM hive)" "Windows Services (SYSTEM hive)")${RESET}"
  echo -e "  ${RED}[12]${RESET} Event Log                     ${DIM}Security/System/PS/RDP (.evtx)${RESET}"
  echo -e "  ${YELLOW}[13]${RESET} Amcache + Shimcache           ${DIM}$(L "Timeline esecuzione binari" "Binary execution timeline")${RESET}"
  echo -e "  ${GREEN}[14]${RESET} Recycle Bin                   ${DIM}$(L "File eliminati" "Deleted files") (\$Recycle.Bin)${RESET}"
  echo -e "  ${RED}[15]${RESET} WMI Subscriptions             ${DIM}$(L "Persistenza invisibile" "Fileless persistence") (T1546.003)${RESET}"
  echo -e "  ${BLUE}[16]${RESET} SRUM                          ${DIM}$(L "Uso risorse per applicazione" "Resource usage per application")${RESET}"
  echo -e "  ${CYAN}[17]${RESET} Browser History               ${DIM}Chrome / Edge / Firefox${RESET}"
  echo -e "  ${MAGENTA}[18]${RESET} UserAssist / RunMRU           ${DIM}$(L "Attività interattiva utente" "Interactive user activity")${RESET}"
  echo -e "  ${CYAN}[19]${RESET} ShellBags                     ${DIM}$(L "Navigazione cartelle (anche cancellate)" "Folder navigation (including deleted)")${RESET}"
  echo -e "  ${RED}[20]${RESET} SAM — $(L "Hash Locali " "Local Hashes")            ${DIM}$(L "Hash NTLM account (impacket)" "NTLM account hashes (impacket)")${RESET}"
  echo -e "  ${YELLOW}[21]${RESET} MFT Timeline                  ${DIM}Master File Table + timestomping${RESET}"
  echo -e "  ${GREEN}[22]${RESET} OpenSave / LastVisited MRU    ${DIM}$(L "File aperti/salvati via dialogo" "Files opened/saved via dialog")${RESET}"
  echo -e "  ${CYAN}[23]${RESET} USN Journal                   ${DIM}\$UsnJrnl:\$J — $(L "change log NTFS" "NTFS change log")${RESET}"
  echo -e "  ${RED}[24]${RESET} NTDS.dit                      ${DIM}Active Directory hash (DC offline)${RESET}"
  echo -e "  ${BLUE}[25]${RESET} Hibernation / Pagefile        ${DIM}hiberfil.sys · pagefile.sys strings${RESET}"
  echo -e "  ${RED}[26]${RESET} WER Files (Error Reports)     ${DIM}ReportArchive · ReportQueue (.wer)${RESET}"
  echo -e "  ${MAGENTA}[27]${RESET} Credential Manager            ${DIM}DPAPI blob offline${RESET}"
  echo -e "  ${CYAN}[28]${RESET} WLAN & VPN Profiles           ${DIM}WiFi · NetworkList · VPN${RESET}"
  echo -e "  ${GREEN}[29]${RESET} AppX / UWP Packages           ${DIM}$(L "App Store + sideload sospetti" "App Store + suspicious sideloads")${RESET}"
  echo -e "  ${CYAN}[30]${RESET} Browser Downloads & Logins    ${DIM}Download + Login Data (DPAPI)${RESET}"
  echo -e "  ${YELLOW}[31]${RESET} Clipboard History             ${DIM}$(L "Cronologia appunti Win10+" "Clipboard history Win10+")${RESET}"
  echo -e "  ${GREEN}[32]${RESET} Office MRU                    ${DIM}$(L "File recenti Word/Excel/PowerPoint" "Recent Word/Excel/PowerPoint files")${RESET}"
  echo -e "  ${RED}[33]${RESET} Defender Quarantine           ${DIM}$(L "File in quarantena + threatname" "Quarantined files + threatname")${RESET}"
  echo -e "  ${MAGENTA}[34]${RESET} PS ScriptBlock Logging        ${DIM}Event ID 4104 — PS Operational.evtx${RESET}"
  echo -e "  ${GREEN}[35]${RESET} JumpLists                     ${DIM}AutomaticDestinations · CustomDestinations${RESET}"
  echo -e "  ${CYAN}[36]${RESET} Network Artifacts             ${DIM}$(L "Profili rete · Interfacce TCP/IP (registry)" "Network profiles · TCP/IP interfaces (registry)")${RESET}"
  echo -e "  ${YELLOW}[37]${RESET} Master Timeline               ${DIM}$(L "Aggregazione cross-moduli con filtri" "Cross-module aggregation with filters")${RESET}"
  echo -e "  ${RED}[38]${RESET} PAD Offline AD Analysis       ${DIM}$(L "NTDS.dit offline — utenti privilegiati, ACL, GPO" "NTDS.dit offline — privileged users, ACL, GPO")${RESET}"
  echo -e "  ${MAGENTA}[39]${RESET} AI Chat History               ${DIM}Claude · ChatGPT · Copilot · Cursor · Gemini · Codex${RESET}"
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

run_module_by_number() {
    case "$1" in
        1)  module_ps_history ;;
        2)  module_notepad_tabstate ;;
        3)  module_ifeo ;;
        4)  module_bam ;;
        5)  module_run_keys ;;
        6)  module_prefetch ;;
        7)  module_scheduled_tasks ;;
        8)  module_usb ;;
        9)  module_lnk ;;
        10) module_rdp_cache ;;
        11) module_services ;;
        12) module_evtx ;;
        13) module_amcache ;;
        14) module_recycle_bin ;;
        15) module_wmi ;;
        16) module_srum ;;
        17) module_browser ;;
        18) module_userassist ;;
        19) module_shellbags ;;
        20) module_sam ;;
        21) module_mft ;;
        22) module_opensave ;;
        23) module_usn ;;
        24) module_ntds ;;
        25) module_hiberfil ;;
        26) module_wer_files ;;
        27) module_credential_manager ;;
        28) module_wlan ;;
        29) module_appx ;;
        30) module_browser_extra ;;
        31) module_clipboard ;;
        32) module_office_mru ;;
        33) module_defender_quarantine ;;
        34) module_ps_scriptblock ;;
        35) module_jumplists ;;
        36) module_network_artifacts ;;
        37) module_master_timeline ;;
        38) module_pad_offline ;;
        39) module_ai_chat ;;
        *)  err "$(L "Modulo sconosciuto:" "Unknown module:") $1" ;;
    esac
}
