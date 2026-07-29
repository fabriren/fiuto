
# ================================================================
#  MODULO 8 — Scheduled Tasks
# ================================================================
module_scheduled_tasks() {
    section_header "Scheduled Tasks" "$YELLOW"
    check_win_root || return 1

    local TASKS_DIR
    TASKS_DIR=$(ci_find_dir "$WIN_ROOT" "Windows/System32/Tasks")
    if [[ -z "$TASKS_DIR" || ! -d "$TASKS_DIR" ]]; then
        warn "$(L "Directory Tasks non trovata" "Tasks directory not found")"
        return 0
    fi

    mapfile -t TASK_FILES < <(find "$TASKS_DIR" -type f ! -iname "*.job" -printf "%T@ %p\n" 2>/dev/null | sort -rn | cut -d' ' -f2-)
    if [[ ${#TASK_FILES[@]} -eq 0 ]]; then
        mapfile -t TASK_FILES < <(find "$TASKS_DIR" -type f ! -iname "*.job" -print0 2>/dev/null | xargs -0 ls -t 2>/dev/null)
    fi
    local TOTAL=${#TASK_FILES[@]}
    info "$(L "Trovati" "Found") $TOTAL task in: $TASKS_DIR"

    declare -a ALL_TASKS=()
    local SUSP_COUNT=0

    for TF in "${TASK_FILES[@]}"; do
        # Parser XML: estrae Command, Arguments, WorkingDirectory, Author, Description, Trigger
        local T_OUT
        T_OUT=$("$PY3" - "$TF" << 'PYEOF' 2>/dev/null || echo "|||||| "
import sys, xml.etree.ElementTree as ET
try:
    tree = ET.parse(sys.argv[1])
    root = tree.getroot()
    def ft(tag):
        for e in root.iter():
            if e.tag.split('}')[-1] == tag and e.text and e.text.strip():
                return e.text.strip()
        return ''
    triggers = []
    for e in root.iter():
        t = e.tag.split('}')[-1]
        if t.endswith('Trigger') and t != 'Triggers':
            sb = ''
            for c in e:
                if c.tag.split('}')[-1] == 'StartBoundary' and c.text:
                    sb = c.text.strip()
            triggers.append(f"{t}:{sb}" if sb else t)
    print('\t'.join([ft('Command'), ft('Arguments'), ft('WorkingDirectory'),
                     ft('Author'), ft('Description')[:100], ', '.join(triggers)]))
except:
    print('\t\t\t\t\t')
PYEOF
)
        IFS=$'\t' read -r T_CMD T_ARGS T_WORKDIR T_AUTHOR T_DESC T_TRIGGERS <<< "$T_OUT"

        local IS_SUSP=false
        local FULL_CMD="${T_CMD} ${T_ARGS}"
        echo "$FULL_CMD" | grep -qi "powershell\|wscript\|cscript\|mshta\|regsvr32\|rundll32\|certutil\|bitsadmin\|temp\|appdata\|programdata\|public\|downloads\|http\|ftp\|base64\|-enc\|-w hidden\|-nop" && IS_SUSP=true

        local RELPATH="${TF#${TASKS_DIR}/}"
        # Salva i 7 campi + flag sospetto separati da §  (carattere che non appare nei path)
        ALL_TASKS+=("${RELPATH}§${T_CMD}§${T_ARGS}§${T_WORKDIR}§${T_AUTHOR}§${T_DESC}§${T_TRIGGERS}§${IS_SUSP}")

        if $IS_SUSP; then
            echo -e "  ${RED}[!] $RELPATH${RESET}"
            [[ -n "$T_CMD"      ]] && echo -e "      ${YELLOW}Cmd:     ${RESET}${BOLD}$T_CMD${RESET}"
            [[ -n "$T_ARGS"     ]] && echo -e "      ${RED}Args:    $T_ARGS${RESET}"
            [[ -n "$T_WORKDIR"  ]] && echo -e "      ${DIM}WorkDir: $T_WORKDIR${RESET}"
            [[ -n "$T_AUTHOR"   ]] && echo -e "      ${DIM}Author:  $T_AUTHOR${RESET}"
            [[ -n "$T_TRIGGERS" ]] && echo -e "      ${DIM}Trigger: $T_TRIGGERS${RESET}"
            SUSP_COUNT=$((SUSP_COUNT + 1))
        else
            echo -e "  ${DIM}$RELPATH${RESET}"
            [[ -n "$T_CMD" ]] && echo -e "      ${DIM}↳ $T_CMD${RESET}"
        fi
    done

    separator
    info "$(L "Task totali:" "Total tasks:") ${BOLD}$TOTAL${RESET}  |  Sospetti: ${RED}${BOLD}$SUSP_COUNT"

    ask_yn "Generare report HTML?" || return 0

    local REPORT_HTML; REPORT_HTML=$(prepare_report_dir "scheduled_tasks")
    local SCAN; SCAN=$(date "+%d/%m/%Y %H:%M:%S")

    # Funzione escape HTML — definita fuori dal loop così è disponibile anche nei subshell
    _esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

    local ROWS=""
    for E in "${ALL_TASKS[@]}"; do
        IFS='§' read -r TNAME T_CMD T_ARGS T_WORKDIR T_AUTHOR T_DESC T_TRIGGERS TISUSP <<< "$E"

        local ROW_STYLE="" ARG_EXTRA_CSS=""
        if [[ "$TISUSP" == "true" ]]; then
            ROW_STYLE="style='background:rgba(255,123,114,.07);border-left:3px solid var(--accent2)'"
            ARG_EXTRA_CSS=" bad"
        fi

        # Ogni campo su riga propria, testo sempre wrappabile
        local DETAILS=""
        [[ -n "$T_CMD"      ]] && DETAILS+="<div class='drow'><span class='lbl'>CMD</span><span class='mono fld' style='color:var(--accent)'>$(_esc "$T_CMD")</span></div>"
        [[ -n "$T_ARGS"     ]] && DETAILS+="<div class='drow'><span class='lbl'>ARG</span><span class='mono fld${ARG_EXTRA_CSS}'>$(_esc "$T_ARGS")</span></div>"
        [[ -n "$T_WORKDIR"  ]] && DETAILS+="<div class='drow'><span class='lbl'>DIR</span><span class='mono fld mid'>$(_esc "$T_WORKDIR")</span></div>"
        [[ -n "$T_TRIGGERS" ]] && DETAILS+="<div class='drow'><span class='lbl'>TRG</span><span class='mono fld' style='color:var(--accent4)'>$(_esc "$T_TRIGGERS")</span></div>"
        [[ -n "$T_AUTHOR"   ]] && DETAILS+="<div class='drow'><span class='lbl'>AUT</span><span class='mono fld dim'>$(_esc "$T_AUTHOR")</span></div>"
        [[ -n "$T_DESC"     ]] && DETAILS+="<div class='drow'><span class='lbl'>DSC</span><span class='fld' style='font-size:.72rem;color:var(--text-mid)'>$(_esc "$T_DESC")</span></div>"
        [[ -z "$DETAILS"    ]] && DETAILS="<span class='dim' style='font-size:.72rem'>—</span>"

        ROWS+="<tr ${ROW_STYLE}>
          <td class='tname' style='vertical-align:top;padding-top:.5rem'>$(_esc "$TNAME")</td>
          <td style='padding:.35rem .9rem'>${DETAILS}</td>
        </tr>"
    done

    {
        html_header "Scheduled Tasks"
        html_page_header "ST" "Scheduled <span>Tasks</span> Forensics" \
            "Windows\\System32\\Tasks" "$SCAN" "$WIN_ROOT"
        echo "<div class='statsbar'>
          <div class='stat'><div class='label'>Task trovati</div><div class='value'>${TOTAL}</div></div>
          <div class='stat'><div class='label'>Sospetti</div><div class='value' style='color:var(--accent2)'>${SUSP_COUNT}</div></div>
        </div>
        <main>
        <style>
          /* label badge fisso, testo sempre wrappabile */
          .lbl {
            display: inline-block;
            font-family: var(--mono);
            font-size: .58rem;
            color: var(--text-dim);
            width: 2.6rem;
            margin-right: .5rem;
            text-transform: uppercase;
            vertical-align: top;
            padding-top: .1rem;
            flex-shrink: 0;
          }
          .drow {
            display: flex;
            align-items: flex-start;
            margin-bottom: .22rem;
          }
          .drow:last-child { margin-bottom: 0; }
          /* campo testo: si wrappa sempre, niente overflow nascosto */
          .fld {
            word-break: break-all;
            overflow-wrap: anywhere;
            white-space: pre-wrap;
            flex: 1;
          }
          /* colonna nome task: wrappabile, font piccolo */
          .tname {
            font-family: var(--mono);
            font-size: .7rem;
            color: var(--text-mid);
            word-break: break-all;
            min-width: 160px;
            max-width: 280px;
            vertical-align: top;
            padding: .5rem .8rem .5rem 1rem;
          }
        </style>
        <div class='stitle'>Task pianificati — CMD · ARG · DIR · TRG · AUT · DSC</div>
        <div class='card'><table>
          <thead><tr>
            <th style='width:22%;min-width:160px'>Task</th>
            <th>$(L "Dettagli" "Details")</th>
          </tr></thead>
          <tbody>${ROWS}</tbody>
        </table></div></main>"
        html_footer "$SCAN" "$WIN_ROOT"
    } > "$REPORT_HTML"

    register_report "$REPORT_HTML"
    ok "$(L "Report salvato:" "Report saved:") ${BOLD}$REPORT_HTML"
    open_report_prompt "$REPORT_HTML"
}
