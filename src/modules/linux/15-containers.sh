
# --- LINUX 16 — Container (Docker / Podman) ---
#
# Un host con container ha un intero piano di esecuzione che gli altri moduli
# non vedono: processi, filesystem e persistenza vivono dentro le immagini.
# Qui si ricostruisce l'inventario offline dai metadati sul disco e si
# evidenziano le configurazioni che permettono la fuga dal container verso
# l'host — privileged, mount di / o del socket Docker, CAP_SYS_ADMIN,
# condivisione del namespace PID/rete dell'host.
module_linux_containers() {
    section_header "Linux — Container (Docker / Podman)" "$BLUE"
    check_target_root || return 1

    local DOCKER PODMAN
    DOCKER=$(ci_find_dir "$WIN_ROOT" "var/lib/docker")
    PODMAN=$(ci_find_dir "$WIN_ROOT" "var/lib/containers/storage")
    if [[ -z "$DOCKER" && -z "$PODMAN" ]]; then
        warn "$(L "Nessun runtime container trovato (/var/lib/docker, /var/lib/containers)." \
                 "No container runtime found (/var/lib/docker, /var/lib/containers).")"
        return 0
    fi
    [[ -n "$DOCKER" ]] && info "Docker: ${BOLD}${DOCKER}"
    [[ -n "$PODMAN" ]] && info "Podman: ${BOLD}${PODMAN}"

    local OUT; OUT=$(mktemp); register_tmp "$OUT"
    "$PY3" - "$OUT" "${DOCKER:-}" "${PODMAN:-}" << 'PYEOF' 2>/dev/null
import sys, os, json, glob

out_path = sys.argv[1]
docker   = sys.argv[2] if len(sys.argv) > 2 else ''
podman   = sys.argv[3] if len(sys.argv) > 3 else ''

rows = []

def risk_of(hostconfig, config):
    """Indicatori di fuga dal container verso l'host."""
    risks = []
    hc = hostconfig or {}
    if hc.get('Privileged'):
        risks.append('PRIVILEGED')
    for b in (hc.get('Binds') or []):
        src = str(b).split(':')[0]
        if src == '/':
            risks.append('MOUNT_ROOT_HOST')
        elif 'docker.sock' in str(b):
            risks.append('DOCKER_SOCKET')
        elif src in ('/etc', '/var/run', '/proc', '/sys', '/boot'):
            risks.append(f'MOUNT_{src.strip("/").upper()}')
    caps = hc.get('CapAdd') or []
    for c in caps:
        if str(c).upper().replace('CAP_', '') in ('SYS_ADMIN', 'SYS_PTRACE', 'SYS_MODULE', 'ALL'):
            risks.append(f'CAP_{str(c).upper().replace("CAP_", "")}')
    if str(hc.get('PidMode', '')) == 'host':
        risks.append('PID_HOST')
    if str(hc.get('NetworkMode', '')) == 'host':
        risks.append('NET_HOST')
    if str(hc.get('IpcMode', '')) == 'host':
        risks.append('IPC_HOST')
    return risks

# ---- Docker -------------------------------------------------------------
if docker and os.path.isdir(docker):
    for cdir in sorted(glob.glob(os.path.join(docker, 'containers', '*'))):
        cfg_path = os.path.join(cdir, 'config.v2.json')
        hc_path  = os.path.join(cdir, 'hostconfig.json')
        if not os.path.isfile(cfg_path):
            continue
        try:
            cfg = json.load(open(cfg_path, encoding='utf-8', errors='replace'))
        except Exception:
            continue
        hc = {}
        if os.path.isfile(hc_path):
            try:
                hc = json.load(open(hc_path, encoding='utf-8', errors='replace'))
            except Exception:
                hc = {}
        name  = str(cfg.get('Name', '')).lstrip('/')
        image = cfg.get('Config', {}).get('Image') or cfg.get('Image', '')
        created = str(cfg.get('Created', ''))[:19].replace('T', ' ')
        state = cfg.get('State', {}) or {}
        started  = str(state.get('StartedAt', ''))[:19].replace('T', ' ')
        finished = str(state.get('FinishedAt', ''))[:19].replace('T', ' ')
        running = 'running' if state.get('Running') else 'stopped'
        cmd = ' '.join((cfg.get('Config', {}) or {}).get('Cmd') or [])
        entry = ' '.join((cfg.get('Config', {}) or {}).get('Entrypoint') or [])
        risks = risk_of(hc, cfg)
        rows.append((
            'docker', created or started, name, str(image), running,
            (entry + ' ' + cmd).strip()[:300], ';'.join(risks), os.path.basename(cdir)[:12],
            started, finished,
        ))

# ---- Podman -------------------------------------------------------------
if podman and os.path.isdir(podman):
    cjson = os.path.join(podman, 'overlay-containers', 'containers.json')
    if os.path.isfile(cjson):
        try:
            for c in json.load(open(cjson, encoding='utf-8', errors='replace')):
                rows.append((
                    'podman', str(c.get('created', ''))[:19].replace('T', ' '),
                    str(c.get('names', [''])[0] if c.get('names') else ''),
                    str(c.get('image', '')), '', '', '', str(c.get('id', ''))[:12], '', '',
                ))
        except Exception:
            pass

with open(out_path, 'w', encoding='utf-8') as fh:
    for r in rows:
        fh.write("\t".join(str(x).replace("\t", " ").replace("\n", " ") for x in r) + "\n")
PYEOF

    local TOTAL=0
    [[ -s "$OUT" ]] && TOTAL=$(wc -l < "$OUT")
    if [[ "$TOTAL" -eq 0 ]]; then
        warn "$(L "Nessun container ricostruibile dai metadati." "No container reconstructable from metadata.")"
        return 0
    fi
    local NRISK; NRISK=$(awk -F'\t' '$7!=""' "$OUT" | wc -l)

    ok "$(L "Container trovati:" "Containers found:") ${BOLD}$TOTAL"
    if [[ "$NRISK" -gt 0 ]]; then
        warn "$(L "Container con configurazione a rischio di fuga:" "Containers with escape-prone configuration:") ${BOLD}$NRISK"
        awk -F'\t' '$7!=""{printf "      %s  [%s]  %s\n", $3, $7, $4}' "$OUT" | head -20 | while IFS= read -r LN; do
            echo -e "      ${RED}${LN}${RESET}"
        done
    else
        info "$(L "Nessun indicatore di fuga rilevato." "No escape indicator detected.")"
    fi

    # Log stdout dei container: spesso contengono l'attivita' dell'attaccante.
    local NLOGS=0
    if [[ -n "$DOCKER" ]]; then
        NLOGS=$(find "$DOCKER/containers" -maxdepth 2 -name '*-json.log' -size +0 2>/dev/null | wc -l)
        [[ "$NLOGS" -gt 0 ]] && info "$(L "Log stdout disponibili:" "stdout logs available:") ${BOLD}${NLOGS}"
    fi

    ask_yn "Generare report HTML?" || return 0

    local ROWS; ROWS=$(awk -F'\t' '{print $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7"\t"$8}' "$OUT")
    local TABLE; TABLE=$(_rows_to_table "$ROWS" \
        "Runtime" "$(L "Creato" "Created")" "$(L "Nome" "Name")" "Image" "$(L "Stato" "State")" \
        "$(L "Comando" "Command")" "$(L "Rischi" "Risks")" "ID")

    local NOTE=""
    if [[ "$NRISK" -gt 0 ]]; then
        NOTE="<div class='card' style='margin-bottom:1rem;border-color:rgba(255,123,114,.5)'><div style='padding:1rem 1.5rem;font-size:.8rem;line-height:1.7'>"
        NOTE+="<b>$(L "Indicatori di fuga dal container" "Container escape indicators")</b><br>"
        NOTE+="$(L "PRIVILEGED e MOUNT_ROOT_HOST danno di fatto accesso completo all'host. DOCKER_SOCKET consente di creare nuovi container privilegiati. CAP_SYS_ADMIN e CAP_SYS_MODULE permettono di caricare moduli kernel. PID_HOST espone i processi dell'host." \
            "PRIVILEGED and MOUNT_ROOT_HOST effectively grant full host access. DOCKER_SOCKET allows spawning new privileged containers. CAP_SYS_ADMIN and CAP_SYS_MODULE allow loading kernel modules. PID_HOST exposes host processes.")"
        NOTE+="</div></div>"
    fi

    local STATS
    STATS="$(stat_box "Container" "$TOTAL")"
    STATS+="$(stat_box "$(L "A rischio" "At risk")" "$NRISK" "$([[ "$NRISK" -gt 0 ]] && echo warn || echo info)")"
    STATS+="$(stat_box "$(L "Log stdout" "stdout logs")" "$NLOGS" "info")"
    finish_report "linux_containers" "Linux Container Forensics" "CNT" "/var/lib/docker · /var/lib/containers" "$STATS" \
        "${NOTE}<div class='cards'>$(generic_card_html "$(L "Inventario container" "Container inventory")" "${DOCKER:-$PODMAN}" "$TOTAL" "$TABLE" "▣")</div>"
}
