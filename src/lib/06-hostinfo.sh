
# ================================================================
#  RECUPERO INFORMAZIONI MACCHINA
# ================================================================

gather_host_info() {
    [[ -n "$WIN_ROOT" ]] || return 1

    # Per i volumi non-Windows usa una raccolta info dedicata e termina qui.
    if [[ "$OS_TYPE" == "linux" ]]; then
        gather_host_info_linux
        return 0
    elif [[ "$OS_TYPE" == "macos" ]]; then
        gather_host_info_macos
        return 0
    fi

    local SYSTEM_HIVE; SYSTEM_HIVE=$(get_hive "SYSTEM")
    local SOFTWARE_HIVE; SOFTWARE_HIVE=$(get_hive "SOFTWARE")
    
    if [[ -z "$SYSTEM_HIVE" && -z "$SOFTWARE_HIVE" ]]; then
        warn "$(t hive_not_found)"
    else
        info "$(t retrieving_info)"
        
        local INFO_JSON
        INFO_JSON=$("$PY3" - "$SYSTEM_HIVE" "$SOFTWARE_HIVE" << 'PYEOF' 2>/dev/null
import sys, json
try:
    from regipy.registry import RegistryHive
    system_path = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] else None
    software_path = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] else None

    info = {"hostname": "", "os": "", "ip": "", "domain": ""}

    def get_key(hive, path):
        """Prova il path diretto e con prefisso ROOT\\ (regipy >= 6.x)."""
        for p in (path, 'ROOT\\' + path):
            try:
                return hive.get_key(p)
            except Exception:
                pass
        return None

    if system_path:
        try:
            sys_hive = RegistryHive(system_path)
            # Hostname
            try:
                hk = get_key(sys_hive, 'ControlSet001\\Control\\ComputerName\\ComputerName')
                if hk: info["hostname"] = hk.get_value('ComputerName') or ""
            except: pass

            # Network Info (IP / Domain)
            try:
                tk = get_key(sys_hive, 'ControlSet001\\Services\\Tcpip\\Parameters')
                if tk:
                    info["domain"] = tk.get_value('Domain') or tk.get_value('NV Domain') or ""

                # IP (cerca il primo con un IP valido)
                ik = get_key(sys_hive, 'ControlSet001\\Services\\Tcpip\\Parameters\\Interfaces')
                if ik:
                    for sub in ik.iter_subkeys():
                        ip = sub.get_value('DhcpIPAddress') or sub.get_value('IPAddress')
                        if ip and ip != '0.0.0.0':
                            if isinstance(ip, list): ip = ip[0]
                            info["ip"] = ip
                            break
            except: pass
        except: pass

    if software_path:
        try:
            soft_hive = RegistryHive(software_path)
            try:
                cvk = get_key(soft_hive, 'Microsoft\\Windows NT\\CurrentVersion')
                if cvk:
                    prod = cvk.get_value('ProductName') or ""
                    ver = cvk.get_value('DisplayVersion') or cvk.get_value('ReleaseId') or ""
                    build = cvk.get_value('CurrentBuild') or ""
                    info["os"] = f"{prod} {ver} (Build {build})".strip()
            except: pass
        except: pass

    print(json.dumps(info))
except:
    print("{}")
PYEOF
)
        HOST_NAME=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('hostname',''))" 2>/dev/null)
        OS_VER=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('os',''))" 2>/dev/null)
        IP_ADDR=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('ip',''))" 2>/dev/null)
        DOMAIN_NAME=$(echo "$INFO_JSON" | "$PY3" -c "import sys,json; print(json.load(sys.stdin).get('domain',''))" 2>/dev/null)
    fi

    # Fallback Hostname
    if [[ -z "$HOST_NAME" ]]; then
        local UNABLE_MSG="$([ "$LANG" = "it" ] && echo "Impossibile rilevare il nome macchina automaticamente." || echo "Unable to detect machine name automatically.")"
        warn "$UNABLE_MSG"
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} $(t hostname_prompt) [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi

    print_host_info_table
}

# Raccolta info host per volumi Linux (hostname + distro da /etc)
gather_host_info_linux() {
    info "$(t retrieving_info)"
    local ETC; ETC=$(ci_find_dir "$WIN_ROOT" "etc")
    if [[ -n "$ETC" ]]; then
        local HN; HN=$(ci_find_file "$ETC" "hostname")
        [[ -n "$HN" && -f "$HN" ]] && HOST_NAME=$(head -1 "$HN" 2>/dev/null | tr -d '[:space:]')
        local OSR; OSR=$(ci_find_file "$ETC" "os-release")
        if [[ -n "$OSR" && -f "$OSR" ]]; then
            OS_VER=$(grep -E '^PRETTY_NAME=' "$OSR" 2>/dev/null | head -1 | cut -d= -f2- | tr -d '"')
        fi
    fi
    [[ -z "$OS_VER" ]] && OS_VER="Linux"
    DOMAIN_NAME=""; IP_ADDR=""
    if [[ -z "$HOST_NAME" ]]; then
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} $(t hostname_prompt) [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi
    print_host_info_table
}

# Raccolta info host per volumi macOS (da SystemVersion.plist)
gather_host_info_macos() {
    info "$(t retrieving_info)"
    local SV
    SV=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "System/Library/CoreServices")" "SystemVersion.plist")
    if [[ -n "$SV" && -f "$SV" ]]; then
        local PLIST_TXT; PLIST_TXT=$(read_plist "$SV" 2>/dev/null)
        local PROD VER BUILD
        PROD=$(echo "$PLIST_TXT"  | grep -i 'ProductName:'        | head -1 | sed "s/.*: '\\?//; s/'\\?$//")
        VER=$(echo "$PLIST_TXT"   | grep -i 'ProductVersion:'     | head -1 | sed "s/.*: '\\?//; s/'\\?$//")
        BUILD=$(echo "$PLIST_TXT" | grep -i 'ProductBuildVersion:'| head -1 | sed "s/.*: '\\?//; s/'\\?$//")
        OS_VER=$(echo "${PROD} ${VER} (${BUILD})" | sed 's/  */ /g; s/ ()//')
    fi
    [[ -z "$OS_VER" ]] && OS_VER="macOS"
    # hostname: prova preferences.plist di SystemConfiguration
    local PREF
    PREF=$(ci_find_file "$(ci_find_dir "$WIN_ROOT" "Library/Preferences/SystemConfiguration")" "preferences.plist")
    if [[ -n "$PREF" && -f "$PREF" ]]; then
        HOST_NAME=$(read_plist "$PREF" 2>/dev/null | grep -iE 'HostName:|LocalHostName:|ComputerName:' | head -1 | sed "s/.*: '\\?//; s/'\\?$//")
    fi
    DOMAIN_NAME=""; IP_ADDR=""
    if [[ -z "$HOST_NAME" ]]; then
        local _SUGGESTED_NAME; _SUGGESTED_NAME=$(basename "$WIN_ROOT")
        echo -ne "  ${YELLOW}[?]${RESET} $(t hostname_prompt) [${BOLD}${_SUGGESTED_NAME}${RESET}]: "
        read -r HOST_NAME
        [[ -z "$HOST_NAME" ]] && HOST_NAME="$_SUGGESTED_NAME"
    fi
    print_host_info_table
}

print_host_info_table() {
    local TITLE="$([ "$LANG" = "it" ] && echo "INFORMAZIONI MACCHINA TARGET" || echo "TARGET MACHINE INFORMATION")"
    echo -e "  ${CYAN}${BOLD}┌────────────────────────────────────────────────────────────────────────────┐${RESET}"
    local _T=${#TITLE} _LEFT=$(( (76 - ${#TITLE}) / 2 )) _RIGHT
    _RIGHT=$(( 76 - ${#TITLE} - _LEFT ))
    printf "  ${CYAN}${BOLD}│%*s%*s│${RESET}\n" $(( _LEFT + _T )) "$TITLE" $_RIGHT ""
    echo -e "  ${CYAN}${BOLD}├────────────────────────────────────────────────────────────────────────────┤${RESET}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : ${WHITE}${BOLD}%-52s${RESET} ${CYAN}${BOLD}│${RESET}\n" "$(t hostname)" "${HOST_NAME:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "$(t os)" "${OS_VER:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "$(t ip)" "${IP_ADDR:-N/A}"
    printf "  ${CYAN}${BOLD}│${RESET}  %-18s : %-52s ${CYAN}${BOLD}│${RESET}\n" "$(t domain)" "${DOMAIN_NAME:-N/A}"
    echo -e "  ${CYAN}${BOLD}└────────────────────────────────────────────────────────────────────────────┘${RESET}"
    echo ""
}
