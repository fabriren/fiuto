
# Escape HTML — usare questa invece delle funzioni _esc_X locali nei moduli
# (per contenuto testuale tra i tag: basta neutralizzare & < >)
html_esc() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

# Escape HTML per valori dentro un ATTRIBUTO (title='...', data-*='...'):
# oltre a & < > neutralizza anche gli apici, che altrimenti chiuderebbero l'attributo.
html_attr() { printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g; s/"/\&quot;/g; s/'"'"'/\&#39;/g'; }

# Calcola SHA256 di un file (per chain of custody); ritorna stringa vuota se fallisce
sha256_file() { sha256sum "$1" 2>/dev/null | awk '{print $1}' || true; }

# Scrive un messaggio nel log di sessione (se LOG_FILE è impostato)
log_msg() {
    [[ -n "$LOG_FILE" ]] || return 0
    local _ld; _ld=$(dirname "$LOG_FILE")
    [[ -d "$_ld" ]] || mkdir -p "$_ld" 2>/dev/null || return 0
    printf '[%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$LOG_FILE" 2>/dev/null || true
}
