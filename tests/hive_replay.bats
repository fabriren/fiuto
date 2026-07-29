#!/usr/bin/env bats
#
# Test del replay dei transaction log del registro.
#
# Il percorso "successo" richiede regipy e hive dirty reali, quindi non è
# riproducibile in CI: qui si verificano gli invarianti che valgono sempre,
# in particolare che recover_hive non fallisca MAI in modo distruttivo —
# qualunque problema deve degradare sull'hive originale, non interrompere
# l'analisi.

setup() {
    export FIUTO_LIB_ONLY=1
    REPO_ROOT="$(cd "$(dirname "$BATS_TEST_FILENAME")/.." && pwd)"
    # shellcheck disable=SC1091
    source "$REPO_ROOT/fiuto.sh"
    FIXTURE="$(mktemp -d)"
    export TMPDIR="$FIXTURE/tmp"
    mkdir -p "$TMPDIR"
}

teardown() {
    [[ -n "${FIXTURE:-}" && -d "$FIXTURE" ]] && rm -rf "$FIXTURE"
}

_make_hive() {
    mkdir -p "$FIXTURE/Windows/System32/config"
    printf 'regf-fake-hive' > "$FIXTURE/Windows/System32/config/SYSTEM"
}

@test "recover_hive su path vuoto ritorna vuoto" {
    run recover_hive ""
    [ -z "$output" ]
}

@test "recover_hive su file inesistente ritorna vuoto" {
    run recover_hive "$FIXTURE/non-esiste"
    [ -z "$output" ]
}

@test "recover_hive senza transaction log ritorna l'hive originale" {
    _make_hive
    run recover_hive "$FIXTURE/Windows/System32/config/SYSTEM"
    [ "$output" = "$FIXTURE/Windows/System32/config/SYSTEM" ]
}

@test "recover_hive con log a zero byte ritorna l'originale" {
    _make_hive
    : > "$FIXTURE/Windows/System32/config/SYSTEM.LOG1"
    : > "$FIXTURE/Windows/System32/config/SYSTEM.LOG2"
    run recover_hive "$FIXTURE/Windows/System32/config/SYSTEM"
    [ "$output" = "$FIXTURE/Windows/System32/config/SYSTEM" ]
}

@test "recover_hive rispetta HIVE_REPLAY=false" {
    _make_hive
    printf 'HvLE-fake-log' > "$FIXTURE/Windows/System32/config/SYSTEM.LOG1"
    HIVE_REPLAY=false
    run recover_hive "$FIXTURE/Windows/System32/config/SYSTEM"
    [ "$output" = "$FIXTURE/Windows/System32/config/SYSTEM" ]
}

# Con log presenti ma regipy assente o hive non valido il replay fallisce:
# deve comunque restituire un percorso utilizzabile, mai una stringa vuota.
#
# NB: si cattura solo lo stdout — `run` di bats fonde stderr dentro $output,
# mentre qui la garanzia da verificare è proprio che sullo stdout ci sia
# unicamente il percorso.
@test "recover_hive degrada sull'originale se il replay non riesce" {
    _make_hive
    printf 'HvLE-fake-log-non-valido' > "$FIXTURE/Windows/System32/config/SYSTEM.LOG1"
    out="$(recover_hive "$FIXTURE/Windows/System32/config/SYSTEM" 2>/dev/null)"
    [ -n "$out" ]
    [ -f "$out" ]
    [ "$out" = "$FIXTURE/Windows/System32/config/SYSTEM" ]
}

@test "recover_hive non scrive mai sul volume di evidenza" {
    _make_hive
    printf 'HvLE-fake-log' > "$FIXTURE/Windows/System32/config/SYSTEM.LOG1"
    before="$(find "$FIXTURE/Windows" -type f | sort)"
    run recover_hive "$FIXTURE/Windows/System32/config/SYSTEM"
    after="$(find "$FIXTURE/Windows" -type f | sort)"
    [ "$before" = "$after" ]
}

@test "_hive_recovered_path è deterministico per lo stesso hive" {
    _make_hive
    a="$(_hive_recovered_path "$FIXTURE/Windows/System32/config/SYSTEM")"
    b="$(_hive_recovered_path "$FIXTURE/Windows/System32/config/SYSTEM")"
    [ "$a" = "$b" ]
    [ -n "$a" ]
}

# Più utenti hanno hive con lo stesso basename (NTUSER.DAT): le copie
# ricostruite non devono sovrascriversi a vicenda.
@test "_hive_recovered_path distingue hive omonimi di utenti diversi" {
    mkdir -p "$FIXTURE/Users/mario" "$FIXTURE/Users/anna"
    touch "$FIXTURE/Users/mario/NTUSER.DAT" "$FIXTURE/Users/anna/NTUSER.DAT"
    a="$(_hive_recovered_path "$FIXTURE/Users/mario/NTUSER.DAT")"
    b="$(_hive_recovered_path "$FIXTURE/Users/anna/NTUSER.DAT")"
    [ "$a" != "$b" ]
}

@test "la directory degli hive ricostruiti è accessibile solo al proprietario" {
    _make_hive
    d="$(_hive_replay_tmpdir)"
    [ -d "$d" ]
    [ "$(stat -c '%a' "$d")" = "700" ]
}

@test "get_hive risolve l'hive con case diverso nel percorso" {
    mkdir -p "$FIXTURE/WINDOWS/system32/CONFIG"
    printf 'regf' > "$FIXTURE/WINDOWS/system32/CONFIG/system"
    WIN_ROOT="$FIXTURE"
    run get_hive SYSTEM
    [ -n "$output" ]
    [ -f "$output" ]
}

@test "get_hive ritorna vuoto se l'hive non esiste" {
    WIN_ROOT="$FIXTURE"
    run get_hive SOFTWARE
    [ -z "$output" ]
}

@test "get_user_hive ritorna vuoto se l'utente non ha l'hive" {
    mkdir -p "$FIXTURE/Users/vuoto"
    run get_user_hive "$FIXTURE/Users/vuoto" "NTUSER.DAT"
    [ -z "$output" ]
}

@test "recover_hive non emette diagnostica sullo stdout" {
    # recover_hive è invocata in command substitution: qualunque messaggio
    # su stdout finirebbe dentro il percorso restituito.
    _make_hive
    printf 'HvLE-fake' > "$FIXTURE/Windows/System32/config/SYSTEM.LOG1"
    out="$(recover_hive "$FIXTURE/Windows/System32/config/SYSTEM" 2>/dev/null)"
    [ "$out" = "$FIXTURE/Windows/System32/config/SYSTEM" ]
    [[ "$out" != *"["* ]]
}
