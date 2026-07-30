#!/usr/bin/env python3
"""Compila ogni parser Python incorporato in fiuto.sh.

`bash -n` valida la sintassi shell ma tratta gli heredoc come testo opaco:
un errore di sintassi in uno dei parser Python inline resta invisibile fino
al momento in cui quel modulo viene eseguito su un disco reale — cioè nel
punto peggiore possibile.

Questo linter estrae ogni blocco

    ... << 'PYEOF'
    <codice python>
    PYEOF

e lo passa a compile(). Esce con stato 1 al primo blocco non compilabile.
"""
from __future__ import annotations

import re
import sys

# Cattura l'apertura di un heredoc quotato il cui delimitatore contiene "PY".
# Il quoting (<< 'PYEOF') garantisce che bash non espanda il contenuto, quindi
# quello che leggiamo è esattamente ciò che riceverà l'interprete Python.
OPEN = re.compile(r"<<\s*'(?P<delim>[A-Za-z_][A-Za-z0-9_]*PY[A-Za-z0-9_]*|PY[A-Za-z0-9_]*)'")


def extract_blocks(path: str) -> list[tuple[int, str, str]]:
    """Ritorna [(riga_iniziale, delimitatore, sorgente)] per ogni heredoc Python."""
    with open(path, encoding="utf-8", errors="replace") as fh:
        lines = fh.readlines()

    blocks: list[tuple[int, str, str]] = []
    i = 0
    while i < len(lines):
        # Le righe di commento non aprono heredoc: un esempio d'uso in un
        # commento (`# ... << 'PYEOF'`) faceva credere al linter di essere
        # dentro un blocco mai chiuso.
        if lines[i].lstrip().startswith("#"):
            i += 1
            continue
        match = OPEN.search(lines[i])
        if not match:
            i += 1
            continue
        delim = match.group("delim")
        start = i + 1
        body: list[str] = []
        i += 1
        while i < len(lines) and lines[i].rstrip("\n").strip() != delim:
            body.append(lines[i])
            i += 1
        if i >= len(lines):
            print(f"{path}:{start}: heredoc '{delim}' non chiuso", file=sys.stderr)
            sys.exit(1)
        blocks.append((start + 1, delim, "".join(body)))
        i += 1
    return blocks


def main() -> int:
    if len(sys.argv) < 2:
        print("uso: lint_embedded_python.py <script.sh>...", file=sys.stderr)
        return 2

    failures = 0
    total = 0
    for path in sys.argv[1:]:
        for lineno, delim, source in extract_blocks(path):
            total += 1
            try:
                compile(source, f"{path}:{lineno}", "exec")
            except SyntaxError as exc:
                failures += 1
                offset = lineno + (exc.lineno or 1) - 1
                print(
                    f"{path}:{offset}: SyntaxError nel blocco '{delim}': {exc.msg}",
                    file=sys.stderr,
                )
                if exc.text:
                    print(f"    {exc.text.rstrip()}", file=sys.stderr)

    print(f"Parser Python incorporati compilati: {total - failures}/{total}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
