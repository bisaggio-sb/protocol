#!/bin/bash
# PreToolUse hook na Bash: blokuje 'git push' jeśli tests/regression.py czerwona.
# To ostatnia linia obrony — bug który ominął post-edit hook tu się zatrzyma.

set -uo pipefail

input=$(cat)
cmd=$(printf '%s' "$input" | python3 -c "
import sys, json
try:
    d = json.load(sys.stdin)
    print(d.get('tool_input', {}).get('command', ''))
except Exception:
    print('')
" 2>/dev/null)

# Tylko jeśli komenda zawiera 'git push' (i nie jest np. 'echo git push')
case "$cmd" in
    *"git push"*)
        proj="${CLAUDE_PROJECT_DIR:-/home/user/protocol}"
        reg="$proj/tests/regression.py"
        if [ ! -f "$reg" ]; then
            exit 0
        fi
        # Sprawdź czy 'git push' faktycznie odpalany jako komenda (nie np. w komentarzu)
        # Heurystyka: trim whitespace, check że nie zaczyna się od 'echo'
        if ! out=$(PYTHONPATH="$proj" python3 "$reg" 2>&1); then
            printf '✗ git push ZABLOKOWANY — regresja FAIL:\n%s\n' "$out" >&2
            printf '\nNapraw kod i odpal regresję ręcznie:\n' >&2
            printf '  PYTHONPATH=%s python3 %s\n' "$proj" "$reg" >&2
            exit 2
        fi
        printf '✓ regresja zielona, push dozwolony\n' >&2
        ;;
esac
exit 0
