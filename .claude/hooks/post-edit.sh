#!/bin/bash
# PostToolUse hook: walidacja krytycznych plików Python po Edit/Write/MultiEdit.
# Wywoływany przez harness Claude Code z JSON tool-event na stdin.
#
# Walidacja:
#   1) py_compile (składnia)
#   2) tests/regression.py (build_document na 9 szablonach + placeholder filter)
#      — tylko gdy edytowano generate_docx.py
# Exit 2 = błąd widoczny dla Claude (do natychmiastowej naprawy).

set -uo pipefail

input=$(cat)
file=$(printf '%s' "$input" | python3 -c "
import sys, json
try:
    d = json.load(sys.stdin)
    print(d.get('tool_input', {}).get('file_path', ''))
except Exception:
    print('')
" 2>/dev/null)

case "$file" in
    */app.py|*/generate_docx.py)
        proj="${CLAUDE_PROJECT_DIR:-/home/user/protocol}"
        fail=0

        # 1) py_compile (składnia)
        if ! out=$(python3 -m py_compile "$file" 2>&1); then
            printf '✗ py_compile FAIL %s:\n%s\n' "$file" "$out" >&2
            fail=1
        fi

        # 2) regresja — tylko dla generate_docx.py (app.py wymaga Streamlit + sieci)
        case "$file" in
            */generate_docx.py)
                reg="$proj/tests/regression.py"
                if [ -f "$reg" ]; then
                    if ! reg_out=$(PYTHONPATH="$proj" python3 "$reg" 2>&1); then
                        printf '✗ regresja FAIL po edycji %s:\n%s\n' "$file" "$reg_out" >&2
                        fail=1
                    fi
                fi
                ;;
        esac

        if [ "$fail" -ne 0 ]; then
            exit 2
        fi
        printf '✓ post-edit OK: %s\n' "$(basename "$file")" >&2
        ;;
esac
exit 0
