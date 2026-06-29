#!/usr/bin/env python3
"""Keep-alive prod (protocol.streamlit.app).

Streamlit Community Cloud usypia apki po okresie bez ruchu. Zwykły GET NIE budzi
(liczy się sesja WebSocket — realnie otwarta apka w przeglądarce), dlatego tu
headless Chromium: wchodzi na URL, jeśli widać ekran „Zzzz" klika „wake up"
i czeka aż apka się zbootuje (cold boot ~3 min — apka ma ciężkie zależności,
m.in. libreoffice). Odpalane z GitHub Actions co 30 min (repo publiczne = darmowe).

Świadomie NIE failuje pipeline'u przy niepewności (żeby nie spamować maili o
czerwonym CI) — twardy błąd tylko gdy przeglądarka w ogóle nie wstanie.
"""
import sys
from playwright.sync_api import sync_playwright

URL = "https://protocol.streamlit.app/"
APP_MARKER = "Mölkky"            # tekst z headera działającej apki
SLEEP_MARKER = "get this app back up"  # przycisk na ekranie uśpienia
BOOT_TIMEOUT_MS = 200_000        # ~3.3 min na cold boot ciężkiej apki


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=60_000)
        except Exception as e:
            print(f"[keepalive] Nie udało się załadować {URL}: {e}")
            browser.close()
            return 1  # twardy błąd — sieć/przeglądarka padła

        # Ekran uśpienia? Kliknij „wake up".
        try:
            wake = page.get_by_text(SLEEP_MARKER, exact=False)
            if wake.count() > 0:
                wake.first.click()
                print("[keepalive] Apka spała — kliknięto wake-up, budzę…")
            else:
                print("[keepalive] Brak ekranu uśpienia — apka już aktywna lub bootuje.")
        except Exception as e:
            print(f"[keepalive] Próba kliknięcia wake pominięta: {e}")

        # Potwierdź boot: czekaj aż pojawi się treść apki.
        try:
            page.wait_for_selector(f"text={APP_MARKER}", timeout=BOOT_TIMEOUT_MS)
            print("[keepalive] OK — apka obudzona i aktywna.")
        except Exception:
            # Nie potwierdzono w czasie — ping i tak wykonany, sesja nawiązana.
            print("[keepalive] Nie potwierdzono treści w limicie czasu "
                  "(ping wykonany; apka może jeszcze bootować).")

        browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
