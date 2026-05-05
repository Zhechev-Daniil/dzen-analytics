# -*- coding: utf-8 -*-
"""
dzen_export_cookies.py — экспорт кук Дзен для GitHub Actions.

Запускай на своём ПК:
  py -3.14 dzen_export_cookies.py

Откроется браузер — войди в аккаунт Дзен, затем нажми Enter в терминале.
Куки сохранятся в dzen_cookies.json.
"""

import json
import sys
from pathlib import Path

SCRIPT_DIR  = Path(__file__).parent
OUTPUT_FILE = SCRIPT_DIR / "dzen_cookies.json"


def main():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright не установлен. Выполни:")
        print("  py -3.14 -m pip install playwright")
        print("  py -3.14 -m playwright install chromium")
        sys.exit(1)

    print("Открываю браузер...")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--no-sandbox", "--start-maximized"]
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        try:
            page.goto("https://dzen.ru", wait_until="domcontentloaded", timeout=20000)
        except Exception:
            pass  # продолжаем даже если таймаут

        print()
        print("=" * 60)
        print("Браузер открылся.")
        print()
        print("1. Убедись что ты залогинен на dzen.ru")
        print("   (если нет — войди в аккаунт прямо сейчас)")
        print()
        print("2. Когда залогинен — вернись сюда и нажми Enter")
        print("=" * 60)
        input("  >>> Нажми Enter когда готово: ")

        all_cookies = context.cookies(["https://dzen.ru", "https://www.dzen.ru"])
        browser.close()

    if not all_cookies:
        print()
        print("ОШИБКА: куки не найдены. Попробуй снова и убедись что залогинен.")
        sys.exit(1)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(all_cookies, f, ensure_ascii=False, indent=2)

    print()
    print(f"Сохранено {len(all_cookies)} кук в файл:")
    print(f"  {OUTPUT_FILE}")
    print()
    print("=" * 60)
    print("СЛЕДУЮЩИЙ ШАГ — добавь куки в GitHub Secret:")
    print()
    print("  1. Открой файл dzen_cookies.json (он рядом со скриптом)")
    print("  2. Скопируй ВСЁ содержимое (Ctrl+A, Ctrl+C)")
    print("  3. Зайди на github.com/Zhechev-Daniil/dzen-analytics")
    print("  4. Settings → Secrets and variables → Actions")
    print("  5. New repository secret")
    print("  6. Name:  DZEN_COOKIES")
    print("  7. Value: вставь скопированное (Ctrl+V)")
    print("  8. Add secret")
    print("=" * 60)
    print()
    print("Куки живут 1-3 месяца. Когда Actions начнёт падать")
    print("с ошибкой DZEN_SESSION_EXPIRED — повтори этот скрипт.")


if __name__ == "__main__":
    main()
