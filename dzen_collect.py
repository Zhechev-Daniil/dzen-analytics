# -*- coding: utf-8 -*-
"""
dzen_collect.py — ежедневный сбор данных с Дзена и обновление Excel
Запускается автоматически через планировщик Cowork или вручную.

Требования:
  pip install playwright openpyxl
  playwright install chromium
"""

import json
import subprocess
import sys
import os
from pathlib import Path

PUBLISHER_ID = "69cd140eea571e1d07752e89"
SCRIPT_DIR   = Path(__file__).parent
UPDATE_SCRIPT = SCRIPT_DIR / "dzen_update.py"
EXCEL_PATH    = SCRIPT_DIR / "dzen_analytics.xlsx"


JS_EXTRACTOR = """
async (publisherId) => {
  const csrf = window._csrfToken;
  if (!csrf) throw new Error('No CSRF token');
  const headers = {'X-CSRF-Token': csrf};
  const base = 'https://dzen.ru/editor-api/v2/publisher/' + publisherId;

  const [statsResp, dailyResp, subResp] = await Promise.all([
    fetch(base + '/stats2?allPublications=true&fields=typeSpecificViews&fields=deepViewsRate&fields=ctr&fields=shares&fields=comments&fields=subscriptions&fields=likes&fields=impressions&fields=deepViews&fields=sumInvolvedViewTimeSec&fields=pageViews&groupBy=flight&sortBy=addTime&sortOrderDesc=true&total=true&pageSize=100&page=0&clid=320', {headers}),
    fetch(base + '/stats2?fields=typeSpecificViews&fields=impressions&fields=deepViews&fields=pageViews&fields=subscriptions&fields=likes&fields=sumInvolvedViewTimeSec&dateGroupBy=day&total=true&premiumOnly=false&clid=320', {headers}),
    fetch(base + '/stats2?fields=subscribers&fields=subscribersDiff&dateGroupBy=day&total=true&clid=320', {headers})
  ]);

  const stats = await statsResp.json();
  const daily = await dailyResp.json();
  const subs  = await subResp.json();

  const pubs = stats.publications.map(p => {
    const s = p.stats || {}, pub = p.publication || {};
    // Пытаемся найти публичный URL в разных полях которые может вернуть API
    const link = pub.shareUrl || pub.publicationUrl || pub.publicUrl || pub.webUrl
              || pub.pageUrl || pub.uri || pub.url || pub.link || pub.clickUrl
              || pub.originalUrl || pub.articleUrl || null;
    return {
      title: pub.title, id: pub.publicationId, type: pub.publicationType,
      link: link,
      pub_keys: Object.keys(pub).join(','),  // отладка — какие поля вообще есть
      pub_date: new Date(pub.addTime).toISOString().substring(0,10),
      impressions: s.impressions||0, page_views: s.pageViews||0,
      deep_views: s.deepViews||0, type_specific_views: s.typeSpecificViews||0,
      subscriptions: s.subscriptions||0, likes: s.likes||0,
      comments: s.comments||0, shares: s.shares||0,
      view_time_min: Math.round((s.sumInvolvedViewTimeSec||0)/60*10)/10,
      ctr_pct: Math.round((s.ctr||0)*10000)/100,
      deep_views_rate_pct: Math.round((s.deepViewsRate||0)*10000)/100
    };
  });

  const dailyIntervals = (daily.total?.intervals||[]).slice(-35).map(i => ({
    date: new Date(i.timestamp).toISOString().substring(0,10),
    impressions: i.impressions||0, page_views: i.pageViews||0,
    deep_views: i.deepViews||0, type_views: i.typeSpecificViews||0,
    subscriptions: i.subscriptions||0, likes: i.likes||0,
    view_time_min: Math.round((i.sumInvolvedViewTimeSec||0)/60*10)/10
  }));

  const subIntervals = (subs.total?.intervals||[]).filter(i=>i.subscribers>0).slice(-35).map(i=>({
    date: new Date(i.timestamp).toISOString().substring(0,10),
    subscribers: i.subscribers||0
  }));

  const t = stats.total?.stats||{};
  return JSON.stringify({
    pubCount: stats.publicationCount,
    pubs, dailyIntervals, subIntervals,
    channelTotal: {
      impressions: t.impressions||0, page_views: t.pageViews||0,
      deep_views: t.deepViews||0, type_views: t.typeSpecificViews||0,
      subscriptions: t.subscriptions||0, likes: t.likes||0,
      comments: t.comments||0, shares: t.shares||0,
      view_time_min: Math.round((t.sumInvolvedViewTimeSec||0)/60)
    }
  });
}
"""


def collect_via_playwright(publisher_id: str) -> str:
    """Открывает Дзен в браузере с профилем пользователя и собирает данные."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright не установлен. Установите: pip install playwright && playwright install chromium")
        sys.exit(1)

    # Путь к профилю Chrome пользователя
    chrome_profile = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=chrome_profile,
            channel="chrome",
            headless=False,  # видимый браузер, чтобы не потерять сессию
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        page = context.new_page()
        stats_url = f"https://dzen.ru/profile/editor/id/{publisher_id}/publications-stat?statType=publications&intervalType=infinite"
        page.goto(stats_url, wait_until="networkidle", timeout=30000)

        # Ждём загрузки данных
        page.wait_for_timeout(3000)

        js_call = f"({JS_EXTRACTOR.strip()})('{publisher_id}')"
        result = page.evaluate(js_call)
        context.close()
        return result


def collect_via_playwright_ci(publisher_id: str) -> str:
    """Headless-режим для GitHub Actions: куки берутся из DZEN_COOKIES."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright не установлен. Установите: pip install playwright && playwright install chromium")
        sys.exit(1)

    cookies_raw = os.environ.get("DZEN_COOKIES", "")
    if not cookies_raw:
        raise RuntimeError("Переменная окружения DZEN_COOKIES не задана")

    cookies = json.loads(cookies_raw)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context()
        context.add_cookies(cookies)
        page = context.new_page()

        stats_url = (
            f"https://dzen.ru/profile/editor/id/{publisher_id}"
            f"/publications-stat?statType=publications&intervalType=infinite"
        )
        page.goto(stats_url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(4000)

        js_call = f"({JS_EXTRACTOR.strip()})('{publisher_id}')"
        try:
            result = page.evaluate(js_call)
        except Exception as e:
            if "No CSRF token" in str(e):
                raise RuntimeError(
                    "DZEN_SESSION_EXPIRED: сессия истекла — "
                    "обнови куки через dzen_export_cookies.py и загрузи в GitHub Secret DZEN_COOKIES"
                )
            raise

        context.close()
        browser.close()
        return result


def update_excel(json_data: str):
    """Вызывает dzen_update.py для обновления Excel."""
    result = subprocess.run(
        [sys.executable, str(UPDATE_SCRIPT), json_data],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print(result.stdout)
    else:
        print("ОШИБКА обновления Excel:")
        print(result.stderr)
        sys.exit(1)


if __name__ == "__main__":
    print(f"Сбор данных Дзен для канала {PUBLISHER_ID}...")
    if os.environ.get("CI") or os.environ.get("DZEN_COOKIES"):
        print("Режим CI: используем headless Chromium с куками из DZEN_COOKIES")
        data_json = collect_via_playwright_ci(PUBLISHER_ID)
    else:
        data_json = collect_via_playwright(PUBLISHER_ID)
    print(f"Данные получены ({len(data_json)} символов). Обновляю Excel...")
    update_excel(data_json)
    print(f"Готово! Файл: {EXCEL_PATH}")
