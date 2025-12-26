from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
from datetime import datetime

EXCEL_FILE = "Гол_во_втором_тайме_с_LIVE_HT.xlsx"
SHEET_NAME = "LIVE_HT"
URL = "https://www.flashscore.com/"

# открываем Excel (НЕ создаём!)
wb = load_workbook(EXCEL_FILE)
ws = wb[SHEET_NAME]

# очищаем старые данные, заголовки оставляем
if ws.max_row > 1:
    ws.delete_rows(2, ws.max_row)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    )

    page.goto(URL, timeout=60000)
    page.wait_for_timeout(5000)

    # ждём появления live-матчей
    page.wait_for_selector(".event__match", timeout=60000)

    matches = page.query_selector_all(".event__match")

    match_number = 1

    for match in matches:
        try:
            home = match.query_selector(".event__participant--home").inner_text().strip()
            away = match.query_selector(".event__participant--away").inner_text().strip()

            score_el = match.query_selector(".event__scores")
            score = score_el.inner_text().replace(":", "-") if score_el else ""

            ws.append([
                match_number,
                home,
                away,
                score,
                None, None, None,      # П1, П2, ТМ4.5
                None, None,            # владение
                None, None,            # удары
                None, None, None, None,
                None, None, None, None
            ])

            match_number += 1

        except Exception:
            continue

    browser.close()

# служебная информация
ws["T1"] = "Обновлено:"
ws["T2"] = datetime.now().strftime("%d.%m.%Y %H:%M")
ws["T3"] = f"Live матчей: {match_number - 1}"

wb.save(EXCEL_FILE)
