import requests
from openpyxl import load_workbook
from datetime import datetime

# =====================
# НАСТРОЙКИ
# =====================
EXCEL_F
EXCEL_FILE = "Гол_во_втором_тайме_с_LIVE_HT.xlsx"

SHEET_NAME = "LIVE_HT"

SOFA_LIVE_API = "https://api.sofascore.com/api/v1/sport/football/events/live"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json"
}

# =====================
# ЗАГРУЗКА EXCEL
# =====================
wb = load_workbook(EXCEL_FILE)
ws = wb[SHEET_NAME]

# очищаем старые данные, оставляя заголовки
if ws.max_row > 1:
    ws.delete_rows(2, ws.max_row)

# =====================
# ЗАГРУЗКА LIVE МАТЧЕЙ
# =====================
resp = requests.get(SOFA_LIVE_API, headers=HEADERS, timeout=30)
data = resp.json()

events = data.get("events", [])

match_number = 1

for ev in events:
    try:
        home = ev["homeTeam"]["name"]
        away = ev["awayTeam"]["name"]

        # текущий счёт
        home_score = ev["homeScore"].get("current", 0)
        away_score = ev["awayScore"].get("current", 0)
        score = f"{home_score}-{away_score}"

        # статус матча
        status = ev["status"]["description"]

        # статистика (может отсутствовать)
        stats = ev.get("statistics", {})
        poss_home = None
        poss_away = None
        shots_home = None
        shots_away = None

        ws.append([
            match_number,
            home,
            away,
            score,
            None,   # П1
            None,   # П2
            None,   # ТМ4.5
            poss_home,
            poss_away,
            shots_home,
            shots_away,
            None, None, None, None,   # последние 2 игры К1
            None, None, None, None    # последние 2 игры К2
        ])

        match_number += 1

    except Exception:
        continue

# =====================
# СЛУЖЕБНАЯ ИНФОРМАЦИЯ
# =====================
ws["T1"] = "Обновлено:"
ws["T2"] = datetime.now().strftime("%d.%m.%Y %H:%M")
ws["T3"] = f"Live матчей: {match_number - 1}"

# =====================
# СОХРАНЕНИЕ
# =====================
wb.save(EXCEL_FILE)

# =====================
# СЛУЖЕБНАЯ ИНФОРМАЦИЯ
# =====================
ws["T1"] = "Обновлено:"
ws["T2"] = datetime.now().strftime("%d.%m.%Y %H:%M")
ws["T3"] = f"Матчей найден
