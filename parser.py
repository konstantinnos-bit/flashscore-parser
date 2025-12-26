
# FlashScore LIVE HT parser
# Writes data ONLY to LIVE_HT sheet

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

BASE_URL = "https://www.flashscore.com"
HEADERS = {"User-Agent": "Mozilla/5.0"}
EXCEL_FILE = "Гол_во_втором_тайме_с_LIVE_HT.xlsx"
SHEET_NAME = "LIVE_HT"


def get_live_matches():
    url = BASE_URL + "/football/"
    html = requests.get(url, headers=HEADERS, timeout=20).text
    soup = BeautifulSoup(html, "lxml")

    matches = []
    for row in soup.select("div.event__match"):
        status = row.select_one(".event__stage")
        score = row.select_one(".event__score")

        if not status or not score:
            continue

        if status.text.strip().lower() not in ["ht", "half-time"]:
            continue

        if score.text.strip() not in ["0-0", "1-0", "0-1"]:
            continue

        match_id = row.get("id", "").replace("g_1_", "")
        if match_id:
            matches.append(match_id)

    return matches


def parse_match(match_id):
    url = f"{BASE_URL}/match/{match_id}/#/match-summary"
    soup = BeautifulSoup(requests.get(url, headers=HEADERS).text, "lxml")

    teams = soup.select(".participant__participantName")
    team1, team2 = teams[0].text, teams[1].text

    score = soup.select_one(".detailScore__wrapper").text.strip()

    odds = soup.select(".oddsValueInner")
    p1 = int(float(odds[0].text))
    p2 = int(float(odds[2].text))
    tm45 = float(odds[5].text)

    shots = soup.select(".stat__value")
    try:
        s1, s2 = int(shots[2].text), int(shots[3].text)
    except:
        s1 = s2 = ""

    return team1, team2, score, p1, p2, tm45, s1, s2


def write_excel(rows):
    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    ws.delete_rows(2, ws.max_row)

    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=val)

    wb.save(EXCEL_FILE)


def main():
    rows = []
    for i, match_id in enumerate(get_live_matches(), start=1):
        team1, team2, score, p1, p2, tm45, s1, s2 = parse_match(match_id)

        rows.append([
            i, team1, team2, score, p1, p2, tm45,
            "", "", s1, s2,
            "", "", "", "",
            "", "", "", ""
        ])

    write_excel(rows)


if __name__ == "__main__":
    main()
