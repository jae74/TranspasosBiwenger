# python
import requests


class SofaScoreScraper:
    def __init__(self):
        self.base_url = "https://www.sofascore.com/api"


url = "https://www.sofascore.com/api/v1/unique-tournament/54/season/62048/statistics"
params = {
    "limit": 20,
    "order": "-rating",
    "accumulation": "total",
    "fields": "goals,successfulDribbles,tackles,assists,accuratePassesPercentage,rating",
    "filters": "position.in.G~D~M~F",
}

# stronger browser-like headers
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.sofascore.com/",
    "Origin": "https://www.sofascore.com",
    "Connection": "keep-alive",
    "X-Requested-With": "XMLHttpRequest",
}

with requests.Session() as s:
    s.headers.update(headers)
    resp = s.get(url, params=params, timeout=10, verify=True)
    print("status:", resp.status_code)
    print("response headers:", resp.headers)
    # server message often in body; print first 2000 chars to inspect
    print("body preview:", resp.text[:2000])

# If still 403, try cloudscraper which handles some anti-bot checks:
try:
    import cloudscraper
    scraper = cloudscraper.create_scraper()
    resp2 = scraper.get(url, params=params, headers=headers, timeout=15)
    print("cloudscraper status:", resp2.status_code)
    print("cloudscraper body preview:", resp2.text[:2000])
except Exception as e:
    print("cloudscraper unavailable or failed:", e)
