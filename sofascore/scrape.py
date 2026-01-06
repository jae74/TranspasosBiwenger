# python
from urllib.parse import urlencode
import json

class SofaScoreScraper:
    def __init__(self, fields="goals,successfulDribbles,tackles,assists,accuratePassesPercentage,rating"):
        self.base_url = "https://www.sofascore.com/api/v1/unique-tournament/54/season/62048/statistics"
        self.params = {
            "limit": 20,
            "order": "-rating",
            "accumulation": "total",
            "fields": fields,
            "filters": "position.in.G~D~M~F",
            "group": "summary",
        }
        self.user_agent = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )

    def fetch_statistics(self):
        """
        Use Playwright to load the main site (solve JS/anti-bot) then request the API.
        Prints status, headers preview and body preview for debugging. Returns parsed JSON on 200.
        """
        try:
            from playwright.sync_api import sync_playwright
        except ModuleNotFoundError:
            print("Playwright not installed. Run: pip install playwright && playwright install")
            return None

        full = f"{self.base_url}?{urlencode(self.params)}"
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(user_agent=self.user_agent, locale="en-US")
            page = context.new_page()

            # Visit main page to obtain cookies / solve JS challenges
            page.goto("https://www.sofascore.com/", wait_until="networkidle")

            # Request the API endpoint
            resp = page.goto(full, wait_until="networkidle")
            if not resp:
                print("No response object from Playwright when requesting the API.")
                browser.close()
                return None

            status = resp.status
            headers = resp.all_headers()
            body = resp.text()

            print("status:", status)
            print("response headers preview:", {k: headers.get(k) for k in ("content-type", "server", "cache-control") if k in headers})
            print("body preview:", body[:2000])

            browser.close()

            if status == 200:
                try:
                    # Save JSON response into ./json/sofascore.json with indent
                    with open("./json/sofascore.json", "w", encoding="utf-8") as f:
                        json.dump(json.loads(body), f, indent=4, ensure_ascii=False)
                    return json.loads(body)
                except json.JSONDecodeError:
                    print("Response body is not valid JSON.")
                    return None
            return None


if __name__ == "__main__":
    scraper = SofaScoreScraper()
    stats = scraper.fetch_statistics()
    print(stats)
