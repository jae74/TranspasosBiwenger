from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from lxml import html as lxml_html
import re
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

class ConnectBiwenger:

    @staticmethod
    def get_parsed_html(url, cookies_dict=None, local_storage_dict=None):

        default_cookies = {
            'AMCV_2387401053DB208C0A490D4C%40AdobeOrg': '1176715910%7CMCIDTS%7C20333%7CMCMID%7C83191080408405748529212420616018517824%7CMCAID%7CNONE%7CMCOPTOUT-1756827643s%7CNONE%7CvVersion%7C5.4.0',
            'AMCVS_2387401053DB208C0A490D4C%40AdobeOrg': '1',
            'cmp573Consent': 'true',
            'cmpyoutubeConsent': 'true',
            'didomi_token': 'eyJ1c2VyX2lkIjoiMTk5MDk1MWEtZTVlYS02OWVkLTg1NzUtZWNjZDAwOTRkZTVmIiwiY3JlYXRlZCI6IjIwMjUtMDktMDJUMDc6MjY6MTIuMDYyWiIsInVwZGF0ZWQiOiIyMDI1LTA5LTAyVDA3OjI2OjE0Ljk2NFoiLCJ2ZW5kb3JzIjp7ImVuYWJsZWQiOlsiZ29vZ2xlIiwiYzpmYWNlYm9vay1ZeUpSQXllZCIsImM6eW91dHViZSIsImM6YW51bmNpYW50ZV9sYV9saWdhIiwiYzphZG9iZWF1ZGktQVdlN3J3cWQiLCJjOmJlc29jeS1tRlVFYVpCTSJdfSwicHVycG9zZXMiOnsiZW5hYmxlZCI6WyJkZXZpY2VfY2hhcmFjdGVyaXN0aWNzIiwiZ2VvbG9jYXRpb25fZGF0YSIsImRhdGFfc2hhcmluZ193ZWIiLCJkYXRhX3NoYXJpbmciXX0sInZlcnNpb24iOjIsImFjIjoiREU2QVFBRVlBTmdBbFFEekFJY0FpU0NCZ0dKdy5BQUFBIn0=',
            'euconsent-v2': 'CQXHgMAQXHgMAAHABBENB6FsAP_gAAAAAAAALhNR_G__bXlr-b736ftkeYxf9_hr7sQxBgbJk24FzLvW_JwW32E7NAzatqYKmRIAu3TBIQNlHJDURVCgKIgVrzDMaEyUoTtKJ6BkiFMRY2JYCFxvm4tjeQCY5vr99ld9mR-N7dr82dzyy6hnv3a9_-S1WJCdIYetDfv8ZBKT-9IE9_x8v4v4_N7pE2-eS1n_tGvp6D9-Yvv_dBn99_baffzPn__rl_e7X__f_n37v943X77_____f_-7AAAAMSgAwABBcIpABgACC4Q6ADAAEFwiEAGAAILhBIAMAAQXCLQAYAAguEAA.f_wAAAAAAAAA',
            'hpage': 'ES',
            'pmuser': '{"UT":"ANONYMOUS","lastUpdated":1756797971144}',
        }

        default_local_storage = {
            'lastSession': '{"account":{"id":4366698,"name":"Julio Agüero","email":"juliusague@gmail.com","phone":null,"locale":"es","birthday":0,"status":"valid","credits":3,"created":1755283595,"newsletter":false,"unreadMessages":false,"lastAccess":1756818835,"source":null,"devices":[{"type":"fcm","token":"eAKIQ7u3QzunN26MdQesi9:APA91bHW_8rsJAL-m-TaYPeNLJkuJgMunDzQiiyfycFuA17XTCU9WtuOr2LXGw_nyAwyto67RxR0RNr-zBVX6fNqHpIOpykLhNgBw0WxFZZd-LHIj5dyOps","updated":1756468776}]},"leagues":[{"id":1997867,"name":"Ligarrucha","competition":"segunda-division","scoreID":3,"type":"normal","mode":"league","marketMode":"normal","created":1755283697,"icon":"i/l/1997867.png?v=12","cover":"i/c/1997867.jpg?v=3","user":{"id":12989427,"name":"Llaule","balance":16980900,"icon":"i/u/12989427.png?v=33","role":"manager","type":"normal","joinDate":1755283697,"status":{"offers":1,"bids":2},"favorites":[],"points":34,"position":5},"settings":{"secret":"xwLSGtKnN02p","privacy":"private","onlyAdminPosts":false,"clause":"salePrice","clauseIncrement":1,"immediateSales":0,"balance":"hidden","userOffers":"always","loans":"disabled","loansMinRounds":1,"loansMaxRounds":5,"maxPurchasePrice":200,"challengesAllow":true,"roundDelayed":"recalculation","marketShowBids":false,"lineupMultiPos":false,"lineupAllowExtra":true,"lineupCoach":false,"lineupCaptain":false,"lineupStriker":false,"lineupReserves":false,"lineupMaxClubPlayers":false,"favoritesAllow":false,"auctions":false,"customScore":false},"upgrades":{"premium":{"id":"PremiumLeague","currency":"EUR","price":29.99,"google":"premium_30","huawei":"premium_30","apple":"premium_30"},"ultra":{"id":"UltraLeague","currency":"EUR","price":49.99,"google":"premium_50","huawei":"premium_50","apple":"premium_50"}}},{"id":2007349,"name":"Bellum Montis","competition":"la-liga","scoreID":5,"type":"normal","mode":"league","marketMode":"normal","created":1756109309,"icon":"","cover":"","user":{"id":13062802,"name":"Juliobriga FC","balance":6125900,"icon":"i/u/12989427.png?v=33","role":"manager","type":"normal","joinDate":1756109309,"status":{"offers":0,"bids":0},"favorites":[],"points":39,"position":5},"settings":{"secret":"hC0t74FwzWPp","privacy":"private","onlyAdminPosts":false,"clause":"salePrice","clauseIncrement":1,"immediateSales":0,"balance":"hidden","userOffers":"always","loans":"allow","loansMinRounds":1,"loansMaxRounds":5,"maxPurchasePrice":0,"challengesAllow":true,"roundDelayed":"recalculation","marketShowBids":false,"lineupMultiPos":false,"lineupAllowExtra":false,"lineupCoach":false,"lineupCaptain":false,"lineupStriker":false,"lineupReserves":false,"lineupMaxClubPlayers":false,"favoritesAllow":false,"auctions":false,"customScore":false},"upgrades":{"premium":{"id":"PremiumLeague","currency":"EUR","price":29.99,"google":"premium_30","huawei":"premium_30","apple":"premium_30"},"ultra":{"id":"UltraLeague","currency":"EUR","price":44.99,"google":"premium_50","huawei":"premium_50","apple":"premium_50"}}}],"notifications":[{"type":"homeStory","title":"🏆","campaign":"duplicar-liga-champions-26","coverImg":"https://cdn.biwenger.com/cdn-cgi/image/f=avif/campaigns/2025/football/crear-liga-champions-25-26/icono-1.jpg","id":574171,"body":"<img style=\"max-height: 80vh;\" src=\"https://cdn.biwenger.com/cdn-cgi/image/f=avif/campaigns/2025/football/crear-liga-champions-25-26/story1.jpg\" alt=\"\" />","date":1756807740},{"type":"homeStory","title":"⚽️","campaign":"duplicar-liga-champions-26","coverImg":"https://cdn.biwenger.com/cdn-cgi/image/f=avif/campaigns/2025/football/crear-liga-champions-25-26/icono-2.jpg","id":574172,"body":"<img style=\"max-height: 80vh;\" src=\"https://cdn.biwenger.com/cdn-cgi/image/f=avif/campaigns/2025/football/crear-liga-champions-25-26/story2.jpg\" alt=\"\" />","date":1756807680},{"type":"homeStory","title":"♠️","campaign":"duplicar-liga-champions-26","coverImg":"https://cdn.biwenger.com/cdn-cgi/image/f=avif/campaigns/2025/football/crear-liga-champions-25-26/icono-3.jpg","id":574173,"body":"<img style=\"max-height: 80vh;\" src=\"https://cdn.biwenger.com/cdn-cgi/image/f=avif/campaigns/2025/football/crear-liga-champions-25-26/s',
            'euconsent-v2': 'CQWzukAQWzukAAHABBENB5FsAP_gAAAAAAAALgtR_G__bXlr-b736ftkeYxf9_hr7sQxBgbJk24FzLvW_JwW32E7NAzatqYKmRIAu3TBIQNlHJDURVCgKIgVrzDMaEyUoTtKJ6BkiFMRY2JYCFxvm4tjeQCY5vr99ld9mR-N7dr82dzyy6hnv3a9_-S1WJCdIYetDfv8ZBKT-9IE9_x8v4v4_N7pE2-eS1n_tGvp6D9-Yvv_dBn99_baffzPn__rl_e7X__f_n37v943X77_____f_-6AAAAYlABgACC4JSADAAEFwR0AGAAILgkIAMAAQXBCQAYAAguCWgAwABBcEAA.f_wAAAAAAAAA',
            'didomi_token': 'eyJ1c2VyX2lkIjoiMTk4ZWQzMDMtYTNkYS02ZjY0LWFjZjAtN2JiYzRiN2IwYTA5IiwiY3JlYXRlZCI6IjIwMjUtMDgtMjdUMjA6MjA6MTcuNTk3WiIsInVwZGF0ZWQiOiIyMDI1LTA4LTI3VDIwOjIwOjE4LjkzMFoiLCJ2ZW5kb3JzIjp7ImVuYWJsZWQiOlsiZ29vZ2xlIiwiYzpmYWNlYm9vay1ZeUpSQXllZCIsImM6eW91dHViZSIsImM6YW51bmNpYW50ZV9sYV9saWdhIiwiYzphZG9iZWF1ZGktQVdlN3J3cWQiLCJjOmJlc29jeS1tRlVFYVpCTSJdfSwicHVycG9zZXMiOnsiZW5hYmxlZCI6WyJkZXZpY2VfY2hhcmFjdGVyaXN0aWNzIiwiZ2VvbG9jYXRpb25fZGF0YSIsImRhdGFfc2hhcmluZ193ZWIiLCJkYXRhX3NoYXJpbmciXX0sInZlcnNpb24iOjIsImFjIjoiREU2QVFBRVlBTmdBbFFEekFJY0FpU0NCZ0dKdy5BQUFBIn0=',
            'layout': 'table',
            'league': '1997867',
            'leagueLayout': 'table',
            'locale': 'es',
            'prebidPriorityImp': '{"data":11,"until":1756850400}',
            'satellizer_token': 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOjI4NDg4NzEyLCJpYXQiOjE3NTYzMjYwMzB9.-WcC0Nfra44pYArsQv4uR2rzlgP2eCLZVzTPqaHjBTk',
        }

        if cookies_dict is None:
            cookies_dict = default_cookies
        if local_storage_dict is None:
            local_storage_dict = default_local_storage

        chrome_options = Options()
        #chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        driver = webdriver.Chrome(options=chrome_options)

        try:
            driver.get(url)
            # Set cookies if provided
            for name, value in cookies_dict.items():
                driver.add_cookie({'name': name, 'value': value})
            # Set local storage
            for key, value in local_storage_dict.items():
                driver.execute_script(f"window.localStorage.setItem('{key}', '{value}');")
            driver.refresh()
            time.sleep(3)  # Wait for page to load; adjust as needed

            html = driver.page_source
            lxml_tree=[]
            lxml_tree.append({
                'feed': lxml_html.fromstring(html)
            })
            button_name = '/html/body/app-root/app-nav/nav/linear-tabs/ul/li[2]/a'
            driver.find_element(By.XPATH, button_name).click()
            time.sleep(2)
            html = driver.page_source
            lxml_tree.append({
                'equipos': lxml_html.fromstring(html)
            })
            return lxml_tree
        finally:
            driver.quit()

    @staticmethod
    def access_biwenger(update=False):
        global now_biwenger  # Declarar que usarás la global
        lxml_tree = []
        if update:
            lxml_tree = ConnectBiwenger.get_parsed_html('https://biwenger.as.com/')  # lxml_tree
            now_biwenger = datetime.now()
            print(f"Última actualización de Biwenger a las: {now_biwenger}")
            with open("test.html", "w", encoding="utf-8") as f:
                f.write(lxml_html.tostring(lxml_tree[0]['feed'], pretty_print=True, encoding='unicode'))
            with open("test_Equipos.html", "w", encoding="utf-8") as f:
                f.write(lxml_html.tostring(lxml_tree[1]['equipos'], pretty_print=True, encoding='unicode'))
        else:
            with open("test.html", "r", encoding="utf-8") as f:
                lxml_tree += [{
                    'feed': lxml_html.fromstring(f.read()),
                }]

            with open("test_Equipos.html", "r", encoding="utf-8") as f:
                lxml_tree += [{
                    'equipos': lxml_html.fromstring(f.read()),
                }]
        return lxml_tree

class Excell_Manager:

    @staticmethod
    def insert_dato_excel(sheet_name, col=1, dato='No introducido'):
        file_path = 'Pasta_Biwenger_Portatil_Gris.xlsx'

        wb = load_workbook(file_path)  # Abrir el archivo
        ws = wb[sheet_name]

        ultima_fila = ws.max_row
        while ultima_fila > 0 and ws.cell(row=ultima_fila, column=col).value is None:
            ultima_fila -= 1
        ultima_fila += 1  # Mover a la siguiente fila vacía

        ws.cell(row=ultima_fila, column=col, value=dato)  # Insertar el dato

        celda = ws.cell(row=ultima_fila, column=col)

        if col % 2 == 0:
            # Fuente (negrita, color, tamaño)
            celda.font = Font(name='Liberation Sans', bold=True, size=10)

            # # Fondo (relleno)
            # celda.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

            # Alineación centrada
            celda.alignment = Alignment(horizontal="left", vertical="center")

            # Formato de número (si es necesario)
            celda.number_format = "#,##0"
        else:
            # Fuente (negrita, color, tamaño)
            celda.font = Font(name='Liberation Sans', bold=False, size=10)

            # # Fondo (relleno)
            # celda.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

            # Alineación centrada
            celda.alignment = Alignment(horizontal="left", vertical="center")

        wb.save(file_path)

    @staticmethod
    def check_venta_excel(trade):

        file_path = 'Pasta_Biwenger_Portatil_Gris.xlsx'
        wb = load_workbook(file_path)  # Abrir el archivo
        usuario = None
        if trade.get("vendedor", None):
            usuario = trade['vendedor']
        elif trade.get("comprador", None):
            usuario = trade['comprador']
        ws = wb[usuario]
        date = ws["K2"].value
        print(usuario, trade['jugador'], date)
        fecha_historial = datetime.strptime(str(date), "%Y-%m-%d %H:%M:%S")
        fecha_trade = datetime.strptime(str(trade['fecha']), "%Y-%m-%d %H:%M:%S")
        if fecha_trade >= fecha_historial:
            return True
        else:
            return False

    @staticmethod
    def update_date_to_now(contador):
        file_path = 'Pasta_Biwenger_Portatil_Gris.xlsx'
        wb = load_workbook(file_path)  # Abrir el archivo
        sheets = wb.sheetnames
        now = str(contador)
        now = now.split('.')[0]
        now = datetime.strptime(now, "%Y-%m-%d %H:%M:%S")
        for sheet in sheets:
            ws = wb[sheet]
            ws.cell(row=2, column=11, value=now)  # Insertar el dato

        wb.save(file_path)

    @staticmethod
    def update_date(sheet_name, date):
        file_path = 'Pasta_Biwenger_Portatil_Gris.xlsx'

        wb = load_workbook(file_path)  # Abrir el archivo
        ws = wb[sheet_name]

        ws.cell(row=2, column=11, value=date)  # Insertar el dato

        wb.save(file_path)

    @staticmethod
    def new_transfer_excel(trade):

        print(Excell_Manager.check_venta_excel(trade))
        if not Excell_Manager.check_venta_excel(trade):
            return None
        else:
            if trade.get("vendedor", None):
                print(f"El vendedor de {trade["jugador"]} fue: {trade['vendedor']}")
                Excell_Manager.insert_dato_excel(
                    trade['vendedor'],
                    5,
                    trade['jugador'],
                )
                Excell_Manager.insert_dato_excel(
                    trade['vendedor'],
                    6,
                    trade['precio'],
                )
                Excell_Manager.update_date(trade['vendedor'], trade['fecha'])
            if trade.get("comprador", None):
                print(f"El comprador de {trade["jugador"]} fue: {trade['comprador']}")
                Excell_Manager.insert_dato_excel(
                    trade['comprador'],
                    3,
                    trade['jugador'],
                )
                Excell_Manager.insert_dato_excel(
                    trade['comprador'],
                    4,
                    trade['precio'],
                )
                Excell_Manager.update_date(trade['comprador'], trade['fecha'])
            return None

    @staticmethod
    def update_squadvalue(details):
        file_path = 'Pasta_Biwenger_Portatil_Gris.xlsx'
        sheet_name = details['user']

        wb = load_workbook(file_path)  # Abrir el archivo
        ws = wb[sheet_name]

        ws.cell(row=2, column=8, value=details['value'])  # Insertar el dato

        wb.save(file_path)

class ParseTrades:
    @staticmethod
    def fichajes(elem_post):
        trades = []
        if elem_post.xpath('.//div[contains(concat(" ", @class, " "), " header ")][./user-link]'):
            # Parse the "Fichajes" section when a user-link is present in the Header
            for pc in elem_post.xpath('.//player-card'):

                trade = {
                        'tipo': 'venta',
                        'fecha': parse_date(elem_post.xpath('.//div[@class="date"]/@title')[0]),
                        'vendedor': elem_post.xpath('.//user-link')[0].text_content().strip(),
                        'jugador': pc.xpath('.//div[contains(@class, "main")]')[0].text_content().strip(),
                        'precio': int(re.sub(r'\s+|€|\.', '', pc.xpath('.//*[contains(text(), "€")]')[0].text_content())),
                }
                trades.append(trade)
                Excell_Manager.new_transfer_excel(trade)

            print(f"Las ventas puras al mercado fueron: {len(trades)}")
            for trade in trades:
                print(trade["jugador"])

            return trades

        else:
            # Parse the "Fichajes" section when None in Header
            from_to = elem_post.xpath('.//player-card')[0].xpath('.//div[contains(@class, "from-to")]')[0]
            for pc in elem_post.xpath('.//player-card'):
                try:
                    pc.xpath('.//user-link')[1].text_content().strip()
                except IndexError:
                    trade ={
                            'tipo': 'venta',
                            'fecha': parse_date(elem_post.xpath('.//div[@class="date"]/@title')[0]),
                            'vendedor': pc.xpath('.//user-link')[0].text_content().strip(),
                            'jugador': pc.xpath('.//div[contains(@class, "main")]')[0].text_content().strip(),
                            'precio': int(re.sub(r'\s+|€|\.', '', pc.xpath('.//*[contains(text(), "€")]')[0].text_content())),
                        }
                    trades.append(trade)
                    Excell_Manager.new_transfer_excel(trade)
                    continue
                trade = {
                        'tipo': 'negociación',
                        'fecha': parse_date(elem_post.xpath('.//div[@class="date"]/@title')[0]),
                        'vendedor': pc.xpath('.//user-link')[0].text_content().strip(),
                        'comprador': pc.xpath('.//user-link')[1].text_content().strip(),
                        'jugador': pc.xpath('.//div[contains(@class, "main")]')[0].text_content().strip(),
                        'precio': int(re.sub(r'\s+|€|\.', '', pc.xpath('.//*[contains(text(), "€")]')[0].text_content())),
                    }
                trades.append(trade)
                Excell_Manager.new_transfer_excel(trade)
            return trades

    @staticmethod
    def clausulas(elem_post):
        trades = []
        for pc in elem_post.xpath('.//player-card'):
            trade = {
                'tipo': 'cláusula',
                'fecha': parse_date(elem_post.xpath('.//div[@class="date"]/@title')[0]),
                'vendedor': pc.xpath('.//div[contains(@class, "from-to")]')[0].xpath('.//user-link')[
                    0].text_content().strip(),
                'comprador': pc.xpath('.//div[contains(@class, "from-to")]')[0].xpath('.//user-link')[
                    1].text_content().strip(),
                'jugador': pc.xpath('.//div[contains(@class, "main")]')[0].text_content().strip(),
                'precio': int(re.sub(r'\s+|€|\.', '', pc.xpath('.//*[contains(text(), "€")]')[0].text_content())),
            }
            trades.append(trade)
            Excell_Manager.new_transfer_excel(trade)
        return trades
    @staticmethod
    def mercado_de_fichajes(elem_post):
        trades = []
        for pc in elem_post.xpath('.//player-card'):
            trade = {
                'tipo': 'compra',
                'fecha': parse_date(elem_post.xpath('.//div[@class="date"]/@title')[0]),
                'comprador': pc.xpath('.//user-link')[0].text_content().strip(),
                'jugador': pc.xpath('.//div[contains(@class, "main")]')[0].text_content().strip(),
                'precio': int(re.sub(r'\s+|€|\.', '', pc.xpath('.//*[contains(text(), "€")]')[0].text_content())),
            }
            trades.append(trade)
            Excell_Manager.new_transfer_excel(trade)
        return trades

def parse_date(date_string):
    # Example: '2 sept 2025, 12:29:37'
    months = {
        'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'ago': 8, 'sept': 9, 'oct': 10, 'nov': 11, 'dic': 12
    }
    try:
        date_str, time_str = date_string.split(',', 1)
        day, month_str, year = date_str.split(' ', 2)
        month = months[month_str.lower()]
        dt_str = f"{day.zfill(2)}/{month:02d}/{year.strip()} {time_str.strip()}"
        return datetime.strptime(dt_str, "%d/%m/%Y %H:%M:%S")
    except Exception as e:
        print(e)
        return None

def parse_fichaje(elem_post):
    """
    Parse a single <league-board-post> element to extract transfer information.
    Handles "Fichajes", "Cláusulas", and "Mercado de fichajes" sections.
    1. Fichajes: Can be sales or negotiations.
    2. Cláusulas: Transfers via release clauses.
    3. Mercado de fichajes: Purchases from the market.
    4. If none of these sections are found, returns None.
    """
    if elem_post.xpath('.//h3[normalize-space()="Fichajes"]'):
        return ParseTrades.fichajes(elem_post)

    elif elem_post.xpath('.//h3[normalize-space()="Cláusulas"]'):
        return ParseTrades.clausulas(elem_post)

    elif elem_post.xpath('.//h3[normalize-space()="Mercado de fichajes"]'):
        return ParseTrades.mercado_de_fichajes(elem_post)

    return None

def extract_movements(lxml_tree):
    # Find the next <league-board-post> element that contains an <h3> with text " Fichajes "
    fichajes_posts = lxml_tree.xpath(
        '//league-board-post[.//h3[normalize-space()="Fichajes"]|.//h3[normalize-space()="Cláusulas"]|.//h3[normalize-space()="Mercado de fichajes"]]'
    )
    fichajes_posts = fichajes_posts[::-1]
    return sum([fichaje for fichaje in [parse_fichaje(post) for post in fichajes_posts] if fichaje], [])



def update_squad_value(lxml_tree):
    details_users = lxml_tree.xpath('//div[contains(@class, "table-responsive section-xs light")]//tbody/*') # //td[contains(@class, "text-left user-name")]//a[contains(@role, "button")]
    details=[]
    for details_user in details_users:
        detail = {
            'user': details_user.xpath('.//td[contains(@class, "text-left user-name")]//a[contains(@role, "button")]')[0].text_content().strip(),
            'value': int(re.sub(r'\s+|€|\.|[a-zA-Z]', '', details_user.xpath('.//*[contains(text(), "€")]/@aria-label')[0]))
        }
        Excell_Manager.update_squadvalue(detail)
        details.append(detail)

if __name__ == "__main__":
    now_biwenger = datetime.now()
    update = True
    lxml_tree_list = ConnectBiwenger.access_biwenger(update)
    lxml_tree = lxml_tree_list[0]

    # fichajes_posts = extract_movements(lxml_tree)
    #
    # Excell_Manager.update_date_to_now(now_biwenger)
    #
    # print(f"Found {len(fichajes_posts)} Fichajes:")
    # post_str = ""
    # for post in fichajes_posts:
    #     # print(post)
    #     if post['tipo'] == 'compra':
    #         post_str += f"{post['tipo']},{post['fecha']},,{post['comprador']},{post['jugador']},{post['precio']}\n"
    #     elif post['tipo'] == 'venta':
    #         post_str += f"{post['tipo']},{post['fecha']},{post['vendedor']}, ,{post['jugador']},{post['precio']}\n"
    #     elif post['tipo'] == 'negociación':
    #         post_str += f"{post['tipo']},{post['fecha']},{post['vendedor']}, ,{post['jugador']},{post['precio']}\n"
    #     elif post['tipo'] == 'cláusula':
    #         post_str += f"{post['tipo']},{post['fecha']},{post['vendedor']},{post['comprador']},{post['jugador']},{post['precio']}\n"
    # print(post_str)

    update_squad_value(lxml_tree_list[1]['equipos'])