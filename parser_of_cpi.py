import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import re
import urllib3
import ssl
import certifi
import pyodbc
import pandas as pd
import warnings
import win32com.client as win32
import os
import subprocess
from ollama import Client
import os
import logging

warnings.filterwarnings("ignore", category=UserWarning)


def create_session():
    """Создает сессию с отключенной проверкой SSL."""
    session = requests.Session()

    # Решение 1: Отключаем проверку SSL (самое простое)
    session.verify = False

    # Решение 2: Альтернатива - используем certifi (раскомментируйте при необходимости)
    # session.verify = certifi.where()

    # Отключаем предупреждения SSL
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    return session


def get_current_month_end():
    """Возвращает название последнего загруженного месяца."""

    cnxn = pyodbc.connect(
        "driver={SQL Server};"
        "server=AMDBRATES\DBMVTBAM01;"
        "database=riski;"
        "trusted_connection=yes"
    )
    last_date = (
        pd.read_sql(
            "select top 1 date from riski.[dbo].[RR_INDEX_HISTORY] where index_id = 'RUCPI' order by date desc",
            cnxn,
        )
        .date[0]
        .month
    )
    cnxn.close()

    month_names = {
        1: "январ",
        2: "феврал",
        3: "март",
        4: "апрел",
        5: "ма",
        6: "июн",
        7: "июл",
        8: "август",
        9: "сентябр",
        10: "октябр",
        11: "ноябр",
        12: "декабр",
    }
    return month_names[last_date + 1]


def get_current_month_date():
    cnxn = pyodbc.connect(
        "driver={SQL Server};"
        "server=AMDBRATES\DBMVTBAM01;"
        "database=riski;"
        "trusted_connection=yes"
    )
    last_date = pd.read_sql(
        "select top 1 date from riski.[dbo].[RR_INDEX_HISTORY] where index_id = 'RUCPI' order by date desc",
        cnxn,
    ).date[0]
    cnxn.close()

    cnxn = pyodbc.connect(
        "driver={SQL Server};"
        "server=AMDBRATES\DBMVTBAM01;"
        "database=riski;"
        "trusted_connection=yes"
    )
    dates = pd.read_sql(
        "select CONVERT(varchar,tradedate,104) from rates.[dbo].[MOEX_Index] where SECID = 'RUCBTRNS' order by tradedate",
        cnxn,
    )
    dates[""] = pd.to_datetime(dates[""], format="%d.%m.%Y")
    cnxn.close()

    last_date = last_date.to_pydatetime()
    next_month = last_date.replace(day=28) + timedelta(days=34)
    # new_date = next_month.replace(day=1) - timedelta(days=1)
    new_date = pd.to_datetime("2026-08-31")

    while pd.to_datetime(new_date) not in dates.values:
        new_date -= timedelta(days=1)

    return new_date.strftime("%Y-%m-%d %H:%M:%S")


def parse_page(session, url):
    """Парсит страницу новостей."""
    try:
        response = session.get(
            url,
            timeout=10,
            proxies={"http": "http://127.0.0.1:3128", "https": "http://127.0.0.1:3128"},
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        news_items = []
        # Адаптированные селекторы для Росстата
        items = soup.find_all(
            ["div", "article"], class_=re.compile(r"news|item|post|entry")
        )

        for item in items:
            title_elem = item.find(["h1", "h2", "h3", "a", ".title"])
            link_elem = item.find("a", href=True)

            if title_elem and link_elem:
                title = title_elem.get_text(strip=True)
                link = link_elem["href"]
                if link.startswith("/"):
                    link = "https://rosstat.gov.ru" + link
                elif not link.startswith("http"):
                    link = "https://rosstat.gov.ru/" + link

                news_items.append((title, link))
        return news_items
    except Exception as e:
        print(f"Ошибка парсинга {url}: {e}")
        return []


# def ai_parser(text):
# ОТКЛЮЧАЕМ прокси для localhost + логи
#    os.environ['NO_PROXY'] = 'localhost,127.0.0.1,::1'
#    os.environ['HTTP_PROXY'] = ''
#    os.environ['HTTPS_PROXY'] = ''
#    logging.getLogger("ollama").setLevel(logging.CRITICAL)

#    client = Client(host='http://127.0.0.1:11434')  # ТОЧНО 127.0.0.1!

# Тест
#    response = client.chat(
#        model='qwen25-7b-smart',
#        messages=[{'role': 'user', 'content': 'В следующем тексте найди значение ИПЦ, которое появляется первым, и верни ее значение в виде одного числа с десятичным разделителем в виде точки (кроме значения больше ничего не пиши): ' + text}]
#    )
#    return float(response['message']['content'].replace(",", "."))


def extract_cpi_from_page(session, link, month_end):
    """Извлекает данные ИПЦ из публикации."""
    try:
        response = session.get(
            link,
            timeout=10,
            proxies={"http": "http://127.0.0.1:3128", "https": "http://127.0.0.1:3128"},
        )
        response.raise_for_status()

        response.encoding = "utf-8"
        soup = BeautifulSoup(response.text, "html.parser")

        text = soup.get_text(separator=" ", strip=True)

        # cpi_value = ai_parser(text[100:300])

        # Паттерны для поиска ИПЦ
        pattern = r"составил\s*(\d+(?:,\d+)?)\s*%"

        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            cpi_value = float(match.group(1).replace(",", "."))
            print(f"✅ Найдено: 'составил {cpi_value}%'")
            return cpi_value

        print(f"✅ Найдено: 'составил {cpi_value}%'")
        return cpi_value

    except Exception:
        return None


def send_outlook_alert(cpi_value, month_end, source_url):
    """Отправка через локальный Outlook (без SMTP)."""
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)  # olMailItem = 0

    mail.Subject = f"RUCPI: {month_end} 2026 = {cpi_value}%"
    mail.To = "Alexander.Miloserdov@wiminvest.ru"
    mail.Body = f"""
Новые данные ИПЦ от Росстата:

Дата: {month_end}
Значение: {cpi_value}%
Источник: {source_url}

Автоматическое уведомление.
                """
    mail.Send()  # Отправляет сразу
    print("✅ Отправлено через Outlook!")


def main():
    session = create_session()
    base_url = "https://rosstat.gov.ru/central-news"
    pages = [f"{base_url}?page={i}" for i in range(1, 6)]
    month_end = get_current_month_end()

    print(f"Ищем ИПЦ за конец {month_end}...")

    for page_url in pages:
        print(f"Сканируем {page_url}...")
        news = parse_page(session, page_url)

        for title, link in news:
            title_lower = title.lower()
            if (
                "об индексе потребительских цен в" in title_lower
                and month_end.lower() in title_lower
            ):

                print(f"Найдена релевантная новость: {title[:80]}...")
                print(link)
                cpi = extract_cpi_from_page(session, link, month_end)

                if cpi:
                    print(f"✅ ИПЦ за конец {month_end}: {cpi}%", flush=True)
                    print(f"Источник: {link}")

                    date_BD = get_current_month_date()

                    # 1. Подключение к БД
                    cnxn = pyodbc.connect(
                        "driver={SQL Server};"
                        "server=AMDBRATES\DBMVTBAM01;"
                        "database=riski;"
                        "trusted_connection=yes"
                    )
                    cursor = cnxn.cursor()

                    # 2. INSERT одной записи
                    cursor.execute(
                        f"INSERT INTO riski.[dbo].[RR_INDEX_HISTORY] (date, index_id, px_last) VALUES (CONVERT(datetime,'{date_BD}',120), 'RUCPI', {cpi})"
                    )

                    # 3. COMMIT (ОБЯЗАТЕЛЬНО!)
                    cnxn.commit()

                    # 4. Закрытие
                    cursor.close()
                    cnxn.close()

                    send_outlook_alert(cpi, date_BD, link)

                    return
                else:
                    print(f"Данные ИПЦ не найдены в статье")

    print("❌ Данные по ИПЦ за конец текущего месяца не найдены")


if __name__ == "__main__":
    main()
