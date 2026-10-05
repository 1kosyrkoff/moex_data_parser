# -*- coding: utf-8 -*-
"""Облигации с MOEX ISS по списку ISIN.

  get_bond_yields    — YTM и YTP (к put-оферте) на дату оценки
  get_coupon_schedule — полное расписание выплат: купоны, амортизации, оферты, погашение
  get_bond_card      — все, что ISS знает о выпуске, словарем читаемых таблиц
  get_issuer_sector  — сектор экономики эмитента (по отраслевым индексам MOEX)

Схема расчета доходностей:
  ISIN -> SECID -> цена закрытия на дату оценки (история ISS) -> купонное расписание
  (bondization) -> денежные потоки -> эффективная доходность (ACT/365) методом бисекции.

Конвенции (подобраны сверкой с биржевыми YIELDCLOSE и YIELDTOOFFER):
  * цена — CLOSE (последняя сделка) дня оценки. Если сделок не было, откатываемся
    назад на LOOKBACK дней, а если выпуск не торговался и там — берем его первую
    цену на рынке; фактическая дата цены видна в колонке PRICE_DATE;
  * дисконтируем от даты расчетов T+1 — следующего ТОРГОВОГО дня после даты оценки.
    Биржа считает доходность именно так: на пятничной оценке расчеты уходят на
    понедельник, и без этого сдвига короткие выпуски промахиваются на проценты
    годовых (медиана расхождения падает с 0.43 до 0.002 п.п.). НКД биржа при этом
    публикует на дату торгов, поэтому поле ACCINT с расчетом не сверяется;
  * НКД и непогашенный номинал считаем сами из расписания: так они корректны на
    любую календарную дату и выражены в валюте выпуска (у валютных бумаг ISS отдает
    НКД в рублях, а цену и номинал — в валюте номинала). Масштаб амортизаций при
    этом сверяется с биржевым FACEVALUE: в расписании ISS суммы бывают удвоены;
  * неизвестные будущие купоны (флоатеры, бумаги с купоном только до оферты)
    считаем по последней объявленной ставке — колонка COUPONS помечает такие
    выпуски как 'unknown coupons'. Если не объявлено ни одного купона, переносить
    нечего и доходность не считается вовсе;
  * YTP — только к put-оферте. Оферту по решению эмитента (call) не считаем: тип
    оферты в ISS их не различает, поэтому call отсекаем по CALLOPTIONDATE.
"""
import datetime as dt
from typing import Iterable, Optional, Union

import requests
import pandas as pd

ISS = "https://iss.moex.com/iss"
YEAR = 365.0        # база ACT/365, как в эффективной доходности MOEX
LOOKBACK = 20       # на сколько дней назад искать цену, если в дату оценки сделок не было

# какое поле истории считать ценой; берем первое непустое из цепочки
PRICE_CHAINS = {
    "close": ["CLOSE"],
    "legal": ["LEGALCLOSEPRICE", "CLOSE"],
    "mp3":   ["MARKETPRICE3", "LEGALCLOSEPRICE", "CLOSE"],
    "wap":   ["WAPRICE", "CLOSE"],
}

KNOWN, UNKNOWN = "known all coupons", "unknown coupons"

COLUMNS = ["ISIN", "YTM", "YTP", "COUPONS", "SECID", "SHORTNAME", "DATE", "SETTLE",
           "PRICE", "PRICE_DATE", "PRICE_FIELD", "ACCRUED", "FACEVALUE", "FACEUNIT",
           "MATDATE", "PUTDATE", "EST_COUPONS", "STATUS"]

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (research; MOEX ISS bonds client)"})

_CACHE = {"market": None, "bondization": {}, "static": {}, "columns": {},
          "sectors": None, "emitters": {}}


def _iss(path: str, **params) -> dict:
    """GET к ISS с отключенной метаинформацией -> json."""
    params.setdefault("iss.meta", "off")
    r = S.get(f"{ISS}/{path}", params=params, timeout=60)
    r.raise_for_status()
    return r.json()


def _block(payload: dict, name: str) -> pd.DataFrame:
    """Блок ответа ISS ({'columns': [...], 'data': [...]}) -> DataFrame."""
    b = payload.get(name) or {"columns": [], "data": []}
    return pd.DataFrame(b["data"], columns=b["columns"])


def _date(value) -> Optional[dt.date]:
    """'YYYY-MM-DD' -> date. Пустые значения и '0000-00-00' (бессрочные) -> None."""
    if value is None or value == "" or value == "0000-00-00":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if pd.isna(value):
        return None
    stamp = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(stamp) else stamp.date()


def _num(value) -> Optional[float]:
    """Число ISS -> float. None, пустая строка и NaN -> None."""
    if value is None or value == "":
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_date(value, name="date") -> dt.date:
    """Приводит дату к date: принимает date/datetime, 'YYYY-MM-DD' или 'DD.MM.YYYY'."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y%m%d", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Не понимаю {name}={value!r}, ожидается 'YYYY-MM-DD' или 'DD.MM.YYYY'")


# ---------------------------------------------------------------- справочники

MARKET_COLUMNS = ("SECID,ISIN,SHORTNAME,BOARDID,MATDATE,OFFERDATE,PUTOPTIONDATE,"
                  "CALLOPTIONDATE,BUYBACKDATE,BUYBACKPRICE,FACEVALUE,FACEUNIT,"
                  "COUPONPERCENT,BONDTYPE,BONDSUBTYPE")


def market_snapshot() -> pd.DataFrame:
    """Справочник всех торгуемых облигаций одним запросом (кешируется на сессию)."""
    if _CACHE["market"] is None:
        payload = _iss("engines/stock/markets/bonds/securities.json",
                       **{"iss.only": "securities", "securities.columns": MARKET_COLUMNS})
        df = _block(payload, "securities")
        _CACHE["market"] = df.drop_duplicates("SECID") if not df.empty else df
    return _CACHE["market"]


def find_bond(isin: str) -> Optional[dict]:
    """ISIN -> статические параметры выпуска. У ОФЗ SECID != ISIN, поэтому ищем по базе.

    Сначала в справочнике торгуемых бумаг (один общий запрос на все ISIN), затем —
    поиском по всей базе ISS: так находятся погашенные и делистингованные выпуски.
    """
    isin = isin.strip().upper()
    market = market_snapshot()
    if not market.empty:
        hit = market[market["ISIN"].fillna("").str.upper() == isin]
        if not hit.empty:
            row = hit.iloc[0].to_dict()
            row["ISIN"] = isin
            return row

    if isin in _CACHE["static"]:
        return _CACHE["static"][isin]
    found = _block(_iss("securities.json", q=isin, **{
        "iss.only": "securities",
        "securities.columns": "secid,isin,shortname,primary_boardid,is_traded"}), "securities")
    if found.empty:
        return None
    found = found[found["isin"].fillna("").str.upper() == isin]
    if found.empty:
        return None
    secid = found.sort_values("is_traded", ascending=False).iloc[0]["secid"]

    desc = _block(_iss(f"securities/{secid}.json", **{"iss.only": "description"}), "description")
    fields = dict(zip(desc["name"], desc["value"])) if not desc.empty else {}
    row = {"SECID": secid, "ISIN": isin, "SHORTNAME": fields.get("SHORTNAME"),
           "BOARDID": None, "MATDATE": fields.get("MATDATE"), "OFFERDATE": None,
           "PUTOPTIONDATE": None, "CALLOPTIONDATE": None, "BUYBACKDATE": None,
           "BUYBACKPRICE": None, "FACEVALUE": _num(fields.get("FACEVALUE")),
           "FACEUNIT": fields.get("FACEUNIT"), "COUPONPERCENT": _num(fields.get("COUPONPERCENT")),
           "BONDTYPE": fields.get("BOND_TYPE"), "BONDSUBTYPE": fields.get("BOND_SUBTYPE")}
    _CACHE["static"][isin] = row
    return row


def get_bondization(secid: str) -> dict:
    """Расписание выпуска -> {'coupons', 'amortizations', 'offers'} как DataFrame (кеш)."""
    if secid not in _CACHE["bondization"]:
        payload = _iss(f"securities/{secid}/bondization.json", limit="unlimited")
        _CACHE["bondization"][secid] = {name: _block(payload, name)
                                        for name in ("coupons", "amortizations", "offers")}
    return _CACHE["bondization"][secid]


# ---------------------------------------------------------------------- цена

HISTORY_COLUMNS = ("TRADEDATE,BOARDID,CLOSE,LEGALCLOSEPRICE,MARKETPRICE3,WAPRICE,"
                   "ACCINT,FACEVALUE,BUYBACKDATE,NUMTRADES")


def _history(secid: str, start: dt.date, till: dt.date,
             board: str = None) -> pd.DataFrame:
    """Дневная история торгов выпуска за период -> DataFrame (может быть пустым).

    Без указания доски ISS отдает строки всех режимов сразу (у валютных выпусков это
    TQCB и TQOD с разными ценами), поэтому спрашиваем историю основной доски.
    """
    path = (f"history/engines/stock/markets/bonds/boards/{board}/securities/{secid}.json"
            if board else f"history/engines/stock/markets/bonds/securities/{secid}.json")
    payload = _iss(path, **{"iss.only": "history", "history.columns": HISTORY_COLUMNS,
                            "from": start.isoformat(), "till": till.isoformat()})
    return _block(payload, "history")


def _first_traded_row(secid: str, till: dt.date, board: str = None) -> Optional[dict]:
    """Первая рыночная цена выпуска — запасной вариант, когда сделок не было совсем."""
    dates = _block(_iss(f"history/engines/stock/markets/bonds/securities/{secid}/dates.json"),
                   "dates")
    start = _date(dates.iloc[0]["from"]) if not dates.empty else None
    if start is None or start > till:
        return None
    window = _history(secid, start, min(start + dt.timedelta(days=90), till), board)
    if window.empty:
        return None
    traded = window[window["CLOSE"].notna()]
    return None if traded.empty else traded.iloc[0].to_dict()


def get_price(secid: str, date: dt.date, price: Union[str, float] = "close",
              lookback: int = LOOKBACK, board: str = None) -> dict:
    """Цена на дату оценки -> {'price', 'date', 'field', 'face', 'sessions'}.

    Ищем цену в дату оценки, при отсутствии сделок откатываемся назад на lookback дней,
    а если выпуск не торговался и там — берем его первую цену на рынке. Заодно окно
    захватывает несколько дней вперед: из него берется биржевой календарь для T+1,
    а из найденной строки — номинал, который биржа показывает на эту дату.
    """
    found = {"price": None, "date": None, "field": None, "face": None, "sessions": []}
    if isinstance(price, (int, float)):
        return dict(found, price=float(price), date=date, field="задана вручную")
    fields = PRICE_CHAINS.get(str(price).lower(), PRICE_CHAINS["close"])

    window = _history(secid, date - dt.timedelta(days=lookback),
                      date + dt.timedelta(days=10), board)
    if not window.empty:
        window = window.sort_values("TRADEDATE")
        found["sessions"] = [d for d in (_date(t) for t in window["TRADEDATE"]) if d]
        traded = window[[(_date(t) or date) <= date for t in window["TRADEDATE"]]]
        for _, row in traded.iloc[::-1].iterrows():          # от даты оценки назад
            for field in fields:
                value = _num(row.get(field))
                if value:
                    return dict(found, price=value, date=_date(row["TRADEDATE"]),
                                field=field, face=_num(row.get("FACEVALUE")))

    first = _first_traded_row(secid, date, board)
    if first:
        return dict(found, price=_num(first["CLOSE"]), date=_date(first["TRADEDATE"]),
                    field="первая цена на рынке", face=_num(first.get("FACEVALUE")))
    return found


def next_trading_day(date: dt.date, sessions: Iterable[dt.date] = ()) -> dt.date:
    """Дата расчетов T+1: следующий торговый день после даты оценки.

    Биржа считает доходность к дате расчетов, а не к дате торгов: на пятничной оценке
    расчеты уходят на понедельник, и без этого сдвига короткие выпуски промахиваются
    на проценты годовых. Календарь берем из истории торгов выпуска, а когда будущих
    сессий в ней еще нет (оценка на сегодня) — откатываемся на ближайший будний день.
    """
    future = sorted(d for d in sessions if d > date)
    if future:
        return future[0]
    nxt = date + dt.timedelta(days=1)
    while nxt.weekday() >= 5:                        # суббота и воскресенье
        nxt += dt.timedelta(days=1)
    return nxt


# ------------------------------------------------------- расписание и потоки

def _amort_pairs(amortizations: pd.DataFrame) -> list:
    """Погашения номинала -> [(дата, сумма), ...]."""
    out = []
    for _, r in amortizations.iterrows():
        when, value = _date(r.get("amortdate")), _num(r.get("value"))
        if when and value:
            out.append((when, value))
    return sorted(out)


def normalize_amortizations(amortizations: pd.DataFrame, face: Optional[float],
                            date: dt.date) -> pd.DataFrame:
    """Приводит суммы амортизаций к номиналу, который биржа показывает на дату.

    В расписании ISS суммы бывают несогласованы с номиналом выпуска: у замещающих ОФЗ
    выплаты продублированы (сумма valueprc = 200% вместо 100%), у бумаг с копеечным
    номиналом — огрублены округлением. Форме расписания доверяем, масштаб берем с рынка.
    """
    scheduled = face_at(amortizations, date)
    if not face or not scheduled:
        return amortizations
    scale = face / scheduled
    if abs(scale - 1) < 0.005:
        return amortizations
    scaled = amortizations.copy()
    scaled["value"] = pd.to_numeric(scaled["value"], errors="coerce") * scale
    return scaled


def face_at(amortizations: pd.DataFrame, date: dt.date,
            fallback: Optional[float] = None) -> float:
    """Непогашенный номинал на дату = сумма всех амортизаций после этой даты.

    fallback подставляется только при ПУСТОМ расписании — так бессрочные выпуски не
    выглядят погашенными. Ноль по непустому расписанию законен: у быстро амортизируемых
    бумаг номинал гасится задолго до даты погашения, и купоны после этого равны нулю.
    """
    pairs = _amort_pairs(amortizations)
    if not pairs:
        return fallback or 0.0
    return sum(v for d, v in pairs if d > date)


def value_decimals(coupons: pd.DataFrame) -> int:
    """Сколько знаков после запятой ISS использует в суммах купонов этого выпуска.

    У рублевых бумаг это копейки, у замещающих ОФЗ купон бывает 0.0015 — округление
    до двух знаков обнулило бы и купон, и НКД.
    """
    decimals = 2
    for value in coupons.get("value", []):
        value = _num(value)
        if value is None:
            continue
        text = f"{value:.10f}".rstrip("0")
        decimals = max(decimals, len(text.partition(".")[2]))
    return min(decimals, 8)


def _period_days(row) -> Optional[int]:
    """Длина купонного периода в днях."""
    start, end = _date(row.get("startdate")), _date(row.get("coupondate"))
    if start is None or end is None:
        return None
    days = (end - start).days
    return days if days > 0 else None


def last_known_rate(coupons: pd.DataFrame, amortizations: pd.DataFrame) -> Optional[float]:
    """Последняя известная купонная ставка, % годовых.

    Берем самый свежий объявленный купон, включая будущие: после оферты эмитент часто
    переставляет ставку, и уже объявленный купон — более свежая информация, чем последний
    выплаченный. Если ISS не заполнил valueprc (у флоатеров он почти всегда пустой),
    восстанавливаем ставку из суммы купона: rate = value / (номинал * дней / 365).
    """
    known = []
    for _, r in coupons.iterrows():
        when, value = _date(r.get("coupondate")), _num(r.get("value"))
        if when is None or value is None:
            continue
        known.append((when, r, value))
    if not known:
        return None
    known.sort(key=lambda x: x[0])
    when, row, value = known[-1]

    rate = _num(row.get("valueprc"))
    if rate:
        return rate
    days = _period_days(row)
    start = _date(row.get("startdate"))
    face = face_at(amortizations, start) if start else None
    if not days or not face:
        return None
    return value / face * YEAR / days * 100


def accrued_interest(coupons: pd.DataFrame, amortizations: pd.DataFrame, date: dt.date,
                     rate: Optional[float], decimals: int = 2,
                     face: Optional[float] = None) -> float:
    """НКД на дату оценки: часть текущего купона, накопленная с начала периода.

    Считаем сами, а не берем ACCRUEDINT/ACCINT из ISS: биржевое поле привязано к дате
    торгов (или к T+1 в онлайне), а у валютных выпусков еще и выражено в рублях.
    """
    for _, r in coupons.iterrows():
        start, end = _date(r.get("startdate")), _date(r.get("coupondate"))
        if start is None or end is None or not (start <= date < end):
            continue
        value = _num(r.get("value"))
        if value is None:
            days = _period_days(r)
            outstanding = face_at(amortizations, start, face)
            if not rate or not days or not outstanding:
                return 0.0
            value = round(outstanding * rate / 100 * days / YEAR, decimals)
        return round(value * (date - start).days / (end - start).days, decimals)
    return 0.0


def bond_cashflows(bondization: dict, date: dt.date, till: dt.date,
                   rate: Optional[float] = None, redemption_price: float = 100.0,
                   decimals: int = 2, face: Optional[float] = None) -> pd.DataFrame:
    """Денежные потоки на одну облигацию после даты оценки -> DataFrame(date, amount, kind).

    till — дата выбытия бумаги (погашение или put-оферта): потоки после нее отбрасываются,
    а непогашенный к этой дате номинал возвращается по redemption_price (% от номинала).
    Купоны считаются от непогашенного номинала, поэтому амортизация уменьшает их сама собой.

    НКД к выкупу НЕ добавляется: биржа считает доходность к оферте по цене выкупа без
    накопленного купона, и проверка это подтверждает — с добавлением НКД расходятся
    даже те выпуски, что иначе совпадают с YIELDTOOFFER до четвертого знака.
    """
    amortizations = bondization["amortizations"]
    rows, estimated, lost = [], 0, 0

    for _, r in amortizations.iterrows():
        when, value = _date(r.get("amortdate")), _num(r.get("value"))
        if when is None or value is None or when <= date or when >= till or value == 0:
            continue
        rows.append({"date": when, "amount": value, "kind": "amortization"})

    for _, r in bondization["coupons"].iterrows():
        when = _date(r.get("coupondate"))
        if when is None or when <= date or when > till:
            continue
        value = _num(r.get("value"))
        if value is None:                                   # купон еще не объявлен
            start, days = _date(r.get("startdate")), _period_days(r)
            outstanding = face_at(amortizations, start, face) if start else None
            if outstanding == 0:
                estimated += 1      # номинал уже погашен: купон нулевой, поток пуст
                continue
            if not rate or not days or outstanding is None:
                lost += 1           # ставку перенести не с чего
                continue
            value = round(outstanding * rate / 100 * days / YEAR, decimals)
            estimated += 1
        rows.append({"date": when, "amount": value, "kind": "coupon"})

    pairs = _amort_pairs(amortizations)
    residual = round(sum(v for d, v in pairs if d >= till) if pairs else (face or 0.0), 6)
    if residual > 0:
        rows.append({"date": till, "amount": residual * redemption_price / 100.0,
                     "kind": "redemption"})

    df = pd.DataFrame(rows, columns=["date", "amount", "kind"])
    if not df.empty:
        df = df.sort_values("date").reset_index(drop=True)
    df.attrs["estimated"], df.attrs["lost"] = estimated, lost
    return df


def effective_yield(cashflows: pd.DataFrame, dirty_price: float, date: dt.date,
                    lo: float = -0.999, hi: float = 1000.0, tol: float = 1e-12) -> Optional[float]:
    """Эффективная доходность, % годовых: NPV(потоки) = грязная цена, ACT/365, бисекция."""
    if cashflows is None or cashflows.empty or not dirty_price or dirty_price <= 0:
        return None
    times = [(d - date).days / YEAR for d in cashflows["date"]]
    amounts = list(cashflows["amount"])
    if min(times) <= 0:
        return None

    def npv(y):
        return sum(a / (1.0 + y) ** t for a, t in zip(amounts, times)) - dirty_price

    if npv(lo) * npv(hi) > 0:      # корня на отрезке нет: цена вне диапазона потоков
        return None
    for _ in range(300):
        mid = (lo + hi) / 2
        if npv(lo) * npv(mid) <= 0:
            hi = mid
        else:
            lo = mid
        if hi - lo < tol:
            break
    return round((lo + hi) / 2 * 100, 4)


def put_date(info: dict, offers: pd.DataFrame, date: dt.date, maturity: Optional[dt.date],
             buyback: Optional[dt.date] = None) -> tuple:
    """Ближайшая put-оферта после даты оценки -> (дата, цена выкупа в % от номинала).

    offertype в ISS не отличает put от call ('Оферта' стоит у обоих), поэтому call
    отсекаем по CALLOPTIONDATE выпуска и по типу 'Оферта/Погашение' — им помечена
    оферта, совпадающая с выкупом по решению эмитента или с погашением.
    """
    call = _date(info.get("CALLOPTIONDATE"))
    price = _num(info.get("BUYBACKPRICE")) or 100.0

    def valid(when):
        return (when and when > date and when != call
                and (maturity is None or when < maturity))

    candidates = []
    for _, r in offers.iterrows():
        when = _date(r.get("offerdate"))
        if not valid(when) or "погашение" in str(r.get("offertype") or "").lower():
            continue
        candidates.append((when, _num(r.get("price")) or price))
    for when in (buyback, _date(info.get("BUYBACKDATE")), _date(info.get("PUTOPTIONDATE"))):
        if valid(when) and when not in [c[0] for c in candidates]:
            candidates.append((when, price))
    if not candidates:
        return None, price

    nearest = min(c[0] for c in candidates)
    cluster = [c for c in candidates if (c[0] - nearest).days <= 5]
    return max(cluster, key=lambda x: x[0])          # дата расчетов, а не сбора заявок


# --------------------------------------------------------------------- API

def get_bond_yields(isins: Iterable[str], date=None, price: Union[str, float, dict] = "close",
                    lookback: int = LOOKBACK, assumed_rate: Optional[float] = None,
                    details: bool = False) -> pd.DataFrame:
    """YTM и YTP по списку ISIN на дату оценки -> DataFrame.

    isins        — список ISIN, например ['RU000A0ZZ5H3', 'RU000A1032P1', 'RU000A1038V6']
    date         — дата оценки; по умолчанию сегодня
    price        — 'close' (по умолчанию), 'legal', 'mp3', 'wap', число (% от номинала)
                   или словарь {ISIN: цена}
    lookback     — на сколько дней откатываться назад, если в дату оценки сделок не было
    assumed_rate — ставка (% годовых) для еще не объявленных купонов; по умолчанию
                   используется последняя известная ставка самого выпуска
    details      — True добавляет цену, НКД, номинал, даты и статус расчета

    Колонки: ISIN, YTM, YTP, COUPONS. YTM — эффективная доходность к погашению,
    YTP — к ближайшей put-оферте (NaN, если оферты нет). Обе в процентах годовых,
    ACT/365. COUPONS = 'known all coupons', если все купоны до погашения объявлены
    биржей, и 'unknown coupons', если часть из них посчитана по перенесенной ставке.
    """
    if isinstance(isins, str):
        isins = [isins]
    date = parse_date(date, "date") if date is not None else dt.date.today()
    prices = price if isinstance(price, dict) else {}
    rows = []

    for isin in isins:
        isin = str(isin).strip().upper()
        rec = {"ISIN": isin, "YTM": None, "YTP": None, "COUPONS": None,
               "SECID": None, "SHORTNAME": None, "DATE": date, "SETTLE": None, "PRICE": None,
               "PRICE_DATE": None, "PRICE_FIELD": None, "ACCRUED": None, "FACEVALUE": None,
               "FACEUNIT": None, "MATDATE": None, "PUTDATE": None, "EST_COUPONS": 0,
               "STATUS": "ok"}
        try:
            info = find_bond(isin)
            if not info:
                rec["STATUS"] = "не найден на MOEX"
                rows.append(rec)
                continue
            secid = info["SECID"]
            rec.update({"SECID": secid, "SHORTNAME": info.get("SHORTNAME"),
                        "FACEUNIT": info.get("FACEUNIT")})

            maturity = _date(info.get("MATDATE"))
            if maturity is not None and maturity <= date:
                rec["STATUS"] = "выпуск погашен на дату оценки"
                rows.append(rec)
                continue
            rec["MATDATE"] = maturity

            bondization = get_bondization(secid)
            coupons, amortizations = bondization["coupons"], bondization["amortizations"]
            if coupons.empty and amortizations.empty:
                rec["STATUS"] = "нет купонного расписания в ISS"
                rows.append(rec)
                continue

            quote = get_price(secid, date, prices.get(isin, price), lookback,
                              info.get("BOARDID"))
            clean = quote["price"]
            rec.update({"PRICE": clean, "PRICE_DATE": quote["date"],
                        "PRICE_FIELD": quote["field"]})
            if clean is None:
                rec["STATUS"] = f"нет цены за {lookback} дней до даты оценки"
                rows.append(rec)
                continue

            settle = next_trading_day(date, quote["sessions"])
            rec["SETTLE"] = settle
            if maturity is not None and maturity <= settle:
                rec["STATUS"] = "погашение до даты расчетов T+1"
                rows.append(rec)
                continue

            amortizations = normalize_amortizations(amortizations, quote["face"],
                                                    quote["date"] or settle)
            bondization = dict(bondization, amortizations=amortizations)
            rate = assumed_rate or last_known_rate(coupons, amortizations)
            decimals = value_decimals(coupons)
            reported = quote["face"] or _num(info.get("FACEVALUE"))
            face = face_at(amortizations, settle, reported)
            nkd = accrued_interest(coupons, amortizations, settle, rate, decimals, face)
            rec.update({"FACEVALUE": face, "ACCRUED": nkd})
            if not face:
                rec["STATUS"] = "номинал полностью погашен на дату расчетов"
                rows.append(rec)
                continue
            dirty = clean * face / 100.0 + nkd

            if maturity is not None:
                flows = bond_cashflows(bondization, settle, maturity, rate,
                                       decimals=decimals, face=face)
                rec["EST_COUPONS"] = flows.attrs["estimated"]
                rec["COUPONS"] = UNKNOWN if (flows.attrs["estimated"] or
                                             flows.attrs["lost"]) else KNOWN
                if flows.attrs["lost"]:
                    # переносить нечего: ни одного объявленного купона у выпуска нет
                    rec["STATUS"] = (f"ставка купона неизвестна, {flows.attrs['lost']} "
                                     f"купонов не восстановлены")
                else:
                    rec["YTM"] = effective_yield(flows, dirty, settle)
            else:
                rec["COUPONS"] = UNKNOWN
                rec["STATUS"] = "бессрочная бумага: доходность к погашению не определена"

            put, put_price = put_date(info, bondization["offers"], settle, maturity)
            rec["PUTDATE"] = put
            if put:
                put_flows = bond_cashflows(bondization, settle, put, rate,
                                           redemption_price=put_price,
                                           decimals=decimals, face=face)
                if not put_flows.attrs["lost"]:
                    rec["YTP"] = effective_yield(put_flows, dirty, settle)
        except requests.RequestException as e:
            rec["STATUS"] = f"ошибка запроса: {e}"
        except Exception as e:                  # noqa: BLE001 — одна битая бумага не роняет пачку
            rec["STATUS"] = f"ошибка расчета: {e}"
        rows.append(rec)

    df = pd.DataFrame(rows, columns=COLUMNS)          # columns держит порядок и пустой список
    return df if details else df[COLUMNS[:4]]


# ------------------------------------------------------- купонное расписание

def _total(*values) -> Optional[float]:
    """Сумма непустых слагаемых; все пустые -> None, а не 0: выплаты в этот день нет."""
    known = [v for v in values if v is not None]
    return round(sum(known), 8) if known else None


SCHEDULE_COLUMNS = ["ISIN", "SECID", "SHORTNAME", "date", "coupon_amount", "rate",
                    "amort_amount", "total_amount", "FACEVALUE", "FACEUNIT", "STARTDATE",
                    "RECORDDATE", "OFFERPRICE", "offer_amount", "OFFERSTART", "OFFEREND",
                    "CALLOPTIONDATE", "PUTOPTIONDATE", "STATUS"]

SCHEDULE_DATES = ["date", "STARTDATE", "RECORDDATE", "OFFERSTART", "OFFEREND",
                  "CALLOPTIONDATE", "PUTOPTIONDATE"]

SCHEDULE_NUMBERS = ["coupon_amount", "rate", "amort_amount", "total_amount", "FACEVALUE",
                    "OFFERPRICE", "offer_amount"]


def get_coupon_schedule(isins: Iterable[str]) -> pd.DataFrame:
    """Полное купонное расписание выпусков с MOEX ISS -> DataFrame (одна строка = дата).

    isins — список ISIN или один ISIN строкой.

    Отдается все, что ISS знает по выпуску, включая уже прошедшие выплаты. Дата в пределах
    выпуска встречается ровно один раз: купон, транш амортизации и оферта одного дня лежат
    в одной строке, каждый в своей колонке.

    Суммы выплат дня (наш расчет, snake_case; сырые поля ISS остались в ВЕРХНЕМ регистре):
      coupon_amount — купон;
      amort_amount  — погашение номинала: транш амортизации, а в дату погашения выпуска
                      последний транш (у досрочно амортизированной бумаги там 0.0);
      total_amount  — вся выплата дня, coupon_amount + amort_amount;
      offer_amount  — выкуп по оферте: цена оферты от непогашенного номинала.

    Тип события отдельной колонкой не дублируется, он читается из самих колонок:
      купон на дату      — заполнен STARTDATE (ISS отдает купонный период и у еще не
                           объявленного купона, так что купонная дата видна и при пустом
                           coupon_amount), у некупонных дат он пустой;
      погашение номинала — заполнен amort_amount;
      оферта             — заполнены OFFERPRICE и/или OFFERSTART/OFFEREND (цены у оферты
                           может не быть — 62 случая на 630 оферт, — но границы периода
                           ISS отдает всегда, так что дата оферты не теряется);
      погашение выпуска  — последняя дата расписания: дату погашения мы ставим в него
                           всегда, даже если номинал разошелся амортизацией раньше.

    coupon_amount и rate отдаются СЫРЫМИ: если купон еще не объявлен (флоатеры, бумаги
    с купоном только до оферты), там NaN, а не оценка. Доходности в get_bond_yields при
    этом считаются с переносом последней ставки — расписание и доходность намеренно
    расходятся в этом месте.

    total_amount пустой, если купон дня не объявлен, даже при известном транше: сумма
    выплаты неизвестна, а слагаемые видны в своих колонках. Выкуп по оферте в total_amount
    не входит — он условный: оферту еще надо предъявить.

    FACEVALUE — непогашенный номинал, действующий на дату события, посчитанный из расписания
    амортизаций; транш этой же даты в нем еще не погашен. Одноименное поле ISS для этого не
    годится: оно отдает текущий номинал во всех строках, включая те, что относятся к периодам
    с совсем другим номиналом.

    CALLOPTIONDATE и PUTOPTIONDATE в ISS — скаляры уровня выпуска, указывающие на ближайшую
    оферту, а не свойство периода. Поэтому они проставлены только в той строке, чья дата с
    ними совпала: так видно, какая из оферт расписания классифицирована биржей. Тип остальных
    оферт не восстановить — offertype в ISS не отличает put от call.

    У ненайденного или сломавшегося выпуска будет одна строка с пустой date и причиной
    в STATUS: одна битая бумага не роняет пачку. Отфильтровать: df[df.STATUS.isna()].
    """
    if isinstance(isins, str):
        isins = [isins]
    today = dt.date.today()
    rows = []

    for isin in isins:
        isin = str(isin).strip().upper()
        base = dict.fromkeys(SCHEDULE_COLUMNS)
        base["ISIN"] = isin
        try:
            info = find_bond(isin)
            if not info:
                rows.append(dict(base, STATUS="не найден на MOEX"))
                continue
            base.update({"SECID": info["SECID"], "SHORTNAME": info.get("SHORTNAME"),
                         "FACEUNIT": info.get("FACEUNIT")})

            bondization = get_bondization(info["SECID"])
            maturity = _date(info.get("MATDATE"))
            put, call = _date(info.get("PUTOPTIONDATE")), _date(info.get("CALLOPTIONDATE"))
            amortizations = normalize_amortizations(bondization["amortizations"],
                                                    _num(info.get("FACEVALUE")), today)
            pairs = _amort_pairs(amortizations)
            events, coupon_dates = {}, set()

            def event(when):
                """Строка дня: события одной даты попадают в нее, а не двоят дату."""
                return events.setdefault(when, dict(base, date=when))

            def outstanding(when, pairs=pairs):
                """Номинал, действующий на дату: транш этой даты еще не выплачен."""
                return round(sum(v for d, v in pairs if d >= when), 8) if pairs else None

            for _, r in bondization["coupons"].iterrows():
                when = _date(r.get("coupondate"))
                if when is None:
                    continue
                coupon_dates.add(when)
                row = event(when)
                row.update(coupon_amount=_total(row["coupon_amount"], _num(r.get("value"))),
                           rate=_num(r.get("valueprc")),
                           STARTDATE=_date(r.get("startdate")),
                           RECORDDATE=_date(r.get("recorddate")))

            redeemed = False
            for when, value in pairs:
                redeemed = redeemed or when == maturity
                row = event(when)
                row["amort_amount"] = _total(row["amort_amount"], value)
            if maturity and not redeemed:
                # номинал разошелся амортизацией до даты погашения — выпуск все равно гасится
                row = event(maturity)
                row["amort_amount"] = _total(row["amort_amount"],
                                             outstanding(maturity) or 0.0)

            for _, r in bondization["offers"].iterrows():
                when = _date(r.get("offerdate"))
                if when is None:
                    continue
                row = event(when)
                price, face = _num(r.get("price")), outstanding(when)
                buyback = round(face * price / 100, 8) if price and face else None
                row.update(OFFERPRICE=price, offer_amount=buyback,
                           OFFERSTART=_date(r.get("offerdatestart")),
                           OFFEREND=_date(r.get("offerdateend")))

            for when, row in sorted(events.items()):
                row["FACEVALUE"] = outstanding(when)
                # необъявленный купон делает неизвестной всю выплату дня, а не только себя
                unknown = when in coupon_dates and row["coupon_amount"] is None
                row["total_amount"] = None if unknown else \
                    _total(row["coupon_amount"], row["amort_amount"])
                row["CALLOPTIONDATE"] = call if call == when else None
                row["PUTOPTIONDATE"] = put if put == when else None
                rows.append(row)
        except requests.RequestException as e:
            rows.append(dict(base, STATUS=f"ошибка запроса: {e}"))
        except Exception as e:              # noqa: BLE001 — одна битая бумага не роняет пачку
            rows.append(dict(base, STATUS=f"ошибка разбора: {e}"))

    df = pd.DataFrame(rows, columns=SCHEDULE_COLUMNS)
    for column in SCHEDULE_DATES:
        df[column] = pd.to_datetime(df[column], errors="coerce")
    for column in SCHEDULE_NUMBERS:      # иначе пустые ячейки остаются None, а не NaN
        df[column] = pd.to_numeric(df[column], errors="coerce")
    return df.sort_values(["ISIN", "date"]).reset_index(drop=True)


# ------------------------------------------------------------ карточка бумаги

def iss_columns(engine: str = "stock", market: str = "bonds") -> dict:
    """Метаданные колонок рынка -> {блок: {поле: {short_title, title, type, ...}}} (кеш).

    ISS описывает все свои колонки сам, вместе с русскими подписями и точностью. Берем
    подписи оттуда, а не из захардкоженного словаря: биржа периодически добавляет поля,
    и карточка подхватит их без правок.
    """
    key = f"{engine}/{market}"
    if key not in _CACHE["columns"]:
        payload = _iss(f"engines/{engine}/markets/{market}/securities/columns.json")
        titles = {}
        for name in payload:
            frame = _block(payload, name)
            if not frame.empty:
                titles[name] = {r["name"]: r for r in frame.to_dict("records")}
        _CACHE["columns"][key] = titles
    return _CACHE["columns"][key]


def _typed(value, kind):
    """Значение ISS приходит строкой — приводим к типу, который биржа сама и указала."""
    if value is None or value == "":
        return None
    if kind == "number":
        return _num(value)
    if kind == "date":
        return _date(value)
    if kind == "boolean":
        return bool(_num(value)) if str(value).strip().isdigit() else value
    return value


def _as_rows(frame: pd.DataFrame, titles: dict, block: str) -> list:
    """Однострочный блок ISS -> длинные строки [Показатель, Значение, Поле, Блок]."""
    if frame.empty:
        return []
    rows = []
    for name, value in frame.iloc[0].to_dict().items():
        meta = titles.get(name, {})
        rows.append({"Показатель": meta.get("short_title") or meta.get("title") or name,
                     "Значение": value if not pd.isna(value) else None,
                     "Поле": name, "Блок": block})
    return rows


def _rename(frame: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Переименовывает колонки в русские подписи, сохраняя порядок известных полей."""
    if frame.empty:
        return pd.DataFrame(columns=list(mapping.values()))
    known = [c for c in mapping if c in frame.columns]
    return frame[known + [c for c in frame.columns if c not in mapping]].rename(columns=mapping)


BOARD_TITLES = {"boardid": "Режим", "title": "Название режима", "market": "Рынок",
                "engine": "Движок", "is_traded": "Торгуется", "is_primary": "Основной",
                "currencyid": "Валюта", "history_from": "История с",
                "history_till": "История по", "listed_from": "В листинге с",
                "listed_till": "В листинге по"}

INDEX_TITLES = {"SECID": "Индекс", "SHORTNAME": "Название", "FROM": "Включена",
                "TILL": "Исключена", "CURRENTVALUE": "Значение индекса",
                "LASTCHANGE": "Изменение", "LASTCHANGEPRC": "Изменение, %"}

AGGREGATE_TITLES = {"market_title": "Рынок", "tradedate": "Дата", "value": "Оборот",
                    "volume": "Объем, штук", "numtrades": "Сделок",
                    "updated_at": "Обновлено", "market_name": "Код рынка"}


def get_bond_card(isin: str) -> dict:
    """Все, что MOEX ISS знает о выпуске -> словарь читаемых таблиц.

    isin — один ISIN строкой. Для пачки бумаг вызывайте в цикле: карточка это досье
    на одну бумагу, а не выгрузка.

    Блоки словаря:
      passport   — паспорт выпуска: 41 поле от эмитента и регномера до флагов дефолта
      quotes     — котировки и параметры торгов основного режима: 130 полей трех блоков
                   ISS (параметры инструмента, рынок, доходности биржи)
      schedule   — полное расписание выплат, ровно как отдает get_coupon_schedule
      boards     — все режимы торгов выпуска с диапазонами истории
      indices    — индексы MOEX, в которые бумага входила, с датами включения
      aggregates — обороты по рынкам за последний торговый день

    Подписи берутся из метаданных самой биржи (iss_columns), поэтому новые поля ISS
    появятся в карточке сами. Значения приведены к типам, которые указала биржа.

    Истории торгов в карточке нет сознательно: это досье на бумагу, а не котировочная
    лента. Дневные свечи берутся из moex_cls.get_candles, доходности — из get_bond_yields.
    В quotes лежат доходности в расчете БИРЖИ (YIELD, YIELDTOOFFER, EFFECTIVEYIELD);
    наши, с переносом ставки и расчетами от T+1, живут только в get_bond_yields.

    Пустых ячеек в quotes много — биржа отдает все поля рынка, даже неприменимые
    к выпуску. Отфильтровать: card['quotes'].dropna(subset=['Значение']).
    """
    if not isinstance(isin, str):
        raise TypeError("get_bond_card принимает один ISIN строкой; для пачки — в цикле")
    isin = isin.strip().upper()

    info = find_bond(isin)
    if not info:
        raise ValueError(f"ISIN {isin} не найден на MOEX")
    secid = info["SECID"]

    security = _iss(f"securities/{secid}.json", **{"iss.only": "description,boards"})
    description, boards = _block(security, "description"), _block(security, "boards")

    market = _iss(f"engines/stock/markets/bonds/securities/{secid}.json",
                  **{"iss.only": "securities,marketdata,marketdata_yields"})
    titles = iss_columns()

    passport = pd.DataFrame(
        [{"Показатель": r.get("title") or r["name"],
          "Значение": _typed(r.get("value"), r.get("type")),
          "Поле": r["name"]}
         for r in description.sort_values("sort_order").to_dict("records")],
        columns=["Показатель", "Значение", "Поле"]) if not description.empty else \
        pd.DataFrame(columns=["Показатель", "Значение", "Поле"])

    board = info.get("BOARDID")
    quotes = []
    for name in ("securities", "marketdata", "marketdata_yields"):
        frame = _block(market, name)
        if not frame.empty and board and "BOARDID" in frame.columns:
            picked = frame[frame["BOARDID"] == board]
            frame = picked if not picked.empty else frame
        quotes += _as_rows(frame, titles.get(name, {}), name)

    return {
        "passport": passport,
        "quotes": pd.DataFrame(quotes, columns=["Показатель", "Значение", "Поле", "Блок"]),
        "schedule": get_coupon_schedule(isin),
        "boards": _rename(boards.sort_values(["is_traded", "is_primary"], ascending=False)
                          if not boards.empty else boards, BOARD_TITLES),
        "indices": _rename(_block(_iss(f"securities/{secid}/indices.json"), "indices"),
                           INDEX_TITLES),
        "aggregates": _rename(_block(_iss(f"securities/{secid}/aggregates.json"),
                                     "aggregates"), AGGREGATE_TITLES),
    }


# ------------------------------------------------- сектор экономики эмитента

# Единственная отраслевая классификация, которую публикует MOEX. Отдельного справочника
# секторов в ISS нет: поле SECTORID пустое у всех 3472 облигаций и всех акций, а блок
# emitters несет только ИНН, ОГРН и адрес. Поэтому сектор выводим через акции эмитента.
SECTOR_INDICES = {
    "MOEXOG": "Нефть и газ",
    "MOEXEU": "Электроэнергетика",
    "MOEXTL": "Телекоммуникации",
    "MOEXMM": "Металлы и добыча",
    "MOEXFN": "Финансы",
    "MOEXCN": "Потребительский сектор",
    "MOEXCH": "Химия и нефтехимия",
    "MOEXTN": "Транспорт",
    "MOEXIT": "Информационные технологии",
    "MOEXRE": "Недвижимость",
}

# У государственных и муниципальных выпусков сектор эмитента задан самим типом бумаги
TYPE_SECTORS = {
    "ofz_bond": "Государственный сектор",
    "cb_bond": "Государственный сектор",
    "subfederal_bond": "Субфедеральный сектор",
    "municipal_bond": "Муниципальный сектор",
    "ifi_bond": "Международные финансовые организации",
}

BY_INDEX, BY_TYPE = "отраслевой индекс MOEX", "тип выпуска"

SECTOR_COLUMNS = ["ISIN", "SECTOR", "SOURCE", "EMITENT", "INN", "SECID", "TICKER",
                  "BONDTYPE", "STATUS"]


def sector_map() -> dict:
    """{emitent_id: (сектор, тикер акции)} по составам отраслевых индексов (кеш на сессию).

    Стоит 21 запрос при первом обращении: 10 составов индексов и 11 страниц справочника
    торгуемых акций, из которого берется привязка тикера к эмитенту. Дальше бесплатно.
    """
    if _CACHE["sectors"] is not None:
        return _CACHE["sectors"]

    shares = {}
    for index, sector in SECTOR_INDICES.items():
        payload = _iss(f"statistics/engines/stock/markets/index/analytics/{index}.json",
                       **{"iss.only": "analytics"})
        frame = _block(payload, "analytics")
        for secid in frame.get("secids", []):
            if secid:
                shares[secid] = sector

    mapping, start = {}, 0
    while True:                       # справочник акций отдается страницами по 100
        frame = _block(_iss("securities.json", **{
            "iss.only": "securities", "group_by": "group",
            "group_by_filter": "stock_shares", "is_trading": 1,
            "limit": 100, "start": start}), "securities")
        if frame.empty:
            break
        for row in frame.to_dict("records"):
            sector = shares.get(row.get("secid"))
            if sector and row.get("emitent_id"):
                mapping[row["emitent_id"]] = (sector, row["secid"])
        if len(frame) < 100:
            break
        start += 100

    _CACHE["sectors"] = mapping
    return mapping


def find_emitter(isin: str) -> Optional[dict]:
    """ISIN -> строка справочника с эмитентом: emitent_id, название, ИНН, тип бумаги."""
    isin = isin.strip().upper()
    if isin not in _CACHE["emitters"]:
        frame = _block(_iss("securities.json", q=isin, **{"iss.only": "securities"}),
                       "securities")
        if not frame.empty and "isin" in frame.columns:
            frame = frame[frame["isin"].fillna("").str.upper() == isin]
        _CACHE["emitters"][isin] = None if frame.empty else frame.iloc[0].to_dict()
    return _CACHE["emitters"][isin]


def get_issuer_sector(isins: Iterable[str]) -> pd.DataFrame:
    """Сектор экономики эмитента по списку ISIN -> DataFrame.

    Колонки: ISIN, SECTOR, SOURCE, EMITENT, INN, SECID, TICKER, BONDTYPE, STATUS.

    ВАЖНО про полноту. Своего справочника секторов MOEX не публикует, поэтому сектор
    определяется двумя способами, и оба неполные:

      SOURCE='отраслевой индекс MOEX' — эмитент облигации сам торгуется акциями, которые
        входят в один из десяти отраслевых индексов биржи. Так классифицируется около
        четверти торгуемых облигаций, причем 700 из 990 попаданий — 'Финансы': банки
        занимают под собственным юрлицом.
      SOURCE='тип выпуска' — ОФЗ, субфедеральные и муниципальные бумаги: там сектор
        эмитента задан самим типом выпуска. Это еще 165 бумаг.

    У остальных SECTOR пуст, и это не сбой, а структура рынка: облигации обычно выпускают
    дочерние SPV со своим ИНН (ООО 'Газпром капитал', ООО 'РЕСО-Лизинг'), у которых акций
    нет и быть не может, а связи 'дочка — материнская компания' в ISS нет вообще. Для таких
    выпусков в колонке INN лежит ИНН эмитента: по нему сектор добирается из внешнего
    источника с ОКВЭД (ЕГРЮЛ, dadata), MOEX тут больше ничего не знает.
    """
    if isinstance(isins, str):
        isins = [isins]
    mapping = sector_map()
    rows = []

    for isin in isins:
        isin = str(isin).strip().upper()
        rec = dict.fromkeys(SECTOR_COLUMNS)
        rec["ISIN"] = isin
        try:
            emitter = find_emitter(isin)
            if not emitter:
                rec["STATUS"] = "не найден на MOEX"
                rows.append(rec)
                continue
            rec.update({"SECID": emitter.get("secid"), "EMITENT": emitter.get("emitent_title"),
                        "INN": emitter.get("emitent_inn"), "BONDTYPE": emitter.get("type"),
                        "STATUS": "ok"})

            found = mapping.get(emitter.get("emitent_id"))
            if found:
                rec.update({"SECTOR": found[0], "SOURCE": BY_INDEX, "TICKER": found[1]})
            elif emitter.get("type") in TYPE_SECTORS:
                rec.update({"SECTOR": TYPE_SECTORS[emitter["type"]], "SOURCE": BY_TYPE})
            else:
                rec["STATUS"] = "эмитент не представлен в отраслевых индексах MOEX"
        except requests.RequestException as e:
            rec["STATUS"] = f"ошибка запроса: {e}"
        except Exception as e:              # noqa: BLE001 — одна битая бумага не роняет пачку
            rec["STATUS"] = f"ошибка разбора: {e}"
        rows.append(rec)

    return pd.DataFrame(rows, columns=SECTOR_COLUMNS)


if __name__ == "__main__":
    ISINS = ["RU000A0ZZ5H3", "RU000A1032P1", "RU000A1038V6"]
    print(get_bond_yields(ISINS, date="2026-09-04"), "\n")
    print(get_bond_yields(ISINS, date="2026-09-04", details=True)[
        ["ISIN", "SHORTNAME", "SETTLE", "PRICE", "PRICE_DATE", "ACCRUED",
         "FACEVALUE", "MATDATE", "PUTDATE", "STATUS"]], "\n")
    print(get_coupon_schedule("RU000A0ZZ5H3")[
        ["date", "coupon_amount", "rate", "amort_amount", "total_amount", "FACEVALUE",
         "OFFERPRICE", "offer_amount", "PUTOPTIONDATE"]].tail(8), "\n")

    card = get_bond_card("RU000A0ZZ5H3")
    print("карточка:", {name: len(frame) for name, frame in card.items()}, "\n")
    print(card["passport"].head(8).to_string(index=False), "\n")

    print(get_issuer_sector(["RU000A1043H5", "RU000A1038V6", "RU000A0ZZ5H3"])[
        ["ISIN", "SECTOR", "SOURCE", "EMITENT", "TICKER"]].to_string(index=False))
