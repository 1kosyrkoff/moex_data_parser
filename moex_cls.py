# -*- coding: utf-8 -*-
"""Выгрузка котировок (свечей) с MOEX ISS: ticker, start_date, finish_date, period.

custom_index — доходность взвешенного композита тикеров (спецификация: spec-custom-index.md).
series_index — то же по переданным pd.Series (спецификация: spec-series-index.md).
"""
import datetime as dt
import math
import numbers
import warnings
from typing import Literal, Union

import requests
import pandas as pd

ISS = "https://iss.moex.com/iss"
PAGE = 500  # ISS отдает не более 500 свечей за запрос

# псевдоним периода -> код интервала ISS ("m" — минуты, месяц это "1mo")
PERIODS = {
    "1m": 1, "10m": 10, "1h": 60, "1d": 24, "1w": 7, "1mo": 31, "1q": 4,
    "min": 1, "hour": 60, "day": 24, "week": 7, "month": 31, "quarter": 4,
}

# подсказка редактору: эти значения всплывают в автодополнении period=...
# (в PERIODS их больше — слова 'day', 'week' и коды ISS тоже принимаются)
Period = Literal["1m", "10m", "1h", "1d", "1w", "1mo", "1q"]

MSK = dt.timezone(dt.timedelta(hours=3))  # без перехода на летнее время, tzdata не нужна
LAG = pd.Timedelta(days=10)  # сужение периода больше этого — предупреждение в custom_index и series_index

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (research; MOEX ISS client)"})


def interval_code(period: Union[Period, int]) -> int:
    """Приводит период к коду интервала ISS: 1, 10, 60, 24, 7, 31, 4."""
    if isinstance(period, int):
        code = period
    else:
        code = PERIODS.get(str(period).strip().lower())
    if code not in set(PERIODS.values()):
        raise ValueError(f"Неизвестный period={period!r}. Доступно: {sorted(PERIODS)}")
    return code


def parse_date(value, name="date"):
    """Приводит дату к формату ISS 'YYYY-MM-DD'.

    ISS понимает только ISO-формат: на '31.07.2026' он не выдает ошибку,
    а молча возвращает пустой ответ, поэтому дату нормализуем сами.
    """
    if isinstance(value, (dt.date, dt.datetime)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y%m%d", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"Не понимаю {name}={value!r}, ожидается 'YYYY-MM-DD' или 'DD.MM.YYYY'")


def find_board(ticker):
    """Ищет (engine, market, board) для тикера: акции -> TQBR, валюта -> CETS и т.д."""
    r = S.get(f"{ISS}/securities/{ticker}.json",
              params={"iss.meta": "off", "iss.only": "boards"}, timeout=30)
    r.raise_for_status()
    b = r.json()["boards"]
    df = pd.DataFrame(b["data"], columns=b["columns"])
    if df.empty:
        raise ValueError(f"Тикер {ticker} не найден на MOEX")
    df = df.sort_values(["is_traded", "is_primary"], ascending=False)
    row = df.iloc[0]
    return row["engine"], row["market"], row["boardid"]


def get_candles(ticker: str, start_date, finish_date,
                period: Union[Period, int] = "1d",
                board: str = None) -> pd.DataFrame:
    """Свечи по тикеру за период -> DataFrame (индекс begin, колонки OHLC + value/volume).

    ticker      — 'SBER', 'CNYRUB_TOM', 'IMOEX', 'SU26238RMFS4', ...
    start_date  — 'YYYY-MM-DD', 'DD.MM.YYYY' или date/datetime
    finish_date — то же, включительно
    period      — 1m / 10m / 1h / 1d / 1w / 1mo / 1q
    board       — режим торгов, по умолчанию определяется автоматически
    """
    ticker = ticker.strip().upper()
    start_date = parse_date(start_date, "start_date")
    finish_date = parse_date(finish_date, "finish_date")
    if start_date > finish_date:
        raise ValueError(f"start_date ({start_date}) позже finish_date ({finish_date})")
    engine, market, primary_board = find_board(ticker)
    board = board or primary_board
    url = f"{ISS}/engines/{engine}/markets/{market}/boards/{board}/securities/{ticker}/candles.json"

    rows, columns, start = [], None, 0
    while True:  # ISS отдает свечи страницами по 500 штук
        r = S.get(url, params={"iss.meta": "off", "from": start_date, "till": finish_date,
                               "interval": interval_code(period), "start": start}, timeout=60)
        r.raise_for_status()
        c = r.json()["candles"]
        columns = c["columns"]
        rows += c["data"]
        if len(c["data"]) < PAGE:
            break
        start += len(c["data"])

    df = pd.DataFrame(rows, columns=columns)
    if df.empty:
        return df
    df["begin"] = pd.to_datetime(df["begin"])
    df["end"] = pd.to_datetime(df["end"])
    df = df.drop_duplicates("begin").sort_values("begin").set_index("begin")
    df.insert(0, "ticker", ticker)
    return df[["ticker", "open", "high", "low", "close", "value", "volume", "end"]]


def custom_index(weights: dict, start_date, finish_date, name: str, *,
                 intraday: bool = False) -> pd.Series:
    """Накопленная доходность взвешенного композита тикеров -> Series (индекс date, в долях).

    weights     — {'RUCBTRNS': 0.5, 'IMOEX': 0.5}: каждый вес > 0, сумма = 1
    start_date  — 'DD.MM.YYYY', 'YYYY-MM-DD' или date/datetime
    finish_date — то же, включительно
    name        — имя возвращаемой Series
    intraday    — оставить свечу за сегодня (по Москве), хотя день еще не закрыт

    Берется close дневных свечей с борды из find_board и только на датах, где котировка есть
    у всех тикеров. На этих датах считаются доходности тикеров, доходность композита за день —
    sum(w * r) (ежедневная ребалансировка), результат — их цепочка минус 1: 0 на первой общей
    дате, последнее значение — доходность за весь период.
    У акций и облигаций close — цена без дивидендов и купонов, а не полная доходность.

    Тикер без котировок или нет общих дат — ValueError. Если из-за тикера период сузился
    больше чем на 10 дней (индекс перестали считать, еще не считали) — UserWarning.
    """
    if not isinstance(weights, dict) or not weights:
        raise ValueError("weights: нужен непустой словарь {тикер: вес}")
    w = {}
    for ticker, weight in weights.items():
        key = str(ticker).strip().upper()
        if key in w:
            raise ValueError(f"weights: тикер {key} указан дважды")
        if isinstance(weight, bool) or not isinstance(weight, numbers.Real) \
                or not math.isfinite(weight) or weight <= 0:
            raise ValueError(f"weights: вес {key} должен быть числом > 0, получено {weight!r}")
        w[key] = float(weight)
    total = math.fsum(w.values())
    if abs(total - 1) > 1e-9:  # допуск: 0.2 + 0.2 + 0.6 во float не равно 1
        raise ValueError(f"weights: сумма весов {total:.10g}, а должна быть 1")

    start = pd.Timestamp(parse_date(start_date, "start_date"))
    finish = pd.Timestamp(parse_date(finish_date, "finish_date"))
    today = pd.Timestamp(dt.datetime.now(MSK).date())

    closes = {}
    for ticker in w:
        candles = get_candles(ticker, start, finish, "1d")  # борду определяет find_board
        if not intraday:
            # у закрытых свечей end бывает и 23:59:55, поэтому незакрытую узнаем по дате
            candles = candles[candles.index.normalize() < today] if len(candles) else candles
        close = candles["close"].astype("float64").dropna() if len(candles) else pd.Series(dtype="float64")
        if close.empty:
            hint = ", свеча за сегодня отброшена (intraday=False)" if not intraday and finish >= today else ""
            raise ValueError(f"{ticker}: нет котировок с {start:%d.%m.%Y} по {finish:%d.%m.%Y}{hint} — "
                             "возможно, индекс больше не считается или еще не считался")
        if (close <= 0).any():
            raise ValueError(f"{ticker}: котировка {close[close <= 0].iloc[0]} "
                             f"на {close[close <= 0].index[0]:%d.%m.%Y}, доходность не определена")
        closes[ticker] = close

    # доходности только после join: иначе, например, движение облигационного индекса
    # за субботу пропадет вместе с субботой, а не перейдет в доходность понедельника
    levels = pd.concat(closes, axis=1, join="inner").sort_index()
    if levels.empty:
        ranges = "; ".join(f"{t} {s.index[0]:%d.%m.%Y}–{s.index[-1]:%d.%m.%Y}" for t, s in closes.items())
        raise ValueError(f"{name}: у тикеров нет общих дат ({ranges})")

    first, last = levels.index[0], levels.index[-1]
    end = min(finish, today)
    if first - start > LAG or end - last > LAG:
        causes = [f"{t} с {s.index[0]:%d.%m.%Y}" for t, s in closes.items() if s.index[0] - start > LAG]
        causes += [f"{t} по {s.index[-1]:%d.%m.%Y}" for t, s in closes.items() if end - s.index[-1] > LAG]
        warnings.warn(f"{name}: котировки {', '.join(causes) or 'у тикеров не пересекаются'}; "
                      f"доходность посчитана с {first:%d.%m.%Y} по {last:%d.%m.%Y}", stacklevel=2)

    daily = (levels / levels.shift(1) - 1).fillna(0.0) @ pd.Series(w)  # первая дата — точка отсчета
    growth = (1 + daily).cumprod() - 1
    growth.index = growth.index.normalize()
    growth.index.name = "date"
    growth.name = name
    return growth


def _clean_series(s, label: str, positive: bool = True) -> pd.Series:
    """Ряд пользователя -> float64 по дням: время и пояс отброшены, за день последнее значение, без NaN.

    positive=True — значения обязаны быть > 0 (уровни); ставки могут быть нулевыми и отрицательными.
    """
    if not isinstance(s, pd.Series):
        raise ValueError(f"{label}: ожидается pd.Series, получено {type(s).__name__}")
    dtype = s.dtype
    if pd.api.types.is_bool_dtype(dtype):
        raise ValueError(f"{label}: значения должны быть float/int, получен bool")
    if not (pd.api.types.is_integer_dtype(dtype) or pd.api.types.is_float_dtype(dtype)):
        for key, v in s.items():
            if v is None or v is pd.NA or (isinstance(v, float) and math.isnan(v)):
                continue
            if isinstance(v, bool) or not isinstance(v, numbers.Real):
                day = f"{key:%d.%m.%Y}" if isinstance(key, dt.date) else repr(key)
                raise ValueError(f"{label}: значение не float/int: {day} → {v!r} ({type(v).__name__})")

    if pd.api.types.is_datetime64_any_dtype(s.index.dtype) \
            or (s.index.dtype == object and all(isinstance(x, dt.date) for x in s.index)):
        dates = pd.DatetimeIndex(s.index)
    else:
        raise ValueError(f"{label}: индекс должен быть датами, получен {s.index.dtype}")
    if dates.tz is not None:
        dates = dates.tz_localize(None)

    v = pd.Series(s.to_numpy(), index=dates).dropna()
    if v.empty:
        raise ValueError(f"{label}: ряд пуст после удаления NaN")
    v = v.astype("float64").sort_index(kind="stable")
    v.index = v.index.normalize()
    v = v[~v.index.duplicated(keep="last")]  # после стабильной сортировки last = позднее время дня
    bad = v.abs() == math.inf
    if bad.any():
        raise ValueError(f"{label}: значение {v[bad].iloc[0]} на {v[bad].index[0]:%d.%m.%Y}")
    if positive and (v <= 0).any():
        raise ValueError(f"{label}: значение {v[v <= 0].iloc[0]} "
                         f"на {v[v <= 0].index[0]:%d.%m.%Y}, доходность не определена")
    return v


def series_index(pairs, start_date=None, finish_date=None, name=None, *,
                 cumulative: bool = False) -> pd.Series:
    """Накопленная доходность взвешенного композита переданных рядов -> Series (индекс date, в долях).

    pairs       — [(ряд, вес), ...]: ряд — pd.Series с датами в индексе, вес > 0, сумма весов = 1
    start_date  — 'DD.MM.YYYY', 'YYYY-MM-DD', date/datetime; None — с начала общего периода
    finish_date — то же, включительно; None — до конца общего периода
    name        — имя возвращаемой Series
    cumulative  — False: ряды — уровни (индекс, пай, цена); True: ряды — ставки overnight
                  в % годовых (21.35), результат — доход от вложения под эти ставки

    Расчет как в custom_index: только даты, где значение есть у всех рядов; на них доходность
    композита за шаг — sum(w * r) (ребалансировка на каждой общей дате), результат — цепочка
    шагов минус 1, в долях: 0 на первой общей дате, последнее — доходность за период.
    Шаг при cumulative=False — r = P_t / P_prev - 1. При cumulative=True — r = ставка_prev / 100 *
    дни / 365: ставка даты действует календарные дни до следующей даты ряда, то есть в выходные
    и пропуски — последняя известная ставка; ставка последней даты уже не начисляется.
    Ряды приводятся как в value_growth: время и часовой пояс отбрасываются, из нескольких
    значений за день берется последнее по времени, NaN удаляются.

    Ряд без данных за период или нет общих дат — ValueError. Если из-за ряда период сузился
    больше чем на 10 дней относительно заданной (не None) границы — UserWarning.
    В сообщениях ряд называется по Series.name, без имени — «ряд №N».
    """
    if isinstance(pairs, (dict, pd.Series, pd.DataFrame, str)) or not hasattr(pairs, "__iter__"):
        raise ValueError("pairs: нужен список пар [(ряд, вес), ...] "
                         "(запись {ряд: вес} невозможна — pd.Series не может быть ключом словаря)")
    pairs = list(pairs)
    if not pairs:
        raise ValueError("pairs: нужен непустой список пар [(ряд, вес), ...]")
    labels, weights = [], []
    for i, pair in enumerate(pairs, 1):
        if not isinstance(pair, (tuple, list)) or len(pair) != 2:
            got = f"{type(pair).__name__} длины {len(pair)}" if isinstance(pair, (tuple, list)) else type(pair).__name__
            raise ValueError(f"пара №{i}: ожидается (pd.Series, вес), получено {got}")
        s, weight = pair
        label = str(s.name) if isinstance(s, pd.Series) and s.name is not None else f"ряд №{i}"
        if isinstance(weight, bool) or not isinstance(weight, numbers.Real) \
                or not math.isfinite(weight) or weight <= 0:
            raise ValueError(f"{label}: вес должен быть числом > 0, получено {weight!r}")
        labels.append(label)
        weights.append(float(weight))
    total = math.fsum(weights)
    if abs(total - 1) > 1e-9:  # допуск: 0.2 + 0.2 + 0.6 во float не равно 1
        raise ValueError(f"pairs: сумма весов {total:.10g}, а должна быть 1")

    start = None if start_date is None else pd.Timestamp(parse_date(start_date, "start_date"))
    finish = None if finish_date is None else pd.Timestamp(parse_date(finish_date, "finish_date"))
    if start is not None and finish is not None and start > finish:
        raise ValueError(f"start_date ({start:%d.%m.%Y}) позже finish_date ({finish:%d.%m.%Y})")
    period = " ".join(p for p in (start and f"с {start:%d.%m.%Y}", finish and f"по {finish:%d.%m.%Y}") if p)

    values = []
    for (s, _), label in zip(pairs, labels):
        v = _clean_series(s, label, positive=not cumulative).loc[start:finish]
        if v.empty:
            raise ValueError(f"{label}: нет данных {period}")
        values.append(v)

    # доходности только после join: иначе движение за даты, которых нет у других рядов, пропадет
    levels = pd.concat(dict(enumerate(values)), axis=1, join="inner").sort_index()
    if levels.empty:
        ranges = "; ".join(f"{l} {v.index[0]:%d.%m.%Y}–{v.index[-1]:%d.%m.%Y}" for l, v in zip(labels, values))
        raise ValueError(f"{name}: у рядов нет общих дат ({ranges})")

    first, last = levels.index[0], levels.index[-1]
    end = None if finish is None else min(finish, pd.Timestamp(dt.datetime.now(MSK).date()))
    late = start is not None and first - start > LAG
    early = end is not None and end - last > LAG
    if late or early:
        causes = [f"{l} с {v.index[0]:%d.%m.%Y}" for l, v in zip(labels, values) if late and v.index[0] - start > LAG]
        causes += [f"{l} по {v.index[-1]:%d.%m.%Y}" for l, v in zip(labels, values) if early and end - v.index[-1] > LAG]
        warnings.warn(f"{name}: данные {', '.join(causes) or 'у рядов не пересекаются'}; "
                      f"доходность посчитана с {first:%d.%m.%Y} по {last:%d.%m.%Y}", stacklevel=2)

    if cumulative:
        # ставка предыдущей даты работает все календарные дни до текущей (пятница — 3 дня)
        days = levels.index.to_series().diff().dt.days
        step = (levels.shift(1) / 100 @ pd.Series(weights)) * days / 365
    else:
        step = (levels / levels.shift(1) - 1) @ pd.Series(weights)
    growth = (1 + step.fillna(0.0)).cumprod() - 1  # первая дата — точка отсчета
    growth.index.name = "date"
    growth.name = name
    return growth


if __name__ == "__main__":
    print(get_candles("SBER", "2024-01-01", "2024-03-01", "1d").head())
