#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Таблица «Альфа по фондам и периодам» с листа «Сводка» — отдельным xlsx.

Из своего кода (вход — DataFrame в формате выгрузки all_funds_perform):

    from alpha_table import make_alpha_table

    df = my_func(...)                    # moment + накопленные доходности рядов (0 на базовую дату)
    path = make_alpha_table(df, "alpha_by_fund_30092026.xlsx")

Из терминала (вход — файл выгрузки или готовый отчёт funds_vs_benchmarks_*.xlsx):

    python ./alpha_table.py 'path_to_excel' [out.xlsx]

DataFrame: даты — в колонке «moment» (или «Дата»/«date») либо в DatetimeIndex; остальные колонки —
ряды фондов и индексов по кодам (RU000A0JR290, MCF2TR, …); значения — накопленная доходность
в долях (−0.012 = −1,2%). Если в df уже уровни (1 + доходность), передайте levels=True.
Лишние колонки не мешают; недостающие ряды — ошибка со списком.

Результат: один лист «Сводка», на нём только таблица (шапка в A1, 19 фондов × периоды 1 мес … 5 лет,
«Периодов с альфой > 0», замечание) и строка-пояснение двумя строками ниже. Валюта не учитывается:
курс CNY/RUB не нужен, «Ликвидность. Юань» в таблицу не входит.

Методика — как в полном отчёте: прирост за период = уровень на конец / уровень на начало − 1,
граница — последнее значение на дату или раньше (допуск 10 дней), «2 года» — максимальное окно
данных, если история покрывает ≥ 90% периода; альфа = прирост фонда − прирост бенчмарка.
Модуль самодостаточный: нужны только pandas, numpy и openpyxl.
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pandas.tseries.offsets import DateOffset

__all__ = ["make_alpha_table", "alpha_dataframe", "FUNDS", "PERIODS"]

# ============================================================================ настройки
# Фонды в порядке листа «Сводка»: (название, тип, ряд фонда, бенчмарк, замечание)
FUNDS = [
    ("Индекс МосБиржи (ОПИФ)", "ОПИФ", "RU000A0JR290", "MCF2TR", None),
    ("Акции", "ОПИФ", "RU000A0JR282", "MCFTR", None),
    ("Перспективные размещения", "ОПИФ", "RU000A0JU8K6", "VTBFG0013", None),
    ("Денежный рынок. Рубли", "ОПИФ", "RU000A1056R6", "RUSFAR", None),
    ("Казначейский", "ОПИФ", "RU000A0JR2C1", "RUCBTRNS", None),
    ("Металлургия", "ОПИФ", "RU000A0JS9R3", "MOEXMM", None),
    ("Нефтегазовый сектор", "ОПИФ", "RU000A0JS9Q5", "MOEXOG", None),
    ("Облигации. Ответственные инвестиции", "ОПИФ", "RU000A0JU8M2", "RUCBTRNS", None),
    ("Сбалансированные инвестиции", "ОПИФ", "RU000A0JR2A5", "VTBFG0005", None),
    ("Сбалансированный", "ОПИФ", "RU000A0JNWK1", "VTBFG0006", None),
    ("Сбалансированный. Консервативный", "ОПИФ", "RU000A0JS9T9", "VTBFG0011", None),
    ("Инновационный", "ОПИФ", "RU000A0JS9P7", "MOEXIT", "смена мандата: Электроэнергетика → Инновационный"),
    ("Накопительный резерв", "ОПИФ", "RU000A0JT4S1", "VTBFG0010", "скорректировано на дробление 1:100"),
    ("Облигации. Российские эмитенты", "ОПИФ", "RU000A104R48", "RUCBTRNS", None),
    ("Акции. Российские эмитенты", "ОПИФ", "RU000A104R55", "MCFTR", None),
    ("Индекс МосБиржи (БПИФ)", "БПИФ", "RU000A101EJ5", "VTBFG0012", None),
    ("Ликвидность", "БПИФ", "RU000A1014L8", "RUSFAR", None),
    ("Корпоративные облигации", "БПИФ", "RU000A1002S8", "RUCBTRNS", None),
    ("Устойчивое развитие российских компаний", "БПИФ", "RU000A103LL2", "MRSVRT", None),
]
# Дробления паёв: (ряд, дата, множитель). Применяется, только если скачок есть в данных.
SPLITS = [("RU000A0JT4S1", "2026-08-14", 100.0)]
PERIODS = [("1 месяц", 1), ("2 месяца", 2), ("3 месяца", 3), ("6 месяцев", 6), ("12 месяцев", 12),
           ("2 года", 24), ("3 года", 36), ("4 года", 48), ("5 лет", 60)]
STALE_DAYS = 10
MAXWIN_COVERAGE = 0.90
DATE_COLS = ("moment", "Дата", "date", "Date")

COLS = ["Фонд", "Тип", "Бенчмарк"] + [p for p, _ in PERIODS] + ["Периодов с альфой > 0", "Замечание"]
DASH = "—"

# оформление как на листе «Сводка»
ARIAL = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="FF3A3A38")
ZEBRA_FILL = PatternFill("solid", fgColor="FFF2F2EF")
THIN = Side(style="thin", color="FFD6D5CE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREY = "FF52514E"
WIDTHS = {"A": 40, "B": 8, "C": 17, "M": 14, "N": 75}
PERIOD_WIDTH = 11


# ============================================================================ данные
def _as_dates(values) -> pd.DatetimeIndex | None:
    """Даты, если значения на них похожи (≥ 90% распознаётся), иначе None."""
    s = pd.Series(values)
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_datetime64_any_dtype(s):
        return None                         # RangeIndex / числа — не даты
    d = pd.to_datetime(s, dayfirst=True, errors="coerce")
    return pd.DatetimeIndex(d) if len(d) and d.notna().mean() >= 0.9 else None


def prepare_levels(data: pd.DataFrame, levels: bool = False) -> pd.DataFrame:
    """DataFrame выгрузки → уровни рядов (1 + накопленная доходность) с датами в индексе."""
    df = data.copy()
    idx = df.index if isinstance(df.index, pd.DatetimeIndex) else None
    if idx is None:
        col = next((c for c in DATE_COLS if c in df.columns), None)
        if col is not None:
            idx, df = _as_dates(df[col]), df.drop(columns=col)
        elif (idx := _as_dates(df.index)) is None:
            idx, df = _as_dates(df.iloc[:, 0]), df.iloc[:, 1:]     # даты в первой колонке
        if idx is None:
            raise ValueError("Не нашёл даты: нужна колонка moment (или Дата/date) либо DatetimeIndex")
    df.index = pd.DatetimeIndex(idx)
    df = df[df.index.notna()]
    df = df[~df.index.duplicated(keep="last")].sort_index()
    df = df.apply(pd.to_numeric, errors="coerce")
    start = df.apply(lambda c: c.dropna().iloc[0] if c.notna().any() else np.nan).median()
    if not levels and start > 0.5:
        raise ValueError("Похоже, в DataFrame уже уровни (≈1 на старте), а не доходности — передайте levels=True")
    if levels and start < 0.5:
        raise ValueError("Похоже, в DataFrame накопленные доходности (≈0 на старте) — уберите levels=True")
    lv = df if levels else 1.0 + df
    for code, d, k in SPLITS:
        d = pd.Timestamp(d)
        if code not in lv or not (lv.index.min() < d <= lv.index.max()):
            continue
        before = lv.loc[:d - pd.Timedelta(days=1), code].dropna()
        after = lv.loc[d:, code].dropna()
        if len(before) and len(after) and after.iloc[0] / before.iloc[-1] < 2.0 / k:   # скачок есть — правим
            lv.loc[d:, code] = lv.loc[d:, code] * k
    x = lv.ffill()
    jumps = (x / x.shift(1) - 1).abs()
    big = [(c, d.date()) for c in lv.columns for d in jumps.index[jumps[c] > 0.5]]
    if big:
        print("ВНИМАНИЕ: дневные изменения > 50% (новое дробление?):", big)
    return lv


# ============================================================================ расчёт
@dataclass
class PeriodDef:
    label: str
    nominal_start: pd.Timestamp
    start: pd.Timestamp | None
    end: pd.Timestamp
    status: str          # ok | maxwin | short


def period_defs(dates: pd.DatetimeIndex) -> list[PeriodDef]:
    end, first = dates.max(), dates.min()
    out = []
    for label, m in PERIODS:
        nom = end - DateOffset(months=m)
        if nom >= first:
            out.append(PeriodDef(label, nom, dates[dates <= nom].max(), end, "ok"))
        elif (end - first).days / (end - nom).days >= MAXWIN_COVERAGE:
            out.append(PeriodDef(label, nom, first, end, "maxwin"))
        else:
            out.append(PeriodDef(label, nom, None, end, "short"))
    return out


def _asof_pos(s: pd.Series, pos: int, dates: pd.DatetimeIndex) -> int | None:
    """Позиция последнего значения на строке pos или раньше, не старше STALE_DAYS дней."""
    d0 = dates[pos]
    for p in range(pos, -1, -1):
        if (d0 - dates[p]).days > STALE_DAYS:
            return None
        if not pd.isna(s.iat[p]):
            return p
    return None


def period_alpha(lv: pd.DataFrame, fcode: str, bcode: str, p: PeriodDef) -> float | None:
    """Альфа (ариф.) фонда к бенчмарку за период; None — период не покрыт данными."""
    if p.status == "short":
        return None
    dates, f, b = lv.index, lv[fcode], lv[bcode]
    end_pos = dates.get_loc(p.end)
    start_pos = int(np.argmax((f.notna() & b.notna()).values)) if p.status == "maxwin" else dates.get_loc(p.start)
    fp, bp = _asof_pos(f, start_pos, dates), _asof_pos(b, start_pos, dates)
    fe, be = _asof_pos(f, end_pos, dates), _asof_pos(b, end_pos, dates)
    if None in (fp, bp, fe, be) or fe <= fp or be <= bp:
        return None
    return (f.iat[fe] / f.iat[fp] - 1) - (b.iat[be] / b.iat[bp] - 1)


def alpha_dataframe(data: pd.DataFrame, *, levels: bool = False) -> tuple[pd.DataFrame, list[PeriodDef]]:
    """Таблица «Альфа по фондам и периодам» как DataFrame (столбцы — как на листе «Сводка»)."""
    lv = prepare_levels(data, levels)
    missing = sorted({c for _, _, f, b, _ in FUNDS for c in (f, b)} - set(lv.columns))
    if missing:
        raise ValueError(f"В данных нет рядов: {missing}")
    pdefs = period_defs(lv.index)
    rows = []
    for name, ftype, fcode, bcode, note in FUNDS:
        alphas = [period_alpha(lv, fcode, bcode, p) for p in pdefs]
        nums = [a for a in alphas if a is not None]
        rows.append([name, ftype, bcode, *[DASH if a is None else a for a in alphas],
                     f"{sum(a > 0 for a in nums)} из {len(nums)}", note])
    return pd.DataFrame(rows, columns=COLS), pdefs


# ============================================================================ xlsx
def _save_table(df: pd.DataFrame, pdefs: list[PeriodDef], out: Path, source: str | None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    ws.sheet_view.showGridLines = False
    first_p, last_p = 4, 3 + len(PERIODS)                # D … L
    last_row = len(df) + 1

    for c, h in enumerate(COLS, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name=ARIAL, sz=9, b=True, color="FFFFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 31.5

    for i, row in enumerate(df.itertuples(index=False), 2):
        zebra = (i % 2 == 1)                              # как в оригинале: каждая вторая строка
        for c, v in enumerate(row, 1):
            cell = ws.cell(i, c)
            if c == len(COLS) - 1:                       # «Периодов с альфой > 0» — формула от строки
                a, b = get_column_letter(first_p), get_column_letter(last_p)
                cell.value = f'=COUNTIF({a}{i}:{b}{i},">0")&" из "&COUNT({a}{i}:{b}{i})'
                cell.font = Font(name=ARIAL, sz=10)
            elif first_p <= c <= last_p and v != DASH:
                cell.value = float(v)
                cell.number_format = "0.00%"
                cell.font = Font(name=ARIAL, sz=10)
            elif v == DASH or c == len(COLS):
                cell.value = v if isinstance(v, str) else None
                cell.font = Font(name=ARIAL, sz=9, color=GREY)
                if v == DASH:
                    cell.alignment = Alignment(horizontal="center")
            else:
                cell.value = v
                cell.font = Font(name=ARIAL, sz=10)
            cell.border = BORDER
            if zebra:
                cell.fill = ZEBRA_FILL

    rng = f"{get_column_letter(first_p)}2:{get_column_letter(last_p)}{last_row}"
    tl = f"{get_column_letter(first_p)}2"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({tl}),{tl}>0)"],
                                                   font=Font(name=ARIAL, sz=10, b=True, color="FF1F5FA9")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({tl}),{tl}<0)"],
                                                   font=Font(name=ARIAL, sz=10, color="FFB3372F")))

    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    for c in range(first_p, last_p + 1):
        ws.column_dimensions[get_column_letter(c)].width = PERIOD_WIDTH
    ws.freeze_panes = "B2"
    ws.page_setup.orientation = "landscape"                # печать: одна страница в ширину
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    end = pdefs[0].end
    maxwin = next((p for p in pdefs if p.status == "maxwin"), None)
    note = (f"Альфа = прирост фонда − прирост бенчмарка за период (арифметическая). Конец всех периодов — "
            f"{end:%d.%m.%Y}.")
    if maxwin is not None:
        yrs = f"{(maxwin.end - maxwin.start).days / 365.25:.2f}".replace(".", ",")
        note += f" «{maxwin.label}» — максимальное окно данных {maxwin.start:%d.%m.%Y} – {end:%d.%m.%Y} ({yrs} г)."
    note += " Синий — фонд обогнал бенчмарк, красный — отстал, «—» — период не покрыт данными."
    if source:
        note += f" Источник: {source}."
    c = ws.cell(last_row + 2, 1, note)
    c.font = Font(name=ARIAL, sz=9, color=GREY)

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def make_alpha_table(data: pd.DataFrame, out: str | Path | None = None, *, levels: bool = False,
                     source: str | None = None) -> Path:
    """DataFrame выгрузки → xlsx с таблицей «Альфа по фондам и периодам». Возвращает путь к файлу.

    data   — DataFrame как all_funds_perform: даты (колонка moment или индекс) + накопленные доходности рядов;
    out    — куда сохранить; по умолчанию alpha_by_fund_<ДДММГГГГ последней даты>.xlsx в текущей папке;
    levels — True, если в data уже уровни (1 + доходность), а не доходности;
    source — подпись источника в строке-пояснении под таблицей (необязательно).
    """
    if not isinstance(data, pd.DataFrame):
        raise TypeError(f"Ожидался pandas.DataFrame, получен {type(data).__name__}")
    table, pdefs = alpha_dataframe(data, levels=levels)
    if out is None:
        out = Path.cwd() / f"alpha_by_fund_{pdefs[0].end:%d%m%Y}.xlsx"
    return _save_table(table, pdefs, Path(out), source)


# ============================================================================ запуск из терминала
def _read_excel(path: Path) -> tuple[pd.DataFrame, bool, str]:
    """Файл → (DataFrame, levels, подпись источника): выгрузка или лист «Данные» готового отчёта."""
    xl = pd.ExcelFile(path)
    if "Данные" in xl.sheet_names:
        df = xl.parse("Данные")
        keep = [c for c in df.columns if isinstance(c, str) and not c.startswith("Unnamed") and len(c) < 40]
        return df[keep], True, f"лист «Данные» файла {path.name}"
    return xl.parse(xl.sheet_names[0]), False, f"выгрузка {path.name}"


def main(argv: list[str]) -> Path:
    if not argv:
        raise SystemExit("Запуск: python ./alpha_table.py 'path_to_excel' [out.xlsx]")
    src = Path(argv[0]).expanduser().resolve()
    if not src.exists():
        raise SystemExit(f"Нет файла {src}")
    data, levels, source = _read_excel(src)
    if len(argv) > 1:
        out = Path(argv[1])
    else:
        m = re.search(r"(\d{8})", src.stem)
        out = src.with_name(f"alpha_by_fund_{m.group(1)}.xlsx") if m else None
    res = make_alpha_table(data, out, levels=levels, source=source)
    print(f"Готово: {res}")
    return res


if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except ValueError as e:
        raise SystemExit(f"Ошибка: {e}")
