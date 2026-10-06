#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Таблица «Альфа по фондам и периодам» с листа «Сводка» — отдельным xlsx.

    python ./alpha_table.py 'path_to_excel' [out.xlsx]

path_to_excel — выгрузка all_funds_perform_ДДММГГГГ.xlsx (или готовый отчёт
funds_vs_benchmarks_*.xlsx). Результат — alpha_by_fund_ДДММГГГГ.xlsx рядом с исходным
файлом: один лист «Сводка», на нём только таблица (шапка в A1, 19 фондов × периоды
1 мес … 5 лет, «Периодов с альфой > 0», замечание) и строка-пояснение под ней.

Методика расчёта альфы та же, что в полном отчёте: period_defs и fund_metrics берутся из
build_funds_vs_benchmarks.py, он должен лежать в той же папке. Валюта не учитывается:
курс CNY/RUB не нужен, «Ликвидность. Юань» в таблицу не входит.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_funds_vs_benchmarks as fvb  # noqa: E402  (PERIODS, period_defs, fund_metrics)

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
# дробления паёв: (ряд, дата, множитель) — берутся из основного скрипта
SPLITS = getattr(fvb, "SPLITS", [("RU000A0JT4S1", "2026-08-14", 100.0)])
COLS = ["Фонд", "Тип", "Бенчмарк"] + [p for p, _ in fvb.PERIODS] + ["Периодов с альфой > 0", "Замечание"]
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


# ----------------------------------------------------------------------------- данные
def _report_data(path: Path) -> pd.DataFrame | None:
    """Лист «Данные» готового отчёта (уровни рядов, уже с поправкой на дробление) или None."""
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return None
    if "Данные" not in xl.sheet_names:
        return None
    df = xl.parse("Данные")
    df = df.rename(columns={df.columns[0]: "Дата"}).dropna(subset=["Дата"]).set_index("Дата")
    df.index = pd.DatetimeIndex(df.index)
    keep = [c for c in df.columns if isinstance(c, str) and not c.startswith("Unnamed") and len(c) < 40]
    return df[keep].apply(pd.to_numeric, errors="coerce").sort_index()


def load_levels(path: Path) -> tuple[pd.DataFrame, str]:
    """Уровни рядов (1 + накопленная доходность): из выгрузки или с листа «Данные» отчёта."""
    rep = _report_data(path)
    if rep is not None:
        return rep, f"лист «Данные» файла {path.name}"
    raw = pd.read_excel(path)
    raw = raw.rename(columns={raw.columns[0]: "Дата"}).dropna(subset=["Дата"]).set_index("Дата").sort_index()
    raw.index = pd.DatetimeIndex(raw.index)
    lv = 1.0 + raw.apply(pd.to_numeric, errors="coerce")
    for code, d, k in SPLITS:
        d = pd.Timestamp(d)
        if code in lv and lv.index.min() < d <= lv.index.max():
            lv.loc[d:, code] *= k
    return lv, f"выгрузка {path.name}"


# ----------------------------------------------------------------------------- таблица
def alpha_table(lv: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """DataFrame таблицы «Альфа по фондам и периодам» (альфа = прирост фонда − прирост бенчмарка)."""
    missing = sorted({c for _, _, f, b, _ in FUNDS for c in (f, b)} - set(lv.columns))
    if missing:
        raise SystemExit(f"В данных нет рядов: {missing}")
    pdefs = fvb.period_defs(lv.index)
    rows = []
    for name, ftype, fcode, bcode, note in FUNDS:
        alphas = []
        for p in pdefs:
            m = fvb.fund_metrics(lv, fcode, bcode, p)
            alphas.append(m.G if m.status in ("ok", "maxwin") else DASH)
        nums = [a for a in alphas if a != DASH]
        rows.append([name, ftype, bcode, *alphas, f"{sum(a > 0 for a in nums)} из {len(nums)}", note])
    return pd.DataFrame(rows, columns=COLS), pdefs


def save_table(df: pd.DataFrame, pdefs: list, out: Path, source: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    ws.sheet_view.showGridLines = False
    n_per = len(fvb.PERIODS)
    first_p, last_p = 4, 3 + n_per                       # D … L
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
    note += f" Синий — фонд обогнал бенчмарк, красный — отстал, «—» — период не покрыт данными. Источник: {source}."
    c = ws.cell(last_row + 2, 1, note)
    c.font = Font(name=ARIAL, sz=9, color=GREY)

    wb.save(out)
    return out


# ----------------------------------------------------------------------------- точка входа
def make_alpha_table(path_to_excel: str, out: str | None = None) -> Path:
    """Строит xlsx с таблицей «Альфа по фондам и периодам» и возвращает путь к нему."""
    src = Path(path_to_excel).expanduser().resolve()
    if not src.exists():
        raise SystemExit(f"Нет файла {src}")
    lv, source = load_levels(src)
    df, pdefs = alpha_table(lv)
    if out is None:
        m = re.search(r"(\d{8})", src.stem)
        tag = m.group(1) if m else f"{lv.index.max():%d%m%Y}"
        out = src.with_name(f"alpha_by_fund_{tag}.xlsx")
    return save_table(df, pdefs, Path(out), source)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("Запуск: python ./alpha_table.py 'path_to_excel' [out.xlsx]")
    res = make_alpha_table(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(f"Готово: {res}")
