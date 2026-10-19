import calc_series_change
import pandas as pd
import pyodbc
from datetime import datetime, timedelta
import numpy as np
import psycopg2
import moex_cls as mcls
import alpha_table
import sys


def compose_otchet(inp_period):
    user = "tmp_skosyrkov"
    password = "A65h9bmstu12sg9!"
    start_date = datetime.today() - timedelta(days=365 * 2 - 3)
    finish_date = datetime.today() - timedelta(days=inp_period)
    con = pyodbc.connect(
        "driver={SQL Server};"
        "server=AMDBRATES\DBMVTBAM01;"
        "database=Rates;"
        "trusted_connection=yes"
    )
    con_dl = psycopg2.connect(
        host="SMSK01DB749",
        port="15432",
        database="WIM_datalake_prod",
        user=user,
        password=password,
    )

    date_str = start_date.strftime("%d.%m.%Y")

    funds = pd.read_sql(
        f"""
        SELECT DISTINCT period, 
        naimenovanie,
        (CASE WHEN nams.naimenovanie = 'ОПИФРФИ "ВИМ - Накопительный резерв"' and costs.period <= TO_DATE('13.08.2026', 'DD.MM.YYYY') THEN costs.stoimost_paya / 100 ELSE costs.stoimost_paya END) as stoimost_paya
        FROM pif.stoimost_chistykh_aktivov costs
        INNER JOIN pif.pif nams ON costs.pif = nams.ssylka
        WHERE costs.period >= TO_DATE('{date_str}', 'DD.MM.YYYY')
        ORDER BY period DESC
        """,
        con=con_dl,
    )

    isin_and_bench = [
        ('БПИФРФИ "ВИМ - Индекс МосБиржи"', "MCF2TR", "1 | "),
        ('БПИФРФИ "ВИМ - Устойчивое развитие российских компаний"', "MRSVRT", "2 | "),
        ('БПИФРФИ "Ликвидность"', "RUSFAR", "3 | "),
        ('БПИФРФИ "ВИМ - Корпоративные облигации"', "RUCBTRNS", "4 | "),
        ('ИПИФФИ "2х Акции"', "VTBST7", "5 | "),
        ('ИПИФФИ "2х ОФЗ"', "VTBST5", "6 | "),
        ('ИПИФФИ "Валютное плечо"', "VTBST6", "7 | "),
        ('ОПИФРФИ "ВИМ - Акции"', "MCFTR", "8 | "),
        ('ОПИФРФИ "ВИМ - Акции"', "MCFWTR", "9 | "),
        ('ОПИФРФИ "ВИМ - Акции. Российские эмитенты"', "MCFTR", "10 | "),
        ('ОПИФРФИ "ВИМ - Акции. Российские эмитенты"', "MCFWTR", "11 | "),
        ('ОПИФРФИ "ВИМ - Накопительный резерв"', "VTBFG0010", "12 | "),
        ('ОПИФРФИ "ВИМ - Металлургия"', "MOEXMM", "13 | "),
        ('ОПИФРФИ "ВИМ - Нефтегазовый сектор"', "MOEXOG", "14 | "),
        ('ОПИФРФИ "ВИМ - Перспективные размещения"', "MIPOTR", "15 | "),
        ('ОПИФРФИ "ВИМ - Инновационный"', "MOEXIT", "16 | "),
        ('ОПИФРФИ "ВИМ - Денежный рынок. Рубли"', "RUSFAR", "17 | "),
        ('ОПИФРФИ "ВИМ - Сбалансированный"', "VTBFG0006", "18 | "),
        ('ОПИФРФИ "ВИМ - Сбалансированный. Консервативный"', "VTBFG0011", "19 | "),
        ('ОПИФРФИ "ВИМ - Сбалансированный. Российские эмитенты"', "VTBFG0005", "20 | "),
        ('ОПИФРФИ "ВИМ - Сбалансированные инвестиции"', "VTBFG0005", "21 | "),
        ('ОПИФРФИ "ВИМ - Ключевой+"', "RUFLCBTR", "22 | "),
        ('ОПИФРФИ "ВИМ - Облигации. Российские эмитенты"', "RUCBTRNS", "23 | "),
        ('ОПИФРФИ "ВИМ - Казначейский"', "RUCBTRNS", "24 |"),
        ('ОПИФРФИ "ВИМ - Облигации. Ответственные инвестиции"', "RUCBTRNS", "25 | "),
        ('ОПИФРФИ "ВИМ - Умный капитал"', "LQDT", "26 | "),
        ('ОПИФРФИ "ВИМ - Накопительный в юанях"', "RUCNYTR", "27 | "),
        ('ОПИФРФИ "ВИМ - Новые валютные облигации"', "RUEU10", "28 | "),
        ('ОПИФРФИ "Яндекс Пэй. Фонд рублёвых облигаций"', "RUCBTRNS", "29 | "),
        ('ОПИФРФИ "Яндекс Пэй. Фонд акций и облигаций"', "VTBST3", "30 | "),
        ('ОПИФРФИ "Яндекс Пэй. Фонд золото"', "VTBST4", "31 | "),
    ]

    all_funds_perform = pd.DataFrame()

    # --- ДОБАВЛЯЕМ ОЧИСТКУ ДАННЫХ ПЕРЕД ЦИКЛОМ ---
    funds = funds.dropna(subset=["period", "stoimost_paya"])
    funds = funds.drop_duplicates(subset=["naimenovanie", "period"], keep="last")
    # funds["isin"] = funds["isin"].str.extract(r"(RU[0-9A-Z]{10})", expand=False)
    # ---------------------------------------------

    def fetch_micex_index(security_id, start, finish):
        df = pd.read_sql(
            f""" 
            select TRADEDATE, [CLOSE] as [{security_id}] from rates.dbo.dbo_MICEX_Index
            where securityid = ? and tradedate >= CONVERT(date, ?, 104) and tradedate <= CONVERT(date, ?, 104)
            order by TRADEDATE
            """,
            con=con,
            params=[security_id, start, finish],
        )
        return df.set_index("TRADEDATE").squeeze()

    def fetch_fund_index(security_id, start, finish):
        df = pd.read_sql(
            """
                select period, stoimost_paya as temp_column_name 
                from pif.stoimost_chistykh_aktivov costs
                INNER JOIN pif.pif nams ON costs.pif = nams.ssylka
                where naimenovanie = %s 
                and period >= %s
                and period <= %s
                order by period
            """,
            # and period >= TO_DATE(%s, 'DD.MM.YYYY')
            # and period <= TO_DATE(%s, 'DD.MM.YYYY')
            con=con_dl,
            params=[security_id, start, finish],
        )
        # Переименовываем колонку в самом pandas, здесь кавычки ничего не ломают
        df = df.rename(columns={"temp_column_name": security_id})
        return df.set_index("period").squeeze()

    for one, one_bench, number in isin_and_bench:
        # 1. Фильтруем данные
        sub_df = funds[funds["naimenovanie"] == one]

        # 2. ПРОВЕРКА: если данных для этого ISIN нет, пропускаем
        if sub_df.empty:
            print(
                f"Предупреждение: для {number + one} не найдено данных в датафрейме funds. Пропускаем."
            )
            continue

        # Сортируем строго по дате перед расчетом
        sub_df = sub_df.sort_values("period")

        # 3. Если данные есть, спокойно делаем расчет
        one_series = calc_series_change.value_growth(
            sub_df.set_index("period")["stoimost_paya"]
        )
        all_funds_perform[number + one] = one_series

        # Границы дат для текущего фонда
        start_dt = all_funds_perform[number + one].first_valid_index()
        finish_dt = all_funds_perform[number + one].last_valid_index()

        # теперь грузим бенч (сразу рядом). match используется т.к. есть кастомные индексы и я их считаю на лету
        try:
            match one_bench:

                case "VTBFG0005":
                    b_df_1 = fetch_micex_index("MCFTR", start_dt, finish_dt)
                    b_df_2 = fetch_micex_index("RUCBTRNS", start_dt, finish_dt)
                    all_funds_perform[f"50%MCFTR + 50%RUCBTRNS, {number + one}"] = (
                        mcls.quarterly_index(
                            pairs=[(b_df_1, 0.5), (b_df_2, 0.5)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                            cumulative_yield=True,
                        )
                        .reset_index()[["date", one_bench]]
                        .set_index("date")[one_bench]
                    )  # функция возвращает большой датафрейм (
                    # полезно для проверки). Я переделываю датафрейм в series чтобы можно было легко вязать все остальное

                case "VTBFG0006":
                    b_df_1 = fetch_micex_index("MCFTR", start_dt, finish_dt)
                    b_df_2 = fetch_micex_index("RUCBTRNS", start_dt, finish_dt)
                    all_funds_perform[f"65%MCFTR + 35%RUCBTRNS, {number + one}"] = (
                        mcls.quarterly_index(
                            pairs=[(b_df_1, 0.65), (b_df_2, 0.35)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                            cumulative_yield=True,
                        )
                        .reset_index()[["date", one_bench]]
                        .set_index("date")[one_bench]
                    )  # функция возвращает большой датафрейм (
                    # полезно для проверки). Я переделываю датафрейм в series чтобы можно было легко вязать все остальное

                case "VTBFG0011":
                    b_df_1 = fetch_micex_index("MCFTR", start_dt, finish_dt)
                    b_df_2 = fetch_micex_index("RUCBTRNS", start_dt, finish_dt)
                    all_funds_perform[f"20%MCFTR + 80%RUCBTRNS, {number + one}"] = (
                        mcls.quarterly_index(
                            pairs=[(b_df_1, 0.2), (b_df_2, 0.8)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                            cumulative_yield=True,
                        )
                        .reset_index()[["date", one_bench]]
                        .set_index("date")[one_bench]
                    )  # функция возвращает большой датафрейм (
                    # полезно для проверки). Я переделываю датафрейм в series чтобы можно было легко вязать все остальное

                case "VTBFG0010":
                    b_df_1 = fetch_micex_index("MCFTR", start_dt, finish_dt)
                    b_df_2 = fetch_micex_index("RUCBTRNS", start_dt, finish_dt)
                    all_funds_perform[f"40%MCFTR + 60%RUCBTRNS, {number + one}"] = (
                        mcls.quarterly_index(
                            pairs=[(b_df_1, 0.4), (b_df_2, 0.6)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                            cumulative_yield=True,
                        )
                        .reset_index()[["date", one_bench]]
                        .set_index("date")[one_bench]
                    )  # функция возвращает большой датафрейм (
                    # полезно для проверки). Я переделываю датафрейм в series чтобы можно было легко вязать все остальное

                case "VTBFG0017":
                    b_df_1 = fetch_micex_index("IMOEX", start_dt, finish_dt)
                    b_df_2 = fetch_micex_index("RUCBTRNS", start_dt, finish_dt)
                    all_funds_perform[f"50%IMOEX + 50%RUCBTRNS, {number + one}"] = (
                        mcls.quarterly_index(
                            pairs=[(b_df_1, 0.5), (b_df_2, 0.5)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                            cumulative_yield=True,
                        )
                        .reset_index()[["date", one_bench]]
                        .set_index("date")[one_bench]
                    )  # функция возвращает большой датафрейм (
                    # полезно для проверки). Я переделываю датафрейм в series чтобы можно было легко вязать все остальное

                case "LQDT":
                    b_df = fetch_fund_index(
                        f'БПИФРФИ "Ликвидность"', start_dt, finish_dt
                    )
                    all_funds_perform[f'БПИФРФИ "Ликвидность", {number + one}'] = (
                        mcls.series_index(
                            pairs=[(b_df, 1)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                        )
                    )

                case "VTBST3":
                    b_df_1 = fetch_micex_index("MCFTR", start_dt, finish_dt)
                    b_df_2 = fetch_micex_index("RUCBTRNS", start_dt, finish_dt)
                    all_funds_perform[f"40%MCFTR+60%RUCBTRNS, {number + one}"] = (
                        mcls.quarterly_index(
                            pairs=[(b_df_1, 0.4), (b_df_2, 0.6)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                            cumulative_yield=True,
                        )
                        .reset_index()[["date", one_bench]]
                        .set_index("date")[one_bench]
                    )  # функция возвращает большой датафрейм (
                    # полезно для проверки). Я переделываю датафрейм в series чтобы можно было легко вязать все остальное

                case "VTBST4":
                    b_df = fetch_fund_index(
                        'БПИФРФИ "ВИМ - Золото"', start_dt, finish_dt
                    )
                    all_funds_perform[f"БПИФРФИ ВИМ - Золото, {number + one}"] = (
                        mcls.series_index(
                            pairs=[(b_df, 1)],
                            start_date=start_dt,
                            finish_date=finish_dt,
                            name=one_bench,
                        )
                    )

                case (
                    "VTBST5"
                ):  # сложный индекс, считается как 200% r(RGBITR) - 100% r(LQDT), типо LQDT = стоиомсть фондирования. daily reweight
                    b_df_1 = fetch_micex_index("RGBITR", start_dt, finish_dt)
                    b_df_2 = fetch_fund_index(
                        'БПИФРФИ "Ликвидность"', start_dt, finish_dt
                    )
                    # считаем накопленную доходность каждого timeseries, далее вычитаем из одного другое
                    b_df_1 = mcls.series_index(
                        pairs=[(b_df_1, 1)],
                        start_date=start_dt,
                        finish_date=finish_dt,
                        name=one_bench + f"RGBITR",
                    )
                    b_df_2 = mcls.series_index(
                        pairs=[(b_df_2, 1)],
                        start_date=start_dt,
                        finish_date=finish_dt,
                        name=one_bench + f"LQDT",
                    )
                    common_index = b_df_1.index.intersection(
                        b_df_2.index
                    )  # ищем пересечения и считаем доху чисто по ним. логика не рушится, т.к. дохи считаются относительно первого дня выборок
                    result = (
                        2 * b_df_1.loc[common_index] - b_df_2.loc[common_index]
                    )  # считаем по формуле 200% r(RGBITR) - 100% r(LQDT)
                    all_funds_perform[f"2RGBITR-LQDT, {number + one}"] = result

                case "VTBST6":  # сам фонд отдает значения в рублях
                    b_df_1 = fetch_micex_index(
                        "RUEU10", start_dt, finish_dt
                    )  # оно в баксах
                    b_df_2 = fetch_fund_index(
                        'БПИФРФИ "ВИМ - Ликвидность. Юань"', start_dt, finish_dt
                    )  # оно в рублях
                    # тащим курсы usdrub
                    b_df_rubusd = pd.read_sql(
                        f"""
                        select date as tradedate, rate as USD from rates.dbo.dbo_usd
                        where ticket='usd' and date >= CONVERT(date, ?, 104) and date <=CONVERT(date, ?, 104)
                        """,
                        con=con,
                        params=[start_dt, finish_dt],
                    )  # не определяется по понедельникам, надо сделать полный по датам датафрейм и дальше протягивать курс с сб по пн:
                    b_df_rubusd["tradedate"] = pd.to_datetime(b_df_rubusd["tradedate"])
                    b_df_rubusd = b_df_rubusd.set_index("tradedate")
                    b_df_full_usd = b_df_rubusd.asfreq("D")
                    b_df_full_usd["USD"] = b_df_full_usd["USD"].ffill()
                    b_df_usd = b_df_full_usd["USD"]
                    # делим rueu10 на курс чтобы учесть динамику
                    common_index = b_df_1.index.intersection(
                        b_df_usd.index
                    )  # ищем пересечения и считаем доху чисто по ним. логика не рушится, т.к. дохи считаются относительно первого дня выборок
                    result = b_df_1.loc[common_index] * b_df_usd.loc[common_index]
                    # считаем накопленную доходность каждого timeseries, далее вычитаем из одного другое
                    b_df_1 = mcls.series_index(
                        pairs=[(result, 1)],
                        start_date=start_dt,
                        finish_date=finish_dt,
                        name=one_bench + f"RUEU10",
                    )
                    b_df_2 = mcls.series_index(
                        pairs=[(b_df_2, 1)],
                        start_date=start_dt,
                        finish_date=finish_dt,
                        name=one_bench + f"CNYM",
                    )
                    common_index = b_df_1.index.intersection(
                        b_df_2.index
                    )  # ищем пересечения и считаем доху чисто по ним. логика не рушится, т.к. дохи считаются относительно первого дня выборок
                    result = (
                        2 * b_df_1.loc[common_index] - b_df_2.loc[common_index]
                    )  # считаем по формуле 200% r(RUEU10) - 100% r(CNYM)
                    all_funds_perform[f"2RUEU10-CNYM, {number + one}"] = result

                case (
                    "VTBST7"
                ):  # сложный индекс, считается как 200% r(MEBCTR) - 100% r(RUSFAR)
                    b_df_1 = fetch_micex_index("MEBCTR", start_dt, finish_dt)
                    # b_df_1 = mcls.series_index(
                    #     pairs=[(b_df_1, 1)],
                    #     start_date=start_dt,
                    #     finish_date=finish_dt,
                    #     name=one_bench + f"MEBCTR",
                    # )
                    # тяжелый случай. отдельную функцию не буду писать - сори мне лень
                    b_df_2 = fetch_micex_index("RUSFAR", start_dt, finish_dt)
                    b_df_2 = (
                        b_df_2 + 0.3
                    )  # ЭТО ДОБАВЛЕНИЕ НЕКОТОРОГО СПРЕДА К РУСФАРУ (ПРЕМИЯ ЗА ФОНДИРОВАНИЕ)
                    # нахожу совпадения по дням
                    common_index = b_df_1.index.intersection(b_df_2.index)
                    diff_in_days = common_index.index.to_series().diff().dt.days

                    b_df_2 = 1 / 100 * b_df_2 * diff_in_days / 365

                    # 1. Находим общие даты и сортируем
                    common_index = b_df_1.index.intersection(b_df_2.index).sort_values()

                    # 2. Выравниваем обе Series по общему индексу
                    idx_one = b_df_1.loc[common_index]
                    idx_two = b_df_2.loc[common_index]

                    # 3. Считаем разницу в днях (исправленная строка)
                    diff_in_days = common_index.to_series().diff().dt.days

                    # 4. Создаем массив под результат и задаем начальное значение
                    composite_values = np.zeros(len(common_index))
                    composite_values = 100.0  # <-- Укажите здесь ваше стартовое значение для первой даты

                    # 5. Считаем индекс в цикле
                    for t in range(1, len(common_index)):
                        # Формула: 2 * (index_one(t)/index_one(t-1) - index_two(t)/100 * diff_days/365) * composite(t-1)
                        ratio_one = idx_one.iloc[t] / idx_one.iloc[t - 1]
                        rate_part = (idx_two.iloc[t] / 100) * (
                            diff_in_days.iloc[t] / 365
                        )
                        comp_prev = composite_values[t - 1]

                        composite_values[t] = (2 * ratio_one - rate_part) * comp_prev

                    # 6. Переводим в итоговый Series
                    composite_index = pd.Series(composite_values, index=common_index)
                    all_funds_perform[f"2MEBCTR-RUSFAR, {number + one}"] = (
                        composite_index
                    )

                case _:
                    # Для стандартных бенчмарков, которые загружаются напрямую целиком
                    one_bench_df = fetch_micex_index(one_bench, start_dt, finish_dt)
                    if one_bench == "RUSFAR":
                        all_funds_perform[one_bench + ", " + number + one] = (
                            mcls.series_index(
                                pairs=[(one_bench_df, 1)],
                                start_date=start_dt,
                                finish_date=finish_dt,
                                name=one_bench,
                                cumulative=True,
                            )
                        )
                    elif one_bench == "RUSFARCNY":
                        all_funds_perform[one_bench + ", " + number + one] = (
                            mcls.series_index(
                                pairs=[(one_bench_df, 1)],
                                start_date=start_dt,
                                finish_date=finish_dt,
                                name=one_bench,
                                cumulative=True,
                            )
                        )
                    else:
                        all_funds_perform[one_bench + ", " + number + one] = (
                            mcls.series_index(
                                pairs=[(one_bench_df, 1)],
                                start_date=start_dt,
                                finish_date=finish_dt,
                                name=one_bench,
                            )
                        )

        except Exception as e:
            print(f"Ошибка при обработке бенчмарка {one_bench} для фонда {one}: {e}")
    return all_funds_perform


if __name__ == "__main__":
    delta = int(sys.argv[1])
    x = compose_otchet(delta)
    x.to_excel(
        f'data_for_{(datetime.today() - timedelta(delta)).strftime("%d%m%Y")}.xlsx'
    )
    alpha_table.make_alpha_table(
        x,
        f'fund_vs_bench_{(datetime.today() - timedelta(delta)).strftime("%d%m%Y")}.xlsx',
    )
