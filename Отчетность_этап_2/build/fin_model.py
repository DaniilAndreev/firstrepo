# -*- coding: utf-8 -*-
"""Финансовая модель бизнес-плана (базовый сценарий 2027–2029), тыс. руб.

Все показатели считаются из допущений ниже; суммы получены сложением округленных до 0,1 слагаемых,
поэтому арифметика в таблицах сходится.
"""
Q = ["I кв.", "II кв.", "III кв.", "IV кв."]

PRICE = {"12": 1.100, "30": 4.200}          # средняя цена реализации, тыс. руб. за изделие
VAR = {"12": 0.374, "30": 0.961}            # переменные затраты на изделие (сырье, упаковка, сдельный труд, энергия)
VOL = {
    2027: {"12": [100, 300, 800, 500], "30": [10, 30, 60, 50]},
    2028: {"12": [400, 800, 1600, 1000], "30": [30, 70, 130, 100]},
    2029: {"12": [7000], "30": [600]},
}
FIXED = {   # постоянные затраты по кварталам (2029 — год)
    2027: {
        "Аренда производственной площадки": [60, 60, 60, 60],
        "Оплата специалистов по договорам ГПХ": [90, 90, 90, 90],
        "Маркетинг и реклама": [50, 75, 100, 75],
        "Оборудование для серийного изготовления": [450, 0, 0, 0],
        "Декларирование и лабораторные испытания": [60, 0, 0, 0],
        "Хостинг, домены, внешние сервисы": [10, 10, 10, 10],
        "Бухгалтерское и юридическое сопровождение": [15, 15, 15, 15],
    },
    2028: {
        "Аренда производственной площадки": [90, 90, 90, 90],
        "Оплата специалистов по договорам ГПХ": [180, 180, 180, 180],
        "Маркетинг и реклама": [100, 125, 175, 100],
        "Оборудование для серийного изготовления": [200, 0, 0, 0],
        "Декларирование и лабораторные испытания": [0, 0, 0, 0],
        "Хостинг, домены, внешние сервисы": [15, 15, 15, 15],
        "Бухгалтерское и юридическое сопровождение": [20, 20, 20, 20],
    },
    2029: {
        "Аренда производственной площадки": [480],
        "Оплата специалистов по договорам ГПХ": [1200],
        "Маркетинг и реклама": [800],
        "Оборудование для серийного изготовления": [300],
        "Декларирование и лабораторные испытания": [0],
        "Хостинг, домены, внешние сервисы": [100],
        "Бухгалтерское и юридическое сопровождение": [120],
    },
}
TAX_RATE, TAX_MIN = 0.20, 0.03              # АУСН «доходы минус расходы»: 20 %, минимальный налог 3 % доходов
LOAN = {"amount": 700.0, "in": (2027, 0), "out": (2028, 1)}   # заем учредителя: поступление I кв. 2027, возврат II кв. 2028
OPENING_CASH_2027 = 0.0
GRANT = 1000.0


def r1(x):
    return round(x + 1e-9, 1)


def year(y):
    v = VOL[y]
    n = len(v["12"])
    rev12 = [r1(v["12"][i] * PRICE["12"]) for i in range(n)]
    rev30 = [r1(v["30"][i] * PRICE["30"]) for i in range(n)]
    income = [r1(rev12[i] + rev30[i]) for i in range(n)]
    var = [r1(v["12"][i] * VAR["12"] + v["30"][i] * VAR["30"]) for i in range(n)]
    fixed = {k: [float(x) for x in vals] for k, vals in FIXED[y].items()}
    pre_tax_exp = [r1(var[i] + sum(f[i] for f in fixed.values())) for i in range(n)]
    inc_y = r1(sum(income))
    exp_y = r1(sum(pre_tax_exp))
    tax_y = r1(max(TAX_RATE * (inc_y - exp_y), TAX_MIN * inc_y))
    tax = [r1(tax_y * income[i] / inc_y) for i in range(n)]
    tax[-1] = r1(tax_y - sum(tax[:-1]))
    total_exp = [r1(pre_tax_exp[i] + tax[i]) for i in range(n)]
    result = [r1(income[i] - total_exp[i]) for i in range(n)]
    return dict(n=n, vol12=v["12"], vol30=v["30"], rev12=rev12, rev30=rev30, income=income, var=var,
                fixed=fixed, tax=tax, total_exp=total_exp, result=result, tax_y=tax_y,
                inc_y=inc_y, exp_y=r1(sum(total_exp)), res_y=r1(sum(result)), pre_tax_y=exp_y)


Y = {y: year(y) for y in (2027, 2028, 2029)}


def cash(y, opening):
    d = Y[y]
    n = d["n"]
    loan_in = [0.0] * n
    loan_out = [0.0] * n
    if LOAN["in"][0] == y:
        loan_in[LOAN["in"][1] if n > 1 else 0] = LOAN["amount"]
    if LOAN["out"][0] == y:
        loan_out[LOAN["out"][1] if n > 1 else 0] = LOAN["amount"]
    rows = []
    bal = opening
    for i in range(n):
        start = bal
        inflow = r1(d["income"][i] + loan_in[i])
        outflow = r1(d["total_exp"][i] + loan_out[i])
        saldo = r1(inflow - outflow)
        bal = r1(start + saldo)
        rows.append(dict(start=start, sales=d["income"][i], loan_in=loan_in[i], inflow=inflow,
                         ops=d["total_exp"][i], loan_out=loan_out[i], outflow=outflow, saldo=saldo, end=bal))
    return rows


CF = {}
_open = OPENING_CASH_2027
for _y in (2027, 2028, 2029):
    CF[_y] = cash(_y, _open)
    _open = CF[_y][-1]["end"]


def year_cf(y):
    rows = CF[y]
    return dict(start=rows[0]["start"], sales=r1(sum(r["sales"] for r in rows)),
                loan_in=r1(sum(r["loan_in"] for r in rows)), inflow=r1(sum(r["inflow"] for r in rows)),
                ops=r1(sum(r["ops"] for r in rows)), loan_out=r1(sum(r["loan_out"] for r in rows)),
                outflow=r1(sum(r["outflow"] for r in rows)), saldo=r1(sum(r["saldo"] for r in rows)),
                end=rows[-1]["end"])


def indicators():
    rev = [Y[y]["inc_y"] for y in (2027, 2028, 2029)]
    res = [Y[y]["res_y"] for y in (2027, 2028, 2029)]
    cagr = (rev[2] / rev[0]) ** 0.5 - 1
    margin = [res[i] / rev[i] for i in range(3)]
    d = Y[2027]
    fixed_y = r1(sum(sum(v) for v in d["fixed"].values()))
    var_y = r1(sum(d["var"]))
    cm_ratio = (d["inc_y"] - var_y) / d["inc_y"]
    breakeven = fixed_y / (cm_ratio - TAX_MIN)
    # окупаемость средств гранта: накопленный финансовый результат достигает суммы гранта
    cum, payback = 0.0, None
    for y in (2027, 2028):
        for i, x in enumerate(Y[y]["result"]):
            cum = r1(cum + x)
            if payback is None and cum >= GRANT:
                payback = (y, i)
    if payback is None:
        cum29 = cum + Y[2029]["res_y"]
        payback = (2029, None) if cum29 >= GRANT else None
    return dict(rev=rev, res=res, cagr=cagr, margin=margin, breakeven=r1(breakeven),
                breakeven_share=breakeven / d["inc_y"], payback=payback, fixed_2027=fixed_y, var_2027=var_y,
                cum=[r1(sum(res[:i + 1])) for i in range(3)])


IND = indicators()

if __name__ == "__main__":
    for y in (2027, 2028, 2029):
        d = Y[y]
        print(y, "доходы", d["income"], "=", d["inc_y"], "| расходы", d["total_exp"], "=", d["exp_y"],
              "| налог", d["tax"], "| результат", d["result"], "=", d["res_y"])
        print("   CF:", [(r["start"], r["inflow"], r["outflow"], r["end"]) for r in CF[y]])
    print(IND)
