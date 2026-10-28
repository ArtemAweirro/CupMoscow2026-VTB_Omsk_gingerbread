#!/usr/bin/env python3
"""Графики для слайдов 3, 6, 7 (PNG, 220 dpi) и CSV с данными графиков.

Источники: out/sim_*.csv (симуляция, синтетические данные) и out/model_values.json
(значения Excel-модели после model/check_model.py).
Запуск: .venv/bin/python sim/charts.py
"""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "out"
DPI = 220

# Палитра (проверена валидатором: CVD и normal-vision PASS; у двух светлых цветов контраст < 3:1,
# поэтому везде прямые подписи). Цвет закреплен за вариантом на всех графиках.
COLOR = {"В3": "#2a78d6", "В1": "#eb6834", "В2": "#1baf7a", "В4": "#eda100"}
NAME = {"В1": "В1 Правила + усил. аутентификация", "В2": "В2 ML-скоринг + сегментация",
        "В3": "В3 Адаптивное трение", "В4": "В4 Внешние участники"}
ORDER = ["В1", "В2", "В3", "В4"]
SHORT = {"В1": "В1 Правила", "В2": "В2 ML-скоринг", "В3": "В3 Адаптивное трение", "В4": "В4 Внешние участники"}
OP_LABEL = {"В1": (8, 6, "left"), "В2": (-8, 4, "right"), "В3": (-8, 0, "right"), "В4": (8, -8, "left")}
GAIN, COST, TOTAL = "#2a78d6", "#eb6834", "#52514e"
SCEN = {"p": ("Пессимист", "#86b6ef"), "b": ("База", "#2a78d6"), "o": ("Оптимист", "#104281")}
TEXT, TEXT2, MUTED, GRID, SURFACE = "#0b0b0b", "#52514e", "#8a8984", "#e6e5e0", "#ffffff"

SYNTH = "Синтетические данные: иллюстрация механизма, не прогноз."
MODEL = ("Расчет команды на допущениях (гипотезы помечены в Excel): иллюстрация механизма, не прогноз. "
         "Данные рынка: Банк России, II кв. 2026.")

plt.rcParams.update({
    "font.family": "Arial", "font.size": 11, "text.color": TEXT,
    "axes.edgecolor": MUTED, "axes.labelcolor": TEXT2, "axes.titlecolor": TEXT,
    "axes.spines.top": False, "axes.spines.right": False, "axes.linewidth": 0.8,
    "xtick.color": TEXT2, "ytick.color": TEXT2, "axes.grid": True, "grid.color": GRID,
    "grid.linewidth": 0.6, "axes.axisbelow": True,
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
})


def ru(x, d=1):
    """Число по-русски: пробел в тысячах, запятая в дробной части."""
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def pct_axis(ax, axis="y", d=0):
    fmt = matplotlib.ticker.FuncFormatter(lambda v, _: ru(v, d) + "%")
    (ax.yaxis if axis == "y" else ax.xaxis).set_major_formatter(fmt)


def finish(fig, name, note, title, subtitle=None):
    fig.suptitle(title, x=0.01, ha="left", fontsize=15, fontweight="bold", y=0.985)
    if subtitle:
        fig.text(0.01, 0.915, subtitle, fontsize=10.5, color=TEXT2, ha="left")
    fig.text(0.01, 0.012, note, fontsize=8.5, color=MUTED, ha="left")
    fig.savefig(OUT / name, dpi=DPI)
    plt.close(fig)
    print("saved", OUT / name)


def envelope(g):
    """Нижняя огибающая: для каждой доли ограничений лучшие найденные потери (недоминируемые пороги)."""
    g = g.sort_values("legit_restricted_rate")
    keep, best = [], np.inf
    for _, r in g.iterrows():
        if r.loss_share < best - 1e-12:
            keep.append(r)
            best = r.loss_share
    return pd.DataFrame(keep)


# ------------------------------------------------------------------ симуляция
def loss_vs_friction(curves, res):
    fig, ax = plt.subplots(figsize=(11, 6.2))
    fig.subplots_adjust(left=0.08, right=0.97, top=0.86, bottom=0.14)
    xmax = 1.6
    budget = res.loc[res.variant_short == "В1", "legit_restricted_rate"].iloc[0] * 100
    ax.axvline(budget, color=MUTED, lw=1, ls=(0, (4, 3)))
    ax.text(budget + 0.02, 74, f"Равное трение:\n{ru(budget, 2)}% законных операций", fontsize=9, color=TEXT2, va="top")
    ax.axhline(35, color=MUTED, lw=1, ls=(0, (1, 2)))
    ax.text(0.02, 35.8, "Равные потери: 35%", fontsize=9, color=TEXT2, ha="left")
    for v in ORDER:
        e = envelope(curves[curves.variant_short == v])
        e = e[e.legit_restricted_rate * 100 <= xmax]
        x, y = e.legit_restricted_rate * 100, e.loss_share * 100
        ax.plot(x, y, color=COLOR[v], lw=2.6 if v == "В3" else 2, solid_capstyle="round", zorder=3)
        op = res[res.variant_short == v].iloc[0]
        ax.scatter([op.legit_restricted_rate * 100], [op.loss_share * 100], s=70, color=COLOR[v],
                   edgecolor=SURFACE, linewidth=2, zorder=4)
        dx, dy, ha = OP_LABEL[v]
        ax.annotate(ru(op.loss_share * 100, 0) + "%", (op.legit_restricted_rate * 100, op.loss_share * 100),
                    xytext=(dx, dy), textcoords="offset points", ha=ha, va="center", fontsize=9.5, color=TEXT,
                    fontweight="bold" if v == "В3" else "normal")
        above = v == "В1"  # подпись В1 над линией: под ней проходит В2
        ax.text(x.iloc[-1], y.iloc[-1] + (1.2 if above else -1.6), SHORT[v], ha="right",
                va="bottom" if above else "top", color=TEXT, fontsize=10,
                fontweight="bold" if v == "В3" else "normal")
    b4 = envelope(curves[curves.variant_short == "В4"])
    ax.annotate("В4: трение не ниже ~0,5%\n(обязательное подтверждение\nдоверенным лицом для 55+)",
                (b4.legit_restricted_rate.iloc[0] * 100, b4.loss_share.iloc[0] * 100), xytext=(0.95, 62),
                textcoords="data", fontsize=8.5, color=TEXT2, va="bottom",
                arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.8))
    ax.set_xlim(0, xmax)
    ax.set_ylim(0, 80)
    pct_axis(ax, "y")
    pct_axis(ax, "x", 1)
    ax.set_xlabel("Доля законных операций с ограничением (пауза, звонок, отказ)")
    ax.set_ylabel("Непредотвращенные потери, % от суммы попыток хищения")
    finish(fig, "01_poteri_vs_trenie.png",
           SYNTH + " Точки: рабочие пороги при равном трении. Линии: лучшие настройки порогов каждого варианта.",
           "Потери против трения: при том же трении адаптивное трение теряет меньше",
           "Каждая линия: варианты порогов одного подхода. Ниже и левее лучше. Задержка данных до суток учтена.")


def variants_compare(res, model):
    fig, axes = plt.subplots(1, 3, figsize=(13, 5.4))
    fig.subplots_adjust(left=0.22, right=0.97, top=0.78, bottom=0.12, wspace=0.3)
    rows = list(reversed(ORDER))
    panels = [
        ("Потери при равном трении", "% от суммы попыток, синтетика",
         {v: res.set_index("variant_short").loc[v, "loss_share"] * 100 for v in ORDER}, 0, "%"),
        ("Трение при равных потерях (35%)", "% законных операций с ограничением, синтетика",
         {v: res.set_index("variant_short").loc[v, "restricted_rate_at_equal_loss"] * 100 for v in ORDER}, 2, "%"),
        ("Стоимость внедрения", "млн руб., оценка команды",
         {v: model["variant_cost"][v] for v in ORDER}, 1, ""),
    ]
    for ax, (t, sub, vals, d, unit) in zip(axes, panels):
        y = np.arange(len(rows))
        ax.barh(y, [vals[v] for v in rows], color=[COLOR[v] for v in rows], height=0.62, edgecolor=SURFACE, linewidth=2)
        for yi, v in zip(y, rows):
            ax.text(vals[v], yi, "  " + ru(vals[v], d) + unit, va="center", fontsize=10.5, color=TEXT,
                    fontweight="bold" if v == "В3" else "normal")
        ax.set_yticks(y)
        ax.set_yticklabels([NAME[v] for v in rows] if ax is axes[0] else [""] * len(rows))
        ax.set_title(t + "\n", loc="left", fontsize=12, fontweight="bold")
        ax.text(0, 1.02, sub, transform=ax.transAxes, fontsize=9, color=TEXT2)
        ax.set_xlim(0, max(vals.values()) * 1.3)
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _, d=d, u=unit: ru(v, 1 if d == 2 else 0) + u))
        ax.grid(axis="y", visible=False)
        ax.tick_params(axis="y", length=0)
    finish(fig, "02_sravnenie_variantov.png",
           SYNTH + " Стоимость: оценка команды (Excel, лист «Варианты»); В3 по смете ≤ 50 млн руб.",
           "Сравнение вариантов: В3 теряет меньше при том же трении и требует меньше ограничений при тех же потерях")


# ---------------------------------------------------------------- экономика
def economics(model):
    b = model["scenarios"]["b"]
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6), gridspec_kw={"width_ratios": [1, 2.3]})
    fig.subplots_adjust(left=0.06, right=0.98, top=0.8, bottom=0.2, wspace=0.25)

    # потери клиентов: было -> стало
    vals = [b["loss_before"], b["loss_after"]]
    ax1.bar([0, 1], vals, color=[TOTAL, GAIN], width=0.6, edgecolor=SURFACE, linewidth=2)
    for i, v in enumerate(vals):
        ax1.text(i, v, ru(v, 0), ha="center", va="bottom", fontsize=11, fontweight="bold")
    ax1.text(0.5, vals[0] * 1.12, f"−{ru(b['saved'], 0)} сохранено\n(KPI клиента, не доход банка)",
             ha="center", fontsize=9.5, color=TEXT2)
    ax1.set_xticks([0, 1], ["Было", "Стало"])
    ax1.set_ylim(0, vals[0] * 1.35)
    ax1.set_title("Потери клиентов ВТБ в фокусе\n", loc="left", fontsize=12, fontweight="bold")
    ax1.text(0, 1.02, "млн руб./год", transform=ax1.transAxes, fontsize=9, color=TEXT2)
    ax1.grid(axis="x", visible=False)

    # водопад эффекта для банка
    steps = [("Избежанный\nрасход банка", b["avoided"], "gain"),
             ("Экономия на\nошибочных\nограничениях", b["fp_saving"], "gain"),
             ("Удержание\nклиентов", b["retention"], "gain"),
             ("Эффект\n(валовый)", b["gross"], "total"),
             ("Эксплуатация\n(Opex)", -b["opex"], "cost"),
             ("Чистый\nгодовой\nэффект", b["net"], "total")]
    run = 0.0
    for i, (lab, v, kind) in enumerate(steps):
        if kind == "total":
            bottom, h, color, run = 0, v, TOTAL, v
        else:
            bottom = run if v >= 0 else run + v
            h, color = abs(v), GAIN if v >= 0 else COST
            run += v
        ax2.bar(i, h, bottom=bottom, color=color, width=0.62, edgecolor=SURFACE, linewidth=2)
        sign = "" if kind == "total" else ("+" if v >= 0 else "−")
        ax2.text(i, bottom + h, sign + ru(abs(v), 0), ha="center", va="bottom", fontsize=11,
                 fontweight="bold" if kind == "total" else "normal")
        if i < len(steps) - 1:
            ax2.plot([i + 0.31, i + 0.69], [run, run], color=MUTED, lw=0.8)
    ax2.set_xticks(range(len(steps)), [s[0] for s in steps], fontsize=9.5)
    ax2.set_ylim(0, b["gross"] * 1.22)
    ax2.set_title("Эффект для банка, базовый сценарий\n", loc="left", fontsize=12, fontweight="bold")
    ax2.text(0, 1.02, "млн руб./год после выхода на полный эффект", transform=ax2.transAxes, fontsize=9, color=TEXT2)
    ax2.grid(axis="x", visible=False)
    pb = int(b["payback_month"]) if isinstance(b["payback_month"], (int, float)) else b["payback_month"]
    finish(fig, "03_ekonomika_vodopad.png", MODEL,
           f"Экономика: ≈{ru(b['net'], 0)} млн руб./год чистого эффекта при внедрении {ru(b['capex'], 1)} млн руб. (≤ 50)",
           f"Окупаемость: {pb}-й месяц от старта проекта. Сохраненные средства клиентов не считаются доходом банка: "
           f"в эффект входит только избежанный расход ({ru(model_k(model) * 100, 0)}% от них).")


def model_k(model):
    b = model["scenarios"]["b"]
    return b["avoided"] / b["saved"] if b["saved"] else 0


def tornado(model):
    s = model["sensitivity"]
    base = s["base_net"]
    df = pd.DataFrame(s["rows"])
    df["range"] = (df.high - df.low).abs()
    df = df.sort_values("range")
    fig, ax = plt.subplots(figsize=(12, 6.2))
    fig.subplots_adjust(left=0.36, right=0.95, top=0.82, bottom=0.14)
    y = np.arange(len(df))
    step = s["step"] * 100
    for yi, (_, r) in zip(y, df.iterrows()):
        for val, color in ((r.low, COST), (r.high, GAIN)):
            ax.barh(yi, val - base, left=base, color=color, height=0.62, edgecolor=SURFACE, linewidth=1.5)
            ax.text(val, yi, ("  " if val >= base else "") + ru(val, 0) + ("  " if val < base else ""),
                    va="center", ha="left" if val >= base else "right", fontsize=9.5, color=TEXT)
    ax.axvline(base, color=TEXT, lw=1)
    ax.set_yticks(y, [n.replace(" (", "\n(", 1) if len(n) > 45 else n for n in df.name], fontsize=9.5)
    lo, hi = df[["low", "high"]].min().min(), df[["low", "high"]].max().max()
    pad = (hi - lo) * 0.15
    ax.set_xlim(lo - pad, hi + pad)
    ax.grid(axis="y", visible=False)
    ax.tick_params(axis="y", length=0)
    ax.set_xlabel("Чистый годовой эффект, млн руб./год (база " + ru(base, 0) + ")")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=COST, label=f"допущение −{ru(step, 0)}%"),
                       Patch(color=GAIN, label=f"допущение +{ru(step, 0)}%")],
              loc="lower right", frameon=False, fontsize=10)
    finish(fig, "04_chuvstvitelnost_tornado.png", MODEL,
           f"Чувствительность: эффект остается положительным при ±{ru(step, 0)}% по каждому допущению",
           "Изменение одного допущения при остальных базовых. Сверху самые влиятельные.")
    df.sort_values("range", ascending=False).assign(base=base).to_csv(OUT / "chuvstvitelnost.csv", index=False, encoding="utf-8")


def cumulative(model):
    fig, ax = plt.subplots(figsize=(11, 5.8))
    fig.subplots_adjust(left=0.08, right=0.86, top=0.84, bottom=0.14)
    months = np.arange(1, 37)
    ax.axvspan(0.5, 6.5, color=GRID, alpha=0.6, lw=0)
    ax.text(3.5, 0.97, "внедрение\n(М5-М6 пилот 20%)", transform=ax.get_xaxis_transform(), ha="center", va="top",
            fontsize=9, color=TEXT2)
    ax.axhline(0, color=TEXT, lw=0.9)
    rows, labelled = [], {}
    for sc in ("p", "b", "o"):
        name, color = SCEN[sc]
        cf = model["months_cf"][sc]
        ax.plot(months, cf, color=color, lw=2.4 if sc == "b" else 2)
        ax.text(36.6, cf[-1], f"{name}: {ru(cf[-1], 0)}", va="center", fontsize=10, color=TEXT,
                fontweight="bold" if sc == "b" else "normal")
        pb = model["scenarios"][sc]["payback_month"]
        if isinstance(pb, (int, float)):
            ax.scatter([pb], [cf[int(pb) - 1]], s=55, color=color, edgecolor=SURFACE, linewidth=2, zorder=4)
            labelled.setdefault(int(pb), []).append((name, cf[int(pb) - 1]))
        rows += [{"scenario": name, "month": int(m), "cumulative_net_mln": v} for m, v in zip(months, cf)]
    note = "; ".join(f"{' и '.join(n.lower() for n, _ in items)}: М{pb}" for pb, items in sorted(labelled.items()))
    ax.text(7.6, 0.93, "Окупаемость (точки):\n" + note, transform=ax.get_xaxis_transform(), va="top",
            fontsize=9.5, color=TEXT2)
    ax.set_xlim(0.5, 36.5)
    ax.set_xticks([1, 6, 7, 12, 18, 24, 30, 36])
    ax.set_xlabel("Месяц от старта проекта")
    ax.set_ylabel("Накопленный чистый эффект, млн руб.")
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: ru(v, 0)))
    finish(fig, "05_nakoplennyi_effekt.png", MODEL,
           "Накопленный эффект: проект окупается во всех трех сценариях",
           "Чистый поток = эффект − внедрение − эксплуатация. С М7 полный эффект.")
    pd.DataFrame(rows).to_csv(OUT / "nakoplennyi_effekt.csv", index=False, encoding="utf-8")


def tables(model, res):
    rows, labelled = [], {}
    for sc in ("p", "b", "o"):
        s = model["scenarios"][sc]
        rows.append({"scenario": SCEN[sc][0], **s})
    pd.DataFrame(rows).to_csv(OUT / "scenarii_excel.csv", index=False, encoding="utf-8")
    cols = ["variant", "TP", "FP", "FN", "TN", "precision", "recall", "fpr", "legit_restricted",
            "legit_restricted_rate", "signals_cleared_share", "prevented_share", "loss_share",
            "restricted_rate_at_equal_loss"]
    res[cols].to_csv(OUT / "slaid3_metriki_variantov.csv", index=False, encoding="utf-8")


def main():
    model = json.loads((OUT / "model_values.json").read_text())
    curves = pd.read_csv(OUT / "sim_curves.csv")
    res = pd.read_csv(OUT / "sim_variants_metrics.csv")
    loss_vs_friction(curves, res)
    variants_compare(res, model)
    economics(model)
    tornado(model)
    cumulative(model)
    tables(model, res)


if __name__ == "__main__":
    main()
