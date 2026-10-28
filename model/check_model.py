#!/usr/bin/env python3
"""Проверка Excel-модели без Excel: пересчет формул (библиотека formulas), ошибки, «зашитые» числа, сверка с 00.

Запуск: .venv/bin/python model/check_model.py
Пишет out/model_values.json: значения ключевых ячеек для графиков.
"""
import json
import re
import sys
from pathlib import Path

import formulas
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "model" / "CupMoscow2026-ВТБ-ОмскийПряник.xlsx"
OUT_JSON = ROOT / "out" / "model_values.json"

# числа, допустимые в формулах как явные константы: 0, 1, 2 (полугодие → год), 4 (кварталов),
# 12 (месяцев), 24, 36 (горизонт), 60 (минут), 1000 (единицы), 0.0001/0.001 (допуск сверки)
ALLOWED = {"0", "1", "2", "4", "12", "24", "36", "60", "1000", "0.0001", "0.001"}


def calc():
    model = formulas.ExcelModel().loads(str(XLSX)).finish()
    sol = model.calculate()
    values = {}
    for key, val in sol.items():
        m = re.match(r"^'?\[(.+?)\](.+?)'?!([A-Z]+\d+)$", key)
        if not m:
            continue
        v = val.value[0, 0] if hasattr(val, "value") else val
        values[(m.group(2).upper(), m.group(3))] = v
    return values


def get(values, sheet, cell):
    return values.get((sheet.upper(), cell))


def main():
    wb = load_workbook(XLSX)
    values = calc()
    errors, hard = [], []
    n_formulas = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_formulas += 1
                    v = get(values, ws.title, c.coordinate)
                    if v is None or (isinstance(v, str) and v.startswith("#")) or type(v).__name__ == "XlError":
                        errors.append(f"{ws.title}!{c.coordinate}: {v} <- {c.value}")
                    body = re.sub(r"'[^']*'!", "", c.value)          # имена листов
                    body = re.sub(r'"[^"]*"', "", body)               # строки
                    body = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "", body)   # адреса ячеек
                    for num in re.findall(r"(?<![A-Za-z_])\d+(?:\.\d+)?", body):
                        if num not in ALLOWED:
                            hard.append(f"{ws.title}!{c.coordinate}: {num} in {c.value}")
    print(f"formulas: {n_formulas}, errors: {len(errors)}, hardcoded numbers: {len(hard)}")
    for e in errors[:30]:
        print("  ERR", e)
    for h in hard[:30]:
        print("  HARD", h)

    def find(sheet, label, col):
        ws = wb[sheet]
        for row in ws.iter_rows(min_col=2, max_col=2):
            c = row[0]
            if isinstance(c.value, str) and c.value.strip().startswith(label):
                return get(values, sheet, f"{col}{c.row}")
        raise KeyError(f"{sheet}: {label}")

    sc_cols = {"p": "D", "b": "E", "o": "F"}
    kpi_cols = {"now": "D", "p": "E", "b": "F", "o": "G"}
    out = {"scenarios": {}}
    for sc in ("p", "b", "o"):
        e, k = sc_cols[sc], kpi_cols[sc]
        out["scenarios"][sc] = {
            "loss_before": find("KPI", "1. Потери", "D"),
            "loss_after": find("KPI", "1. Потери", k),
            "saved": find("KPI", "Сохраненные", k),
            "recall_after": find("KPI", "2. Полнота", k),
            "detected_before": find("KPI", "Корректно выявлено", "D"),
            "detected_after": find("KPI", "Корректно выявлено", k),
            "restr_before": find("KPI", "3. Ошибочные", "D"),
            "restr_after": find("KPI", "3. Ошибочные", k),
            "nps_delta": find("KPI", "4. NPS", k),
            "avoided": find("Эффект", "Избежанный расход", e),
            "fp_saving": find("Эффект", "Экономия на ошибочных", e),
            "retention": find("Эффект", "Удержание", e),
            "gross": find("Эффект", "Итого эффект", e),
            "opex": find("Эффект", "Затраты на эксплуатацию", e),
            "net": find("Эффект", "Чистый годовой", e),
            "capex": find("Эффект", "Затраты на внедрение", e),
            "payback_month": find("Эффект", "Срок окупаемости", e),
            "cum36": find("Эффект", "Накопленный чистый эффект за 36", e),
            "npv36": find("Эффект", "NPV", e),
        }
    ws = wb["Чувствительность"]
    out["sensitivity"] = {
        "step": get(values, "Чувствительность", "C4"),
        "base_net": get(values, "Чувствительность", "C5"),
        "check": get(values, "Чувствительность", "C6"),
        "rows": [
            {"name": ws[f"B{r}"].value,
             "low": get(values, "Чувствительность", f"D{r}"),
             "high": get(values, "Чувствительность", f"E{r}")}
            for r in range(9, 18)
        ],
    }
    wsv = wb["Варианты"]
    cost_row = next(c.row for c in wsv["B"] if c.value == "Стоимость внедрения (оценка), млн руб.")
    out["variant_cost"] = {v: get(values, "Варианты", f"{col}{cost_row}") for v, col in zip(["В1", "В2", "В3", "В4"], "DEFG")}
    tot_row = next(c.row for c in wsv["B"] if c.value == "Итого (взвешенная оценка)")
    out["variant_score"] = {v: get(values, "Варианты", f"{col}{tot_row}") for v, col in zip(["В1", "В2", "В3", "В4"], "DEFG")}
    out["budget_ok"] = find("Затраты", "Бюджет ≤", "C")
    out["time_ok"] = find("Затраты", "Срок ≤", "C")
    out["months_cf"] = {
        sc: [get(values, "Эффект", f"{c}{row}") for c in
             [wb["Эффект"].cell(row=1, column=3 + m).column_letter for m in range(1, 37)]]
        for sc, row in (("p", 32), ("b", 35), ("o", 38))
    }

    def conv(o):
        if isinstance(o, dict):
            return {k: conv(v) for k, v in o.items()}
        if isinstance(o, list):
            return [conv(v) for v in o]
        if hasattr(o, "item"):
            return o.item()
        return o if isinstance(o, (int, float, str, type(None))) else str(o)

    out = conv(out)
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(out, ensure_ascii=False, indent=2))

    b = out["scenarios"]["b"]
    print("\nБаза vs ориентиры 00 (разд. 4.5, 5, 6):")
    checks = [
        ("Потери ВТБ было, млн", b["loss_before"], 1964),
        ("Сохраненные средства, млн", b["saved"], 393),
        ("Избежанный расход, млн", b["avoided"], 59),
        ("Экономия на ограничениях, млн", b["fp_saving"], 129),
        ("Удержание, млн", b["retention"], 86),
        ("Итого эффект, млн", b["gross"], 270),
        ("Opex, млн", b["opex"], 24),
        ("Внедрение, млн", b["capex"], 47.6),
        ("Ограничения было, %", b["restr_before"] * 100, 0.525),
        ("Ограничения стало, %", b["restr_after"] * 100, 0.216),
        ("NPS, п.", b["nps_delta"], 0.8),
    ]
    for name, got, ref in checks:
        print(f"  {name:32s} {got:10.3f}  ориентир {ref}")
    for sc in ("p", "b", "o"):
        s = out["scenarios"][sc]
        print(f"  [{sc}] saved {s['saved']:.1f} gross {s['gross']:.1f} net {s['net']:.1f} "
              f"detected {s['detected_before']:.1f}→{s['detected_after']:.1f} тыс. "
              f"payback {s['payback_month']} cum36 {s['cum36']:.1f} npv {s['npv36']:.1f} nps {s['nps_delta']:+.2f}")
    print("  budget:", out["budget_ok"], "time:", out["time_ok"], "sens check:", out["sensitivity"]["check"])
    print("  scores:", out["variant_score"], "costs:", out["variant_cost"])
    return 1 if errors or hard else 0


if __name__ == "__main__":
    sys.exit(main())
