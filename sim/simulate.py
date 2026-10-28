#!/usr/bin/env python3
"""Симуляция на синтетических данных: сравнение вариантов В1-В4.

ИЛЛЮСТРАЦИЯ МЕХАНИЗМА, НЕ ПРОГНОЗ. Все распределения ниже: допущения команды, не данные ВТБ.

Что моделируем:
  * законные переводы физлиц (профили: возраст, типичная сумма, активность, время суток);
  * мошенничество 0,3% операций: сценарий А «безопасный счет» (социальная инженерия, клиент сам
    подтверждает), сценарий Б «захват доступа» (смена устройства/номера, затем перевод);
  * шум: законные операции с похожими признаками (крупная покупка после смены телефона,
    крупный перевод 55+ новому получателю после звонка продавца);
  * задержка данных: часть признаков (смена номера у оператора, входящий звонок, база ФинЦЕРТ)
    в момент операции недоступна и приходит пакетом до суток;
  * регуляторный слой (упрощенно): одинаков во всех вариантах, модель его не отменяет;
  * исходы мер считаем как ожидаемые значения (вероятность остановить мошенничество и
    вероятность ограничить законную операцию), чтобы кривые были гладкими.

Запуск:  .venv/bin/python sim/simulate.py
Выход:   out/sim_variants_metrics.csv, out/sim_curves.csv, out/sim_delay.csv, out/sim_measures_v3.csv,
         out/sim_seeds.csv
"""
from dataclasses import dataclass, replace
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingClassifier

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "out"
SEED = 2026


@dataclass(frozen=True)
class Config:
    n_clients: int = 30_000
    days: int = 60
    train_days: int = 36            # обучение на первых днях, оценка на остальных
    ops_per_client_day: float = 0.35
    fraud_share: float = 0.003      # доля мошеннических операций (0,1-0,5% по инструкции)
    share_a: float = 0.65           # доля сценария А «безопасный счет»
    # доступность признаков в момент операции (остальное приходит с задержкой до суток)
    avail_phone: float = 0.5        # смена номера (данные оператора)
    avail_call: float = 0.4         # входящий звонок перед операцией (данные оператора)
    avail_flag: float = 0.7         # получатель в базе ФинЦЕРТ (обновление базы)
    avail_flag_ext: float = 0.95    # В4: обмен данными с другими банками
    trusted_optin: float = 0.15     # В3: доля клиентов 55+ с подключенной «Защитой близких»
    trusted_amount: float = 15_000  # В4: обязательное подтверждение доверенным лицом для 55+ выше суммы


# Меры и их ожидаемый эффект (допущения): индекс = код меры
MEASURES = ["пропустить", "предупреждение", "усиленное подтверждение", "пауза (охлаждение)", "звонок", "отказ"]
PASS, WARN, STEP, PAUSE, CALL, DECLINE = range(6)
P_STOP_A = np.array([0.00, 0.15, 0.10, 0.45, 0.65, 0.95])   # «безопасный счет»: клиент сам подтверждает
P_STOP_B = np.array([0.00, 0.05, 0.75, 0.60, 0.90, 0.98])   # захват доступа: у мошенника нет биометрии
P_RESTR = np.array([0.00, 0.00, 0.02, 1.00, 1.00, 1.00])    # законная операция реально остановлена
TRUSTED = (0.70, 0.60, 1.00)  # «Защита близких»: стоп А, стоп Б, ограничение законной (ждет подтверждения)


def generate(cfg: Config, rng: np.random.Generator) -> pd.DataFrame:
    nc = cfg.n_clients
    grp = rng.choice(3, size=nc, p=[0.30, 0.38, 0.32])
    age = np.select([grp == 0, grp == 1], [rng.integers(18, 35, nc), rng.integers(35, 55, nc)], rng.integers(55, 86, nc))
    old = age >= 55
    median = np.exp(rng.normal(np.log(3_000), 0.7, nc)) * np.where(old, 0.8, 1.0)
    activity = rng.gamma(2.0, cfg.ops_per_client_day / 2.0, nc) * np.where(old, 0.7, 1.0)
    peak = rng.normal(14, 3, nc)

    # законные операции
    cid = np.repeat(np.arange(nc), rng.poisson(activity * cfg.days))
    n = len(cid)
    d = pd.DataFrame({
        "client": cid,
        "day": rng.integers(0, cfg.days, n),
        "hour": np.clip(rng.normal(peak[cid], 3.5), 0, 23.99),
        "amount": median[cid] * np.exp(rng.normal(0, 0.9, n)),
        "new_rcp": rng.random(n) < 0.18,
        "device_chg": rng.random(n) < 0.006,
        "phone_chg": rng.random(n) < 0.003,
        "call_before": rng.random(n) < 0.04,
        "on_call": rng.random(n) < 0.03,
        "rcp_flag": rng.random(n) < 0.0003,
        "sbp": rng.random(n) < 0.6,
    })
    # шум 1: крупная покупка после смены телефона (как у стажера в кейсе)
    m = rng.random(n) < 0.003
    d.loc[m, ["device_chg", "new_rcp"]] = True
    d.loc[m, "phone_chg"] = rng.random(m.sum()) < 0.4
    d.loc[m, "amount"] = median[cid[m]] * rng.uniform(8, 40, m.sum())
    # шум 2: крупный законный перевод 55+ новому получателю после звонка (ремонт, покупка)
    m = old[cid] & (rng.random(n) < 0.004)
    d.loc[m, "new_rcp"] = True
    d.loc[m, "call_before"] = rng.random(m.sum()) < 0.6
    d.loc[m, "on_call"] = rng.random(m.sum()) < 0.3
    d.loc[m, "amount"] = median[cid[m]] * rng.uniform(6, 30, m.sum())
    d["label"], d["scenario"] = 0, "законная"

    # мошенничество
    nf = int(round(n * cfg.fraud_share / (1 - cfg.fraud_share)))
    na = int(round(nf * cfg.share_a))
    nb = nf - na
    old_ids, young_ids = np.flatnonzero(old), np.flatnonzero(~old)
    va = np.where(rng.random(na) < 0.7, rng.choice(old_ids, na), rng.choice(young_ids, na))
    fa = pd.DataFrame({
        "client": va, "day": rng.integers(0, cfg.days, na),
        "hour": np.clip(rng.normal(13, 3, na), 0, 23.99),
        "amount": np.exp(rng.normal(np.log(20_000), 0.9, na)),
        "new_rcp": rng.random(na) < 0.97, "device_chg": rng.random(na) < 0.05,
        "phone_chg": rng.random(na) < 0.03, "call_before": rng.random(na) < 0.85,
        "on_call": rng.random(na) < 0.60, "rcp_flag": rng.random(na) < 0.30, "sbp": rng.random(na) < 0.65,
        "label": 1, "scenario": "А: безопасный счет",
    })
    vb = rng.integers(0, nc, nb)
    night = rng.random(nb) < 0.4
    fb = pd.DataFrame({
        "client": vb, "day": rng.integers(0, cfg.days, nb),
        "hour": np.where(night, rng.uniform(0, 6, nb), np.clip(rng.normal(15, 4, nb), 0, 23.99)),
        "amount": np.exp(rng.normal(np.log(15_000), 0.9, nb)),
        "new_rcp": rng.random(nb) < 0.95, "device_chg": rng.random(nb) < 0.80,
        "phone_chg": rng.random(nb) < 0.60, "call_before": rng.random(nb) < 0.10,
        "on_call": rng.random(nb) < 0.05, "rcp_flag": rng.random(nb) < 0.30, "sbp": rng.random(nb) < 0.7,
        "label": 1, "scenario": "Б: захват доступа",
    })
    df = pd.concat([d, fa, fb], ignore_index=True)
    df["age"] = age[df.client]
    df["median"] = median[df.client]
    df["night"] = (df.hour < 6) | (df.hour >= 23)
    # задержка данных: в момент операции признак виден только с вероятностью avail_*
    r = rng.random((len(df), 4))
    df["phone_chg_rt"] = df.phone_chg & (r[:, 0] < cfg.avail_phone)
    df["call_before_rt"] = df.call_before & (r[:, 1] < cfg.avail_call)
    df["rcp_flag_rt"] = df.rcp_flag & (r[:, 2] < cfg.avail_flag)
    df["rcp_flag_ext"] = df.rcp_flag & (r[:, 3] < cfg.avail_flag_ext)
    # регуляторный слой видит данные с базовой задержкой во всех прогонах: сравнение «без задержки»
    # меняет только данные для скоринга
    df["phone_chg_reg"] = df.phone_chg & (r[:, 0] < Config.avail_phone)
    df["rcp_flag_reg"] = df.rcp_flag & (r[:, 2] < Config.avail_flag)
    df["trusted_on"] = (df.age >= 55) & (rng.random(len(df)) < cfg.trusted_optin)
    return df.sample(frac=1, random_state=SEED).reset_index(drop=True)


def features(df: pd.DataFrame, mode: str) -> np.ndarray:
    """mode: rt = в момент операции; full = после догрузки данных; ext = rt + обмен данными (В4)."""
    phone = df.phone_chg if mode == "full" else df.phone_chg_rt
    call = df.call_before if mode == "full" else df.call_before_rt
    flag = {"rt": df.rcp_flag_rt, "full": df.rcp_flag, "ext": df.rcp_flag_ext}[mode]
    return np.column_stack([
        np.log(df.amount), np.log(df.amount / df["median"]), df.new_rcp, df.night, df.age, df.age >= 55,
        df.sbp, df.device_chg, phone, call, df.on_call, flag,
    ]).astype(float)


def regulatory(df: pd.DataFrame, flag_col: str = "rcp_flag_reg") -> np.ndarray:
    """Регуляторный слой (упрощенно, 2 признака): получатель в базе ФинЦЕРТ; смена номера за 48 ч
    + перевод новому получателю. Минимальная мера: пауза. Модель не может ее снять."""
    hit = df[flag_col] | (df.phone_chg_reg & df.new_rcp)
    return np.where(hit, PAUSE, PASS)


def outcome(df, action, trusted=None, extra=None):
    """Ожидаемые вероятности: остановить мошенничество и ограничить законную операцию."""
    is_a = (df.scenario == "А: безопасный счет").to_numpy()
    is_b = (df.scenario == "Б: захват доступа").to_numpy()
    legit = (df.label == 0).to_numpy()
    p_stop = np.where(is_a, P_STOP_A[action], np.where(is_b, P_STOP_B[action], 0.0))
    p_restr = np.where(legit, P_RESTR[action], 0.0)
    if extra is not None:  # доп. мера поверх (эскалация после паузы): объединяем как независимые
        stop_x = np.where(is_a, P_STOP_A[CALL], np.where(is_b, P_STOP_B[CALL], 0.0))
        p_stop = np.where(extra, 1 - (1 - p_stop) * (1 - stop_x), p_stop)
    if trusted is not None:
        st = np.where(is_a, TRUSTED[0], np.where(is_b, TRUSTED[1], 0.0))
        p_stop = np.where(trusted, 1 - (1 - p_stop) * (1 - st), p_stop)
        p_restr = np.where(trusted & legit, np.maximum(p_restr, TRUSTED[2]), p_restr)
    signal = (action > PASS) | (trusted if trusted is not None else False)
    return signal, p_stop, p_restr


# ------------------------------------------------------------------ варианты
def v1_rules(df, f=1.0):
    """В1: пороги по сумме и признакам + усиленная аутентификация. Только данные «в моменте»."""
    ratio = df.amount / df["median"]
    act = np.full(len(df), PASS)
    act[((df.amount > 15_000 * f) & df.new_rcp) | (ratio > 10 * f)] = STEP
    risky = (df.device_chg | df.phone_chg_rt) & df.new_rcp & (ratio > 3 * f)
    act[((df.amount > 50_000 * f) & df.new_rcp) | risky] = CALL   # удержание до звонка
    act = np.maximum(act, regulatory(df))
    return outcome(df, act)


def v2_ml(df, s_rt, t):
    """В2: ML-скоринг + сегментация (порог для 55+ ниже), бинарное решение: пропустить или блок до звонка."""
    thr = np.where(df.age >= 55, t * 0.6, t)
    act = np.where(s_rt >= thr, CALL, PASS)
    act = np.maximum(act, regulatory(df))
    return outcome(df, act)


def v3_adaptive(df, s_rt, s_full, t):
    """В3: скоринг + уровни мер. t = порог паузы; ниже: предупреждение и усиленное подтверждение
    (сигнал без ограничения), выше: звонок и отказ. Пауза = перепроверка по догрузившимся данным.
    Для захвата доступа (смена устройства) вместо паузы сначала биометрия."""
    tiers = np.array([t / 5, t / 2.5, t, t * 3, t * 8])
    tiers = np.minimum(tiers, [0.95, 0.96, 0.97, 0.98, 0.99])
    act = np.searchsorted(tiers, s_rt, side="right")  # 0..5
    takeover = df.device_chg.to_numpy() & (act == PAUSE)
    act = np.where(takeover, STEP, act)               # мера под тип риска
    act = np.maximum(act, regulatory(df))
    recheck = (act == PAUSE) & (s_full >= tiers[3])   # данные догрузились: эскалация до звонка
    trusted = df.trusted_on.to_numpy() & (act >= STEP)
    return outcome(df, act, trusted=trusted, extra=recheck)


def v4_external(df, s_ext, t, amount=None):
    """В4: обмен данными с другими банками (база дропов полнее) + обязательная «Защита близких»
    для 55+ при переводе новому получателю выше суммы: доп. подтверждение и задержка."""
    thr = np.where(df.age >= 55, t * 0.6, t)
    act = np.where(s_ext >= thr, CALL, PASS)
    act = np.maximum(act, regulatory(df, "rcp_flag_ext"))
    trusted = ((df.age >= 55) & df.new_rcp & (df.amount > (amount or Config.trusted_amount))).to_numpy()
    return outcome(df, act, trusted=trusted)


def metrics(df, signal, p_stop, p_restr):
    fraud = (df.label == 1).to_numpy()
    tp = int((signal & fraud).sum())
    fp = int((signal & ~fraud).sum())
    fn = int((~signal & fraud).sum())
    tn = int((~signal & ~fraud).sum())
    amt = df.amount.to_numpy()
    total = amt[fraud].sum()
    prevented = (amt * p_stop)[fraud].sum()
    restricted = p_restr[~fraud].sum()
    return {
        "TP": tp, "FP": fp, "FN": fn, "TN": tn,
        "precision": tp / (tp + fp) if tp + fp else 0.0,
        "false_signal_share": fp / (tp + fp) if tp + fp else 0.0,
        "recall": tp / (tp + fn) if tp + fn else 0.0,
        "fpr": fp / (fp + tn) if fp + tn else 0.0,
        "legit_restricted": restricted,
        "legit_restricted_rate": restricted / (~fraud).sum(),
        "signals_cleared_share": 1 - restricted / fp if fp else 0.0,
        "fraud_amount_total": total,
        "prevented_amount": prevented,
        "prevented_share": prevented / total if total else 0.0,
        "loss_share": 1 - prevented / total if total else 0.0,
    }


def train(df, mode, cfg):
    tr = df[df.day < cfg.train_days]
    # неглубокие деревья + регуляризация, без early stopping: при 0,3% мошенничества случайная
    # валидационная выборка early stopping делает модель нестабильной
    model = HistGradientBoostingClassifier(max_iter=150, learning_rate=0.05, max_depth=3, min_samples_leaf=200,
                                           l2_regularization=1.0, early_stopping=False, random_state=SEED)
    model.fit(features(tr, mode), tr.label)
    return model


VARIANTS = {
    "В1": "В1 Правила + усиленная аутентификация",
    "В2": "В2 ML-скоринг + сегментация",
    "В3": "В3 Адаптивное трение",
    "В4": "В4 Внешние участники",
}


def run(cfg: Config, rng_seed=SEED):
    rng = np.random.default_rng(rng_seed)
    df = generate(cfg, rng)
    m_rt, m_full, m_ext = (train(df, mode, cfg) for mode in ("rt", "full", "ext"))
    te = df[df.day >= cfg.train_days].reset_index(drop=True)
    s_rt = m_rt.predict_proba(features(te, "rt"))[:, 1]
    s_full = m_full.predict_proba(features(te, "full"))[:, 1]
    s_ext = m_ext.predict_proba(features(te, "ext"))[:, 1]
    policies = {
        "В1": lambda p: v1_rules(te, p),
        "В2": lambda p: v2_ml(te, s_rt, p),
        "В3": lambda p: v3_adaptive(te, s_rt, s_full, p),
        "В4": lambda p: v4_external(te, s_ext, p),
    }
    return df, te, policies


LOSS_TARGET = 0.35  # сравнение «при равных потерях»: непредотвращенные потери ≤ 35% суммы попыток

GRID = {
    "В1": np.geomspace(0.25, 8, 40),
    "В2": np.geomspace(0.003, 0.95, 40),
    "В3": np.geomspace(0.003, 0.95, 40),
    "В4": np.geomspace(0.003, 0.95, 40),
}


def sweep(policies, te):
    rows = []
    for v, grid in GRID.items():
        for p in grid:
            m = metrics(te, *policies[v](p))
            rows.append({"variant_short": v, "param": p, **m})
    return pd.DataFrame(rows)


def operating_points(curves: pd.DataFrame, budget: float) -> dict:
    """Рабочая точка варианта: минимум потерь при ошибочных ограничениях ≤ budget
    (единый «бюджет трения»); для В1 базовые пороги правил (param = 1)."""
    pts = {"В1": 1.0}
    for v in ("В2", "В3", "В4"):
        c = curves[(curves.variant_short == v) & (curves.legit_restricted_rate <= budget)]
        if c.empty:  # даже минимальное трение выше бюджета: берем точку с минимальным трением
            c = curves[curves.variant_short == v].nsmallest(1, "legit_restricted_rate")
        pts[v] = float(c.nsmallest(1, "loss_share").param.iloc[0])
    return pts


def main():
    OUT.mkdir(exist_ok=True)
    cfg = Config()
    df, te, policies = run(cfg)
    print(f"операций: {len(df):,}, мошеннических: {df.label.sum():,} ({df.label.mean():.2%}); "
          f"тест: {len(te):,} / {te.label.sum():,}")
    curves = sweep(policies, te)
    b1 = metrics(te, *policies["В1"](1.0))
    pts = operating_points(curves, budget=b1["legit_restricted_rate"])

    rows = []
    for v, p in pts.items():
        m = metrics(te, *policies[v](p))
        rows.append({"variant_short": v, "variant": VARIANTS[v], "param": p, **m})
    res = pd.DataFrame(rows)

    # те же варианты при равных потерях: сколько трения нужно, чтобы потери ≤ LOSS_TARGET
    target = LOSS_TARGET
    eq = []
    for v in VARIANTS:
        c = curves[(curves.variant_short == v) & (curves.loss_share <= target + 1e-9)]
        eq.append(c.legit_restricted_rate.min() if not c.empty else np.nan)
    res["restricted_rate_at_equal_loss"] = eq

    # задержка данных: те же рабочие точки без задержки (все признаки доступны в моменте)
    cfg_nd = replace(cfg, avail_phone=1.0, avail_call=1.0, avail_flag=1.0)
    _, te_nd, pol_nd = run(cfg_nd)
    budget = b1["legit_restricted_rate"]
    pts_nd = operating_points(sweep(pol_nd, te_nd), budget=budget)  # тот же бюджет трения
    delay = []
    for v, p in pts_nd.items():
        m_d = res.loc[res.variant_short == v].iloc[0]
        m_nd = metrics(te_nd, *pol_nd[v](p))
        delay.append({"variant_short": v, "loss_share_with_delay": m_d.loss_share,
                      "loss_share_no_delay": m_nd["loss_share"],
                      "restricted_rate_no_delay": m_nd["legit_restricted_rate"],
                      "delay_penalty_pp": (m_d.loss_share - m_nd["loss_share"]) * 100})

    # структура мер В3 на рабочей точке
    tiers_t = pts["В3"]
    m_rt = train(df, "rt", cfg)
    s_rt = m_rt.predict_proba(features(te, "rt"))[:, 1]
    tiers = np.minimum(np.array([tiers_t / 5, tiers_t / 2.5, tiers_t, tiers_t * 3, tiers_t * 8]), [0.95, 0.96, 0.97, 0.98, 0.99])
    act = np.searchsorted(tiers, s_rt, side="right")
    act = np.where(te.device_chg.to_numpy() & (act == PAUSE), STEP, act)
    act = np.maximum(act, regulatory(te))
    meas = (pd.DataFrame({"мера": np.array(MEASURES)[act], "законная": te.label == 0})
            .groupby(["мера", "законная"]).size().unstack(fill_value=0)
            .rename(columns={True: "законные", False: "мошеннические"}).reindex(MEASURES).fillna(0).astype(int))

    # устойчивость: те же правила выбора рабочих точек на других случайных выборках
    seeds = []
    for seed in (SEED, 11, 42, 777):
        _, te_s, pol_s = run(cfg, rng_seed=seed)
        cur_s = sweep(pol_s, te_s)
        pts_s = operating_points(cur_s, budget=metrics(te_s, *pol_s["В1"](1.0))["legit_restricted_rate"])
        for v, p in pts_s.items():
            m = metrics(te_s, *pol_s[v](p))
            seeds.append({"seed": seed, "variant_short": v, "loss_share": m["loss_share"],
                          "legit_restricted_rate": m["legit_restricted_rate"], "precision": m["precision"],
                          "recall": m["recall"]})
    seeds = pd.DataFrame(seeds)
    seeds.to_csv(OUT / "sim_seeds.csv", index=False, encoding="utf-8")
    print(seeds.pivot(index="seed", columns="variant_short", values="loss_share").round(3))

    res.to_csv(OUT / "sim_variants_metrics.csv", index=False, encoding="utf-8")
    curves.to_csv(OUT / "sim_curves.csv", index=False, encoding="utf-8")
    pd.DataFrame(delay).to_csv(OUT / "sim_delay.csv", index=False, encoding="utf-8")
    meas.to_csv(OUT / "sim_measures_v3.csv", encoding="utf-8")

    show = ["variant_short", "TP", "FP", "FN", "precision", "recall", "fpr", "legit_restricted_rate",
            "signals_cleared_share", "loss_share", "restricted_rate_at_equal_loss"]
    with pd.option_context("display.width", 200, "display.float_format", "{:.4f}".format):
        print(res[show].to_string(index=False))
        print(pd.DataFrame(delay).to_string(index=False))
        print(meas)


if __name__ == "__main__":
    main()
