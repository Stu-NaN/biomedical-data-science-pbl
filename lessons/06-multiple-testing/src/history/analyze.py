import hashlib
import json
import math
from pathlib import Path

import numpy as np
import pandas as pd


DATA = Path("work/source/6. multiple-test/data/leukemia.csv")
META = ["sample_id", "original_split", "diagnosis"]


def beta_continued_fraction(a, b, x):
    qab = a + b
    qap = a + 1.0
    qam = a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < 1e-300:
        d = 1e-300
    d = 1.0 / d
    h = d
    for m in range(1, 301):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < 1e-300:
            d = 1e-300
        c = 1.0 + aa / c
        if abs(c) < 1e-300:
            c = 1e-300
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < 1e-300:
            d = 1e-300
        c = 1.0 + aa / c
        if abs(c) < 1e-300:
            c = 1e-300
        d = 1.0 / d
        change = d * c
        h *= change
        if abs(change - 1.0) < 3e-14:
            return h
    raise ArithmeticError("Incomplete beta continued fraction did not converge")


def regularized_beta(x, a, b):
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    log_bt = (
        math.lgamma(a + b)
        - math.lgamma(a)
        - math.lgamma(b)
        + a * math.log(x)
        + b * math.log1p(-x)
    )
    bt = math.exp(log_bt)
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * beta_continued_fraction(a, b, x) / a
    return 1.0 - bt * beta_continued_fraction(b, a, 1.0 - x) / b


def t_two_sided_p(t_value, df):
    if not math.isfinite(t_value) or not math.isfinite(df) or df <= 0:
        raise ValueError("Invalid t statistic or degrees of freedom")
    x = df / (df + t_value * t_value)
    return min(1.0, max(0.0, regularized_beta(x, df / 2.0, 0.5)))


def t_critical_975(df):
    lo, hi = 0.0, 100.0
    for _ in range(70):
        mid = (lo + hi) / 2.0
        if t_two_sided_p(mid, df) > 0.05:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2.0


def beta_continued_fraction_array(a, b, x):
    a = np.asarray(a, dtype=float)
    b = np.asarray(b, dtype=float)
    c = np.ones_like(x)
    d = 1.0 - (a + b) * x / (a + 1.0)
    d = np.where(np.abs(d) < 1e-300, 1e-300, d)
    d = 1.0 / d
    h = d.copy()
    for m in range(1, 301):
        m2 = 2 * m
        aa = m * (b - m) * x / ((a + m2 - 1.0) * (a + m2))
        d = 1.0 + aa * d
        d = np.where(np.abs(d) < 1e-300, 1e-300, d)
        c = 1.0 + aa / c
        c = np.where(np.abs(c) < 1e-300, 1e-300, c)
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1.0))
        d = 1.0 + aa * d
        d = np.where(np.abs(d) < 1e-300, 1e-300, d)
        c = 1.0 + aa / c
        c = np.where(np.abs(c) < 1e-300, 1e-300, c)
        d = 1.0 / d
        change = d * c
        h *= change
        if np.max(np.abs(change - 1.0)) < 3e-14:
            return h
    raise ArithmeticError("Vector incomplete beta continued fraction did not converge")


def regularized_beta_array(x, a, b):
    x = np.clip(np.asarray(x, dtype=float), 1e-300, 1.0 - 1e-16)
    a = np.asarray(a, dtype=float)
    b = np.asarray(b, dtype=float)
    log_bt = (
        math.lgamma(float(np.ravel(a)[0] + np.ravel(b)[0]))
        - math.lgamma(float(np.ravel(a)[0]))
        - math.lgamma(float(np.ravel(b)[0]))
        + a * np.log(x)
        + b * np.log1p(-x)
    )
    bt = np.exp(log_bt)
    direct = x < (a + 1.0) / (a + b + 2.0)
    out = np.empty_like(x)
    if np.any(direct):
        xd = x[direct]
        ad = np.broadcast_to(a, x.shape)[direct]
        bd = np.broadcast_to(b, x.shape)[direct]
        btd = bt[direct]
        out[direct] = (
            btd
            * beta_continued_fraction_array(ad, bd, xd)
            / ad
        )
    if np.any(~direct):
        xc = 1.0 - x[~direct]
        ac = np.broadcast_to(a, x.shape)[~direct]
        bc = np.broadcast_to(b, x.shape)[~direct]
        btc = bt[~direct]
        out[~direct] = 1.0 - (
            btc
            * beta_continued_fraction_array(bc, ac, xc)
            / bc
        )
    return np.clip(out, 0.0, 1.0)


def adjust_values(p):
    m = len(p)
    order = np.argsort(p, kind="stable")
    ranks = np.arange(1, m + 1)
    sorted_bh = np.minimum.accumulate((p[order] * m / ranks)[::-1])[::-1]
    bh = np.empty_like(sorted_bh)
    bh[order] = np.minimum(sorted_bh, 1.0)
    return np.minimum(p * m, 1.0), bh


def simulate(repeats=1000, seed=42, delta=0.0, signals=0):
    m, n_all, n_aml = 1000, 27, 11
    dfree = n_all + n_aml - 2
    rng = np.random.default_rng(seed)
    truth = np.arange(m) < signals
    records = {
        method: {
            "R": np.zeros(repeats),
            "V": np.zeros(repeats),
            "FDP": np.zeros(repeats),
            "any_false": np.zeros(repeats),
            "power": np.full(repeats, np.nan),
        }
        for method in ["raw", "bonferroni", "bh"]
    }
    first = {}
    for run in range(repeats):
        a = rng.normal(size=(n_aml, m))
        b = rng.normal(size=(n_all, m))
        if signals:
            a[:, truth] += delta
        mean_a, mean_b = a.mean(axis=0), b.mean(axis=0)
        var_a, var_b = a.var(axis=0, ddof=1), b.var(axis=0, ddof=1)
        pooled_var = ((n_aml - 1) * var_a + (n_all - 1) * var_b) / dfree
        standard_error = np.sqrt(pooled_var * (1.0 / n_aml + 1.0 / n_all))
        t_values = (mean_a - mean_b) / standard_error
        x = dfree / (dfree + t_values**2)
        p = regularized_beta_array(x, dfree / 2.0, 0.5)
        bonf, bh = adjust_values(p)
        adjusted = {"raw": p, "bonferroni": bonf, "bh": bh}
        for method, values in adjusted.items():
            rejected = values <= 0.05
            r = int(rejected.sum())
            v = int((rejected & ~truth).sum())
            s = int((rejected & truth).sum())
            records[method]["R"][run] = r
            records[method]["V"][run] = v
            records[method]["FDP"][run] = v / max(r, 1)
            records[method]["any_false"][run] = int(v > 0)
            if signals:
                records[method]["power"][run] = s / signals
            if run == 0:
                first[method] = r
    summary = {}
    for method, data in records.items():
        summary[method] = {
            "mean_R": float(data["R"].mean()),
            "sd_R": float(data["R"].std(ddof=1)),
            "p05_R": float(np.quantile(data["R"], 0.05)),
            "p95_R": float(np.quantile(data["R"], 0.95)),
            "mean_V": float(data["V"].mean()),
            "FWER": float(data["any_false"].mean()),
            "FDR": float(data["FDP"].mean()),
            "mean_power": (
                None if not signals else float(np.nanmean(data["power"]))
            ),
        }
    return {"first_run_R": first, "summary": summary}


def test_table(frame, probes=None):
    if probes is None:
        probes = [col for col in frame.columns if col not in META]
    a = frame.loc[frame.diagnosis.eq("AML"), probes].to_numpy(dtype=float)
    b = frame.loc[frame.diagnosis.eq("ALL"), probes].to_numpy(dtype=float)
    na, nb = len(a), len(b)
    mean_a, mean_b = a.mean(axis=0), b.mean(axis=0)
    var_a, var_b = a.var(axis=0, ddof=1), b.var(axis=0, ddof=1)
    delta = mean_a - mean_b
    va, vb = var_a / na, var_b / nb
    se = np.sqrt(va + vb)
    df = (va + vb) ** 2 / (va**2 / (na - 1) + vb**2 / (nb - 1))
    t_values = delta / se
    raw = np.array([t_two_sided_p(float(t), float(d)) for t, d in zip(t_values, df)])
    order = np.argsort(raw, kind="stable")
    ranks = np.arange(1, len(raw) + 1)
    sorted_bh = np.minimum.accumulate((raw[order] * len(raw) / ranks)[::-1])[::-1]
    bh = np.empty_like(sorted_bh)
    bh[order] = np.minimum(sorted_bh, 1.0)
    pooled_sd = np.sqrt(
        ((na - 1) * var_a + (nb - 1) * var_b) / (na + nb - 2)
    )
    return pd.DataFrame(
        {
            "probe": probes,
            "n_AML": na,
            "n_ALL": nb,
            "mean_AML": mean_a,
            "mean_ALL": mean_b,
            "delta": delta,
            "df": df,
            "se": se,
            "raw": raw,
            "bonferroni": np.minimum(raw * len(raw), 1.0),
            "bh": bh,
            "cohen_d": delta / pooled_sd,
        }
    )


def add_ci(frame):
    frame = frame.copy()
    critical = np.array([t_critical_975(float(df)) for df in frame.df])
    frame["ci95_low"] = frame.delta - critical * frame.se
    frame["ci95_high"] = frame.delta + critical * frame.se
    return frame


def main():
    raw_bytes = DATA.read_bytes()
    df = pd.read_csv(DATA)
    x = df.drop(columns=META)
    assert df.shape == (72, 7132)
    assert df.sample_id.is_unique and x.columns.is_unique
    assert np.isfinite(x.to_numpy(dtype=float)).all()

    group_counts = {
        split: {
            diagnosis: int(((df.original_split == split) & (df.diagnosis == diagnosis)).sum())
            for diagnosis in ["ALL", "AML"]
        }
        for split in ["train", "test"]
    }
    train = df.loc[df.original_split.eq("train")]
    test = df.loc[df.original_split.eq("test")]
    discovery = test_table(train)
    counts = {
        method: int((discovery[method] <= 0.05).sum())
        for method in ["raw", "bonferroni", "bh"]
    }

    top20 = discovery.sort_values(["raw", "probe"]).head(20).copy()
    top20["wrong_bh_m20"] = np.minimum.accumulate(
        (top20.raw.to_numpy() * 20 / np.arange(1, 21))[::-1]
    )[::-1]
    top20["wrong_bonf_m20"] = np.minimum(top20.raw.to_numpy() * 20, 1.0)

    selected = discovery.loc[discovery.bh <= 0.05].sort_values(["raw", "probe"]).head(10)
    selected_ci = add_ci(selected)
    validation = test_table(test, selected.probe.tolist())
    validation = validation.merge(
        selected[["probe", "delta"]].rename(columns={"delta": "discovery_delta"}),
        on="probe",
        validate="one_to_one",
    )
    validation["same_direction"] = (
        np.sign(validation.delta) == np.sign(validation.discovery_delta)
    )
    validation["replicated"] = (validation.bonferroni <= 0.05) & validation.same_direction

    bins = np.linspace(0, 1, 21)
    hist, _ = np.histogram(discovery.raw, bins=bins)
    first = selected_ci.iloc[0]
    probe_table = selected_ci[
        ["probe", "delta", "ci95_low", "ci95_high", "cohen_d", "raw", "bonferroni", "bh"]
    ].copy()
    probe_table = probe_table.merge(
        validation[["probe", "delta", "bonferroni", "same_direction", "replicated"]].rename(
            columns={
                "delta": "test_delta",
                "bonferroni": "test_bonferroni",
            }
        ),
        on="probe",
        validate="one_to_one",
    )
    result = {
        "shape": list(df.shape),
        "sha256": hashlib.sha256(raw_bytes).hexdigest(),
        "value_min": float(x.to_numpy().min()),
        "value_max": float(x.to_numpy().max()),
        "group_counts": group_counts,
        "family_size": int(len(discovery)),
        "bonferroni_raw_threshold": 0.05 / len(discovery),
        "counts": counts,
        "bh_max_k": int(counts["bh"]),
        "bh_p_at_k": float(
            discovery.sort_values(["raw", "probe"]).iloc[counts["bh"] - 1].raw
        ),
        "histogram_bins_0.05": hist[:1].tolist(),
        "histogram_bins_0.05_to_0.10": hist[1:2].tolist(),
        "top20_wrong_adjustment": [
            {
                "probe": str(row.probe),
                "raw": float(row.raw),
                "full_bh": float(row.bh),
                "m20_bh": float(row.wrong_bh_m20),
            }
            for row in top20.head(3).itertuples()
        ],
        "candidates": probe_table.to_dict(orient="records"),
        "validation_pass_count": int(validation.replicated.sum()),
        "top_probe_detail": {
            "probe": str(first.probe),
            "n_AML": int(first.n_AML),
            "n_ALL": int(first.n_ALL),
            "mean_AML": float(first.mean_AML),
            "mean_ALL": float(first.mean_ALL),
            "delta": float(first.delta),
            "ci95": [float(first.ci95_low), float(first.ci95_high)],
            "cohen_d": float(first.cohen_d),
            "raw": float(first.raw),
            "bonferroni": float(first.bonferroni),
            "bh": float(first.bh),
        },
        "simulations": {
            "all_null": simulate(),
            "signal_1.0": simulate(delta=1.0, signals=100),
            "signal_0.5": simulate(delta=0.5, signals=100),
        },
    }
    assert counts == {"raw": 1694, "bonferroni": 26, "bh": 535}, counts
    assert result["validation_pass_count"] == 7
    assert abs(result["simulations"]["all_null"]["summary"]["raw"]["mean_R"] - 49.944) < 1e-10
    Path("work/analysis.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
