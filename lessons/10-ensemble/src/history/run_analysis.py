from pathlib import Path
import hashlib
import importlib.util
import json
import platform
import time

import numpy as np
import pandas as pd
import sklearn
from sklearn.base import clone
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.metrics import accuracy_score, balanced_accuracy_score, log_loss, roc_auc_score
from sklearn.tree import DecisionTreeClassifier


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "10. Ensemble"
OUT = ROOT / "analysis"
OUT.mkdir(exist_ok=True)
spec = importlib.util.spec_from_file_location("course_demo", SOURCE / "scripts/demo_ensemble.py")
demo = importlib.util.module_from_spec(spec)
spec.loader.exec_module(demo)
SEED = 1010


def save(frame, name):
    frame.to_csv(OUT / name, index=False, encoding="utf-8-sig")
    print(f"Saved {name}", flush=True)
    return frame


def main():
    started = time.time()
    train, valid, features, label, _ = demo.load("wdbc")
    X, y, Xv, yv = train[features], train[label], valid[features], valid[label]
    dev = pd.concat([train, valid], ignore_index=True)
    info = {
        "python": platform.python_version(),
        "sklearn": sklearn.__version__,
        "numpy": np.__version__,
        "pandas": pd.__version__,
        "seed": SEED,
        "dataset": "WDBC",
        "train_n": len(train),
        "train_positive": int(y.sum()),
        "validation_n": len(valid),
        "validation_positive": int(yv.sum()),
        "features": features,
        "missing_development": int(dev[features].isna().sum().sum()),
        "overlapping_ids": len(set(train.sample_id) & set(valid.sample_id)),
        "test_labels_opened": False,
        "input_sha256": {
            n: hashlib.sha256((SOURCE / "data/student" / n).read_bytes()).hexdigest()
            for n in ("train.csv", "validation.csv")
        },
    }
    (OUT / "run-info.json").write_text(json.dumps(info, indent=2), encoding="utf-8")
    save(demo.fit_table(X, y, Xv, yv), "model-comparison.csv")
    save(demo.development_cv(dev[features], dev[label], None), "cross-validation.csv")
    rows = []
    for depth in (1, 2, 3, 4, None):
        model = DecisionTreeClassifier(max_depth=depth, random_state=SEED).fit(X, y)
        row = {"max_depth": str(depth), "leaves": model.get_n_leaves()}
        for prefix, xx, yy in (("train", X, y), ("validation", Xv, yv)):
            row[prefix + "_accuracy"] = accuracy_score(yy, model.predict(xx))
            row[prefix + "_auc"] = roc_auc_score(yy, model.predict_proba(xx)[:, 1])
            row[prefix + "_ba"] = balanced_accuracy_score(yy, model.predict(xx))
        rows.append(row)
    save(pd.DataFrame(rows), "tree-depth.csv")

    rng = np.random.default_rng(SEED)
    resamples = [rng.choice(len(y), len(y), replace=True) for _ in range(30)]
    rows = []
    for i, idx in enumerate(resamples[:10]):
        model = DecisionTreeClassifier(max_depth=2, random_state=SEED).fit(X.iloc[idx], y.iloc[idx])
        rows.append({"replicate": i + 1, "root_feature": features[model.tree_.feature[0]],
                     "root_threshold": model.tree_.threshold[0]})
    save(pd.DataFrame(rows), "bootstrap-roots.csv")

    rows = []
    for n in (5, len(y)):
        rng = np.random.default_rng(SEED)
        fractions = np.array([len(np.unique(rng.choice(n, n, replace=True))) / n
                              for _ in range(1000)])
        rows.append({"n": n, "replicates": 1000,
                     "unique_mean": fractions.mean(), "unique_sd": fractions.std(ddof=1),
                     "oob_mean": (1 - fractions).mean(),
                     "unique_theory": 1 - (1 - 1 / n) ** n,
                     "oob_theory": (1 - 1 / n) ** n})
    save(pd.DataFrame(rows), "bootstrap-fractions.csv")

    rows, samples = [], {}
    for name in ("single_tree", "bagging", "random_forest", "logistic", "gradient_boosting"):
        probs = np.array([clone(demo.models()[name]).fit(X.iloc[idx], y.iloc[idx])
                          .predict_proba(Xv)[:, 1] for idx in resamples])
        share = (probs >= .5).mean(axis=0)
        row = {"model": name, "prediction_sd_mean": probs.std(axis=0).mean(),
               "prediction_sd_max": probs.std(axis=0).max(),
               "class_flip_rate": ((share > 0) & (share < 1)).mean(),
               "class_flip_count": int(((share > 0) & (share < 1)).sum())}
        rows.append(row)
        samples[name] = probs
        print(f"Stability finished {name}", flush=True)
    save(pd.DataFrame(rows), "prediction-stability.csv")
    np.savez_compressed(OUT / "stability-predictions.npz", **samples)
    save(demo.tree_count_curve(X, y, Xv, yv), "tree-count-curve.csv")
    save(demo.feature_importance(X, y, Xv, yv), "feature-importance.csv")
    X.corr().to_csv(OUT / "train-feature-correlation.csv", encoding="utf-8-sig")

    rows = []
    for lr in (.01, .1, .5):
        model = GradientBoostingClassifier(n_estimators=500, learning_rate=lr, max_depth=3,
                                           random_state=SEED).fit(X, y)
        for i, (pt, pv) in enumerate(zip(model.staged_predict_proba(X),
                                        model.staged_predict_proba(Xv)), start=1):
            rows.append({"learning_rate": lr, "n_estimators": i,
                         "train_log_loss": log_loss(y, pt),
                         "validation_log_loss": log_loss(yv, pv),
                         "validation_auc": roc_auc_score(yv, pv[:, 1])})
        print(f"Boosting finished {lr}", flush=True)
    save(pd.DataFrame(rows), "boosting-curve.csv")
    info["elapsed_seconds"] = time.time() - started
    (OUT / "run-info.json").write_text(json.dumps(info, indent=2), encoding="utf-8")
    print("COMPLETE", json.dumps(info, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
