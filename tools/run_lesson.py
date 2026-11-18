"""Portable archival checks; do not overwrite submitted homework or open locked labels."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import platform
import pickletools
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy
from scipy import stats
from scipy.spatial.distance import pdist
import sklearn
from sklearn.compose import ColumnTransformer
from sklearn.decomposition import PCA
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, balanced_accuracy_score, log_loss, roc_auc_score
from sklearn.model_selection import train_test_split
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.tree import DecisionTreeClassifier
from threadpoolctl import threadpool_limits

ROOT = Path(__file__).resolve().parents[1]
LESSONS = sorted(p.name for p in (ROOT / "lessons").iterdir() if p.is_dir())
SEED = 2026


def save_json(path, data):
    def convert(value):
        if isinstance(value, np.generic):
            return value.item()
        if isinstance(value, np.ndarray):
            return value.tolist()
        raise TypeError(type(value).__name__)

    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=convert) + "\n", encoding="utf-8")


def table(out, name, frame):
    frame.to_csv(out / name, index=False, encoding="utf-8-sig")


def module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def evaluate(y, probability):
    pred = np.asarray(probability) >= 0.5
    return {
        "n": len(y),
        "accuracy": float(accuracy_score(y, pred)),
        "balanced_accuracy": float(balanced_accuracy_score(y, pred)),
        "roc_auc": float(roc_auc_score(y, probability)),
        "log_loss": float(log_loss(y, probability, labels=[0, 1])),
    }


def tools_demo(base, out):
    rng = np.random.default_rng(SEED)
    frame = pd.DataFrame({"sample": np.arange(1, 31), "measurement": rng.normal(10, 2, 30)})
    table(out, "generated-example.csv", frame)
    table(out, "descriptive-statistics.csv", frame[["measurement"]].describe().reset_index())
    return {"data_kind": "synthetic demonstration generated during archive assembly", "n": len(frame)}


def data_types(base, out):
    source = base / "assets/pbl-understanding-data/heart-disease-four-site.csv"
    frame = pd.read_csv(source)
    numeric = frame.select_dtypes("number")
    audit = pd.DataFrame({
        "field": frame.columns, "dtype": [str(t) for t in frame.dtypes],
        "missing": frame.isna().sum().values,
        "unique": frame.nunique().values,
    })
    table(out, "variable-audit.csv", audit)
    table(out, "numeric-summary.csv", numeric.describe().T.reset_index(names="field"))
    site = frame.groupby("site", dropna=False).agg(
        n=("record_id", "size"), missing_thalach=("thalach", lambda s: int(s.isna().sum()))
    ).reset_index()
    table(out, "site-summary.csv", site)
    plt.figure(figsize=(8, 4))
    plt.bar(audit.field, audit.missing, color="#347d86")
    plt.xticks(rotation=60, ha="right")
    plt.ylabel("Missing values")
    plt.tight_layout()
    plt.savefig(out / "missingness.png", dpi=160)
    plt.close()
    return {"n": len(frame), "p": len(frame.columns), "source": source.relative_to(ROOT).as_posix()}


def high_dimension(base, out):
    with np.load(base / "data/cohort120.npz", allow_pickle=False) as archive:
        key = "X"
        x = np.asarray(archive[key], dtype=float)
        labels = archive["y"]
        dev_indices = archive["development"]
    manifest = pd.read_csv(base / "data/patient-manifest.csv")
    assert x.shape[0] == len(manifest)
    # Parse string literals without executing the object-array pickle in the supplied NPZ.
    with zipfile.ZipFile(base / "data/cohort120.npz") as container:
        with container.open("ids.npy") as stream:
            version = np.lib.format.read_magic(stream)
            if version != (1, 0):
                raise ValueError(f"Unexpected NPY header version: {version}")
            np.lib.format.read_array_header_1_0(stream)
            expected = set(manifest.patient_id)
            ids = [arg for op, arg, _ in pickletools.genops(stream.read())
                   if op.name in {"UNICODE", "BINUNICODE", "SHORT_BINUNICODE", "BINUNICODE8"}
                   and isinstance(arg, str) and arg in expected]
    if len(ids) != len(manifest) or len(set(ids)) != len(manifest):
        raise ValueError("Cannot verify expression row IDs against patient manifest")
    manifest = manifest.set_index("patient_id").loc[ids].reset_index()
    development = manifest.split.eq("development").to_numpy()
    np.testing.assert_array_equal(np.sort(dev_indices), np.flatnonzero(development))
    np.testing.assert_array_equal(labels, manifest.target.to_numpy())
    xd = x[development]
    if not np.isfinite(xd).all():
        raise ValueError("Non-finite development expression")
    varying = xd.std(axis=0) > 0
    z = StandardScaler().fit_transform(xd[:, varying])
    pca = PCA(n_components=2, svd_solver="full").fit(z)
    coordinates = pca.transform(z)
    frame = pd.DataFrame(coordinates, columns=["PC1", "PC2"])
    frame.insert(0, "patient_id", manifest.loc[development, "patient_id"].to_numpy())
    table(out, "development-pca.csv", frame)
    rows = []
    for p in [20, 100, 500, 2000, min(5000, z.shape[1])]:
        d = pdist(z[:, :p])
        rows.append({"features": p, "distance_mean": d.mean(), "distance_sd": d.std(ddof=1),
                     "coefficient_of_variation": d.std(ddof=1) / d.mean()})
    table(out, "distance-concentration.csv", pd.DataFrame(rows))
    plt.figure(figsize=(6, 4))
    plt.scatter(frame.PC1, frame.PC2, color="#347d86", s=20)
    plt.xlabel("PC1")
    plt.ylabel("PC2")
    plt.tight_layout()
    plt.savefig(out / "development-pca.png", dpi=160)
    plt.close()
    return {
        "matrix_key": key, "matrix_shape": list(x.shape), "development_n": int(development.sum()),
        "retained_features": int(varying.sum()), "pca_explained_variance_ratio": pca.explained_variance_ratio_,
        "protocol": "Exploratory PCA on development rows only; no UMAP or test evaluation in this smoke check.",
    }


def hypothesis(base, out):
    frame = pd.read_csv(base / "data/heart-disease-hypothesis.csv")
    groups = [frame.loc[frame.disease.eq(y), "thalach"].dropna().to_numpy() for y in [0, 1]]
    no, yes = groups
    result = stats.ttest_ind(yes, no, equal_var=False)
    v1, v0 = yes.var(ddof=1) / len(yes), no.var(ddof=1) / len(no)
    se = np.sqrt(v1 + v0)
    df = (v1 + v0) ** 2 / (v1**2 / (len(yes) - 1) + v0**2 / (len(no) - 1))
    delta = yes.mean() - no.mean()
    half = stats.t.ppf(0.975, df) * se
    table(out, "group-summary.csv", pd.DataFrame([
        {"disease": y, "n": len(values), "mean": values.mean(), "sd": values.std(ddof=1)}
        for y, values in enumerate(groups)
    ]))
    return {
        "raw_n": len(frame), "missing_thalach": int(frame.thalach.isna().sum()),
        "included_n": sum(map(len, groups)), "endpoint": "thalach", "contrast": "disease minus no disease",
        "mean_difference": delta, "welch_t": result.statistic, "degrees_of_freedom": df,
        "p_value_two_sided": result.pvalue, "ci95": [delta - half, delta + half],
        "boundary": "Unadjusted observational group comparison, not a causal or diagnostic claim.",
    }


def bh(pvalues):
    order = np.argsort(pvalues)
    ranked = pvalues[order]
    adjusted = np.minimum.accumulate((ranked * len(ranked) / np.arange(1, len(ranked) + 1))[::-1])[::-1]
    result = np.empty_like(adjusted)
    result[order] = np.minimum(1, adjusted)
    return result


def multiple_testing(base, out):
    frame = pd.read_csv(base / "data/leukemia.csv")
    train = frame[frame.original_split.eq("train")]
    features = [c for c in frame if c not in {"sample_id", "original_split", "diagnosis"}]
    labels = sorted(train.diagnosis.unique())
    if len(labels) != 2:
        raise ValueError(f"Expected two diagnostic groups, found {labels}")
    a = train.loc[train.diagnosis.eq(labels[0]), features].to_numpy(dtype=float)
    b = train.loc[train.diagnosis.eq(labels[1]), features].to_numpy(dtype=float)
    test = stats.ttest_ind(a, b, axis=0, equal_var=False)
    p = test.pvalue
    if not np.isfinite(p).all():
        raise ValueError("Non-finite p-values; define constant-feature policy before proceeding")
    adjusted = bh(p)
    results = pd.DataFrame({
        "feature": features, "mean_difference": a.mean(axis=0) - b.mean(axis=0),
        "welch_t": test.statistic, "raw_p": p, "bonferroni_p": np.minimum(1, p * len(p)),
        "bh_adjusted_p": adjusted,
    })
    table(out, "training-tests.csv", results.sort_values("raw_p"))
    rng = np.random.default_rng(SEED)
    null = rng.uniform(size=(1000, 1000))
    null_counts = np.sum(null < 0.05, axis=1)
    table(out, "null-simulation.csv", pd.DataFrame({"replicate": np.arange(1000), "raw_false_positives": null_counts}))
    return {
        "training_n": len(train), "features": len(features), "contrast": f"{labels[0]} minus {labels[1]}",
        "discoveries": {"raw": int((p < .05).sum()), "bonferroni": int((p < .05 / len(p)).sum()),
                        "bh": int((adjusted < .05).sum())},
        "null_simulation": {"mean_false_positives": null_counts.mean(), "repetitions": 1000, "tests_per_repetition": 1000},
        "protocol": "Discovery split only. Existing held-out results remain historical; no new test evaluation.",
    }


def naive_bayes(base, out):
    lab = module(base / "scripts/lab_utils.py", "nb_lab")
    train = pd.read_csv(base / "data/student/train.csv")
    valid = pd.read_csv(base / "data/student/validation.csv")
    assert set(train.sample_id).isdisjoint(valid.sample_id)
    model = lab.fit_model(train)
    probability = model.predict_proba(valid[lab.FEATURES])[:, 1]
    thresholds = lab.threshold_table(valid.bladder_inflammation, probability)
    table(out, "validation-thresholds.csv", thresholds)
    table(out, "training-dependence.csv", lab.dependence(train))
    return {"training_n": len(train), "validation_n": len(valid),
            "selected_threshold": lab.choose_threshold(thresholds),
            "validation_at_0_5": lab.scores(valid.bladder_inflammation, probability),
            "protocol": "Training and validation only; locked labels are not read."}


def linear_models(base, out):
    source = base / "data/heart.csv"
    frame = pd.read_csv(source)
    target = "heart_disease_present"
    continuous = ["age", "trestbps", "chol", "thalach", "oldpeak"]
    categorical = ["sex", "cp", "fbs", "restecg", "exang", "slope", "ca", "thal"]
    train, valid = train_test_split(frame, test_size=.2, stratify=frame[target], random_state=SEED)
    rows = []
    for name, estimator in [
        ("logistic", LogisticRegression(max_iter=5000, C=1, random_state=SEED)),
        ("lda_shrinkage", LinearDiscriminantAnalysis(solver="lsqr", shrinkage="auto")),
    ]:
        prep = ColumnTransformer([
            ("numeric", make_pipeline(SimpleImputer(strategy="median"), StandardScaler()), continuous),
            ("categorical", make_pipeline(SimpleImputer(strategy="most_frequent"),
                                         OneHotEncoder(handle_unknown="ignore", sparse_output=False)), categorical),
        ])
        model = make_pipeline(prep, estimator).fit(train[continuous + categorical], train[target])
        probability = model.predict_proba(valid[continuous + categorical])[:, 1]
        rows.append({"model": name, **evaluate(valid[target], probability)})
    table(out, "new-validation-comparison.csv", pd.DataFrame(rows))
    assets = base / "assets/pbl-linear-models-lda-logistic-regression/statlog-heart.csv"
    alternate = pd.read_csv(assets)
    return {
        "training_n": len(train), "validation_n": len(valid), "raw_missing_cells": int(frame.isna().sum().sum()),
        "alternate_dataset_rows": len(alternate), "alternate_dataset_sha256": hashlib.sha256(assets.read_bytes()).hexdigest(),
        "primary_dataset_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "protocol": "New archival baseline on data/heart.csv, NOT a recreation of historical statlog-heart results.",
        "boundary": "Existing assets use a different dataset. Labels were previously available; this is not a new blind test.",
    }


def svm(base, out):
    lab = module(base / "scripts/lab_utils.py", "svm_lab")
    train, valid = lab.load_development()
    assert set(train.sample_id).isdisjoint(valid.sample_id)
    grid, models = lab.parameter_experiment(train, valid)
    configs = lab.select_validation_configs(grid)
    table(out, "validation-grid.csv", grid)
    selected = {}
    for family, key in configs.items():
        model = models[key]
        selected[family] = {
            "configuration": str(key),
            "metrics": lab.metrics(valid.abnormal, model.decision_function(valid[lab.FEATURES])),
        }
    return {"training_n": len(train), "validation_n": len(valid), "selected": selected,
            "protocol": "Uses provided SVM helper, training scaling and predeclared validation rule. No locked-test reads."}


def ensemble(base, out):
    train = pd.read_csv(base / "data/student/train.csv")
    valid = pd.read_csv(base / "data/student/validation.csv")
    features = [c for c in train if c not in {"sample_id", "malignant"}]
    assert set(train.sample_id).isdisjoint(valid.sample_id)
    models = {
        "tree": DecisionTreeClassifier(random_state=SEED),
        "random_forest": RandomForestClassifier(n_estimators=100, n_jobs=1, random_state=SEED),
        "gradient_boosting": GradientBoostingClassifier(random_state=SEED),
    }
    rows = []
    for name, model in models.items():
        model.fit(train[features], train.malignant)
        rows.append({"model": name, **evaluate(valid.malignant, model.predict_proba(valid[features])[:, 1])})
    table(out, "validation-models.csv", pd.DataFrame(rows))
    return {"training_n": len(train), "validation_n": len(valid), "features": len(features),
            "protocol": "Fixed archival baselines on WDBC. No hyperparameter search and no locked-test reads."}


FUNCTIONS = {
    "01-tools": tools_demo, "02-data-types": data_types,
    "03-04-high-dimensional-reduction": high_dimension, "05-hypothesis": hypothesis,
    "06-multiple-testing": multiple_testing, "07-naive-bayes": naive_bayes,
    "08-linear-models": linear_models, "09-svm": svm, "10-ensemble": ensemble,
}


def run(lesson):
    base = ROOT / "lessons" / lesson
    out = base / "outputs/reproduced"
    out.mkdir(parents=True, exist_ok=True)
    with threadpool_limits(limits=1):
        result = FUNCTIONS[lesson](base, out)
    result["archive_check"] = True
    result["seed"] = SEED
    result["locked_test_label_files_opened"] = False
    result["held_out_evaluation_performed"] = False
    result["versions"] = {
        "python": platform.python_version(), "numpy": np.__version__, "pandas": pd.__version__,
        "scipy": scipy.__version__, "scikit-learn": sklearn.__version__, "matplotlib": matplotlib.__version__,
    }
    save_json(out / "results.json", result)
    print(f"PASS {lesson}: {out.relative_to(ROOT).as_posix()}", flush=True)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("lesson", choices=[*LESSONS, "all"])
    args = parser.parse_args()
    for lesson in LESSONS if args.lesson == "all" else [args.lesson]:
        run(lesson)


if __name__ == "__main__":
    main()
