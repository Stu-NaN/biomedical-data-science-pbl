from pathlib import Path
import sys
import json
import platform
import zipfile
import warnings
import numpy as np
import pandas as pd
import scipy
from scipy.stats import norm
from scipy.spatial.distance import pdist
import sklearn
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle

warnings.filterwarnings("ignore", category=FutureWarning)
ROOT = Path(__file__).parent
BASE = Path(r"C:\Users\A\AppData\Local\Temp\svm-review-20261111-a\9. SVM")
sys.path.insert(0, str(BASE / "scripts"))
import lab_utils as lab

train, val = lab.load_development()
geo = lab.geometry_experiment(train)
grid, models = lab.parameter_experiment(train, val)
scale, deltas = lab.scaling_experiment(train, val)
configs = lab.select_validation_configs(grid)
frozen = {k: models[v] for k, v in configs.items()}
noise = {k: lab.noise_experiment(m, train, val) for k, m in frozen.items()}

# Only development IDs are joined to the public subtype labels.
development_ids = set(train.sample_id) | set(val.sample_id)
subtype = {}
with zipfile.ZipFile(BASE / "data/raw/vertebral-column.zip") as z:
    name = next(n for n in z.namelist() if n.endswith("column_3C.dat"))
    for i, line in enumerate(z.read(name).decode().splitlines(), 1):
        sid = f"V{i:03d}"
        if sid in development_ids:
            parts = line.split()
            subtype[sid] = parts[-1]
            row = pd.concat([train, val]).set_index("sample_id").loc[sid]
            assert np.allclose([float(parts[0]), float(parts[5])], row[lab.FEATURES].to_numpy(dtype=float))
for d in [train, val]:
    d["subtype"] = d.sample_id.map(subtype)

metrics = {}
for key, model in frozen.items():
    score = model.decision_function(val[lab.FEATURES])
    metrics[key] = lab.metrics(val.abnormal, score)
    metrics[key]["subtype_correct"] = {
        sub: int(np.sum(((score >= 0) == val.abnormal) & (val.subtype == sub)))
        for sub in ["NO", "DH", "SL"]
    }

linear = frozen["linear"]
svc = lab.svc_of(linear)
zv = linear.named_steps["standardscaler"].transform(val[lab.FEATURES])
w = svc.coef_[0]
dist = np.abs(linear.decision_function(val[lab.FEATURES])) / np.linalg.norm(w)
theory = [0] + [float(norm.cdf(-dist / s).mean()) for s in [.05, .1, .2]]
linear_scores = linear.decision_function(val[lab.FEATURES])
score_bins = []
for label, mask in [
    ("f<−1", linear_scores < -1),
    ("−1≤f<0", (linear_scores >= -1) & (linear_scores < 0)),
    ("0≤f<1", (linear_scores >= 0) & (linear_scores < 1)),
    ("f≥1", linear_scores >= 1),
]:
    count = int(mask.sum())
    abnormal = int(val.loc[mask, "abnormal"].sum())
    score_bins.append(dict(label=label, n=count, abnormal=abnormal,
                           fraction=abnormal / count))
small_gamma = []
for C in [1., 10., 100.]:
    linear_approx = lab.make_model("linear", C=.02*C).fit(train[lab.FEATURES], train.abnormal)
    rbf_score = models[("rbf", C, .01)].decision_function(val[lab.FEATURES])
    approx_score = linear_approx.decision_function(val[lab.FEATURES])
    small_gamma.append(dict(rbf_C=C, gamma=.01, equivalent_linear_C=.02*C,
                            prediction_disagreements=int(np.sum((rbf_score >= 0) != (approx_score >= 0))),
                            score_correlation=float(np.corrcoef(rbf_score, approx_score)[0, 1]),
                            linear_validation=lab.metrics(val.abnormal, approx_score)))
zt = linear.named_steps["standardscaler"].transform(train[lab.FEATURES])
extreme_index = int(np.flatnonzero(train.sample_id.to_numpy() == "V116")[0])
extreme_gamma_d2 = .01 * np.sum((zt - zt[extreme_index])**2, axis=1)
new_dh = val[(val.subtype == "DH") & (linear_scores < 0) &
             (frozen["rbf"].decision_function(val[lab.FEATURES]) >= 0)]

result = {
    "environment": {"python": platform.python_version(), "numpy": np.__version__, "pandas": pd.__version__,
                    "scipy": scipy.__version__, "sklearn": sklearn.__version__, "matplotlib": matplotlib.__version__},
    "configs": configs, "metrics": metrics, "scaling_deltas": deltas,
    "noise": {k: d.to_dict("records") for k, d in noise.items()},
    "noise_theory_linear": theory,
    "geometry": geo["table"].to_dict("records"),
    "candidates": geo["candidates"].to_dict("records"),
    "deletion": geo["changes"].to_dict("records"),
    "contradiction": geo["contradiction"].to_dict("records"),
    "grid": grid.to_dict("records"), "scaling": scale.to_dict("records"),
    "score_bins_linear": score_bins,
    "training_subtype_medians": train.groupby("subtype").degree_spondylolisthesis.median().to_dict(),
    "small_gamma_checks": small_gamma,
    "geometry_diagnostics": {
        "training_pair_distance_median": float(np.median(pdist(zt))),
        "V116_z": zt[extreme_index].tolist(),
        "V116_gamma_d2_median": float(np.median(np.delete(extreme_gamma_d2, extreme_index))),
        "cropped_training_count": int(np.any((train[lab.FEATURES].to_numpy() < [20, -15]) |
                                             (train[lab.FEATURES].to_numpy() > [100, 105]), axis=1).sum()),
    },
    "newly_correct_dh": new_dh[["sample_id", *lab.FEATURES]].to_dict("records"),
}
(ROOT / "verified_results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
grid.to_csv(ROOT / "verified_grid.csv", index=False)
print(json.dumps({k: result[k] for k in ["environment", "configs", "metrics", "scaling_deltas", "noise", "noise_theory_linear"]}, ensure_ascii=False, indent=2))

OUT = ROOT / "figures"
OUT.mkdir(exist_ok=True)
plt.rcParams.update({
    "font.family": "Microsoft YaHei", "font.size": 10, "axes.titlesize": 11,
    "axes.labelsize": 10, "legend.fontsize": 9, "xtick.labelsize": 9,
    "ytick.labelsize": 9, "axes.spines.top": False, "axes.spines.right": False,
    "axes.unicode_minus": False, "savefig.dpi": 320,
})
COL = {0: "#247A94", 1: "#D2674E"}
SUB = {"NO": ("#247A94", "o", "NO 正常"), "DH": ("#AA7C13", "s", "DH 椎间盘突出"),
       "SL": ("#BE4D65", "^", "SL 脊椎滑脱")}

def save(fig, name):
    fig.savefig(OUT / name, facecolor="white", bbox_inches="tight", pad_inches=.08)
    plt.close(fig)

def points(ax, d, X=None, support=None, alpha=.72):
    X = d[lab.FEATURES].to_numpy() if X is None else X
    for cls, mark in [(0, "o"), (1, "^")]:
        mask = d.abnormal.to_numpy() == cls
        ax.scatter(X[mask, 0], X[mask, 1], s=19, c=COL[cls], marker=mark, alpha=alpha, linewidths=.3, edgecolors="white")
    if support is not None:
        ax.scatter(X[support, 0], X[support, 1], s=57, facecolors="none", edgecolors="#262626", linewidths=.6)
    ax.grid(alpha=.12, linewidth=.5)
    ax.set_axisbelow(True)

handles = [Line2D([], [], marker="o", color=COL[0], linestyle="", label="NO 正常"),
           Line2D([], [], marker="^", color=COL[1], linestyle="", label="AB 异常"),
           Line2D([], [], color="#222222", label="决策边界 f = 0"),
           Line2D([], [], color="#777777", linestyle="--", label="参考线 f = ±1")]

def boundary(ax, model, title, d=train, limits=((20, 100), (-15, 105)), sv=False):
    xlim, ylim = limits
    xx, yy = np.meshgrid(np.linspace(*xlim, 360), np.linspace(*ylim, 360))
    inp = pd.DataFrame(np.c_[xx.ravel(), yy.ravel()], columns=lab.FEATURES)
    score = model.decision_function(inp).reshape(xx.shape)
    ax.contourf(xx, yy, score >= 0, [-.5, .5, 1.5], colors=["#E7F1F5", "#FBEEE8"], alpha=.75)
    if score.min() < 0 < score.max():
        ax.contour(xx, yy, score, [0], colors="#242424", linewidths=1.25)
    levels = [v for v in [-1, 1] if score.min() < v < score.max()]
    if levels:
        ax.contour(xx, yy, score, levels, colors="#777777", linewidths=.7, linestyles="--")
    points(ax, d, support=lab.svc_of(model).support_ if sv else None)
    ax.set(xlim=xlim, ylim=ylim, title=title, xlabel="骨盆入射角（源尺度）")

# Overview retains the extreme measurement, then gives a central-region enlargement.
fig, axes = plt.subplots(1, 2, figsize=(7, 3.15))
points(axes[0], train)
axes[0].set(xlabel="骨盆入射角（源尺度）", ylabel="滑脱程度（源尺度）", title="a  全部训练数据  n = 186")
extreme = train.loc[train.degree_spondylolisthesis.idxmax()]
axes[0].annotate(extreme.sample_id, xy=extreme[lab.FEATURES].to_numpy(dtype=float),
                 xytext=(-36, -2), textcoords="offset points", fontsize=9)
points(axes[1], train)
axes[1].set(xlim=(20, 100), ylim=(-15, 105), xlabel="骨盆入射角（源尺度）", title="b  主要重叠区域")
fig.legend(handles=handles[:2], loc="lower center", ncol=2, frameon=False)
fig.tight_layout(rect=(0, .08, 1, 1))
save(fig, "fig1.png")

# Geometry uses equal aspect; all eight identifiers remain readable.
g = geo["table"]; z = geo["z"]; gm = geo["model"]
gw, gb = gm.coef_[0], gm.intercept_[0]
fig, axes = plt.subplots(1, 2, figsize=(7, 3.5), sharex=True, sharey=True)
xx = np.linspace(z[:, 0].min()-.25, z[:, 0].max()+.25, 200)
for offset, color in zip([-.6, 0, .6], ["#A04E6B", "#222222", "#C98A25"]):
    axes[0].plot(xx, -(gw[0]*xx+gb+offset)/gw[1], color=color, lw=1.2, label=f"截距偏移 {offset:+.1f}")
for a in axes:
    points(a, geo["subset"], z, gm.support_, alpha=1)
    a.set(xlabel="z1 标准化骨盆入射角", aspect="equal",
          xlim=(z[:, 0].min()-.3, z[:, 0].max()+.3), ylim=(z[:, 1].min()-.3, z[:, 1].max()+.3))
    for i, row in g.iterrows():
        if row.sample_id in ["V061", "V217", "V072"]:
            offset = (14, 0) if row.sample_id == "V217" else (8, 6)
            a.annotate(row.sample_id, (z[i, 0], z[i, 1]), xytext=offset,
                       textcoords="offset points", fontsize=8.5,
                       va="center" if row.sample_id == "V217" else "bottom",
                       bbox=dict(facecolor="white", edgecolor="none", alpha=.94, pad=.12))
for level, style in [(0, "-"), (-1, "--"), (1, "--")]:
    axes[1].plot(xx, -(gw[0]*xx+gb-level)/gw[1], color="#333333", linestyle=style, lw=1.1)
zm, zp = z[gm.support_[0]], z[gm.support_[1]]
axes[1].annotate("", xy=zp, xytext=zm, arrowprops=dict(arrowstyle="<->", color="#B44949", lw=1.3))
axes[0].set(title="a  三条线均为零训练错误", ylabel="z2 标准化滑脱程度")
axes[1].set(title="b  最大间隔由两个支持向量约束")
axes[0].legend(loc="upper left", fontsize=7.5, frameon=True,
               facecolor="white", edgecolor="none", framealpha=.96)
fig.tight_layout()
save(fig, "fig2.png")

fig, axes = plt.subplots(1, 3, figsize=(7, 3))
for a, C in zip(axes, [.01, 1, 100]):
    row = grid[(grid.kernel == "linear") & (grid.C == C)].iloc[0]
    boundary(a, models[("linear", C, .1)], f"C = {C:g}\n验证 BA = {row.val_balanced_accuracy:.3f}")
axes[0].set_ylabel("滑脱程度（源尺度）")
fig.legend(handles=handles, loc="lower center", ncol=4, frameon=False, fontsize=7.5)
fig.tight_layout(rect=(0, .12, 1, 1))
save(fig, "fig3.png")

fig, axes = plt.subplots(2, 3, figsize=(7, 5.2), sharex=True, sharey=True)
for a, (C, gamma) in zip(axes.flat, [(1,.01),(1,1),(1,10),(1,10),(10,10),(100,10)]):
    row = grid[(grid.kernel=="rbf") & (grid.C==C) & (grid.gamma==gamma)].iloc[0]
    boundary(a, models[("rbf", C, gamma)], f"C = {C:g}   γ = {gamma:g}\n训练 {row.train_balanced_accuracy:.3f} / 验证 {row.val_balanced_accuracy:.3f}")
for a in axes[0]:
    a.set_xlabel("")
axes[0, 0].set_ylabel("滑脱程度（源尺度）")
axes[1, 0].set_ylabel("滑脱程度（源尺度）")
fig.legend(handles=handles, loc="lower center", ncol=4, frameon=False, fontsize=7.5)
fig.tight_layout(rect=(0, .06, 1, 1), h_pad=1.4)
save(fig, "fig4.png")

fig, axes = plt.subplots(1, 2, figsize=(7, 3.2))
rbf = grid[grid.kernel=="rbf"]
for a, metric, title in zip(axes, ["train_balanced_accuracy", "val_balanced_accuracy"], ["a  训练平衡准确率", "b  验证平衡准确率"]):
    arr = rbf.pivot(index="C", columns="gamma", values=metric).loc[lab.RBF_C, lab.GAMMA_GRID].to_numpy()
    im = a.imshow(arr, vmin=.48, vmax=.92, cmap="YlGnBu")
    for i in range(4):
        for j in range(4):
            a.text(j, i, f"{arr[i,j]:.3f}", ha="center", va="center", color="white" if arr[i,j]>.78 else "#111111", fontsize=10)
    a.set(xticks=range(4), xticklabels=[".01", ".1", "1", "10"],
          yticks=range(4), yticklabels=[".1", "1", "10", "100"], xlabel="gamma", ylabel="C", title=title)
    a.add_patch(Rectangle((2.5, 1.5), 1, 1, fill=False, ec="#D04B41", lw=2))
fig.subplots_adjust(left=.07, right=.82, bottom=.17, top=.87, wspace=.32)
cax = fig.add_axes([.865, .22, .025, .57])
fig.colorbar(im, cax=cax, label="平衡准确率")
save(fig, "fig5.png")

fig, axes = plt.subplots(1, 3, figsize=(7, 3))
for a, kernel, title in zip(axes, ["linear","poly","rbf"], ["线性核", "二次多项式核", "RBF 核"]):
    row = grid[(grid.kernel==kernel)&(grid.C==1)&(grid.gamma==.1)].iloc[0]
    boundary(a, models[(kernel,1,.1)], f"{title}\n验证 BA = {row.val_balanced_accuracy:.3f}")
axes[0].set_ylabel("滑脱程度（源尺度）")
fig.legend(handles=handles, loc="lower center", ncol=4, frameon=False, fontsize=7.5)
fig.tight_layout(rect=(0, .12, 1, 1))
save(fig, "fig6.png")

# Error analysis zoom shows subtypes and mistakes without support-vector clutter.
fig, axes = plt.subplots(1, 2, figsize=(7, 3.3), sharex=True, sharey=True)
for a, (key, model) in zip(axes, frozen.items()):
    xx, yy = np.meshgrid(np.linspace(20,100,360), np.linspace(-15,47,300))
    ss = model.decision_function(pd.DataFrame(np.c_[xx.ravel(),yy.ravel()], columns=lab.FEATURES)).reshape(xx.shape)
    a.contourf(xx, yy, ss>=0, [-.5,.5,1.5], colors=["#E7F1F5","#FBEEE8"], alpha=.8)
    a.contour(xx,yy,ss,[0],colors="#242424",linewidths=1.2)
    for sub, (color, mark, label) in SUB.items():
        mask = val.subtype==sub
        a.scatter(val.loc[mask,lab.FEATURES[0]],val.loc[mask,lab.FEATURES[1]],s=30,c=color,marker=mark,edgecolors="white",linewidths=.4,label=label)
    error = (model.decision_function(val[lab.FEATURES])>=0) != val.abnormal
    a.scatter(val.loc[error,lab.FEATURES[0]],val.loc[error,lab.FEATURES[1]],s=80,facecolors="none",edgecolors="#B5222C",linewidths=1)
    a.set(xlim=(20,100),ylim=(-15,47),xlabel="骨盆入射角（源尺度）",
          title="a  线性核  C = 1" if key=="linear" else "b  RBF  C = 10  γ = 10")
axes[0].set_ylabel("滑脱程度（源尺度）")
h,l = axes[0].get_legend_handles_labels()
h.append(Line2D([],[],marker="o",markersize=8,markerfacecolor="none",markeredgecolor="#B5222C",linestyle=""));l.append("预测错误")
fig.legend(h,l,loc="lower center",ncol=4,frameon=False,fontsize=8)
fig.tight_layout(rect=(0,.1,1,1))
save(fig,"fig7.png")

fig, axes = plt.subplots(1,2,figsize=(7,2.9))
for key, color, label in [("linear",COL[0],"线性核 C = 1"),("rbf",COL[1],"RBF C = 10  γ = 10")]:
    d=noise[key]
    axes[0].plot(d.noise_sd_fraction*100,d.mean_flip_rate*100,"o-",color=color,label=label)
    axes[1].errorbar(d.noise_sd_fraction*100,d.mean_balanced_accuracy,
                     yerr=d.sd_balanced_accuracy,fmt="o-",color=color,
                     capsize=3,elinewidth=.9,label=label)
axes[0].plot([0,5,10,20],np.asarray(theory)*100,"--",color="#333333",lw=1,label="线性核理论翻转率")
axes[0].set(ylabel="预测翻转率（%）",title="a  预测是否改变")
axes[1].set(ylabel="验证平衡准确率",title="b  分类表现是否保持",ylim=(.62,.83))
for a in axes:
    a.set(xlabel="噪声标准差 / 训练标准差（%）",xticks=[0,5,10,20])
    a.grid(alpha=.15)
fig.legend(*axes[0].get_legend_handles_labels(),loc="lower center",ncol=3,frameon=False,fontsize=8)
fig.tight_layout(rect=(0,.12,1,1))
save(fig,"fig8.png")
print("Figures saved:", OUT)
