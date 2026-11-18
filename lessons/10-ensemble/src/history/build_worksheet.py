from pathlib import Path
import json
import shutil

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
A = ROOT / "analysis"
FIG = ROOT / "figures"
FIG.mkdir(exist_ok=True)
OUT = Path(r"C:\Users\A\Downloads\Ensemble-student-worksheet-completed.docx")
INFO = json.loads((A / "run-info.json").read_text(encoding="utf-8"))
MODELS = pd.read_csv(A / "model-comparison.csv").set_index("model")
CV = pd.read_csv(A / "cross-validation.csv").set_index("model")
STAB = pd.read_csv(A / "prediction-stability.csv").set_index("model")
DEPTH = pd.read_csv(A / "tree-depth.csv", keep_default_na=False)
ROOTS = pd.read_csv(A / "bootstrap-roots.csv")
BOOT = pd.read_csv(A / "bootstrap-fractions.csv").set_index("n")
COUNT = pd.read_csv(A / "tree-count-curve.csv")
IMP = pd.read_csv(A / "feature-importance.csv")
GB = pd.read_csv(A / "boosting-curve.csv")
CORR = pd.read_csv(A / "train-feature-correlation.csv", index_col=0)
NAMES = {
    "single_tree": "单树 深度4",
    "bagging": "Bagging 100棵",
    "random_forest": "随机森林 100棵",
    "adaboost": "AdaBoost 50轮",
    "gradient_boosting": "梯度提升 100轮",
    "logistic": "逻辑回归",
    "rbf_svm": "RBF-SVM",
}
COLORS = ["#176658", "#A14450", "#516677"]
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.labelcolor": "#222222",
    "axes.edgecolor": "#777777",
    "xtick.color": "#444444",
    "ytick.color": "#444444",
    "savefig.facecolor": "white",
})


def savefig(fig, name):
    fig.savefig(FIG / name, dpi=250, bbox_inches="tight")
    plt.close(fig)


def equation(text, name, figsize=(6.5, .45)):
    fig, ax = plt.subplots(figsize=figsize)
    ax.axis("off")
    ax.text(.5, .5, text, ha="center", va="center", fontsize=13)
    savefig(fig, name)


def figures():
    equation(r"$G=1-\sum_{k=0}^{1}p_k^2,\qquad"
             r"\Delta G=G_{\rm parent}-\frac{n_L}{n}G_L-\frac{n_R}{n}G_R$",
             "gini.png")
    equation(r"$\mathrm{Var}(\bar f_B)=\rho\sigma^2+\frac{1-\rho}{B}\sigma^2"
             r"\ \longrightarrow\ \rho\sigma^2\quad(B\to\infty)$", "variance.png")
    equation(r"$E(U/n)=1-(1-1/n)^n,\qquad E(O/n)=(1-1/n)^n$",
             "bootstrap.png")
    equation(r"$-\frac{\partial\ell(y_i,F_i)}{\partial F_i}=y_i-p_i,"
             r"\qquad F_m(x)=F_{m-1}(x)+\eta h_m(x)$", "gradient.png")

    fig, axes = plt.subplots(1, 2, figsize=(7.1, 2.65), constrained_layout=True)
    for i, name in enumerate(("bagging", "random_forest")):
        part = COUNT[COUNT.model == name]
        axes[0].plot(part.n_trees, part.validation_roc_auc, "o-", ms=4,
                     color=COLORS[i], label=("Bagging", "Random forest")[i])
    axes[0].set(xscale="log", xlabel="Number of trees", ylabel="Validation ROC-AUC")
    axes[0].set_xticks([1, 10, 50, 100, 300], ["1", "10", "50", "100", "300"])
    axes[0].set_ylim(.91, 1.005)
    axes[0].grid(axis="y", alpha=.2)
    axes[0].legend(frameon=False, fontsize=8, loc="lower right")
    part = STAB.loc[["single_tree", "bagging", "random_forest"]]
    axes[1].bar(["Single tree", "Bagging", "Random forest"],
                part.prediction_sd_mean, color=[COLORS[2], COLORS[0], COLORS[1]], width=.55)
    for i, val in enumerate(part.prediction_sd_mean):
        axes[1].text(i, val + .002, f"{val:.3f}", ha="center", fontsize=9)
    axes[1].set(ylabel="Mean prediction SD", ylim=(0, float(part.prediction_sd_mean.max()) * 1.25))
    axes[1].tick_params(axis="x", labelsize=8)
    axes[1].grid(axis="y", alpha=.2)
    axes[1].set_axisbelow(True)
    savefig(fig, "ensemble-performance.png")

    fig, axes = plt.subplots(1, 2, figsize=(7.1, 2.9), constrained_layout=True)
    for ax, column, title, color in (
        (axes[0], "impurity_importance", "Impurity importance (MDI)", COLORS[0]),
        (axes[1], "permutation_auc_drop_mean", "Permutation AUC decrease", COLORS[1]),
    ):
        part = IMP.nlargest(6, column).sort_values(column)
        errs = part.permutation_auc_drop_sd if column.startswith("permutation") else None
        ax.barh(part.feature, part[column], xerr=errs, color=color, height=.6,
                error_kw={"lw": .7, "capsize": 2, "ecolor": "#555555"})
        ax.set_title(title, fontsize=10, loc="left", pad=10)
        ax.tick_params(axis="y", labelsize=7.5)
        ax.grid(axis="x", alpha=.15)
        ax.set_axisbelow(True)
    savefig(fig, "feature-importance.png")

    fig, axes = plt.subplots(1, 2, figsize=(7.1, 2.65), constrained_layout=True)
    for i, lr in enumerate((.01, .1, .5)):
        part = GB[GB.learning_rate == lr]
        axes[0].plot(part.n_estimators, part.validation_log_loss,
                     color=COLORS[i], label=f"Validation lr={lr:g}", lw=1.35)
        axes[1].plot(part.n_estimators, part.validation_auc,
                     color=COLORS[i], lw=1.35)
    part = GB[GB.learning_rate == .1]
    axes[0].plot(part.n_estimators, part.train_log_loss, "--",
                 color=COLORS[1], label="Train lr=0.1", lw=1)
    axes[0].set(xlabel="Boosting rounds", ylabel="Log-loss")
    axes[0].legend(frameon=False, fontsize=6.8, loc="upper right")
    axes[1].set(xlabel="Boosting rounds", ylabel="Validation ROC-AUC", ylim=(.95, 1.002))
    for ax in axes:
        ax.grid(alpha=.15)
    savefig(fig, "boosting.png")


doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.top_margin, sec.bottom_margin = Inches(.65), Inches(.65)
sec.left_margin, sec.right_margin = Inches(.78), Inches(.78)
sec.header_distance, sec.footer_distance = Inches(.25), Inches(.25)


def set_font(run, size=None, bold=None, font="Times New Roman", east="宋体"):
    run.font.name = font
    fonts = run._element.get_or_add_rPr().rFonts
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn("w:" + attr), None)
    fonts.set(qn("w:eastAsia"), east)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


for name in ("Normal", "Body Text", "Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3"):
    style = doc.styles[name]
    style.font.name = "Times New Roman"
    fonts = style._element.get_or_add_rPr().rFonts
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn("w:" + attr), None)
    fonts.set(qn("w:eastAsia"), "黑体" if name.startswith("Heading") or name == "Title" else "宋体")
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.font.size = Pt(11)
    pf = style.paragraph_format
    pf.space_before, pf.space_after = Pt(0), Pt(5)
    pf.line_spacing = 1.18
doc.styles["Title"].font.size = Pt(23)
doc.styles["Title"].font.bold = True
doc.styles["Title"].paragraph_format.space_after = Pt(9)
doc.styles["Heading 1"].font.size = Pt(16)
doc.styles["Heading 1"].font.bold = True
doc.styles["Heading 1"].paragraph_format.space_before = Pt(14)
doc.styles["Heading 1"].paragraph_format.space_after = Pt(10)
doc.styles["Heading 2"].font.size = Pt(11.5)
doc.styles["Heading 2"].font.bold = True
doc.styles["Heading 2"].paragraph_format.space_before = Pt(9)
doc.styles["Heading 2"].paragraph_format.space_after = Pt(5)
for element in list(doc.styles.element.iter(qn("w:pBdr"))):
    element.getparent().remove(element)

hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
set_font(hp.add_run("第10讲  集成学习与 Bootstrapping  |  南铂山"), 8.5)
footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(footer.add_run("第 "), 8.5)
field = OxmlElement("w:fldSimple")
field.set(qn("w:instr"), "PAGE")
footer._p.append(field)
set_font(footer.add_run(" 页"), 8.5)
doc.core_properties.title = "第10讲 集成学习与Bootstrapping工作单"
doc.core_properties.author = "南铂山"
doc.core_properties.subject = "WDBC 实证分析与模型选择"
doc.core_properties.keywords = "Ensemble, bootstrap, WDBC, worksheet"


def p(text, bold_lead=None, small=False):
    para = doc.add_paragraph()
    if bold_lead:
        set_font(para.add_run(bold_lead), bold=True)
    set_font(para.add_run(text), 9 if small else 11)
    if small:
        para.paragraph_format.line_spacing = 1.1
        para.paragraph_format.space_after = Pt(4)
    return para


def h(text):
    return doc.add_paragraph(text, "Heading 2")


def page(title):
    doc.add_paragraph(title, "Heading 1")


def table(headers, rows, widths=None, size=9.5, left_cols=(0,)):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    if widths is None:
        widths = [6.94 / len(headers)] * len(headers)
    for col, width in zip(t.columns, widths):
        col.width = Inches(width)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:color"), "D9D9D9")
        borders.append(e)
    t._tbl.tblPr.append(borders)
    repeat = OxmlElement("w:tblHeader")
    t.rows[0]._tr.get_or_add_trPr().append(repeat)
    for ri, values in enumerate([headers] + rows):
        row = t.rows[0] if ri == 0 else t.add_row()
        cant = OxmlElement("w:cantSplit")
        row._tr.get_or_add_trPr().append(cant)
        for ci, (cell, text) in enumerate(zip(row.cells, values)):
            cell.width = Inches(widths[ci])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            pr = cell._tc.get_or_add_tcPr()
            margins = OxmlElement("w:tcMar")
            for side, val in (("top", "65"), ("bottom", "65"), ("left", "85"), ("right", "85")):
                e = OxmlElement("w:" + side)
                e.set(qn("w:w"), val)
                e.set(qn("w:type"), "dxa")
                margins.append(e)
            pr.append(margins)
            if ri == 0:
                shade = OxmlElement("w:shd")
                shade.set(qn("w:fill"), "E8ECEB")
                pr.append(shade)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci in left_cols else WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.08
            set_font(para.add_run(str(text)), size, bold=(ri == 0))
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(3)
    spacer.paragraph_format.line_spacing = Pt(1)
    set_font(spacer.add_run(""), 1)
    return t


def picture(name, width=6.65, caption=None):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run()
    shape = run.add_picture(str(FIG / name), width=Inches(width))
    shape._inline.docPr.set("descr", caption or name)
    if caption:
        cap = p(caption, small=True)
        cap.paragraph_format.space_after = Pt(7)
    return para


def code(text):
    para = doc.add_paragraph()
    para.paragraph_format.line_spacing = 1.05
    para.paragraph_format.space_before = Pt(3)
    para.paragraph_format.space_after = Pt(6)
    set_font(para.add_run(text), 9, font="Consolas")
    return para


def f(x):
    return f"{x:.3f}"


def percent(x):
    return f"{100*x:.1f}%"


def main():
    figures()
    doc.add_paragraph("集成学习与 Bootstrapping", "Title")
    p("第10讲工作单完成版  方案A WDBC", bold_lead=None)
    p("姓名：南铂山    学号：24520261155028    日期：2026年11月17日", small=True)
    h("研究问题与主要发现")
    p("在同一数据与验证协议下，多棵树能否同时改善单树的表现和稳定性？"
      "本次结果支持“树集成优于单树”，但不支持“集成必然优于简单基线”。"
      "随机森林的预测波动较小，逻辑回归与 RBF-SVM 的开发集表现仍有竞争力。"
      "因此优先保留逻辑回归作为可解释基线，SVM 与随机森林作为后续候选；所有结论限于本数据的内部验证。")
    h("数据与协议 逐题作答")
    p("1. 每行是一个乳腺肿块细针穿刺图像的细胞核形态特征记录，不是一个独立细胞，也不是一次未来随访。"
      "输入包含30个连续特征；malignant=1 表示数据中已有的恶性诊断，0 表示良性。"
      "sample_id 仅供追踪，数字或编号本身没有可迁移的形态学含义，纳入模型还可能学到采集与排序痕迹。"
      "本表没有患者分组和时间字段，不能核验患者独立性。标签及模型输出不能直接解释成未来患病风险。[1,2]")
    table(["数据部分", "记录数", "恶性记录", "用途"],
          [["训练集", INFO["train_n"], INFO["train_positive"], "学习参数与树结构"],
           ["验证集", INFO["validation_n"], INFO["validation_positive"], "比较模型与探索超参数"],
           ["开发集", 455, INFO["train_positive"] + INFO["validation_positive"], "训练与验证合并后进行5折CV"],
           ["锁定测试集", 114, "未读取标签", "协议锁定后的一次独立评价"]],
          [1.1, .8, 1.05, 3.99], left_cols=(0, 3))
    p("2. 训练集回答“模型如何拟合”，验证集回答“当前候选如何选”，测试集回答“锁定流程对未用于选择的数据表现如何”。"
      "提前看测试标签，会把测试集变成另一套调参数据，无法再作为独立审计。"
      "本次未打开测试标签，也未用全量教师测量表；没有报告测试排名。")
    h("实际运行与公平比较")
    p(f"全部数值由课程脚本函数与补充实验实际重算。Python {INFO['python']}，"
      f"scikit-learn {INFO['sklearn']}，NumPy {INFO['numpy']}；随机种子1010。"
      "训练与验证的 sample_id 无交集，开发集30个特征无缺失。"
      "所有模型使用相同划分；CV 为开发集455条记录上的5折分层交叉验证，shuffle=True。"
      "LR、SVM 的 StandardScaler 均放在 Pipeline 内，只在每折训练部分拟合。")
    p("固定配置：LR使用L2正则、C=1；SVM使用RBF核、C=1、gamma='scale'。"
      "Bagging为100棵不限深度全特征树；RF为100棵不限深度树、max_features='sqrt'。"
      "GB为100轮、学习率0.1、深度3；AdaBoost为50轮树桩。", small=True)
    p("阅读约定：ROC-AUC 衡量排序，balanced accuracy（BA）为敏感度与特异度的平均；"
      "树与LR的类别阈值为0.5，SVM使用默认决策边界。A–H顺序与原工作单一致。"
      "题目要求的“先预测”在此作为机制预期陈述，未虚构实验前的课堂记录。", small=True)

    page("A 决策树的切分与不稳定性")
    h("A1 读课件示例树")
    p("120 > 105.95，先走根节点右支；0.08 > 0.05，再走右支，落入右下叶节点。"
      "该叶有良性3例、恶性112例，因此预测恶性概率为112/115 = 0.9739，类别为恶性。"
      "这是训练样本在该叶的类别频率，不是独立验证后的风险保证；少量样本的叶尤其容易产生过于极端的概率。[1]")
    h("A2 每个节点怎样寻找切分")
    p("CART 在可用特征及候选阈值中，选择使子节点加权 Gini 不纯度下降最多的一次二分。"
      "Gini 越小，节点类别越单一；可理解为按节点类别比例独立抽取两次时，类别不同的概率。"
      "二分类时最大值是0.5，纯节点为0。这里是局部贪心搜索，不保证整棵树的全局最优。")
    picture("gini.png", width=6.4)
    h("A3 树深实验")
    rows = [[r.max_depth if r.max_depth != "None" else "不限",
             int(r.leaves), f(r.train_accuracy), f(r.validation_accuracy),
             f(r.validation_auc), f(r.validation_ba)] for r in DEPTH.itertuples()]
    table(["最大深度", "叶数", "训练准确率", "验证准确率", "验证AUC", "验证BA"],
          rows, [1.05, .55, 1.22, 1.22, 1.45, 1.45])
    best_depth = DEPTH.iloc[DEPTH.validation_accuracy.astype(float).argmax()]
    p(f"机制预期是深树更易拟合训练样本。本次训练准确率随深度上升，不限深度达到1.000；"
      f"验证准确率在深度{best_depth.max_depth}时最高，为{best_depth.validation_accuracy:.3f}。"
      "因此训练更贴合不等于验证更好。单树叶概率离散，深度改变也会改变 AUC；"
      "不同指标的最优深度不必一致。")
    h("A4 重抽样后的根节点")
    rows = [[i + 1, ROOTS.iloc[i].root_feature, f"{ROOTS.iloc[i].root_threshold:.4f}",
             i + 6, ROOTS.iloc[i + 5].root_feature, f"{ROOTS.iloc[i + 5].root_threshold:.4f}"]
            for i in range(5)]
    table(["次", "根特征", "阈值", "次", "根特征", "阈值"], rows,
          [.35, 2.2, .9, .35, 2.2, .94], size=8.5, left_cols=(1, 4))
    p("10份有放回训练样本各拟合一棵深度2的树，模型种子保持1010。"
      "根特征与阈值发生变化，说明小的样本扰动会改变早期切分并传递到后续分支。"
      "这支持单树具有较高的训练样本敏感性；不意味着每棵浅树都具有低偏差。")
    h("A5 标准化为什么有区别")
    p("树依据阈值两侧的大小顺序切分，严格单调缩放通常不改变可分的样本集合；"
      "LR 的正则化及 SVM 的距离和核函数会受量纲影响，因此本次对它们折内标准化。"
      "“通常不需要”不等于所有数值变换完全等价，舍入与离散化仍可能影响树的切分。")

    page("B 投票的收益来自哪里")
    h("B1 硬投票与软投票手算")
    p("硬投票按类别多数决定；软投票先平均恶性概率，再按0.5分类。下面保留原题的三个模型与真值。")
    table(["样本", "真值", "三票", "硬投票", "平均恶性概率", "软投票"],
          [["A", 1, "1  1  0", "1 正确", "(0.80+0.70+0.45)/3 = 0.6500", "1 正确"],
           ["B", 0, "0  1  1", "1 错误", "(0.20+0.60+0.65)/3 = 0.4833", "0 正确"],
           ["C", 1, "0  1  1", "1 正确", "(0.35+0.75+0.80)/3 = 0.6333", "1 正确"],
           ["D", 0, "0  0  1", "0 正确", "(0.15+0.25+0.70)/3 = 0.3667", "0 正确"]],
          [.48, .48, .88, .8, 3.4, .9], size=9)
    p("关键样本是 B。两张恶性票只是略高于0.5，另一张良性票却给出较强的良性判断，"
      "所以多数票误判为1，而概率平均为0.4833，改判为0。"
      "相对真值0，平均概率也比错误的硬标签1更接近正确答案。")
    p("这个例子只证明“软投票可能纠正多数票”，不证明软投票普遍更好。"
      "如果某个模型置信度失真，极端错误概率也会拖累均值。"
      "实际组合应检查概率尺度与校准；权重、阈值和校准流程不能借测试标签来选择。")
    h("B2 方差公式的含义")
    picture("variance.png", width=6.45)
    p("这里 B 是模型个数，不是上一表的样本 B；σ² 为单个模型预测的方差，ρ 为共同的两两预测相关系数。"
      "公式建立在各模型具有相同方差、共同相关性的简化假设上，用于解释机制，而非直接代替实测。")
    p("1. 树数不断增加，只会把第二项压小，方差趋近 ρσ²，而不是必然趋近0。"
      "例如 σ²=1、ρ=0.5 时，10个模型的平均方差为0.55，100个为0.505；继续增加的收益会变小。")
    p("2. 若三棵树几乎总在相同记录上犯同样的错，ρ 接近1，平均仍保留共同错误。"
      "仅重复同一种模型并不能创造信息；需要改变训练样本、候选特征或学习机制来产生有效多样性。")
    h("B3 平均与偏差")
    p("对同分布预测器做算术平均，不会自动消除共同的系统偏差；"
      "许多模型都把同一类边界学错，投票仍可能一起错。"
      "Bagging 的主要作用是平滑不稳定学习器的预测，尤其适合偏差相对低、方差较高的深树。")
    p("正则化逻辑回归通常比深树稳定，在当前样本规模与拟合设置下，单纯重复训练再平均的额外收益可能有限。"
      "这不是说 bagging 在任何条件下都不影响偏差，也不是说所有逻辑回归都低方差；"
      "高维、小样本、弱正则或类别分离时，线性模型同样可能不稳定。")
    h("本题可检验的机制预期")
    p("若 bagging 的收益确实来自降低方差，那么在同一组 bootstrap 训练扰动下，"
      "其固定验证样本的概率标准差应小于单树。D 与 G 将用相同30份重抽样直接检查这一点，"
      "而不是把一次验证 AUC 的提升直接当作“方差已经降低”的证据。")

    page("C Bootstrap 与袋外评价")
    h("C1 一次抽样的具体含义")
    p("对索引[1,2,3,4,5]有放回抽取5次，取原题示例[2,2,5,1,2]。"
      "索引2出现3次，多出了2份重复副本；本袋独特样本为{1,2,5}，占3/5 = 60%。"
      "未入袋的 OOB 样本为{3,4}，占2/5 = 40%。每袋长度仍是5，并非把重复项删去后再训练。")
    h("C2 一千次模拟及有限样本修正")
    picture("bootstrap.png", width=6.4)
    p("U 为一袋中的独特样本数，O 为未入袋样本数；每袋有放回抽取 n 次。", small=True)
    table(["抽样规模n", "独特比例均值", "独特比例SD", "OOB均值", "理论独特比例", "理论OOB"],
          [[n, percent(BOOT.loc[n, "unique_mean"]), f(BOOT.loc[n, "unique_sd"]),
            percent(BOOT.loc[n, "oob_mean"]), percent(BOOT.loc[n, "unique_theory"]),
            percent(BOOT.loc[n, "oob_theory"])] for n in (5, 341)],
          [.85, 1.2, 1.2, 1.1, 1.4, 1.19], size=9)
    p("同一个原始样本在 n 次抽取中均未抽到的概率是(1−1/n)的n次方。"
      "因此 n=5 的理论独特比例为67.232%，OOB为32.768%；常说的63.2%和36.8%是大样本极限，"
      "不能直接套给五个索引的小例子。n=341时才已很接近这个极限。"
      "每次抽到的重复次数不同，所以两种比例都随机波动，且每次相加等于1。")
    code("import numpy as np\n"
         "for n in (5, 341):\n"
         "    rng = np.random.default_rng(1010)\n"
         "    unique = []\n"
         "    for _ in range(1000):\n"
         "        idx = rng.choice(n, size=n, replace=True)\n"
         "        unique.append(len(np.unique(idx)) / n)\n"
         "    print(np.mean(unique), 1 - np.mean(unique))")
    p("两个 n 各自重置为 seed=1010。表中SD是模拟中独特样本比例的样本标准差，"
      "不是模型预测的SD，也不是均值的置信区间。", small=True)
    h("C3 Bagging 如何制造训练集")
    p("Bagging 从训练集独立有放回抽取多份等长样本，各自拟合一棵树，再平均概率或投票；"
      "一棵树不需要知道另一棵树的结果，所以树之间可以并行训练。"
      "抽样不是把数据分成互不重叠的几份，同一记录可以进入多棵树。")
    h("C4 OOB 误差怎么计算及能替代什么")
    p("对训练记录 i，只收集训练时未抽到 i 的那些树的预测，平均为该记录的 OOB 概率；"
      "按固定阈值转类别后，与真实标签汇总计算 OOB 错误率。"
      "本脚本的 oob_score_ 是准确率，OOB error = 1−oob_score_，不是 ROC-AUC。"
      "树数过少时，有的样本可用袋外树不足，估计不稳定。")
    p("OOB 可作为内部误差监控，辅助粗调树数或其他超参数，节省专门留出的验证数据；"
      "反复按 OOB 调参仍会产生选择偏倚，不能替代最终锁定测试，更不能证明外部中心泛化。"
      "若改做 CTG，按行 OOB 可能让同一录音的其他片段进入训练而产生信息泄漏，"
      "应以 recording_id 分组的验证为准；录音分组也不自动等于孕妇个体独立。")

    page("D 单树与树集成的直接比较")
    h("D 主表与稳定性定义")
    table(["模型", "验证AUC", "验证BA", "概率平均SD", "跨类样本占比"],
          [[NAMES[n], f(MODELS.loc[n, "roc_auc"]), f(MODELS.loc[n, "balanced_accuracy"]),
            f(STAB.loc[n, "prediction_sd_mean"]), percent(STAB.loc[n, "class_flip_rate"])]
           for n in ("single_tree", "bagging", "random_forest")],
          [1.8, 1.05, 1.05, 1.45, 1.59])
    p("稳定性使用同一批30份 bootstrap 训练集，各模型内部种子固定；"
      "对114个固定验证样本分别计算30次概率的总体SD（ddof=0），再取平均。"
      "“跨类样本占比”指30次中至少出现过一次0和一次1的验证样本比例，"
      "不是任意相邻两次之间翻转的概率，也不是相对原模型的误判率。")
    p("本次 bagging 与随机森林均比单树有更高的验证 AUC、更小的概率波动；"
      "随机森林进一步降低波动，但它的验证 BA 不高于 bagging。"
      "这提醒我：排序、默认阈值分类和训练扰动稳定性，是三个不同维度，不能只凭一个指标宣布全面最优。")
    h("D1 特征随机化与相关性")
    p("Bagging 的每棵深树都可用全部30个特征，强势特征往往反复被选中；"
      "随机森林在每个节点只考察随机候选特征，本题 max_features='sqrt'，即√30向下取整为5个候选。"
      "这给其他特征参与切分的机会，有可能降低树间相关性；若错误仍高度相关，增加树数的收益有限。"
      "本次没有直接估计树间ρ，因此“去相关”是机制解释，不是已测得的ρ变化。")
    picture("ensemble-performance.png", caption="图1 左为树数与验证AUC，右为同一30份训练扰动下的概率波动。两图来自本次重跑，指标含义不同。")
    h("D2 树数曲线及是否过拟合")
    rows = []
    for n in (1, 10, 50, 100, 300):
        b = COUNT[(COUNT.model == "bagging") & (COUNT.n_trees == n)].iloc[0]
        r = COUNT[(COUNT.model == "random_forest") & (COUNT.n_trees == n)].iloc[0]
        rows.append([n, f(b.validation_roc_auc), f(r.validation_roc_auc),
                     "—" if n < 50 else f(b.oob_accuracy),
                     "—" if n < 50 else f(r.oob_accuracy)])
    table(["树数", "Bagging AUC", "RF AUC", "Bagging OOB准确率", "RF OOB准确率"],
          rows, [.6, 1.35, 1.25, 1.87, 1.87], size=9)
    p("本次主要提升发生在1到10棵之间，50至300棵已接近平台，末位仍会波动。"
      "在基树与数据机制固定时，增加树数主要让平均预测趋于稳定，不像 boosting 那样不断纠正残差；"
      "但这不是验证指标必须单调或永不过拟合的保证。树深、特征采样、标签噪声及反复按验证集选树数仍会影响泛化。")

    page("D 特征重要性与相关特征")
    h("D3 两种重要性的前列对照")
    p("另拟合300棵随机森林，计算训练中平均不纯度下降（MDI）；"
      "置换重要性在固定验证集上，以 ROC-AUC 为评分，逐列打乱20次。"
      "因此它们分别回答“训练切分用得多不多”和“打乱这列会损失多少验证排序表现”，不是同一个量。")
    mdi = IMP.nlargest(5, "impurity_importance").reset_index(drop=True)
    perm = IMP.nlargest(5, "permutation_auc_drop_mean").reset_index(drop=True)
    table(["MDI前五特征", "MDI", "置换前五特征", "AUC下降均值 ± SD"],
          [[mdi.loc[i, "feature"], f(mdi.loc[i, "impurity_importance"]),
            perm.loc[i, "feature"],
            f"{perm.loc[i, 'permutation_auc_drop_mean']:.4f} ± {perm.loc[i, 'permutation_auc_drop_sd']:.4f}"]
           for i in range(5)], [2.05, .65, 2.05, 2.19], size=9, left_cols=(0, 2))
    picture("feature-importance.png", caption="图2 两种重要性各展示前六名。置换误差线是20次随机置换的SD，反映置换重复的波动，不是人群置信区间。")
    h("训练集相关矩阵提供的线索")
    variables = ["worst radius", "worst perimeter", "worst area"]
    table(["Pearson相关", "worst radius", "worst perimeter", "worst area"],
          [[x] + [f"{CORR.loc[x, y]:.3f}" for y in variables] for x in variables],
          [2.14, 1.6, 1.6, 1.6], size=9)
    imp_sub = IMP.set_index("feature")
    sub = "；".join(f"{x}：MDI {imp_sub.loc[x, 'impurity_importance']:.3f}，"
                   f"置换ΔAUC {imp_sub.loc[x, 'permutation_auc_drop_mean']:.4f}"
                   for x in variables)
    p(sub + "。", small=True)
    p("这三列高度相关。某一棵树使用周长切分，另一棵树可能用半径或面积获得相近效果，"
      "所以 MDI 会在替代特征间分摊。单独打乱一列后，其余相关列仍能补充信息，"
      "使单列置换损失很小；低置换值不等于该类形态信息无用。")
    h("怎样限制解释")
    p("MDI 还可能偏爱可切分点多的连续变量；置换结果则依赖当前模型、验证样本和打乱方式。"
      "本题不能把排名解释为生物机制或因果效应，也不能把最小的置换值直接当作删除特征的依据。"
      "下一步可预先定义相关特征组进行联合置换或在训练折内做分组消融，"
      "再用未参与选择的数据检查是否有可迁移的性能损失。")

    page("E Boosting 的顺序依赖")
    h("E1 AdaBoost 怎样接着上一轮学习")
    p("AdaBoost 提高前一轮错分样本的权重，使下一轮弱学习器更关注当前难点，再按学习器表现加权组合。"
      "二分类时，样本权重按指数形式更新并重新归一化。"
      "每轮数据权重依赖前轮结果，所以轮与轮不能像 bagging 的树一样独立并行。"
      "被错标的样本也可能持续获得高权重，因此提升不等于只学习有用的困难样本。")
    h("E2 梯度提升在拟合什么")
    p("二分类 log-loss 中，用 F 表示 log-odds、p 表示 sigmoid(F)，"
      "负梯度为 y−p。新回归树用全部训练记录拟合这个纠错方向，再乘学习率η加入当前 F；"
      "并非只拟合已经分错的记录。例如 y=1、p=0.2 时残差为0.8，需要增大恶性评分；"
      "y=0、p=0.8 时残差为−0.8，需要减小评分。叶值还可依损失作更新。")
    picture("gradient.png", width=6.3)
    h("E3 学习率与轮数的实测")
    rows = []
    for lr in (.01, .1, .5):
        part = GB[GB.learning_rate == lr]
        best = part.loc[part.validation_log_loss.idxmin()]
        last = part.iloc[-1]
        hundred = part[part.n_estimators == 100].iloc[0]
        rows.append([lr, int(best.n_estimators), f(best.validation_log_loss),
                     f(hundred.validation_log_loss), f(last.validation_log_loss), f(last.validation_auc)])
    table(["学习率", "验证loss最小轮数", "最小验证loss", "100轮loss", "500轮loss", "500轮AUC"],
          rows, [.7, 1.5, 1.3, 1.1, 1.1, 1.24], size=9)
    picture("boosting.png", caption="图3 深度3、最多500轮，其余设置相同。左图显示验证log-loss可先降后升；右图说明高AUC不保证概率质量持续改善。")
    r50 = GB[(GB.learning_rate == .1) & (GB.n_estimators == 50)].iloc[0]
    r500 = GB[(GB.learning_rate == .1) & (GB.n_estimators == 500)].iloc[0]
    p(f"学习率0.1时，50轮验证 log-loss 为{r50.validation_log_loss:.3f}，500轮为"
      f"{r500.validation_log_loss:.3f}。轮数过多可能让少量错误变得更自信，"
      "即使 AUC 几乎不变，概率损失仍会恶化。小学习率通常需要更多轮数；"
      "应预先指定选择指标，利用训练内部的早停或CV确定轮数，而非按训练损失或测试表现决定。"
      "表中最小点只是本验证集上的探索结果，不是无偏的最终性能。")
    h("E4 与 XGBoost 和 LightGBM 的关系")
    p("三者都属于梯度提升树家族，不是互相套用的名字。经典 sklearn GradientBoosting 以负梯度加树；"
      "XGBoost 将树复杂度正则写入目标，并利用一阶与二阶导数构造更新；"
      "LightGBM 以直方图分箱和叶子优先生长提高效率，小数据需约束叶数与深度。"
      "这些实现提供缺失值、并行等工程能力，但不能推断在本题必定更准；"
      "sklearn 的 HistGradientBoosting 也有直方图与缺失值支持。本次只查阅文档，未安装运行后二者。[3–5]")

    page("F 公平比较与模型选择")
    h("F 主比较表")
    rows = []
    costs = {"logistic": "较低 系数与方向可检查", "rbf_svm": "较高 核边界难直接解释",
             "random_forest": "中高 多树依赖重要性分析", "gradient_boosting": "中高 顺序加法树较复杂"}
    for name in ("logistic", "rbf_svm", "random_forest", "gradient_boosting"):
        r, c = MODELS.loc[name], CV.loc[name]
        rows.append([NAMES[name], f(r.roc_auc), f(r.balanced_accuracy),
                     f"{c.roc_auc_mean:.3f} ± {c.roc_auc_sd:.3f}", costs[name]])
    table(["模型", "验证AUC", "验证BA", "开发集CV AUC均值 ± SD", "解释成本"],
          rows, [1.22, .85, .85, 1.65, 2.17], size=9, left_cols=(0, 4))
    p("CV在相同5折中比较，SD使用5个折分数的样本标准差（ddof=1）。"
      "SD描述折间波动，不是标准误，更不是95%置信区间；交叉验证的各折也不是完全独立的重复实验。")
    h("F1 集成真的胜过简单基线吗")
    lr, svm, rf, gb = [CV.loc[n] for n in ("logistic", "rbf_svm", "random_forest", "gradient_boosting")]
    p(f"没有。LR与SVM的验证AUC均为1.000；开发集CV中，LR为{lr.roc_auc_mean:.3f}，"
      f"SVM为{svm.roc_auc_mean:.3f}，RF为{rf.roc_auc_mean:.3f}，GB为{gb.roc_auc_mean:.3f}。"
      f"RF和GB相对LR分别低{lr.roc_auc_mean-rf.roc_auc_mean:.3f}和"
      f"{lr.roc_auc_mean-gb.roc_auc_mean:.3f}，差异小于约0.010量级的折间SD。"
      "这不支持集成领先，也不能仅凭均值与SD宣称模型等效或统计显著。")
    for name in ("logistic", "rbf_svm", "random_forest", "gradient_boosting"):
        r = MODELS.loc[name]
        p(f"{NAMES[name]}：验证混淆计数 TN={int(r.tn)}、FP={int(r.fp)}、"
          f"FN={int(r.fn)}、TP={int(r.tp)}；敏感度{r.sensitivity:.3f}，特异度{r.specificity:.3f}。", small=True)
    p("AUC=1仅表示这114条验证记录的正负分数排序完全分离，不意味着每个阈值均正确，"
      "也不意味着未来病例不会出错。LR默认阈值仍有误判，恰好说明排序与决策不能混为一谈。")
    h("F2 我会优先保留什么")
    p("依照课件第14页先建简单基线、再要求复杂模型证明收益的思路，我优先保留标准化LR。"
      "它的开发集表现不弱、计算代价较低，可检查系数与预测方向；但相关特征会让单个系数不稳定，"
      "系数同样不是因果效应。SVM作为高表现候选保留，RF作为低概率波动的非线性候选保留。")
    p("只有当预先规定的嵌套CV、锁定测试或独立外部数据上，复杂模型在核心指标与实际误判代价上"
      "显示可重复且有实际意义的提升，并且校准、稳定性和计算成本可接受时，才改选复杂模型。"
      "仅在这套验证集上再找一个更高的小数，不足以支持改选。")
    h("公平比较仍有哪些局限")
    p("这里比较的是课程给定配置，不是各模型调参后的上限。固定验证集被用于曲线观察，"
      "随后又参与开发集CV，两套结果并非互相独立的确认。"
      "若依据这些结果反复筛选，最终泛化仍应由未触碰的测试集或独立外部队列审计。")

    page("G 稳定性与不确定性")
    h("G1 同一批训练扰动下的预测变化")
    table(["模型", "平均概率SD", "最大概率SD", "跨类样本数", "跨类样本占比"],
          [[NAMES[n], f(STAB.loc[n, "prediction_sd_mean"]), f(STAB.loc[n, "prediction_sd_max"]),
            f"{int(STAB.loc[n, 'class_flip_count'])}/114", percent(STAB.loc[n, "class_flip_rate"])]
           for n in ("single_tree", "bagging", "random_forest", "logistic", "gradient_boosting")],
          [1.85, 1.2, 1.2, 1.3, 1.39])
    p("前三行直接回应D，后两行为补充实验，使“优先LR”的选择也有训练扰动证据。"
      "SVM默认输出的是 decision_function，未启用概率校准；"
      "为避免把分数SD与概率SD混用，本表不报告SVM的概率稳定性，也不声称它已与RF同样稳定。")
    tree_sd = STAB.loc["single_tree", "prediction_sd_mean"]
    rf_sd = STAB.loc["random_forest", "prediction_sd_mean"]
    p(f"RF的平均概率SD较单树降低约{100*(1-rf_sd/tree_sd):.1f}%。"
      "这是在既定模型配置、固定验证记录与这30份重抽样下观察到的差异。"
      "较小SD仅说明预测对本类训练样本扰动不敏感，不保证偏差小、校准好或临床决策有效；"
      "始终给出同一个错误答案的模型也可以很稳定。")
    h("G2 三类不确定性必须分开")
    table(["要回答的问题", "重抽样或验证对象", "本次能否支持"],
          [["训练样本稍变 模型预测会变多少",
            "重抽样训练记录 每份重新拟合 固定验证样本",
            "能 在本方案下用概率SD和跨类占比描述"],
           ["已拟合模型的评价指标有多不确定",
            "固定模型 重抽样评价记录 重算AUC等",
            "未计算 不能把上表当指标置信区间"],
           ["换医院 时间或人群还能否可靠",
            "独立外部队列或时间外验证",
            "不能 本数据缺少对应设计与分组信息"]],
          [1.8, 2.6, 2.54], size=9.5, left_cols=(0, 1, 2))
    p("第一类中，所有模型共享完全相同的训练重抽样，并固定模型种子，因此主要观察的是训练样本变化，"
      "不是反复改变随机种子产生的算法随机性。外层 bootstrap 中的一条原始记录可重复出现；"
      "bagging与RF内部仍各自执行原本的重抽样，这是两层不同用途的抽样。")
    p("第二类即便今后计算指标 bootstrap 区间，也只覆盖固定模型在当前评价样本分布下的抽样变动，"
      "不自动包含重新训练、调参、模型选择和分布漂移。存在患者重复记录时还应按患者重抽样，"
      "而本表没有足够信息保证这一点。")
    h("G3 哪些话不能从实验推出")
    p("本次仅支持“同一训练扰动下，树集成较单树更稳定”。"
      "不能据此推出外部可靠性、AUC置信区间或单次出错概率；"
      "选型还要考虑表现、解释成本与阈值代价。")

    page("H 机制对照与研究备忘录")
    h("H1 Bagging 与 Boosting 对照")
    table(["方面", "Bagging 含随机森林", "Boosting"],
          [["训练方式", "各基模型可独立并行训练", "轮间顺序依赖 新一轮使用旧模型结果"],
           ["每轮数据或权重", "各取一份有放回训练样本 RF另在节点采样特征",
            "AdaBoost更新样本权重 GB拟合当前损失负梯度"],
           ["常用基学习器", "较深的树 先拟合再平均", "树桩或浅树 逐步小幅纠错"],
           ["主要作用", "主要降低高方差 不自动消除共同偏差",
            "逐步降低拟合偏差 方差仍需正则化控制"],
           ["噪声与轮数", "平均可缓解个别树波动 树数增加通常趋于平台",
            "可能持续追逐噪声 学习率与轮数需联动 用验证或早停控制"],
           ["关键超参数", "n_estimators max_features max_samples max_depth min_samples_leaf",
            "learning_rate n_estimators max_depth min_samples_leaf subsample"]],
          [1.05, 2.94, 2.95], size=9.5, left_cols=(0, 1, 2))
    h("H2 研究备忘录")
    memo = (
        "在WDBC的既定划分与五折开发集验证下，我优先采用标准化逻辑回归，保留RBF-SVM和随机森林。"
        "证据是：树集成明显改善单树，但随机森林和梯度提升的CV AUC未超过LR；候选间差距也小于折间波动。"
        "验证AUC等于1只说明本批样本排序完全，不能证明未来预测完美。"
        "训练扰动实验中，随机森林的概率波动小于单树，体现了平均预测的稳定收益；"
        "LR的稳定性已另行检查，SVM则尚无可比的概率稳定性结果。"
        "LR便于审查系数，解释成本较低，但系数和树的重要性都不是因果效应。"
        "本次结论仅限内部数据与给定配置，未揭示测试标签，也不能核验患者独立性或跨医院泛化。"
        "若预注册的外部验证显示随机森林在漏诊代价、校准与性能上持续优于LR，且收益足以抵偿解释成本，我会改选随机森林。"
    )
    count = sum("\u4e00" <= c <= "\u9fff" for c in memo)
    assert 250 <= count <= 400, count
    p(memo)
    h("材料与资料来源")
    p("[1] 本课程《第10讲工作单》、ensemble-pbl.pptx第5、8、14页、data/README.md、"
      "scripts/demo_ensemble.py。教师参考仅用于核查机制，不作为本次实测结果来源。", small=True)
    p("[2] Wolberg W, Mangasarian O, Street N, Street W. Breast Cancer Wisconsin (Diagnostic). "
      "UCI Machine Learning Repository. DOI: 10.24432/C5DW2B。", small=True)
    p("[3] scikit-learn 官方文档 1.11 Ensembles。"
      "scikit-learn.org/stable/modules/ensemble.html。", small=True)
    p("[4] XGBoost 官方文档 Introduction to Boosted Trees。"
      "xgboost.readthedocs.io/en/stable/tutorials/model.html。", small=True)
    p("[5] LightGBM 官方文档 Features。"
      "lightgbm.readthedocs.io/en/latest/Features.html。资料核验日期2026年11月17日。", small=True)

    # Avoid a blank opening page before a manual section break.
    for para in doc.paragraphs:
        para.paragraph_format.widow_control = True
    doc.save(OUT)
    desktop = Path(r"C:\Users\A\Desktop\集成学习_worksheet_完成版.docx")
    shutil.copy2(OUT, desktop)
    support = Path(r"C:\Users\A\Downloads\Ensemble-worksheet-support")
    support.mkdir(exist_ok=True)
    shutil.copytree(A, support / "analysis", dirs_exist_ok=True)
    shutil.copytree(FIG, support / "figures", dirs_exist_ok=True)
    for subpath in ("scripts/demo_ensemble.py", "data/README.md",
                    "data/student/train.csv", "data/student/validation.csv"):
        target = support / "10. Ensemble" / subpath
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / "10. Ensemble" / subpath, target)
    shutil.copy2(ROOT / "run_analysis.py", support / "run_analysis.py")
    shutil.copy2(ROOT / "build_worksheet.py", support / "build_worksheet.py")
    print(f"Memo Chinese characters: {count}")
    print(OUT)
    print(desktop)


if __name__ == "__main__":
    main()
