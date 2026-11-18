from pathlib import Path
from copy import deepcopy
import json
import re
import hashlib
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).parent
meta = json.loads((ROOT / "source_metadata.json").read_text(encoding="utf-8"))
res = json.loads((ROOT / "verified_results.json").read_text(encoding="utf-8"))
source = Document(meta["source"])
doc = Document()
section = doc.sections[0]
section.page_width, section.page_height = Cm(21), Cm(29.7)
section.top_margin, section.bottom_margin = Cm(1.9), Cm(1.9)
section.left_margin, section.right_margin = Cm(2.15), Cm(2.15)
section.header_distance, section.footer_distance = Cm(.8), Cm(.85)

def style_font(style, latin, east, size, bold=False):
    style.font.name = latin
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    rf = style.element.get_or_add_rPr().get_or_add_rFonts()
    for attr, val in [("ascii", latin), ("hAnsi", latin), ("eastAsia", east)]:
        rf.set(qn("w:" + attr), val)
    for attr in ["asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"]:
        rf.attrib.pop(qn("w:" + attr), None)
    for e in style.element.xpath("./w:pPr/w:pBdr"):
        e.getparent().remove(e)

style_font(doc.styles["Normal"], "Times New Roman", "宋体", 10.5)
fmt = doc.styles["Normal"].paragraph_format
fmt.line_spacing = 1.22
fmt.space_after = Pt(5)
fmt.widow_control = True
for name, size in [("Title", 22), ("Subtitle", 11), ("Heading 1", 14), ("Heading 2", 11)]:
    style_font(doc.styles[name], "Arial", "微软雅黑", size, name != "Subtitle")
    f = doc.styles[name].paragraph_format
    f.space_before = Pt(11 if "Heading" in name else 0)
    f.space_after = Pt(6)
    f.line_spacing = 1.15
    f.keep_with_next = True
style_font(doc.styles["Caption"], "Times New Roman", "宋体", 9)
doc.styles["Caption"].paragraph_format.space_before = Pt(3)
doc.styles["Caption"].paragraph_format.space_after = Pt(7)
doc.styles["Caption"].paragraph_format.line_spacing = 1.12

def p(text="", bold=False, style=None):
    previous = doc.element.body[-2] if len(doc.element.body) > 1 else None
    par = doc.add_paragraph(style=style)
    if previous is not None and previous.tag == qn("w:tbl"):
        par.paragraph_format.space_before = Pt(6)
    run = par.add_run(text)
    if bold:
        run.bold = True
    return par

def h(text):
    return p(text, style="Heading 1")

def sub(text):
    return p(text, style="Heading 2")

def equation(index):
    e = deepcopy(source.paragraphs[index]._p)
    # Preserve the original native Word equation while normalizing paragraph spacing.
    pp = e.find(qn("w:pPr"))
    if pp is not None:
        e.remove(pp)
    pp = OxmlElement("w:pPr")
    pp.append(OxmlElement("w:keepNext"))
    e.insert(0, pp)
    par = p()
    par._p.addnext(e)
    par._p.getparent().remove(par._p)
    return e

def formula(text):
    par = p(text)
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(3)
    par.paragraph_format.space_after = Pt(6)
    par.paragraph_format.keep_with_next = True
    for run in par.runs:
        run.font.name = "Cambria"
        run.font.size = Pt(10.5)
    return par

def code(text):
    par = p(text)
    par.paragraph_format.line_spacing = 1.1
    par.paragraph_format.space_before = Pt(3)
    par.paragraph_format.space_after = Pt(7)
    for run in par.runs:
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    return par

def table(number, caption, heads, rows, widths=None):
    cp = p(f"表 {number}  {caption}", style="Caption")
    cp.paragraph_format.keep_with_next = True
    t = doc.add_table(rows=1, cols=len(heads))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    widths = widths or [16.7 / len(heads)] * len(heads)
    for col, width in zip(t.columns, widths):
        col.width = Cm(width)
    for c, title in zip(t.rows[0].cells, heads):
        c.text = str(title)
    for row in rows:
        cells = t.add_row().cells
        for c, value in zip(cells, row):
            c.text = str(value)
    pr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = OxmlElement("w:" + side)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "D9D9D9")
        borders.append(b)
    pr.append(borders)
    for ri, row in enumerate(t.rows):
        rpr = row._tr.get_or_add_trPr()
        no_split = OxmlElement("w:cantSplit")
        rpr.append(no_split)
        if ri == 0:
            rpr.append(OxmlElement("w:tblHeader"))
        for ci, c in enumerate(row.cells):
            c.width = Cm(widths[ci])
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tcpr = c._tc.get_or_add_tcPr()
            margins = OxmlElement("w:tcMar")
            for side in ["top", "bottom", "left", "right"]:
                mar = OxmlElement("w:" + side)
                mar.set(qn("w:w"), "70" if side in ["top", "bottom"] else "85")
                mar.set(qn("w:type"), "dxa")
                margins.append(mar)
            tcpr.append(margins)
            if ri == 0:
                shade = OxmlElement("w:shd")
                shade.set(qn("w:fill"), "E7EBEF")
                tcpr.append(shade)
            for par in c.paragraphs:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                par.paragraph_format.line_spacing = 1.12
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.space_before = Pt(0)
                par.paragraph_format.keep_with_next = ri < len(t.rows) - 1
                for run in par.runs:
                    run.font.size = Pt(9)
                    run.bold = ri == 0
    return t

def figure(n, caption, width=16.7):
    par = p()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.keep_with_next = True
    run = par.add_run()
    shape = run.add_picture(str(ROOT / f"figures/fig{n}.png"), width=Cm(width))
    shape._inline.docPr.set("descr", caption)
    cp = p(f"图 {n}  {caption}", style="Caption")
    cp.paragraph_format.keep_together = True

def pagebreak():
    doc.add_page_break()

def f3(x):
    return f"{x:.3f}"

p("支持向量机的间隔与分类边界", style="Title")
p("第 9 讲作业  SVM 与 Margin", style="Subtitle")
p("姓名：南铂山    学号：24520261155028    日期：2026 年 11 月 18 日")
p("这份作业从八个可分点出发，检验最大间隔怎样选出边界，再回到完整训练集，观察软间隔、C、核函数和输入尺度的作用。最直观的结果是：三条线都能零错误分类时，最近点距离仍可相差一倍以上；而完整数据上，参考带最宽的模型反而把所有记录判成异常。间隔必须和违例一起看。")
p("全文保留手算和受控对照。参数比较使用 186 条训练记录与 62 条验证记录；保留测试集不参与本次选参和错误分析。")

h("一  数据与实验范围")
p("数据来自 UCI Vertebral Column，每一行是一位患者的一条测量记录。原数据共 310 条、6 个连续测量字段，无缺失；正常 NO 有 100 条，异常 AB 有 210 条。AB 又包含椎间盘突出 DH 60 条和脊椎滑脱 SL 150 条。把两种病合成一个标签，方便二分类，却也把它们在特征上的差别藏了起来。")
p("本课固定使用骨盆入射角 pelvic_incidence 和滑脱程度 degree_spondylolisthesis，主要为了把边界画在二维平面上，不表示它们是最佳特征。sample_id 是课程按源文件行号编的 V001 至 V310，只用于追踪记录。它不是生理测量，不能作为输入。abnormal=1 也只表示当前异常类别，不表示严重程度或未来患病风险。")
table(1, "固定分层划分及各自职责", ["集合", "NO", "AB", "合计", "职责"], [
    ["训练",60,126,186,"拟合标准化参数与模型"],
    ["验证",20,42,62,"比较参数并记录配置"],
    ["保留测试",20,42,62,"教师安排后作一次评价"]], [2,1.3,1.3,1.5,10.6])
p("划分种子为 9。StandardScaler 只在训练集上估计均值和标准差，验证集只接受变换。训练和预处理放进同一条 Pipeline，避免先用全部数据标准化再划分。")
p("滑脱程度较低处，两类明显重叠；较高处主要是 AB。图 1 中的 V116 是一条滑脱程度超过 400 的 AB 记录。它影响画图范围，但这本身不是删掉它的理由：还需核对原始测量、单位和录入情况。数据说明没有给出足够的测量误差资料，因此这里保留原值，坐标轴沿用源尺度，不擅自加上角度或百分比单位。")
p("八点几何子集另按固定规则选取：训练集中滑脱程度小于 5 的 NO，以及大于 30 的 AB，各按编号取前四条。这是有意挑成可分的教学例子。它的零错误只用于说明几何，正式参数比较仍用全部 186 条训练记录。")
figure(1, "左图保留全部 186 条训练记录；右图放大骨盆入射角 20 至 100、滑脱程度 −15 至 105 的区域，7 条记录在范围外。两图使用同一批记录，范围外记录仍参与拟合。蓝色圆点为 NO，红色三角为 AB。")

h("二  零训练错误还不能选出边界")
p("三条候选线的法向量相同，只把截距分别移动 −0.6、0、+0.6。运行前，我选中间那条，理由是它离两边的点都较远。运行后发现三条线的训练错误都是零，只看分对几个点选不出唯一边界；还要把“离两边都远”变成可比较的距离。")
figure(2, "八点均使用完整训练集拟合的 scaler，两个坐标轴等比例显示。左图三条候选线均无训练错误；右图实线为 f=0、虚线为 f=±1，空心圈标出支持向量 V217 与 V061，红色线段连接二者。")
equation(21)
p("这里 z 是标准化后的输入。f(z)>0 判为 AB，f(z)<0 判为 NO。带标签的距离 y f(z)/‖w‖ 在误分时为负；不带标签的 |f(z)|/‖w‖ 只量距离，不判断对错。")
sub("V061 的手算")
p("代入工作单给定的四位小数：w=(0.7287, 0.7492)，b=−0.0200，z=(0.7108, 0.6700)。")
formula("f = 0.7287×0.7108 + 0.7492×0.6700 − 0.0200 = 0.99992")
formula("‖w‖ ≈ 1.0451，   d = |f|/‖w‖ ≈ 0.957")
p("所以预测为 AB。用未舍入的模型参数核对，f=1.000000，d=0.956807。f 约为 1 表示它位于参考线 f=1 上，不能写成 100% 风险。")
p("把 w 和 b 同乘 3，f 变成原来的 3 倍，‖w‖ 也变成 3 倍；边界、预测和距离均不变。分数的绝对值依赖参数的表示方式，几何距离则不受这种同比缩放影响。")
table(2, "三条候选线的最小带符号距离", ["截距偏移","训练错误","最小距离","最近的记录"], [
    ["−0.6",0,"0.3827","V061（AB）"],
    ["0",0,"0.9568","V217 与 V061"],
    ["+0.6",0,"0.3827","V217（NO）"]], [2.5,2.5,3.5,8.2])
p("截距减小时，边界向 AB 一侧移动，先贴近 V061；增大时则先贴近 V217。截距移动 0.6，相当于沿法向移动 0.6/1.0451≈0.574；原来的最小距离 0.957 因而降到 0.383。比较距离之和不合适：少数很远的点可能把总和抬高，掩盖某个点已经贴近边界。最大间隔关注最小距离，给最脆弱的点尽量多留余地。")
p("可分时把约束规范化为 yᵢf(zᵢ)≥1，最大化最小距离就等价于最小化 ½‖w‖²。参考线 f=±1 的总间距为 2/‖w‖=1.9136，最近点到中线的距离为它的一半 0.9568。两者都属于标准化空间，没有原始测量单位。本例用 C=10000 的线性 SVC 近似硬间隔，并检查最小 y·f 与 1 的偏差小于 10⁻⁵。")

h("三  支持向量怎样约束边界")
p("支持向量的定义是对偶系数 αᵢ 非零的训练记录。八点模型只有 V217 和 V061 两个支持向量。删点前，我预计删掉远处的 V072 变化很小，删掉支持向量 V061 变化较大。实验冻结原 scaler，分别删掉一个点，再在相同坐标系中重新拟合：")
table(3, "删除单个训练点后的变化", ["删除记录","类型","‖Δw‖","|Δb|","原八点最大 |Δf|"], [
    ["V072","非支持向量","0.000000","0.000000","0.000000"],
    ["V061","支持向量","0.156617","0.127752","0.427454"]], [2.2,3.1,3.1,3.1,5.2])
p("我预计的方向对了，但删掉 V072 后不只是变化很小，而是在小数点后六位都看不到变化；删掉 V061 后，原八点的分数最大改变 0.427。冻结 scaler 很重要：若同时重新估计均值和标准差，所有点的坐标都会改变，无法单独解释删点的影响。")
equation(40)
p("在这个恰好只有两个支持向量的例子中，Σαᵢyᵢ=0 使两个 α 相等。记正、负支持向量为 z₊、z₋，便有 w=α(z₊−z₋)。二者满足 f(z₊)=1、f(z₋)=−1，两式相减得到 wᵀ(z₊−z₋)=2。")
formula("z₊−z₋ ≈ (1.3342, 1.3718)，   ‖z₊−z₋‖² ≈ 3.6619")
formula("α = 2/‖z₊−z₋‖² ≈ 0.5462，   w ≈ (0.7287, 0.7493)")
p("模型读出的 α=0.546162，w=(0.728685, 0.749228)，b=−0.020003。最后一位的差别来自手算舍入。因此图 2 的边界是两个支持向量连线的垂直平分线，总间距等于它们之间的距离。这个简单结论依赖当前只有两个支持向量，不能照搬到一般多点情形。")
p("硬间隔的互补松弛条件 αᵢ(yᵢfᵢ−1)=0 解释了 V072 为什么不直接出现在 w 中：它严格位于参考带之外，α=0。删掉这个点后，其余点沿用原来的 α，仍满足可行性与全部 KKT 最优条件；w 不变，两个支持向量又固定了 b，因此仍是同一条边界。")
p("也不能反过来说删掉任何支持向量都必然改变模型。若约束存在冗余，剩余样本可能仍维持原解。支持向量表示当前特征和参数下对优化起作用的记录，不等于最典型、最严重或最值得临床关注的患者。")

h("四  矛盾标签为什么需要软间隔")
p("复制 V061，保留坐标却把标签改成 NO，仅用于人工反事实，不修改原始数据。原记录要求 f(z)≥1，复制记录要求 f(z)≤−1。同一个输入不可能同时满足两者，所以硬间隔无解。换成非线性核也不能消除这个矛盾，因为一个确定的输入仍只能得到一个 f(z)。")
p("软间隔允许记录违反参考间隔，并用 hinge loss 计价：ℓᵢ=max(0,1−yᵢfᵢ)。它不会假装把矛盾标签都分对，而是在间隔与违例之间求解有限的权衡。")
table(4, "矛盾九点数据的软间隔结果", ["C","训练错误","总 hinge","支持向量数"], [
    [".01",4,"7.29",8],[1,1,"2.00",4],[100,1,"2.00",4]], [2.4,4,5.2,5.1])
p("设矛盾点的分数为 t，它们的总损失是 max(0,1−t)+max(0,1+t)≥2；在 −1≤t≤1 时恰为 2。C=1 和 100 都达到这个下界，仍至少分错一条。因此零训练错误并非总能实现。")
table(5, "分类错误与间隔违例的区别", ["y·f","−0.5","0.4","1","1.8"], [
    ["hinge","1.5","0.6","0","0"],
    ["预测是否错误","是","否","否","否"],
    ["相对参考带","带内错误侧","带内正确侧","在边线上","带外正确侧"]], [3.3,3.35,3.35,3.35,3.35])
p("y·f=0.4 时预测正确，却仍付出 0.6 的损失，因为它离边界还不够远。分类错误只问正负号，hinge 还记录违例有多深。软间隔图中的 f=±1 之间允许有点，支持向量也可能被误分。注意严重误分的点还可能落在参考带之外的错误侧，不能把所有违例都理解成 |f|<1。")
p("软间隔的 KKT 条件增加 0≤αᵢ≤C：α=0 的点满足 y·f≥1；0<α<C 的点在 y·f=1 上；α=C 的点满足 y·f≤1。边界等号处可能有退化情形，数值判断还需容差，因此仍以非零 α 定义支持向量。完整训练集的 linear C=1 模型有 89 个支持向量，其中 86 个存在超过 10⁻⁵ 的间隔违例，其余 3 个贴在参考线上。")

h("五  C 控制违例的相对代价")
equation(60)
p("第一项限制权重范数，偏好较宽的参考带；第二项累加训练违例。C 增大意味着同样的违例更贵，不是噪声比例或概率阈值。数据线性可分时，若 C 不小于一个硬间隔对偶最优解中的最大 α，该解也满足软间隔的上界约束，因而保留同样的最优边界。八点例子的最大 α 约为 0.546，远小于 10000；第四节的矛盾数据则根本没有硬间隔可行解。")
p("运行前的预测是：C 增大，总 hinge 会降，参考带会变窄；验证表现不一定跟着改善。下面只改变 C，其余条件固定。")
table(6, "只改变线性核的 C 其余设置固定", ["C","带宽","平均 hinge","违例数","训练错误","SV 数","验证 BA"], [
    [f"{r['C']:g}", f"{r['linear_width']:.3f}", f"{r['mean_hinge']:.3f}", r["margin_violations"],r["train_errors"],r["n_sv"],f3(r["val_balanced_accuracy"])]
    for r in res["grid"] if r["kernel"]=="linear"], [1.4,2.2,3,2.2,2.8,2.1,3])
figure(3, "三幅图使用相同坐标范围和完整训练集，每幅有 7 条记录在显示范围外。实线为 f=0，虚线为 f=±1；部分等值线也位于图外。支持向量总数见表 6，不再逐点套圈，以免遮挡两类重叠。BA 指平衡准确率。")
p("C 从 0.01 增大到 100，平均 hinge 从 0.631 降至 0.399，参考带从 12.414 缩至 0.489。验证 BA 却先升后降，C=1 的 0.760 最高。更贴合训练违例并不等于更适合新记录。")
p("C=0.01 尤其值得看：训练错误 60，恰好等于训练集 NO 的人数；模型把所有记录判成 AB。它的 ‖w‖ 很小，带虽然宽，区分能力却丢了。因此不能孤立地把宽带当作泛化好的证据。")
p("在相同数据、核与预处理下，若两次都取得全局最优解，C 增大时总 hinge 不增、‖w‖ 不减。记 R=½‖w‖²、H=Σhinge，对 C₁<C₂ 的最优解有：")
formula("R₁+C₁H₁ ≤ R₂+C₁H₂，   R₂+C₂H₂ ≤ R₁+C₂H₁")
p("相加得到 (C₂−C₁)(H₂−H₁)≤0，所以 H₂≤H₁；代回可得 R₁≤R₂。数值求解存在容差，但方向与表 6 一致。这个推导不约束误分类数、SV 数或验证表现。")
p("训练错误只数预测错了几个点，hinge 还计入正确侧但没达到 y·f=1 的点，并按违例深浅计价。C=1 时，训练错误为 37，间隔违例为 86，其中 49 个点分类正确却仍付出损失。C=0.01 与 0.1 的违例数同为 119，平均 hinge 仍下降，也说明数量不变时深度可以变。")
p("不同 C 的目标值不能直接排名，因为损失项的权重已经换了。逻辑回归也可采用“损失加 L2 正则”的形式，只是把 hinge 换成对数损失；在固定样本数和归一化约定下，C 越大，对应的相对正则越弱。")

h("六  核函数与 gamma 改变了什么")
sub("从显式变换到内积")
p("一个简单例子是正类在 x=±2、负类在 x=±1。原数轴上单个阈值分不开；变换 φ(x)=x² 后，两类分别位于 4 和 1，阈值 2.5 就能分开。核函数计算这种变换后的内积，而不一定把新坐标逐一列出。")
p("本课二次多项式核 K(z,u)=(0.1zᵀu+1)²，可显式写成下列六维映射：")
equation(77)
p("其中交叉项的系数为 0.1√2。把 V061 与 V217 代入，zᵀu≈−0.91325，直接算核函数得到 (1−0.091325)²≈0.82569；先做六维映射再取内积也得到 0.82569。两种路线算的是同一个量。")
equation(82)
p("把线性模型的 w=Σαᵢyᵢzᵢ 代入决策函数，训练点只通过内积出现，所以可用 K 替代内积。预测时各支持向量贡献的是带正负号的 αᵢyᵢK(zᵢ,z)，最后加 b 再看符号，不是简单多数投票。RBF 对应无限维特征空间，核函数省去了显式展开；大样本时还可用随机特征等有限维近似。")
sub("gamma 决定 RBF 相似度衰减多快")
formula("K(z,u) = exp(−γ‖z−u‖²)")
table(7, "两种 gamma 下的 RBF 相似度手算", ["两点距离","γ=0.1","γ=1"], [
    [1,"exp(−0.1)=0.905","exp(−1)=0.368"],
    [2,"exp(−0.4)=0.670","exp(−4)=0.018"]], [3.2,6.75,6.75])
p("gamma 越大，相同距离下的相似度越小，影响越局部。相似度降到 e⁻¹ 的距离为 1/√gamma。C 决定违例有多贵，gamma 决定距离怎样转成相似度，它们可以共同影响边界，但作用不同。")
p("固定 C=1 时，我原先预计 γ=0.01 的边界平滑、接近直线，γ=10 会更弯。这个预测只想到了形状，没有想到模型可能失去对两类的有效区分。")
table(8, "固定 C=1 只改变 gamma", ["gamma","训练 BA","验证 BA","SV 数"], [
    [f"{r['gamma']:g}",f3(r["train_balanced_accuracy"]),f3(r["val_balanced_accuracy"]),r["n_sv"]]
    for r in res["grid"] if r["kernel"]=="rbf" and r["C"]==1], [3.5,4.4,4.4,4.4])
p("结果 γ=0.01、C=1 在训练和验证记录上全部预测 AB，BA 只有 0.500。我没有预料到这一点：边界平滑，不等于能有效区分眼前的数据；全 AB 也是对这两批记录的观察，不是在断言整个输入空间没有边界。")
p("要理解小 gamma，还需看一阶近似的条件：参与训练的点对，以及支持向量与待预测点之间，都要满足 γ‖z−u‖²≪1。此时展开 RBF：")
formula("K(z,u) ≈ 1 − γ‖z‖² − γ‖u‖² + 2γzᵀu")
p("利用 Σαᵢyᵢ=0，常数项和只含待预测点的项抵消，只含训练点的项并入 b，剩下缩放后的线性内积。代回对偶目标并换元，得到近似的线性模型 C′=2γC。补充拟合 linear C=0.02，与 RBF C=1、γ=0.01 的 62 个验证预测完全相同，都为 AB，分数相关系数约为 0.9995。")
p("这仍是有条件的解释：V116 标准化后约为 (3.9, 9.4)，γ=0.01 时，它与其余训练点的 γd² 中位数约为 1，而不是远小于 1。把 C 提到 100 后，RBF 与其对应近似 linear C=2 已有 2 个验证预测不同。小 gamma 的近似能解释趋势，不能当成全空间中的严格等价。")
table(9, "固定 gamma=10 只改变 C", ["C","训练 BA","验证 BA","验证特异度"], [
    [f"{r['C']:g}",f3(r["train_balanced_accuracy"]),f3(r["val_balanced_accuracy"]),f"{r['val_specificity']:.2f}"]
    for r in res["grid"] if r["kernel"]=="rbf" and r["gamma"]==10 and r["C"] in [1,10,100]], [3.5,4.4,4.4,4.4])
figure(4, "上排固定 C=1 改变 gamma；下排固定 gamma=10 改变 C，右上与左下是两组对照共用的基准模型。每幅标题依次列出训练与验证 BA，坐标范围相同，均有 7 条训练记录在范围外，仍参与拟合。")
p("gamma=10 时，相似度降到 e⁻¹ 的距离只有约 0.32 个标准差，支持向量的影响很局部。C 从 10 增到 100，又使同一违例的代价变成 10 倍，模型便更愿意用局部弯曲去迁就个别训练点。训练错误从 25 降到 16，训练 BA 从 0.870 升到 0.910；验证特异度却从 0.80 降到 0.60，BA 从 0.805 降到 0.669。图 4 右下的碎边界没有带来验证改善，提示这组参数过拟合。")
figure(5, "RBF 的 4×4 参数网格，左右共用色标。红框标出按验证规则选中的 C=10、gamma=10。数值保留三位小数，颜色只辅助比较，不代替读数。")
p("若同时改变 C 和 gamma，单凭两幅图无法把变化归因给某一个参数。因此先做表 8、9 的单因素对照，再看联合网格中的交互。训练表整体分数较高，验证表却没有相同趋势。")
figure(6, "同为 C=1、相同训练标准化。多项式核使用 degree=2、gamma=0.1、coef0=1；RBF 使用 gamma=0.1。三图坐标、点型和线型一致，每幅有 7 条训练记录在范围外，仍参与拟合。")
p("线性核只能给直线；二次多项式核能表达二次曲线，例如椭圆、双曲线或其退化形式；RBF 可以形成多个局部区域。本例多项式核和 RBF 的验证 BA 分别为 0.785、0.773。图 6 中两种非线性边界都主要沿滑脱程度分开两类，再在低值重叠区弯曲，因此看起来相近；更灵活的核不一定要画出截然不同的边界。")
p("线性核的 2/‖w‖ 是当前标准化输入空间中的宽度。RBF 的间隔则在特征空间中，原平面上的 f=±1 是分数等值线，不是离 f=0 等距的两条线，不能拿图上间距直接跨核比较 margin。")

h("七  为什么换个单位就会改变预测")
p("把第一个输入乘 100，并不改变记录携带的信息。固定 RBF C=1、gamma=0.1，比较原单位与新单位，各做无标准化和训练集标准化两种设置。我的预测是：标准化那一对几乎不变，不标准化那一对会变，因为后者直接用变大的数值计算距离。")
table(10, "同一信息的四种尺度设置", ["标准化","输入单位","验证 BA","SV 数"], [
    ["无","原单位","0.705",171],["无","x1×100","0.488",186],
    ["训练集拟合","原单位","0.773",116],["训练集拟合","x1×100","0.773",116]], [3.5,5.2,4,4])
p("两个单位版本的验证 decision score 最大差，无标准化时为 1.538，标准化后为 1.78×10⁻¹⁴，已在浮点误差水平。RBF 的距离平方中，第一维贡献被放大了 10000 倍；固定 gamma 下，多数不同点之间的相似度因而变得很小，模型的局部性大幅改变。")
formula("(100x−100μ)/(100s) = (x−μ)/s")
p("StandardScaler 同时把该列训练均值和标准差放大 100 倍，正好抵消这次线性单位变化，结果与预测一致。标准化不是唯一的办法；关键是让距离和正则所隐含的尺度符合问题，而不是由某个字段的数值大小决定。")
p("不能在全部 310 条记录上拟合 scaler。即使不看标签，验证和测试输入的分布也会影响训练坐标系，保留集不再真正独立。若做交叉验证，每折都要只在该折训练部分拟合 scaler；Pipeline 正是为此保留预处理与模型的先后关系。")
sub("代码里的数据流")
code("model = make_pipeline(StandardScaler(),\n"
     "    SVC(kernel='rbf', C=10, gamma=10, probability=False,\n"
     "        tol=1e-7, max_iter=200000))\n"
     "model.fit(train[FEATURES], train.abnormal)\n"
     "score = model.decision_function(validation[FEATURES])\n"
     "prediction = (score >= 0).astype(int)")
p("fit 先从训练输入学习均值、标准差，再把标准化结果交给 SVC。decision_function 对验证输入只调用已有的变换和决策函数，不重新学习。公式中 y=−1/+1，代码里保留 NO=0、AB=1；二分类 decision score 为正对应 classes_[1]，即 AB。")

h("八  验证分数背后具体错了谁")
p("验证集始终预测 AB，accuracy=42/62=0.677，却一个 NO 也识别不出；平衡准确率 BA=(敏感度+特异度)/2=(1+0)/2=0.500。用 BA 可以把两类识别率放在相同权重下比较。")
p("按课程规则在线性与 RBF 家族内分别选取验证 BA 最高的配置；若相同，先取较小 C，再取较小 gamma。阈值固定为 0，输入与 scaler 拟合范围均不变。")
table(11, "冻结配置在 62 条验证记录上的结果", ["配置","TN / FP","FN / TP","敏感度","特异度","BA","AUC"], [
    ["linear C=1","18 / 2","16 / 26","0.619","0.900","0.760","0.830"],
    ["RBF C=10 γ=10","16 / 4","8 / 34","0.810","0.800","0.805","0.781"]], [4.1,2.2,2.2,2.1,2.1,2,2])
p("TN、FP 对应正常记录判对、判错；FN、TP 对应异常记录漏判、判对。RBF 比线性模型多识别 8 条 AB，同时多误判 2 条 NO，BA 高约 4.5 个百分点。它的 AUC 反而低约 0.049，因为固定阈值下的分类与整个分数排序不是同一个问题。")
p("为解释错误，把公开三分类标签按源文件行号与训练、验证 ID 对齐，并检查两项输入逐行一致。亚型只用于事后解释，不用于训练、选参或改阈值。训练集滑脱程度中位数：NO 为 1.56，DH 为 2.65，SL 为 42.89。DH 与 NO 在这项输入上明显更接近。")
table(12, "冻结模型在验证集各亚型中的判对数", ["模型","NO 共20条","DH 共16条","SL 共26条"], [
    ["linear C=1","18 / 20","1 / 16","25 / 26"],
    ["RBF C=10 γ=10","16 / 20","9 / 16","25 / 26"]], [5.9,3.6,3.6,3.6])
figure(7, "验证集按原始亚型着色，红圈表示预测错误。图中放大滑脱程度 −15 至 47 的重叠区域；范围外另有 11 条 SL，两个模型均判对。边界由训练集拟合，图上的验证记录未加入拟合。")
p("线性模型的 16 个漏判中，15 个是 DH，另 1 个是 SL。它主要利用较大的滑脱程度区分异常，因此对这批 DH 的识别很弱。RBF 在 DH 上从判对 1 条增加到 9 条，同时 NO 从判对 18 条降到 16 条。图 7 右图在两类混杂处伸出一个细长的 AB 区域，纳入了部分 DH，也把部分 NO 划了进去。这说明改进具体来自哪里，也暴露了只用这两项输入的局限。")
sub("decision score 不是概率")
table(13, "线性模型验证分数与真实 AB 比例", ["f 范围","人数","其中 AB","观察到的 AB 比例"], [
    [r["label"],r["n"],r["abnormal"],f"{r['fraction']*100:.1f}%"]
    for r in res["score_bins_linear"]], [4.9,2.4,3.4,6])
p("AB 比例整体随分数上升，说明 f 带有排序信息。但 f<−1 的 9 条记录中仍有 4 条 AB，和下一箱的 48% 很接近，可见 f 更负也不能直接换算成更可信的 NO 概率。各箱只有 4 至 25 条记录，这些是当前分箱的经验比例，不是每个新患者的概率。SVM 的 f 可取任意实数，hinge 目标没有要求它等于概率；f=1 只是参考间隔的约定。")
p("线性模型中 f 除以 ‖w‖ 可转为几何距离；核模型的间隔在特征空间里，不能直接照搬这个原输入空间公式。")
p("AUC 可直接用 decision score 排序计算。若需要概率，可在与拟合分离的数据或交叉验证流程中做 sigmoid 或 isotonic 校准，并检查校准效果。逻辑回归在其模型假设下把线性预测子对应到对数几率，但也不自动保证现实数据上的概率校准。本课保持 probability=False，不把校准问题混进当前参数实验。")
sub("测试前记录与保留问题")
p("两组配置均冻结为上述两个输入、仅训练集拟合的 StandardScaler，以及表 11 的 SVC 参数，不把验证记录重新加入拟合。测试前要检验的具体问题是：线性模型是否仍主要漏判 DH；RBF 在 DH 上的改善能否保持，以及正常记录的误判是否增加。")
p("测试评价留待教师安排，不以验证结果冒充测试结果。打开测试后应报告 62 条保留记录的混淆矩阵、敏感度、特异度、BA 与 AUC，并解释至少一种错误；评价后不追加网格、阈值或特征搜索。同源测试还不能代替新医院、不同亚型比例下的外部验证。")
p("表 11 的验证集已参与选配置，分数存在选择乐观。线性家族从 5 组中选，RBF 从 16 组中选，两边搜索范围也不等。样本数同样影响读数：只有 20 条 NO 和 42 条 AB，多错 1 条 NO 就使 BA 降 0.025，多对 1 条 AB 则升约 0.012；当前 0.045 的差距正来自多对 8 条 AB、同时多错 2 条 NO。")

h("九  冻结模型后输入扰动会怎样")
p("我原先的预测是近处容易翻、远处不会翻。更准确地说，远处也有机会翻，只是概率更低。实验冻结两个模型和训练 scaler，只对验证输入添加独立高斯噪声。每一维的噪声标准差设为对应训练标准差的 5%、10%、20%，每档重复 100 次，随机种子为 9，不改标签也不重新训练。翻转率比较加噪前后同一记录的预测是否改变。")
figure(8, "左图比较平均预测翻转率与线性理论值；右图比较平均 BA，误差棒为重复间的 ±1 标准差，不是置信区间。除零噪声基线外，每档重复 100 次。连线只连接已计算水平，不代表真实仪器误差曲线。")
noise_rows = []
for i, level in enumerate([0,5,10,20]):
    l, r = res["noise"]["linear"][i], res["noise"]["rbf"][i]
    noise_rows.append([f"{level}%", f"{l['mean_flip_rate']*100:.2f}%",
                       f"{res['noise_theory_linear'][i]*100:.2f}%",
                       f"{r['mean_flip_rate']*100:.2f}%",
                       f"{l['mean_balanced_accuracy']:.3f}",
                       f"{r['mean_balanced_accuracy']:.3f}"])
table(14, "人工输入噪声下的平均结果", ["噪声强度","线性翻转率","线性理论值","RBF 翻转率","线性 BA","RBF BA"], noise_rows, [2.4,3.1,3.2,3.2,2.4,2.4])
p("标准化空间中的噪声为 ε∼N(0,σ²I)，线性分数变化 wᵀε∼N(0,σ²‖w‖²)。一个距边界 d=|f|/‖w‖ 的点，跨到另一侧的概率为 Φ(−d/σ)。把 62 个点的概率取平均，理论翻转率依次为 2.75%、5.42%、11.53%，与模拟的 3.00%、5.44%、12.11% 接近。")
p("因此距离较近的点更容易改变预测；但预测稳定不等于预测正确。原本误分的点被噪声推过边界，反而可能变对，所以 BA 不必逐档严格下降。这里线性模型 5% 和 10% 噪声下的 BA 就有小幅波动。")
p("当前 RBF 配置的翻转率更高。仅加 5% 噪声，其平均 BA 就从 0.805 降到 0.746，已略低于线性的 0.750；到 20% 时，两者分别为 0.683 和 0.730。图 7 中局部边界换来的识别改善，在这组人工扰动下没有保持。100 次重复反映的是人工噪声的随机性，不是外部人群的稳定性；真实仪器噪声可能有偏差、相关性和不同尺度，需要另行测量。")

h("十  对 SVM 的综合理解")
p("SVM 学到的是决策函数，而不是一条只求把训练点分开的线。可分时，三条线都能零错误分类，最近点距离却分别为 0.383、0.957、0.383；最大间隔提供了进一步选择的原则。分不开时，目标函数用 hinge 记录违例深浅，再用 C 决定愿意付出多大的代价。矛盾点实验说明，这种权衡不是为了让所有训练记录都分对，而是为了在零错误不可能时仍有可求解的目标。")
p("边界与支持向量的联系在八点实验里很具体：w=0.5462(z₊−z₋)，删除远处的非支持向量不改变结果，删除 V061 则使原八点的最大分数变化达到 0.427。完整数据中的支持向量不只贴在参考线上，还包括带内和误分点。核函数把相同思路搬到特征空间；gamma 改变 RBF 相似度的局部范围，而输入尺度会先改变距离，再改变相似度。")
p("一个优势是：有清楚的间隔结构时，支持向量能给出紧凑的预测表示。八点实验只需两个点参与决策函数，其他点的 α 为零；但完整训练集 linear C=1 有 89 个支持向量，这种稀疏性就弱得多。核函数还让同一个优化思路表达非线性结构，验证集上的 RBF 因此多识别了 8 条 DH，也多误判了 2 条 NO。")
p("一个明确的局限是对输入尺度敏感，尤其在单位不同却直接用欧氏距离核时。第一个输入乘 100 后，无标准化的验证 BA 从 0.705 降到 0.488，标准化后则保持 0.773。另一个是复杂局部边界对扰动的敏感性：本次 20% 噪声下，RBF 的 BA 降到 0.683，低于线性的 0.730。宽 margin 或更灵活的核，都不能单独保证泛化。")
p("数据本身的限制也要分开看：这两项输入对 DH 和 NO 的区分不足，换核只能改善其中一部分。面对新任务，我会先检查样本量、核方法的计算成本、输入尺度与距离是否有意义、标签是否合并不同亚型、是否需要概率，以及有没有独立验证数据。核 SVC 的训练成本随样本数呈较差的超线性增长；大样本时应考虑线性求解器或核近似。")
p("这次能从数据里直接说明的，是几何、违例和错误类型怎样相互联系；不能据此声称模型已经适用于临床，更不能把 decision score 当成个体风险。")

h("参考资料与复现要点")
mapping = p("课程资料：student-worksheet.md、svm-lab.ipynb、data/README.md 与 scripts/lab_utils.py。活动 A 对应第一节，B 对应第二节，C 对应第三、四节，D 对应第五节，E 对应第六节，F 对应第七、八节，G 对应第九节；第十节为综合回答。", style=None)
mapping.paragraph_format.keep_with_next = True
refs = [
    "UCI Machine Learning Repository. Vertebral Column. DOI: 10.24432/C5K89B.",
    "Cortes C, Vapnik V. Support-vector networks. Machine Learning. 1995;20:273–297. DOI: 10.1007/BF00994018.",
    "scikit-learn. User Guide: Support Vector Machines; Probability calibration. https://scikit-learn.org/stable/modules/svm.html ; https://scikit-learn.org/stable/modules/calibration.html",
]
for text in refs:
    par = p(text)
    par.paragraph_format.space_after = Pt(3)
    for r in par.runs:
        r.font.size = Pt(9)
p("本次核验使用课程固定 train.csv、validation.csv 与 lab_utils.py 重跑几何、参数、尺度和扰动实验，并重新计算亚型中位数、验证分数分箱、小 gamma 的线性近似对照。未读取锁定测试标签。二分类和三分类源记录只在开发集 ID 范围内核对亚型，不用亚型标签选参。所有新图均由对应拟合模型与实际记录重绘。")
env = res["environment"]
p(f"核验环境：Python {env['python']}；NumPy {env['numpy']}；pandas {env['pandas']}；SciPy {env['scipy']}；scikit-learn {env['sklearn']}；Matplotlib {env['matplotlib']}。一般 SVC 的 tol=10⁻⁷、max_iter=200000；八点近似硬间隔使用 C=10000、tol=10⁻⁹，矛盾点实验 tol=10⁻⁸。噪声标准差使用训练总体标准差（ddof=0）。")
p("几何量采用未舍入模型值计算，展示时再舍入；手算使用工作单给定近似坐标。margin 违例按 hinge>10⁻⁵ 计数，SV 按模型的非零对偶系数识别。BA、敏感度、特异度与 AUC 分开报告，不把图中的显示裁切当作删样本。")

# A page number is useful in a multi-page worksheet; no decorative running header.
footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
field = OxmlElement("w:fldSimple")
field.set(qn("w:instr"), "PAGE")
footer._p.append(field)
for style in doc.styles:
    if style.type == 1:
        for e in style.element.xpath("./w:pPr/w:pBdr"):
            e.getparent().remove(e)
doc.core_properties.title = "支持向量机的间隔与分类边界"
doc.core_properties.subject = "第9讲作业 SVM与Margin"
doc.core_properties.author = "南铂山"
doc.core_properties.comments = ""
out = Path(r"C:\Users\A\Desktop\第9讲作业_SVM与Margin_Claude协作修订版.docx")
doc.save(out)
assert hashlib.sha256(Path(meta["source"]).read_bytes()).hexdigest() == meta["sha256"], "Original was modified"
text = "\n".join(par.text for par in doc.paragraphs)
(ROOT / "revised_text.txt").write_text(text, encoding="utf-8")
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs), "Tables:", len(doc.tables), "Images:", len(doc.inline_shapes))
