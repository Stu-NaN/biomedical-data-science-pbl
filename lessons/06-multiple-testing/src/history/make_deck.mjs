import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const workspaceDir = path.resolve(path.dirname(__filename), "..");
const buildDir = path.join(workspaceDir, "work", "ppt-build");
const previewDir = path.join(workspaceDir, "work", "ppt-preview");
const outputDir = path.join(workspaceDir, "outputs");
const candidatePath = path.join(buildDir, "candidate.pptx");
const finalPath = path.join(outputDir, "第6讲_多重检验校正_课堂展示.pptx");
const data = JSON.parse(
  await fs.readFile(path.join(workspaceDir, "work", "analysis.json"), "utf8"),
);
const runtimeDir =
  "C:/Users/A/.cache/codex-runtimes/codex-primary-runtime/dependencies";
const runtimeNode = path.join(runtimeDir, "node", "bin", "node.exe");
const runtimePython = path.join(runtimeDir, "python", "python.exe");
process.env.RUNTIME_NODE_MODULES = path.join(runtimeDir, "node", "node_modules");
const skillDir =
  "C:/Users/A/.codex/plugins/cache/openai-primary-runtime/presentations/26.1015.12040/skills/presentations";
const { finalizePresentation, applyPresentationChartFont } = await import(
  pathToFileURL(path.join(skillDir, "container_tools/artifact_tool_utils.mjs")).href
);

const W = 1280;
const H = 720;
const FONT = "Arial";
const EAST_ASIAN_FONT = "Microsoft YaHei";
const C = {
  ink: "#20292D",
  muted: "#56666A",
  teal: "#176D66",
  tealLight: "#DCEBE7",
  coral: "#D46C54",
  gold: "#C99728",
  paper: "#F7F8F4",
  white: "#FFFFFF",
  grid: "#D8E0DD",
  pale: "#EEF2EF",
  dark: "#20292D",
};

await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

const presentation = Presentation.create({
  slideSize: { width: W, height: H },
});

function addShape(slide, name, geometry, x, y, w, h, fill, lineFill = "none", lineWidth = 0) {
  return slide.shapes.add({
    geometry,
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: lineFill, width: lineWidth },
  });
}

function text(slide, name, value, x, y, w, h, opts = {}) {
  const shape = addShape(slide, name, "textbox", x, y, w, h, "none");
  shape.text = value;
  shape.text.style = {
    typeface: opts.typeface ?? EAST_ASIAN_FONT,
    fontSize: opts.size ?? 24,
    bold: opts.bold ?? false,
    color: opts.color ?? C.ink,
    alignment: opts.align ?? "left",
    verticalAlignment: opts.valign ?? "middle",
    autoFit: opts.autoFit ?? "shrinkText",
    wrap: "square",
    insets: opts.insets ?? 0,
    lineSpacing: opts.lineSpacing ?? 1.08,
  };
  return shape;
}

function addSlideBase(title, n, { dark = false } = {}) {
  const slide = presentation.slides.add();
  slide.background.fill = dark ? C.dark : C.paper;
  if (title) {
    text(slide, `title-${n}`, title, 68, 42, 1120, 58, {
      size: 42,
      bold: true,
      color: dark ? C.white : C.ink,
    });
  }
  text(
    slide,
    `footer-${n}`,
    "生物医学数据科学  ·  第 6 讲",
    68,
    678,
    440,
    22,
    { size: 15, color: dark ? "#D2DDDA" : C.muted },
  );
  text(slide, `folio-${n}`, String(n).padStart(2, "0"), 1162, 678, 50, 22, {
    size: 15,
    color: dark ? "#D2DDDA" : C.muted,
    align: "right",
  });
  return slide;
}

function addNotes(slide, body, sources = []) {
  const sourceText = sources.length ? `\n来源：${sources.join("；")}` : "";
  slide.speakerNotes.text = `${body}${sourceText}`;
}

function addChart(slide, type, config) {
  const chart = slide.charts.add(type, {
    chartFill: C.paper,
    chartLine: { style: "solid", fill: "none", width: 0 },
    plotAreaFill: C.paper,
    plotAreaLine: { style: "solid", fill: "none", width: 0 },
    ...config,
  });
  applyPresentationChartFont(chart, { fontFamily: EAST_ASIAN_FONT });
  return chart;
}

function styleTable(table, rows, cols, { headerSize = 18, bodySize = 17, firstColBold = false } = {}) {
  table.styleOptions = { headerRow: true, bandedRows: true };
  table.borders.assign({ style: "solid", fill: C.grid, width: 1 });
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r, c);
      cell.fill = r === 0 ? C.teal : r % 2 === 0 ? C.white : C.pale;
      cell.text.style = {
        typeface: EAST_ASIAN_FONT,
        fontSize: r === 0 ? headerSize : bodySize,
        color: r === 0 ? C.white : C.ink,
        bold: r === 0 || (firstColBold && c === 0),
        alignment: c === 0 ? "left" : "center",
        verticalAlignment: "middle",
        autoFit: "shrinkText",
        insets: { left: 10, right: 10, top: 6, bottom: 6 },
      };
    }
  }
}

function addTable(slide, values, x, y, w, h, opts = {}) {
  const rows = values.length;
  const cols = values[0].length;
  const table = slide.tables.add({
    rows,
    columns: cols,
    left: x,
    top: y,
    width: w,
    height: h,
    values,
  });
  styleTable(table, rows, cols, opts);
  return table;
}

function addRule(slide, name, x, y, w, color = C.grid, width = 1) {
  return addShape(slide, name, "line", x, y, w, 0, "none", color, width);
}

function addValue(
  slide,
  name,
  value,
  label,
  x,
  y,
  w,
  color = C.teal,
  labelColor = C.muted,
) {
  text(slide, `${name}-value`, value, x, y, w, 68, {
    size: 52,
    bold: true,
    color,
  });
  text(slide, `${name}-label`, label, x, y + 70, w, 34, {
    size: 20,
    color: labelColor,
  });
}

const sourceFiles = {
  overview: "6. multiple-test/README.md",
  lesson: "6. multiple-test/multiple-test.md",
  design: "6. multiple-test/experiment-design.md",
  data: "6. multiple-test/data/README.md",
  worksheet: "6. multiple-test/student-worksheet.md",
  teacher: "6. multiple-test/teacher-reference.md",
};

// 1. Cover
{
  const slide = addSlideBase(null, 1, { dark: true });
  text(slide, "cover-course", "生物医学数据科学  ·  第 6 讲", 74, 66, 600, 30, {
    size: 18,
    color: "#B9CBC6",
  });
  text(slide, "cover-title", "多重检验校正", 74, 158, 820, 92, {
    size: 66,
    bold: true,
    color: C.white,
  });
  text(slide, "cover-subtitle", "显著的结果，值得相信吗？", 78, 260, 760, 54, {
    size: 30,
    color: "#DCEBE7",
  });
  addRule(slide, "cover-separator", 78, 370, 760, "#53615F", 1);
  addValue(
    slide,
    "cover-probes",
    "7,129",
    "完整检验家族中的探针",
    78,
    405,
    280,
    "#F0BF55",
    "#D2DDDA",
  );
  addValue(
    slide,
    "cover-validation",
    "7 / 10",
    "预先锁定后通过留出判据",
    410,
    405,
    330,
    "#79C0B1",
    "#D2DDDA",
  );
  text(slide, "cover-question", "AML 与 ALL 的差异，哪些值得继续验证？", 790, 450, 410, 92, {
    size: 25,
    color: C.white,
    bold: true,
  });
  addNotes(
    slide,
    "开场先提出问题：当研究者同时检验成千上万个探针，raw p<0.05 的发现有多少值得相信？本次分析不是寻找最多显著项，而是固定检验家族、控制错误并追踪候选证据。这里的 7/10 是同一研究留出样本结果，不是外部独立验证。",
    [sourceFiles.data, sourceFiles.design],
  );
}

// 2. Driving question
{
  const slide = addSlideBase("1,000 次比较与 53 个显著结果", 2);
  text(
    slide,
    "scenario-lead",
    "假设团队同时比较 1,000 个基因，报告 53 个 raw p < 0.05。",
    72,
    132,
    1080,
    42,
    { size: 26, bold: true },
  );
  addValue(slide, "scenario-count", "≈ 50", "若 1,000 个零假设都成立，期望偶然显著数", 74, 226, 360, C.coral);
  addValue(slide, "scenario-fwer", "≈ 1", "独立全零检验时，至少错一次的概率", 494, 226, 340, C.teal);
  text(slide, "scenario-explain", "53 是假设情境里的观察数，不是本次白血病数据的结果。", 74, 354, 890, 38, {
    size: 22,
    color: C.muted,
  });
  const stripX = 74;
  const stripY = 441;
  const stripW = 1095;
  addShape(slide, "scenario-total-strip", "rect", stripX, stripY, stripW, 42, C.tealLight);
  addShape(slide, "scenario-false-strip", "rect", stripX, stripY, stripW * 0.05, 42, C.coral);
  text(slide, "scenario-strip-caption", "每项 α=0.05：全零时，期望假发现数 = m₀α = 1,000 × 0.05 = 50", 74, 506, 1090, 42, {
    size: 21,
    bold: true,
  });
  text(slide, "scenario-boundary", "期望值不是每次必然出现的精确数量；期望假发现数不要求检验彼此独立。", 74, 560, 1090, 44, {
    size: 19,
    color: C.muted,
  });
  addNotes(
    slide,
    "53/1,000 是课程导入的假设案例，不能把它当作真实数据输出。若 1,000 个零假设都成立且每项检验有效，E[V]=m0×alpha=50，这个期望的推导用线性期望，不需要独立。FWER=1-(1-alpha)^m 的公式需要额外的独立且各项恰好以 alpha 错误拒绝的条件，所以这里称为独立全零情境下约等于 1。",
    [sourceFiles.lesson, sourceFiles.worksheet, sourceFiles.teacher],
  );
}

// 3. Data and protocol
{
  const slide = addSlideBase("研究数据与预先固定的分析协议", 3);
  text(slide, "protocol-summary", "比较对象：两种白血病亚型的探针表达，不是患者与健康人的差异。", 74, 126, 1090, 44, {
    size: 24,
    bold: true,
  });
  addValue(slide, "dataset-samples", "72", "份表达样本", 75, 210, 190);
  addValue(slide, "dataset-probes", "7,129", "个微阵列探针", 320, 210, 230);
  addValue(slide, "dataset-family", "m = 7,129", "发现阶段完整家族", 650, 210, 310, C.coral);
  addRule(slide, "protocol-rule", 74, 330, 1090);
  const vals = [
    ["阶段", "ALL", "AML", "用途"],
    ["train · 38", "27", "11", "发现、效应估计、锁定候选"],
    ["test · 34", "20", "14", "锁定后留出检查"],
  ];
  addTable(slide, vals, 74, 358, 1090, 150, { headerSize: 18, bodySize: 19 });
  text(slide, "protocol-method", "每探针双侧 Welch t 检验  ·  H₀：两组总体均值相等  ·  效应方向：AML − ALL", 74, 542, 1100, 38, {
    size: 21,
    bold: true,
    color: C.teal,
  });
  text(slide, "protocol-caveat", "保留处理后的表达尺度；含负值，不取 log。探针不等于独立基因。", 74, 592, 1090, 34, {
    size: 18,
    color: C.muted,
  });
  addNotes(
    slide,
    "说明数据流程。CSV 有 72 行样本与 7,129 列探针；train 的 ALL/AML 为 27/11，用于发现。test 的 20/14 仅在候选锁定后使用。一个探针一项双侧 Welch t 检验，检验家族在看结果前固定为全部探针。文件是处理后的微阵列表达强度，含负数，因此不直接取 log，也不把均值差写成 fold change。",
    [sourceFiles.data, sourceFiles.design, sourceFiles.worksheet],
  );
}

// 4. Error concepts
{
  const slide = addSlideBase("发现集合中的错误：FWER 与 FDR", 4);
  text(slide, "error-example-label", "单次实验示意：宣布 R = 20 项，其中 V = 2 项是假发现", 74, 130, 1090, 38, {
    size: 23,
    bold: true,
  });
  const x = 74;
  const y = 216;
  const total = 1090;
  const item = total / 20;
  for (let i = 0; i < 20; i++) {
    addShape(
      slide,
      `error-dot-${i}`,
      "ellipse",
      x + i * item + 4,
      y,
      37,
      37,
      i < 2 ? C.coral : C.teal,
    );
  }
  text(slide, "error-fdp", "FDP = V / max(R,1) = 2 / 20 = 0.10", 74, 285, 510, 50, {
    size: 25,
    bold: true,
    color: C.coral,
  });
  addRule(slide, "error-divider", 74, 369, 1090);
  text(slide, "fwer-label", "FWER", 74, 405, 210, 38, { size: 28, bold: true, color: C.teal });
  text(slide, "fwer-description", "P(V > 0)：至少一次出现假发现的概率", 74, 451, 520, 62, {
    size: 20,
  });
  text(slide, "fdr-label", "FDR", 670, 405, 200, 38, { size: 28, bold: true, color: C.teal });
  text(slide, "fdr-description", "重复研究中，单次 FDP 的期望：E[FDP]", 670, 451, 450, 62, {
    size: 21,
  });
  text(slide, "error-boundary", "FDR ≤ 0.05 不等于本次 FDP ≤ 0.05，也不表示每项发现有 95% 概率为真。", 74, 555, 1090, 44, {
    size: 20,
    bold: true,
  });
  addNotes(
    slide,
    "图中红色点是一次实验里的两个错误发现，绿色点是其余 18 项。该次 FDP=2/20=0.10。FWER 问的是多次重复中有多少比例至少犯一次错；FDR 是每次发现集错误比例的长期期望。两者控制对象不同，因此 BH 可以有较高的 FWER，同时保持 FDR 在目标附近。FDR 不是单个结果为真的概率。",
    [sourceFiles.lesson, sourceFiles.teacher],
  );
}

// 5. All-null simulation
{
  const slide = addSlideBase("实验 A：没有真信号，也会出现显著结果", 5);
  text(slide, "null-setup", "1,000 个独立正态变量  ·  n=27 与 11  ·  1,000 次重复  ·  种子 42", 74, 124, 1120, 34, {
    size: 20,
    color: C.muted,
  });
  addChart(slide, "bar", {
    position: { left: 60, top: 178, width: 690, height: 380 },
    categories: ["raw", "Bonferroni", "BH"],
    series: [
      {
        name: "FWER / FDR",
        values: [100, 5.1, 5.1],
        fill: C.teal,
      },
    ],
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 90 },
    hasLegend: false,
    xAxis: {
      min: 0,
      max: 100,
      majorUnit: 25,
      numberFormatCode: "0",
      textStyle: { fill: C.muted, fontSize: 17 },
      majorGridlines: { style: "solid", fill: C.grid, width: 1 },
    },
    yAxis: {
      textStyle: { fill: C.ink, fontSize: 20 },
      line: { style: "solid", fill: "none", width: 0 },
    },
    dataLabels: {
      showValue: true,
      position: "outEnd",
      numberFormatCode: "0.0",
      textStyle: { fill: C.ink, fontSize: 17, bold: true },
    },
  });
  text(slide, "null-mean-title", "平均发现数 R", 816, 192, 300, 34, {
    size: 23,
    bold: true,
  });
  addValue(slide, "null-raw-mean", "49.944", "raw p", 816, 247, 280, C.coral);
  addValue(slide, "null-adjusted-mean", "≈ 0.05", "Bonferroni 与 BH", 816, 394, 310, C.teal);
  text(slide, "null-first-run", "第一次：raw 42，Bonferroni 0，BH 0", 816, 552, 360, 44, {
    size: 18,
    color: C.muted,
  });
  text(slide, "null-chart-note", "图示 FWER（%）；全零时有发现即 FDP=1，因此 FDR=FWER。", 72, 612, 1090, 28, {
    size: 17,
    color: C.muted,
  });
  addNotes(
    slide,
    "所有模拟变量真值都为零，故每项发现都是假发现。1,000 次中 raw 平均发现 49.944 项，FWER/FDR 估计为 1；Bonferroni 与 BH 的平均发现约为 0.05，FWER/FDR 估计约 0.051。第一次全零实验恰好 raw 有 42 项，但 Bonferroni 和 BH 为 0，这不能代表长期错误率。全零时本次有发现则 FDP=1，无发现则 FDP=0，所以重复均值的 FDR 与 FWER 相等。",
    [sourceFiles.teacher, sourceFiles.design],
  );
}

// 6. Partial signal simulation
{
  const slide = addSlideBase("实验 B：错误控制与检出力的取舍", 6);
  text(slide, "signal-setup", "1,000 个变量中前 100 个有信号；δ 从 1.0 降到 0.5，噪声标准差不变。", 74, 124, 1100, 34, {
    size: 20,
    color: C.muted,
  });
  const rows = [
    ["偏移 δ", "方法", "平均 R", "平均 V", "FWER", "FDR", "power"],
    ["1.0", "raw", "122.523", "44.799", "1.000", "0.364", "77.72%"],
    ["1.0", "Bonferroni", "6.001", "0.044", "0.044", "0.007", "5.96%"],
    ["1.0", "BH", "30.094", "1.445", "0.754", "0.046", "28.65%"],
    ["0.5", "raw", "72.361", "44.799", "1.000", "0.619", "27.56%"],
    ["0.5", "Bonferroni", "0.270", "0.044", "0.044", "0.039", "0.23%"],
    ["0.5", "BH", "0.407", "0.072", "0.070", "0.045", "0.34%"],
  ];
  addTable(slide, rows, 74, 174, 1090, 330, {
    headerSize: 17,
    bodySize: 17,
    firstColBold: true,
  });
  text(slide, "signal-fdr-not-fwer", "BH：δ=1.0 时 FWER=0.754，但 FDR=0.046", 74, 526, 1090, 44, {
    size: 25,
    bold: true,
    color: C.coral,
  });
  text(slide, "signal-explanation", "BH 控制的是发现集合的 FDR，不是“至少一次错”的 FWER；较弱信号也会降低 power。", 74, 578, 1090, 42, {
    size: 19,
  });
  addNotes(
    slide,
    "前 100 个变量是真信号，剩余 900 个为零；每种效应重复 1,000 次。δ=1 时 BH 平均 power 为 28.65%，Bonferroni 为 5.96%，但 BH 的 FWER 估计为 0.754，FDR 约 0.046。这个例子直观显示 FWER 与 FDR 不可互换。δ=0.5 时两种校正方法 power 都很低；没有检出并不能推翻模拟中已知的真信号。表格使用教师参考中同一套固定种子与生成机制，并由本次数值实现复现。",
    [sourceFiles.design, sourceFiles.teacher],
  );
}

// 7. Real discovery results
{
  const slide = addSlideBase("实验 C：真实数据的三种发现集合", 7);
  text(slide, "real-setup", "train：ALL 27、AML 11  ·  双侧 Welch 检验  ·  完整家族 m=7,129", 74, 124, 1100, 34, {
    size: 20,
    color: C.muted,
  });
  addChart(slide, "bar", {
    position: { left: 65, top: 178, width: 760, height: 390 },
    categories: ["raw p ≤ 0.05", "Bonferroni adjusted p ≤ 0.05", "BH adjusted p ≤ 0.05"],
    series: [
      {
        name: "探针数",
        values: [1694, 26, 535],
        fill: C.teal,
        points: [
          { idx: 0, fill: C.coral },
          { idx: 1, fill: C.gold },
          { idx: 2, fill: C.teal },
        ],
      },
    ],
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 75 },
    hasLegend: false,
    xAxis: {
      min: 0,
      max: 1800,
      majorUnit: 600,
      numberFormatCode: "#,##0",
      textStyle: { fill: C.muted, fontSize: 16 },
      majorGridlines: { style: "solid", fill: C.grid, width: 1 },
    },
    yAxis: {
      textStyle: { fill: C.ink, fontSize: 18 },
      line: { style: "solid", fill: "none", width: 0 },
    },
    dataLabels: {
      showValue: true,
      position: "outEnd",
      numberFormatCode: "#,##0",
      textStyle: { fill: C.ink, fontSize: 18, bold: true },
    },
  });
  addValue(slide, "real-bonf-threshold", "7.01×10⁻⁶", "Bonferroni 对应 raw p 阈值", 860, 242, 330, C.coral);
  text(slide, "real-family-note", "535 是 BH 规则下的发现数，不是真实阳性数。", 860, 416, 330, 74, {
    size: 22,
    bold: true,
  });
  text(slide, "real-truth-boundary", "真实 CSV 没有逐探针真值标签，不能直接计算 V、FDP 或真实 power。", 74, 595, 1090, 42, {
    size: 20,
    color: C.muted,
  });
  addNotes(
    slide,
    "完整发现家族的统计重算得到 raw p≤0.05 的 1,694 个探针，Bonferroni 的 26 个，BH 的 535 个。Bonferroni raw p 阈值为 0.05/7,129=7.0136×10^-6。真实数据里不知道哪些探针的零假设为真，因此不能把 raw 减 BH 的数量解释成假发现，也不能把 535 称为真阳性。探针间相关性与一个基因对应多个探针也限制生物学解释。",
    [sourceFiles.data, sourceFiles.teacher],
  );
}

// 8. Family selection issue
{
  const slide = addSlideBase("先挑最小 p，再按 20 项校正，会漏掉选择机会", 8);
  text(slide, "selection-explain", "首位候选 U22376_cds2_s_at：同一个 raw p，按不同家族计算出不同 adjusted p。", 74, 126, 1100, 48, {
    size: 22,
  });
  const rows = [
    ["校正的家族", "m", "BH adjusted p", "能否支持原始搜索？"],
    ["全部预先定义探针", "7,129", "5.98×10⁻⁵", "可以，覆盖完整发现家族"],
    ["先看结果再选前 20", "20", "1.68×10⁻⁷", "不可以，忽略了选中机会"],
  ];
  addTable(slide, rows, 74, 224, 1090, 194, { headerSize: 18, bodySize: 21 });
  addValue(slide, "selection-ratio", "≈ 356×", "事后 m=20 让调整值变小", 74, 470, 350, C.coral);
  text(slide, "selection-principle", "家族需预先定义，不能事后缩小。", 485, 502, 670, 64, {
    size: 24,
    bold: true,
    color: C.teal,
  });
  text(slide, "selection-exception", "若 20 个探针在看这批数据前已由外部依据指定，m=20 才可能是相应研究的家族。", 74, 598, 1090, 40, {
    size: 19,
    color: C.muted,
  });
  addNotes(
    slide,
    "比较只在同一批数据里事后选 raw p 最小的 20 个，再把这 20 项当作唯一检验家族的错误流程。首位探针的完整家族 BH adjusted p=5.9803e-5；m=20 会报成1.6777e-7，缩小约356倍。即便该探针两种算法下都显著，事后家族也没有提供原 7,129 次搜索的错误控制。只有在分析前就根据独立信息指定并只检验 20 项，才可能合法定义 m=20。",
    [sourceFiles.worksheet, sourceFiles.teacher],
  );
}

// 9. Effect size
{
  const slide = addSlideBase("候选探针：统计差异还要看方向与尺度", 9);
  const d = data.top_probe_detail;
  text(slide, `effect-probe`, d.probe, 74, 124, 1090, 34, {
    size: 21,
    bold: true,
    color: C.teal,
  });
  addChart(slide, "bar", {
    position: { left: 66, top: 174, width: 590, height: 362 },
    categories: ["ALL  (n=27)", "AML  (n=11)"],
    series: [
      {
        name: "平均表达强度",
        values: [d.mean_ALL, d.mean_AML],
        fill: C.teal,
        points: [
          { idx: 0, fill: C.teal },
          { idx: 1, fill: C.coral },
        ],
      },
    ],
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 120 },
    hasLegend: false,
    xAxis: {
      textStyle: { fill: C.ink, fontSize: 19 },
      line: { style: "solid", fill: C.grid, width: 1 },
    },
    yAxis: {
      min: 0,
      max: 4500,
      majorUnit: 1500,
      numberFormatCode: "#,##0",
      textStyle: { fill: C.muted, fontSize: 16 },
      majorGridlines: { style: "solid", fill: C.grid, width: 1 },
    },
    dataLabels: {
      showValue: true,
      position: "outEnd",
      numberFormatCode: "#,##0",
      textStyle: { fill: C.ink, fontSize: 18, bold: true },
    },
  });
  text(slide, "effect-stat", "AML − ALL = −3,160.93", 720, 198, 470, 48, {
    size: 29,
    bold: true,
    color: C.coral,
  });
  text(slide, "effect-ci", "95% 逐项 Welch CI\n[−3,977.85, −2,344.02]", 720, 272, 480, 84, {
    size: 23,
  });
  text(slide, "effect-d", "Cohen d = −1.85", 720, 380, 420, 40, {
    size: 24,
    bold: true,
  });
  text(slide, "effect-p", "BH adjusted p = 5.98×10⁻⁵", 720, 433, 420, 40, {
    size: 21,
    color: C.teal,
  });
  text(slide, "effect-scale", "负值表示 AML 测量值较低；这是强度均值差，不是 fold change。", 720, 501, 465, 76, {
    size: 19,
    color: C.muted,
  });
  text(slide, "effect-ci-caveat", "逐项 CI 未控制 7,129 项家族覆盖率；筛选后的效应可能偏大。", 74, 590, 1090, 38, {
    size: 19,
    color: C.muted,
  });
  addNotes(
    slide,
    `候选 ${d.probe} 的发现组 AML 均值为 ${d.mean_AML.toFixed(2)}，ALL 均值为 ${d.mean_ALL.toFixed(2)}，差值 AML-ALL=${d.delta.toFixed(2)}。95% 区间是该探针的逐项 Welch 区间，不是同时置信区间，也不调整按显著性筛选带来的选择偏倚。Cohen d 使用合并标准差作描述性标准化。原始表达强度范围含负数，均值差不能称为倍数变化。`,
    [sourceFiles.data, sourceFiles.teacher],
  );
}

// 10. Held-out validation
{
  const slide = addSlideBase("预先锁定的 10 个候选：7 个通过留出判据", 10);
  text(slide, "validation-rule", "判据：test 候选家族内 Bonferroni ≤ 0.05，且 AML−ALL 方向与 train 相同", 74, 120, 1110, 42, {
    size: 19,
    color: C.muted,
  });
  const rows = [
    ["序", "探针", "train Δ", "test Δ", "test Bonf p", "结果"],
    ...data.candidates.map((r, i) => [
      String(i + 1),
      r.probe,
      Math.round(r.delta).toLocaleString("en-US"),
      Math.round(r.test_delta).toLocaleString("en-US"),
      r.test_bonferroni < 0.001
        ? r.test_bonferroni.toExponential(2)
        : r.test_bonferroni.toFixed(3),
      r.replicated ? "通过" : "未达阈值",
    ]),
  ];
  const table = addTable(slide, rows, 54, 178, 1172, 412, {
    headerSize: 16,
    bodySize: 16,
    firstColBold: true,
  });
  text(slide, "validation-result", "10 项方向一致；7 项达到预定显著性阈值。", 74, 612, 1090, 34, {
    size: 22,
    bold: true,
    color: C.teal,
  });
  addNotes(
    slide,
    "在开启 test 结果之前，按发现 BH 显著、raw p 升序和探针名破同分规则锁定了 10 项。候选家族大小是 10，因此 test 原始 p 的等价 Bonferroni 阈值为0.005；实际使用 adjusted p≤0.05，并检查方向一致。10 项方向均一致，其中7项通过。其余3项未达预定显著性阈值，这不能证明没有效应，也不能事后换入其他候选。该 test 为同一研究历史留出样本，不是外部队列。",
    [sourceFiles.design, sourceFiles.teacher],
  );
}

// 11. Method choice and limits
{
  const slide = addSlideBase("方法选择取决于错误代价", 11);
  const rows = [
    ["研究目的", "优先考虑", "理由", "代价 / 后续"],
    ["探索性组学筛选", "BH · 控制 FDR", "需要形成可验证候选，通常比 FWER 方法检出力高", "不能把每项显著当作已证实"],
    ["确认性主要终点", "Bonferroni · 控制 FWER", "任一假阳性代价高，且家族应预先定义", "阈值更严格，可能漏掉真信号"],
  ];
  addTable(slide, rows, 74, 154, 1090, 244, {
    headerSize: 17,
    bodySize: 19,
    firstColBold: true,
  });
  text(slide, "limits-title", "本次数据仍未解决的问题", 74, 449, 520, 38, {
    size: 25,
    bold: true,
    color: C.teal,
  });
  const limits = [
    "7,129 个探针不是 7,129 个独立基因证据，探针相关结构未全面验证。",
    "字段不足以排除患者重复、批次或中心混杂；Welch 检验依赖样本独立。",
    "留出验证不等于外部验证；机制、临床意义和因果需要额外证据。",
  ];
  limits.forEach((line, i) => {
    text(slide, `limit-${i}`, line, 74, 504 + i * 48, 1090, 38, {
      size: 19,
      color: C.ink,
    });
  });
  addNotes(
    slide,
    "选择方法要回到研究问题。探索性组学筛选常以 BH 控制 FDR，目标是管理发现集合的长期错误比例，同时保留候选；确认性临床主要终点中，如果任一假阳性代价很高，控制 FWER 的 Bonferroni 更符合目标。Bonferroni 不要求检验独立，但要求单项检验 p 值有效。BH 的经典 FDR 保证依赖独立或适当正依赖结构。数据字段未充分验证探针依赖、患者独立性和批次影响，所以真实数据分析保持探索性。",
    [sourceFiles.lesson, sourceFiles.data, sourceFiles.teacher],
  );
}

// 12. Main conclusion
{
  const slide = addSlideBase("结论：可以支持什么主张？", 12, { dark: true });
  const labels = ["主张", "证据", "推理", "边界"];
  const lines = [
    "AML 与 ALL 之间存在若干值得继续验证的探针差异。",
    "7,129 项完整家族中 535 项 BH 显著；锁定 10 项中 7 项通过 test 候选家族规则。",
    "完整校正减少选择性报告风险，锁定后的留出结果增加同研究内的重复证据。",
    "不能据此证明 535 项都是真实独立基因、疾病因果、临床价值或外部人群可重复。",
  ];
  labels.forEach((label, i) => {
    const y = 140 + i * 115;
    text(slide, `conclusion-label-${i}`, label, 74, y, 150, 42, {
      size: 23,
      bold: true,
      color: i === 3 ? "#F0BF55" : "#79C0B1",
    });
    text(slide, `conclusion-text-${i}`, lines[i], 245, y, 930, 75, {
      size: 24,
      color: C.white,
      bold: i === 0,
    });
  });
  addNotes(
    slide,
    "按主张、证据、推理、边界组织结论。正确的措辞是“支持进一步研究的探索性证据”，而不是“发现了疾病基因”或“证明了因果”。特别强调7/10是同研究内历史留出验证，且候选名单和判据在看 test 前固定。",
    [sourceFiles.data, sourceFiles.teacher],
  );
}

// 13. Oral questions
{
  const slide = addSlideBase("课堂追问备答", 13);
  const items = [
    [
      "BH 显著是否代表每个探针有 95% 概率为真？",
      "不是。FDR 是重复分析中发现集合错误比例的期望，不是单项后验概率。",
    ],
    [
      "为什么 Bonferroni 只找到 26 项，BH 找到 535 项？",
      "两者控制目标不同；FWER 通常更严格，BH 在合适条件下以 FDR 换取更高检出力。",
    ],
    [
      "7 个候选通过 test，能否叫外部验证？",
      "不能。test 来自同一研究；外部验证需要独立收集的队列与预先锁定的检验规则。",
    ],
    [
      "为什么不能把 535 项称为真基因？",
      "真实数据没有逐项真值；探针可能相关，一个基因也可能对应多个探针。",
    ],
  ];
  items.forEach(([q, a], i) => {
    const y = 130 + i * 128;
    text(slide, `qa-q-${i}`, q, 74, y, 1090, 42, {
      size: 22,
      bold: true,
      color: C.teal,
    });
    text(slide, `qa-a-${i}`, a, 74, y + 48, 1090, 58, {
      size: 20,
      color: C.ink,
    });
    if (i < items.length - 1) addRule(slide, `qa-line-${i}`, 74, y + 114, 1090, C.grid, 1);
  });
  addNotes(
    slide,
    "回答追问时优先回到自己的数据。若老师问“这个数反过来结论会变吗”，解释 BH 和 Bonferroni 代表不同的错误控制目标；若问“结论不能说明什么”，要说清真实数据没有逐项真值、探针相关、test 不是外部队列，以及统计关联不能证明因果。",
    [sourceFiles.teacher, sourceFiles.worksheet],
  );
}

// 14. Sources and AI disclosure
{
  const slide = addSlideBase("数据来源与分析说明", 14);
  text(slide, "source-dataset-head", "数据", 74, 130, 140, 32, {
    size: 21,
    bold: true,
    color: C.teal,
  });
  text(slide, "source-dataset", "OpenML leukemia，ID 1104，version 1；Golub et al., Science 286 (1999), 531–537.", 238, 130, 900, 48, {
    size: 18,
  });
  text(slide, "source-method-head", "方法", 74, 202, 140, 32, {
    size: 21,
    bold: true,
    color: C.teal,
  });
  text(slide, "source-method", "Benjamini & Hochberg (1995), JRSS B 57(1), 289–300；Storey (2002), JRSS B 64(3), 479–498.", 238, 202, 900, 55, {
    size: 18,
  });
  text(slide, "source-course-head", "课程材料", 74, 286, 140, 32, {
    size: 21,
    bold: true,
    color: C.teal,
  });
  text(slide, "source-course", "multiple-test.md · experiment-design.md · data/README.md · student-worksheet.md · teacher-reference.md", 238, 286, 900, 64, {
    size: 17,
  });
  addRule(slide, "source-ai-rule", 74, 376, 1090);
  text(slide, "source-ai-head", "AI 使用", 74, 408, 150, 32, {
    size: 21,
    bold: true,
    color: C.teal,
  });
  text(
    slide,
    "source-ai",
    "辅助整理结构、编写统计核验程序与文字校对。真实数据和固定种子模拟均按附件配置重算，并与随包教师参考汇总核对。汇报者应能解释公式、候选名单和验证判据。",
    238,
    402,
    900,
    96,
    { size: 19 },
  );
  addNotes(
    slide,
    "数据出处与报告方式。结果基于提供的 CSV 与元数据快照；学术方法来源见 Benjamini-Hochberg 与 Storey 文献。压缩包中的 teacher-reference 用于核对模拟数字和分析边界。群文件 plan.docx 不在本次提供文件中，因此没有引用或臆测其中未见的内容。",
    [sourceFiles.data, sourceFiles.lesson, sourceFiles.teacher],
  );
}

const slideCount = presentation.slides.items.length;
if (slideCount !== 14) throw new Error(`Expected 14 slides, got ${slideCount}`);

const requiredOutputs = {
  explicitTotalSlideCount: 14,
  requiredNativeChartOwnerSlides: [5, 7, 9],
  requiredNativeTableOwnerSlides: [3, 6, 8, 10, 11],
  materializeLiteralChartWorkbooks: true,
  nativeChartTargetApplication: "powerpoint",
};

for (let i = 0; i < slideCount; i++) {
  const slide = presentation.slides.getItem(i);
  const blob = await presentation.export({ slide, format: "png", scale: 1 });
  const pngPath = path.join(previewDir, `slide-${String(i + 1).padStart(2, "0")}.png`);
  await fs.writeFile(pngPath, new Uint8Array(await blob.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(
    path.join(previewDir, `slide-${String(i + 1).padStart(2, "0")}.layout.json`),
    await layout.text(),
  );
}

const montage = await presentation.export({
  format: "webp",
  montage: true,
  scale: 0.35,
});
await fs.writeFile(
  path.join(previewDir, "montage.webp"),
  new Uint8Array(await montage.arrayBuffer()),
);

const inspect = await presentation.inspect({
  kind: "slide,textbox,shape,chart,table,notes",
  maxChars: 300000,
});
await fs.writeFile(path.join(previewDir, "inspect.ndjson"), inspect.ndjson, "utf8");

const pptxBlob = await PresentationFile.exportPptx(presentation);
await pptxBlob.save(candidatePath);

const finalization = await finalizePresentation({
  ...requiredOutputs,
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(
    skillDir,
    "container_tools/inspect_presentation_package_integrity.py",
  ),
  layoutValidatorPath: path.join(
    skillDir,
    "container_tools/inspect_presentation_layout_geometry.py",
  ),
  layoutArgs: [
    "--expected-slide-size-emu",
    "12192000,6858000",
    "--validate-heading-fit",
    ...requiredOutputs.requiredNativeTableOwnerSlides.flatMap((slideNo) => [
      "--require-native-table-slide",
      String(slideNo),
    ]),
  ],
  fontPolicy: {
    basis: "design",
    families: [FONT, EAST_ASIAN_FONT],
    scriptFonts: { ea: EAST_ASIAN_FONT },
  },
  verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, "validation.json"),
});

console.log(
  JSON.stringify(
    {
      slideCount,
      finalPath,
      candidatePath,
      previewDir,
      finalization,
    },
    null,
    2,
  ),
);
