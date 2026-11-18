import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const SKILL_DIR = "C:/Users/A/.codex/plugins/cache/openai-primary-runtime/presentations/26.927.11222/skills/presentations";
const WORKSPACE_DIR = "C:/Users/A/ppt_high_dimension_reflection";
const BUILD_DIR = path.join(WORKSPACE_DIR, ".codex-build");
const OUT_DIR = path.join(WORKSPACE_DIR, "out");
const CANDIDATE = path.join(BUILD_DIR, "student-package-class3-learning-reflection-candidate.pptx");
const NOTES_JSON = path.join(BUILD_DIR, "speaker-notes.json");
const ASSET_DIR = path.join(BUILD_DIR, "assets-v8");

const { resolvePresentationFont, applyPresentationChartFont } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href,
);

await fs.mkdir(BUILD_DIR, { recursive: true });
await fs.mkdir(OUT_DIR, { recursive: true });

const family = resolvePresentationFont({ fontFamily: "Microsoft YaHei" });
const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
});

const ASSETS = {
  manifold: await fs.readFile(path.join(ASSET_DIR, "actin.png")),
  scatter3d: await fs.readFile(path.join(ASSET_DIR, "microarray.jpg")),
  distanceHist: await fs.readFile(path.join(ASSET_DIR, "actin.png")),
  tech: await fs.readFile(path.join(ASSET_DIR, "tech-network.jpg")),
};

const C = {
  ink: "#102C46",
  deep: "#08243F",
  teal: "#1598C4",
  sea: "#5BD8E7",
  mint: "#E0F5F7",
  cream: "#F3F8FC",
  paper: "#FFFFFF",
  sand: "#D1DEEA",
  muted: "#5E7183",
  coral: "#FF7D68",
  gold: "#F6C84F",
  red: "#D95D68",
  white: "#FFFFFF",
  grid: "#D6E4F2",
  slate: "#48657E",
  lilac: "#E9E4FA",
  peach: "#FFE1DA",
  sky: "#DCEEFF",
};

const EVIDENCE = {
  plan: "C:/Users/A/Downloads/plan.docx",
  worksheet: "C:/Users/A/Downloads/student-package-class3-completed-worksheet.docx",
  sourcePackage: "C:/Users/A/Downloads/student-package-class3.zip",
  readings: "readings/tan-idm2e-ch02-data.pdf; readings/tan-idm2e-appendix-a-b.pdf; readings/tan-idm2e-ch03-overfitting.pdf; readings/cornell-cs4780-lec02-knn-curse-of-dimensionality.pdf",
};

function lineNone() {
  return { style: "solid", fill: "none", width: 0 };
}

function line(color = C.sand, width = 1) {
  return { style: "solid", fill: color, width };
}

function addShape(slide, geometry, x, y, w, h, fill, options = {}) {
  return slide.shapes.add({
    geometry,
    name: options.name,
    position: { left: x, top: y, width: w, height: h },
    fill: fill ?? "none",
    line: options.line ?? lineNone(),
    borderRadius: options.borderRadius,
    shadow: options.shadow,
  });
}

function addText(slide, textValue, x, y, w, h, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: lineNone(),
  });
  shape.text = textValue;
  shape.text.style = {
    typeface: options.typeface ?? family,
    fontSize: options.fontSize ?? 20,
    color: options.color ?? C.ink,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    alignment: options.alignment ?? "left",
    verticalAlignment: options.verticalAlignment ?? "top",
    autoFit: "shrinkTextOnOverflow",
  };
  return shape;
}

function addRichText(slide, runs, x, y, w, h, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: lineNone(),
  });
  shape.text = runs;
  shape.text.style = {
    typeface: options.typeface ?? family,
    fontSize: options.fontSize ?? 20,
    color: options.color ?? C.ink,
    alignment: options.alignment ?? "left",
    verticalAlignment: options.verticalAlignment ?? "top",
    autoFit: "shrinkTextOnOverflow",
  };
  return shape;
}

function addFooter(slide, textValue, dark = false) {
  addText(slide, textValue, 72, 684, 1110, 18, {
    fontSize: 12,
    color: dark ? "#B9CECC" : C.muted,
  });
  addText(slide, String(presentation.slides.items.length).padStart(2, "0"), 1158, 684, 52, 18, {
    fontSize: 12,
    color: dark ? "#B9CECC" : C.muted,
    alignment: "right",
  });
}

function addHeader(slide, index, title, subtitle = "", visual = null) {
  slide.background.fill = C.cream;
  if (visual) {
    addBackdrop(
      slide,
      visual.asset,
      visual.left,
      visual.top,
      visual.width,
      visual.height,
      {
        alt: visual.alt,
        wash: visual.wash ?? "#F7F3EA/72",
        geometry: visual.geometry ?? "roundRect",
        borderRadius: visual.borderRadius ?? "rounded-3xl",
      },
    );
  }
  addShape(slide, "rect", 0, 0, 10, 720, C.teal);
  addText(slide, `${String(index).padStart(2, "0")}  /  学习复盘`, 72, 42, 400, 22, {
    fontSize: 13,
    color: C.teal,
    bold: true,
  });
  addText(slide, title, 72, 78, 1110, 58, {
    fontSize: 34,
    color: C.ink,
    bold: true,
  });
  if (subtitle) {
    addText(slide, subtitle, 72, 140, 1090, 34, {
      fontSize: 17,
      color: C.muted,
    });
  }
}

function addNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
  slide.speakerNotes.setVisible(true);
}

function addStat(slide, value, label, x, y, w, color = C.ink) {
  addText(slide, value, x, y, w, 60, { fontSize: 44, color, bold: true });
  addText(slide, label, x, y + 62, w, 34, { fontSize: 16, color: C.muted });
}

function addDot(slide, x, y, radius, fill) {
  addShape(slide, "ellipse", x, y, radius, radius, fill, { line: line(fill, 0) });
}

function addImage(slide, assetName, x, y, w, h, options = {}) {
  return slide.images.add({
    blob: ASSETS[assetName],
    contentType: options.contentType ?? (assetName === "tech" ? "image/jpeg" : "image/png"),
    alt: options.alt ?? `${assetName} background visual`,
    fit: options.fit ?? "cover",
    position: { left: x, top: y, width: w, height: h },
    crop: options.crop,
    geometry: options.geometry,
    borderRadius: options.borderRadius,
  });
}

function addBackdrop(slide, assetName, x, y, w, h, options = {}) {
  addImage(slide, assetName, x, y, w, h, {
    alt: options.alt,
    fit: options.fit ?? "cover",
    crop: options.crop,
    geometry: options.geometry ?? "roundRect",
    borderRadius: options.borderRadius ?? "rounded-3xl",
  });
  if (options.wash) {
    addShape(slide, options.geometry ?? "roundRect", x, y, w, h, options.wash, {
      borderRadius: options.borderRadius ?? "rounded-3xl",
      line: line(options.lineColor ?? "transparent", options.lineWidth ?? 0),
    });
  }
}

function addFullBleedBackground(slide, assetName, wash = "#F7F3EA/78", alt = "") {
  addImage(slide, assetName, 0, 0, 1280, 720, {
    alt: alt || `${assetName} full-page background`,
    fit: "cover",
  });
  addShape(slide, "rect", 0, 0, 1280, 720, wash);
}

const notes = {};
const TECH_BG_NOTE =
  "背景图：Unsplash 科技网络背景图（photo-1644088379091-d574269d422f），用于整页视觉背景，不作为数据分析证据。";

// Slide 1: cover
{
  const slide = presentation.slides.add();
  addFullBleedBackground(
    slide,
    "tech",
    "#071A38/42",
    "Blue digital network background, Unsplash photo-1644088379091-d574269d422f",
  );
  addShape(slide, "rect", 0, 0, 12, 720, C.coral);
  addText(slide, "COURSE 03  /  HIGH-DIMENSIONAL DATA", 72, 64, 600, 24, {
    fontSize: 15,
    color: "#DDF3EB",
    bold: true,
  });
  addText(slide, "从读讲义到\n有边界的判断", 72, 150, 650, 172, {
    fontSize: 56,
    color: C.white,
    bold: true,
  });
  addText(slide, "高维表达数据、PCA 与 UMAP 的学习过程复盘", 76, 350, 630, 36, {
    fontSize: 21,
    color: "#E6F2ED",
  });
  addText(slide, "我做了什么 · 我怎样学习 · 我如何形成结论", 76, 405, 650, 28, {
    fontSize: 16,
    color: C.coral,
    bold: true,
  });
  addShape(slide, "roundRect", 72, 500, 518, 106, "#08243F/78", {
    borderRadius: "rounded-2xl",
    line: line("#8FB9AE", 1),
  });
  addText(slide, "120 位患者", 98, 520, 150, 30, { fontSize: 17, color: "#BDE8D9", bold: true });
  addText(slide, "23,146 个基因", 270, 520, 190, 30, { fontSize: 17, color: "#FFC0B4", bold: true });
  addText(slide, "p/n ≈ 192.9", 98, 563, 180, 26, { fontSize: 22, color: C.white, bold: true });
  addText(slide, "先问“能不能相信”，再问“图漂不漂亮”", 300, 563, 270, 26, {
    fontSize: 15,
    color: "#D7E7E2",
    bold: true,
  });
  addText(slide, "基于 plan.docx、课程 readings、worksheet 与 evidence pack", 72, 668, 830, 22, {
    fontSize: 12,
    color: "#C7D9D4",
  });
  addText(slide, "01", 1174, 668, 34, 22, { fontSize: 12, color: "#C7D9D4", alignment: "right" });
  notes[1] = `开场先交代这不是对 worksheet 的逐格朗读，而是对一次学习和分析过程的复盘。我这次真正做的事情可以概括成三步：先根据 plan.docx 确认课程强调的概念，再结合 readings 理解这些概念，最后把它们放回固定的高维表达数据里做验证。数据本身很容易让人产生“数字越多、图越复杂、模型越强”的直觉，所以我把本次作业的主问题定成：面对 120 位患者和 23,146 个基因，我怎样展示患者结构，同时不把训练高分或漂亮的降维图误当成可靠发现。右侧三个数字是整套推理的起点。120 是患者数量，23,146 是特征数量，p/n 约为 192.9，说明这是典型的小样本高维场景。后面所有判断，包括距离、PCA、UMAP 和泛化，都要回到这个约束上。视觉素材：Cornell CS4780 Lecture 2 的低维流形示意图，来自课程 readings，仅作背景视觉，未作为本页数据证据。`;
  addNotes(slide, notes[1]);
}

// Slide 2: learning route
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/72", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 2, "我先把学习任务拆成一条证据链", "老师要求的顺序，对我来说不是形式，而是避免“先看图再找解释”的方法");
  const y = 244;
  const boxes = [
    { x: 72, w: 270, n: "01", title: "课程大纲", body: "先确认问题属于哪一类：变量、维度、降维、泛化还是验证。", fill: C.mint, color: C.teal },
    { x: 380, w: 270, n: "02", title: "参考材料", body: "用讲义补概念：距离、标准化、PCA、过拟合与高维邻近。", fill: "#EFE6D6", color: C.gold },
    { x: 688, w: 270, n: "03", title: "数据实践", body: "再进入 120 × 23,146 的表达矩阵，逐个检验核心问题。", fill: "#F2DED7", color: C.coral },
  ];
  for (const b of boxes) {
    addShape(slide, "rect", b.x, y, b.w, 210, b.fill, {
      borderRadius: "rounded-2xl",
      line: line("#D9D4C9", 1),
    });
    addText(slide, b.n, b.x + 24, y + 24, 50, 24, { fontSize: 15, color: b.color, bold: true });
    addText(slide, b.title, b.x + 24, y + 64, b.w - 44, 34, { fontSize: 26, color: C.ink, bold: true });
    addText(slide, b.body, b.x + 24, y + 112, b.w - 44, 68, { fontSize: 17, color: C.slate });
  }
  addShape(slide, "line", 350, y + 104, 26, 0, line(C.teal, 2));
  addShape(slide, "line", 658, y + 104, 26, 0, line(C.teal, 2));
  addText(slide, "核心转换：把“我看到了什么”改写成“这项证据能支持什么，以及不能支持什么”", 72, 522, 1030, 42, {
    fontSize: 23,
    color: C.ink,
    bold: true,
  });
  addFooter(slide, "证据来源：plan.docx；worksheet 的背景阅读与探究问题地图");
  notes[2] = `这一页解释我怎样学习，而不是只罗列我读了哪些文件。plan.docx 把课程重点分成几条线：先理解变量和数据对象，再理解高维带来的统计困难，然后学习 PCA、UMAP、过拟合和验证策略。worksheet 的设计其实和这个顺序一致，所以我没有一开始就直接跑 UMAP，而是先问三个基础问题。第一，矩阵的一行、一列和一个值分别代表什么。第二，距离和相似性在高维里还是否可靠。第三，训练集上的好结果能不能迁移到没见过的患者。参考材料承担的是概念校准作用，例如 Tan 等人的数据与距离讲义帮助我区分欧氏距离、相关性、标准化和降维，Cornell 的材料让我理解有效维数，过拟合讲义提醒我不要把训练表现等同于泛化。等概念有了以后，我才把这些问题落到 B1 到 D1 的数据证据上。对我来说，这一步最大的变化是把作业从“填空”变成一条可追踪的论证链。视觉素材：Cornell CS4780 Lecture 2 的三维点云图，来自课程 readings，仅作背景视觉。`;
  addNotes(slide, notes[2]);
}

// Slide 3: concepts
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/74", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 3, "课程重点决定了我真正要回答的问题", "我没有把“降维”当成单独技巧，而是放回研究设计、统计可靠性和泛化里");
  const rows = [
    ["数据对象", "一行是患者，一列是基因，一个数是 log2(1+FPKM) 表达值", "先定义 n、p、标签和独立观察单位"],
    ["维度灾难", "p 远大于 n 时，距离、模型自由度和检验数量都会变得危险", "先看近邻是否稳定，再谈结构"],
    ["PCA", "线性投影，最大化输入方差，保留多少取决于预处理", "把方差解释和生物学解释分开"],
    ["UMAP", "非线性、偏局部的探索表示，对参数和种子要做敏感性核查", "可以提问，不能直接命名新亚型"],
    ["泛化", "见过的患者 100% 正确，不等于新患者也可靠", "把训练表现和未见表现分开报告"],
  ];
  let y = 220;
  rows.forEach((r, i) => {
    const accent = [C.teal, C.coral, C.gold, C.sea, C.red][i];
    addShape(slide, "rect", 72, y, 7, 72, accent);
    addText(slide, r[0], 102, y + 6, 180, 28, { fontSize: 21, color: C.ink, bold: true });
    addText(slide, r[1], 310, y + 7, 620, 28, { fontSize: 17, color: C.ink });
    addText(slide, r[2], 310, y + 38, 760, 24, { fontSize: 14, color: C.muted });
    addShape(slide, "line", 102, y + 72, 1054, 0, line("#D7D6CB", 1));
    y += 84;
  });
  addFooter(slide, "概念锚点：R1/R2/R3/R4/R5/R6/R7/R9");
  notes[3] = `这里我把 plan.docx 里的学习重点重新组织成五个实际问题。第一是数据对象。教材里对数据矩阵的定义看似基础，但如果连患者和特征都没定义清楚，后面距离和模型都没有解释对象。第二是维度灾难。高维并不只是“列很多”，它会影响邻近关系、模型自由度和可视化。第三是 PCA。PCA 能找到解释方差最多的线性方向，但这句话只说明数学目标，不保证这个方向就是疾病机制。第四是 UMAP。它对局部邻域很有用，但图中的簇距、密度和形状不能直接当作生物学结论。第五是泛化。课程大纲强调训练、验证、测试和外部验证，我在这次小作业里至少把“见过患者”和“未见患者”拆开看，避免只看训练高分。这样整理以后，每个模块都对应一个判断边界：数据对象决定我们在分析谁，维度灾难决定我们能不能相信距离，PCA 和 UMAP 决定我们能怎样展示，泛化决定我们能不能把观察写成预测结论。`;
  addNotes(slide, notes[3]);
}

// Slide 4: data object
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/72", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 4, "先把数据对象说清楚，很多误解会自动消失", "本案例的“高危”是风险分组，不是健康与患病分类");
  const table = slide.tables.add({
    rows: 5,
    columns: 4,
    left: 72,
    top: 214,
    width: 710,
    height: 300,
    columnWidths: [210, 150, 170, 180],
    values: [
      ["数据对象", "n", "p / 标签", "用途边界"],
      ["固定教学队列", "120 位患者", "23,146 基因", "探索 + 预测观察"],
      ["开发集", "80 位患者", "27 高危 / 53 非高危", "拟合尺度与模型"],
      ["未见患者组", "40 位患者", "14 高危 / 26 非高危", "只做固定预测观察"],
      ["表达单位", "一行=患者", "一列=基因", "log2(1+FPKM)"],
    ],
  });
  table.styleOptions = { headerRow: true, bandedRows: true };
  table.borders.assign({ style: "solid", fill: C.sand, width: 1 });
  for (let r = 0; r < 5; r++) {
    for (let c = 0; c < 4; c++) {
      const cell = table.getCell(r, c);
      cell.text.style = {
        typeface: family,
        fontSize: r === 0 ? 16 : 15,
        color: r === 0 ? C.white : C.ink,
        bold: r === 0 || c === 0,
      };
      if (r === 0) cell.fill = C.ink;
      else if (r % 2 === 0) cell.fill = "#F0ECE4";
      else cell.fill = C.paper;
    }
  }
  addShape(slide, "roundRect", 850, 225, 310, 250, C.mint, {
    borderRadius: "rounded-2xl",
    line: line("#B9D8CE", 1),
  });
  addText(slide, "p / n", 886, 260, 120, 30, { fontSize: 18, color: C.teal, bold: true });
  addText(slide, "192.9", 886, 300, 230, 72, { fontSize: 58, color: C.coral, bold: true });
  addText(slide, "特征数约为患者数的 193 倍", 886, 384, 235, 40, { fontSize: 18, color: C.ink, bold: true });
  addText(slide, "所以“变量更多”首先意味着\n需要更严格的验证，而不是更强的结论。", 886, 440, 225, 56, { fontSize: 16, color: C.slate });
  addFooter(slide, "证据：表 A1/A2/A3；数据来源：GEO GSE49711 TUC_G 基因层级文件");
  notes[4] = `这页是 Q1 的核心。数据矩阵一行代表一位患者的一份原发肿瘤样本，一列代表一个基因，一个数是该患者该基因的 log2(1+FPKM) 表达量。固定教学队列有 120 位患者、23,146 个基因，其中 80 位用于开发，40 位作为预先固定的未见患者组。高危标签是研究中的风险分组，1 和 0 只是标签编码，不等于健康和患病。这个区分很重要，因为它决定了我不能把后面的分类准确率说成诊断准确率，也不能把结果直接延伸到临床部署。右边的 p/n 约 192.9 是这份数据最重要的结构性事实。它告诉我：特征维度远高于患者数，模型有很大的记忆空间，距离也可能受到许多相关或低变异基因的影响。表 A3 还提醒我，公开样本 ID 唯一并不能独立证明没有重复患者或隐藏关联。因此我把“患者是否真正独立”列为需要向数据团队追问的来源问题，而不是默认它已经被解决。`;
  addNotes(slide, notes[4]);
}

// Slide 5: distance
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/76", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 5, "Q2 距离：维度越高，绝对距离变大，近远差别反而变小", "我用相对对比度而不是只看距离数值");
  const cats = ["20", "100", "500", "2,000", "5,000", "23,146"];
  const relative = [2.3028, 1.2315, 0.9268, 0.8651, 0.8577, 0.8544];
  const effective = [10.5784, 18.4987, 21.7142, 22.6082, 22.543, 22.5509];
  const chart1 = slide.charts.add("line", {
    position: { left: 72, top: 214, width: 520, height: 300 },
    categories: cats,
    series: [{ name: "相对对比度", values: relative, line: { style: "solid", fill: C.coral, width: 3 }, marker: { symbol: "circle", size: 7 } }],
    hasLegend: false,
    chartFill: C.paper,
    plotAreaFill: C.paper,
    plotAreaLine: line(C.sand, 1),
    xAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: null, title: "基因数" },
    yAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: line(C.grid, 1), min: 0.7, max: 2.5, majorUnit: 0.3, numberFormatCode: "0.0" },
  });
  applyPresentationChartFont(chart1, { fontFamily: family });
  const chart2 = slide.charts.add("line", {
    position: { left: 656, top: 214, width: 520, height: 300 },
    categories: cats,
    series: [{ name: "有效维数", values: effective, line: { style: "solid", fill: C.teal, width: 3 }, marker: { symbol: "circle", size: 7 } }],
    hasLegend: false,
    chartFill: C.paper,
    plotAreaFill: C.paper,
    plotAreaLine: line(C.sand, 1),
    xAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: null, title: "基因数" },
    yAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: line(C.grid, 1), min: 0, max: 25, majorUnit: 5, numberFormatCode: "0" },
  });
  applyPresentationChartFont(chart2, { fontFamily: family });
  addText(slide, "相对对比度", 92, 178, 260, 28, { fontSize: 18, color: C.coral, bold: true });
  addText(slide, "有效维数", 676, 178, 260, 28, { fontSize: 18, color: C.teal, bold: true });
  addText(slide, "20 基因 → 2.3028", 92, 548, 230, 28, { fontSize: 17, color: C.ink, bold: true });
  addText(slide, "23,146 基因 → 0.8544", 330, 548, 270, 28, { fontSize: 17, color: C.coral, bold: true });
  addText(slide, "有效维数在约 22.55 附近平台化，说明 23,146 列并不等于 23,146 个独立方向", 656, 548, 520, 35, { fontSize: 16, color: C.ink, bold: true });
  addFooter(slide, "证据：图 B1、表 B1、表 B2；相对对比度=(最远距离−最近距离)/最近距离");
  notes[5] = `Q2 里我重点学习了三个概念：欧氏距离、相对对比度和有效维数。平均距离随着基因数增加而增加，例如从 20 个基因的 5.9237 增长到全部 23,146 个基因的 204.3323。如果只看这个数字，很容易误以为患者之间变得更容易分开。但相对对比度的趋势相反，从 2.3028 降到 0.8544，说明最近邻和最远邻的相对差异变小，邻近关系更难区分。这正是高维空间稀疏化的一个直观表现。右边的有效维数提供了另一个重要线索：全部基因时约为 22.5509，远远小于列数。这说明基因之间存在相关结构，列数不等于独立信息量。表 B2 也支持这个判断，全部基因时三种距离处理得到的五近邻高度重合，说明在当前数据里存在一定稳定的患者相似性，但这不意味着任何尺度下的近邻都可靠。我的处理是把距离展示从一个绝对数值问题改成一个稳健性问题：标准化方法怎么变，基因子集怎么变，近邻还在不在。视觉素材：Cornell CS4780 Lecture 2 的高维距离分布图，来自课程 readings，仅作背景视觉。`;
  addNotes(slide, notes[5]);
}

// Slide 6: generalization
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/72", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 6, "Q3 见过不等于学会：训练高分必须和未见患者一起看", "我把“能认出开发集”与“能预测新样本”分成两个问题");
  const chart = slide.charts.add("bar", {
    position: { left: 72, top: 210, width: 735, height: 340 },
    categories: ["20 genes", "500 genes", "23,146 genes"],
    series: [
      { name: "见过患者", values: [0.912, 1.0, 1.0], fill: C.teal },
      { name: "未见患者", values: [0.75, 0.875, 0.85], fill: C.coral },
    ],
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 55 },
    hasLegend: true,
    legend: { position: "bottom", textStyle: { fill: C.muted, fontSize: 14 } },
    dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 13, bold: true } },
    chartFill: C.paper,
    plotAreaFill: C.paper,
    plotAreaLine: line(C.sand, 1),
    xAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: null },
    yAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: line(C.grid, 1), min: 0.6, max: 1.05, majorUnit: 0.1, numberFormatCode: "0%" },
  });
  applyPresentationChartFont(chart, { fontFamily: family });
  addShape(slide, "roundRect", 870, 220, 300, 205, C.sky, { borderRadius: "rounded-2xl", line: line("#B9D8E2", 1) });
  addText(slide, "100%", 900, 252, 190, 56, { fontSize: 42, color: C.teal, bold: true });
  addText(slide, "见过患者\n全部基因", 1040, 262, 92, 52, { fontSize: 15, color: C.ink, bold: true });
  addText(slide, "34 / 40", 900, 345, 150, 42, { fontSize: 30, color: C.coral, bold: true });
  addText(slide, "85% 未见患者\n同一排列", 1040, 348, 105, 48, { fontSize: 15, color: C.ink, bold: true });
  addText(slide, "差距不是失败，而是边界：它告诉我模型已经在开发集上饱和，但外推能力仍需要新患者验证。", 870, 486, 294, 84, { fontSize: 17, color: C.ink, bold: true });
  addFooter(slide, "证据：图 B2、表 B3/B4；固定 80/40 划分，40 人结果不用于回选参数");
  notes[6] = `Q3 是我对过拟合最直接的一次实践。表 B3 显示，500 个基因和全部基因时，见过患者都是 80/80，也就是 100%，但未见患者分别是 35/40 和 34/40。这个差距告诉我，训练高分只能说明模型可以解释已经看过的数据，不能自动说明它学到了可以迁移的规律。这里还要细看类别计数：全部基因时高危正确 12/14，非高危正确 22/26。总体 85% 不能替代对不同类别表现的检查，尤其在医学问题里，漏掉高危患者和漏掉非高危患者的代价可能不同。表 B4 进一步说明，20 基因到全部基因，未见患者均值从 81.6% 提升到 85.0%，方向上似乎存在信息，但这 20 次排列共享同一个患者划分，不能把范围看成独立人群的不确定性区间。按照课堂协议，40 名未见患者已经是预先固定的观察对象，所以我不能看到 85% 后再反过来选择基因数或调参。更合理的下一步是把所有预处理放进训练折内，用嵌套交叉验证和真正外部队列来检验泛化。`;
  addNotes(slide, notes[6]);
}

// Slide 7: PCA
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/76", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 7, "Q4 PCA：压缩的是方差，不是生物学真相", "我用 loading、score、解释方差和重建误差分别回答不同问题");
  const chart = slide.charts.add("bar", {
    position: { left: 72, top: 225, width: 440, height: 260 },
    categories: ["PC1", "PC2"],
    series: [{ name: "解释方差比例", values: [0.1449, 0.0826], fill: C.teal }],
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 70 },
    hasLegend: false,
    dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 15, bold: true } },
    chartFill: C.paper,
    plotAreaFill: C.paper,
    plotAreaLine: line(C.sand, 1),
    xAxis: { textStyle: { fill: C.muted, fontSize: 14 }, line: line(C.sand, 1), majorGridlines: null },
    yAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: line(C.grid, 1), min: 0, max: 0.2, majorUnit: 0.05, numberFormatCode: "0%" },
  });
  applyPresentationChartFont(chart, { fontFamily: family });
  addShape(slide, "roundRect", 575, 220, 250, 265, C.mint, { borderRadius: "rounded-2xl", line: line("#B9D8CE", 1) });
  addText(slide, "22.76%", 603, 250, 220, 76, { fontSize: 46, color: C.teal, bold: true });
  addText(slide, "标准化 PCA 的二维累计方差", 615, 348, 180, 42, { fontSize: 16, color: C.ink, bold: true });
  addText(slide, "达到 90% 需要 51 个成分", 615, 418, 190, 30, { fontSize: 16, color: C.slate });
  addShape(slide, "line", 875, 228, 0, 270, line(C.sand, 1));
  addText(slide, "loading", 910, 236, 160, 26, { fontSize: 18, color: C.coral, bold: true });
  addText(slide, "方向上的权重\nMKNK1  0.0165\nPAICS  −0.0162", 910, 276, 250, 104, { fontSize: 19, color: C.ink });
  addText(slide, "score", 910, 410, 160, 26, { fontSize: 18, color: C.teal, bold: true });
  addText(slide, "某个患者投影到这条方向上的坐标\nPC1 = −101.5360", 910, 450, 260, 60, { fontSize: 17, color: C.ink });
  addText(slide, "结论：PC1 先回答“数据变化最多的方向是什么”，不能直接回答“疾病机制是什么”。", 72, 560, 1060, 36, { fontSize: 20, color: C.ink, bold: true });
  addFooter(slide, "证据：表 C1/C2；R5/R5中/R5中B 的 PCA 阅读材料");
  notes[7] = `PCA 是这次作业里最容易被过度解释的部分。它的数学目标是找一个单位长度的方向，使样本投影后的方差最大。这个方向由所有基因的 loading 共同决定，不等于某个基因的效应，也不等于一个已经证实的生物通路。表 C1 里 MKNK1 的权重是 0.0165，PAICS 是 −0.0162。对查询患者 SEQC_NB003 来说，按基因标准化后的前六项贡献和只有 −0.1531，而其余全部基因贡献和是 −101.3829，最后 PC1 得分是 −101.5360。这说明样本得分是许多基因贡献的投影结果，不应该把某两个 loading 直接翻译成因果关系。另一方面，标准化 PCA 的 PC1 和 PC2 解释方差比例分别是 14.49% 和 8.26%，二维累计只有 22.76%，达到 90% 需要 51 个成分。因此 PC1-PC2 图可以帮助观察总体结构，但必须明确它丢掉了大部分输入变化。还要注意预处理的影响：仅中心化和按基因标准化会改变解释方差和重建误差。PCA 优化的是输入方差，不是“信号纯度”；低方差方向也可能包含少数但重要的生物学信号。`;
  addNotes(slide, notes[7]);
}

// Slide 8: UMAP
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/72", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 8, "Q5 UMAP：我把它当作探索图，而不是结论机器", "我核查的是局部近邻、参数敏感性和批次竞争解释");
  const chart = slide.charts.add("bar", {
    position: { left: 72, top: 214, width: 650, height: 330 },
    categories: ["原标准化空间", "PCA 2D", "UMAP 基准", "仅改种子", "仅改近邻数", "仅改 min_dist"],
    series: [{ name: "全体近邻保留率", values: [1.0, 0.2575, 0.4025, 0.405, 0.415, 0.3975], fill: C.teal }],
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 45 },
    hasLegend: false,
    dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 13, bold: true } },
    chartFill: C.paper,
    plotAreaFill: C.paper,
    plotAreaLine: line(C.sand, 1),
    xAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: line(C.grid, 1), min: 0, max: 1.05, majorUnit: 0.2, numberFormatCode: "0%" },
    yAxis: { textStyle: { fill: C.muted, fontSize: 13 }, line: line(C.sand, 1), majorGridlines: null },
  });
  applyPresentationChartFont(chart, { fontFamily: family });
  addShape(slide, "roundRect", 790, 230, 380, 142, "#F2DED7", { borderRadius: "rounded-2xl", line: line("#E4C4BC", 1) });
  addText(slide, "40.25%", 820, 252, 235, 62, { fontSize: 44, color: C.coral, bold: true });
  addText(slide, "UMAP 基准的全体近邻保留率", 830, 330, 290, 28, { fontSize: 16, color: C.ink, bold: true });
  addText(slide, "可以报告的观察", 790, 420, 210, 26, { fontSize: 18, color: C.teal, bold: true });
  addText(slide, "在固定输入、种子和参数下，UMAP 提供了一个局部结构的探索视角。", 790, 456, 330, 50, { fontSize: 17, color: C.ink });
  addText(slide, "不能成立的推断", 790, 530, 210, 26, { fontSize: 18, color: C.coral, bold: true });
  addText(slide, "图上的簇、空隙或高危着色不能直接证明新亚型或批次因果。", 790, 566, 340, 52, { fontSize: 17, color: C.ink });
  addFooter(slide, "证据：图 C2/C3、表 C3/C4；R7/R8/R9 的降维解释边界");
  notes[8] = `我没有因为 UMAP 图看起来更有结构，就把它当作主要证据。表 C3 里，查询患者在原标准化空间的五近邻完全保留，而二维 PCA 的全体近邻保留率只有 25.75%，UMAP 基准为 40.25%。这说明降维会改变邻域关系，也说明 PCA 二维和 UMAP 都是对原空间的有损表示。UMAP 的价值是帮助我提出结构问题，而不是直接回答生物学问题。我还看了单因素对照：只改种子时保留率 40.50%，只改近邻数时 41.50%，只改最小间距时 39.75%，这些数和基准接近，但查询患者的具体近邻仍然会变化。因此我可以说，在本次输入和参数范围内，局部探索结果没有完全崩掉，但不能把这个小范围稳定性扩大成外部可靠性。表 C4 还提醒我，来源列名里可以解析出 12 种测序运行代码，但完整的建库日期、处理批次、中心和标本前处理对应关系并不完整。某些运行代码的高危/非高危比例差异明显，所以图中的分离既可能包含生物信号，也可能包含技术批次。需要补充完整批次信息、独立标志物和外部患者复现，才能把探索观察升级为生物学结论。视觉素材：Cornell CS4780 Lecture 2 的低维流形图，来自课程 readings，用来提示“低维结构只是类比”，不替代本次 UMAP 结果。`;
  addNotes(slide, notes[8]);
}

// Slide 9: decision
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/70", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 9, "我的最终判断：A 做主展示，PCA 与 UMAP 各司其职", "展示、压缩、探索和建模需要不同的证据标准");
  addText(slide, "A", 72, 215, 160, 124, { fontSize: 116, color: C.teal, bold: true });
  addText(slide, "主要展示", 88, 342, 170, 28, { fontSize: 18, color: C.ink, bold: true });
  addText(slide, "标准化原始特征的距离、相关性和近邻稳定性。先把可核查的结构放在主位。", 88, 390, 270, 100, { fontSize: 18, color: C.slate });
  addShape(slide, "line", 380, 220, 0, 350, line(C.sand, 1));
  addText(slide, "PCA", 435, 235, 160, 42, { fontSize: 28, color: C.gold, bold: true });
  addText(slide, "压缩与解释", 435, 280, 250, 28, { fontSize: 17, color: C.ink, bold: true });
  addText(slide, "适合展示线性方向、解释方差和重建损失。二维只解释 22.76%，所以必须把压缩损失写出来。", 435, 330, 270, 112, { fontSize: 17, color: C.slate });
  addText(slide, "UMAP", 780, 235, 160, 42, { fontSize: 28, color: C.coral, bold: true });
  addText(slide, "结构探索", 780, 280, 250, 28, { fontSize: 17, color: C.ink, bold: true });
  addText(slide, "适合提出局部结构问题。基准近邻保留率 40.25%，不能直接命名亚型或替代验证。", 780, 330, 300, 112, { fontSize: 17, color: C.slate });
  addShape(slide, "roundRect", 435, 500, 645, 88, C.lilac, { borderRadius: "rounded-2xl", line: line("#C8C6E5", 1) });
  addText(slide, "停止条件：近邻对尺度/参数高度敏感，或批次与标签混杂无法排除", 465, 526, 580, 28, { fontSize: 18, color: C.ink, bold: true });
  addFooter(slide, "综合证据：表 B1/B2/B3/B4、表 C2/C3/C4、表 D1");
  notes[9] = `综合 Q1 到 Q6 后，我最后没有在 A、B、C、D 之间简单地选一个“最好看的方法”。我的选择是 A 作为主要展示方式，因为它最接近原始数据对象，能够直接报告标准化距离、相关性和近邻，同时把尺度、基因子集和稳定性问题暴露出来。PCA 作为辅助展示，原因是它有明确的线性方向和解释方差，但二维只保留 22.76% 的标准化输入变化，不能把图当成无损空间。UMAP 作为探索图，帮助我看局部结构和提出假设，但基准全体近邻保留率只有 40.25%，所以它不适合承担“证明亚型”的任务。后续建模建议与主要展示方式分开：如果要预测风险，需要在每个训练折内拟合标准化、特征选择或 PCA，并用正则化、嵌套交叉验证和外部队列评估；不能因为固定未见组有 85% 就声称临床有效。我的停止条件也很具体：如果换距离尺度或基因子集后近邻大幅改变，或者 UMAP 对参数非常敏感，或者批次和高危标签的关系无法排除，我就停止做生物学解释。恢复判断需要独立技术批次、临床协变量、分子标志物和新患者验证。视觉素材：Cornell CS4780 Lecture 2 的三维点云图，来自课程 readings，仅作背景视觉。`;
  addNotes(slide, notes[9]);
}

// Slide 10: AI role
{
  const slide = presentation.slides.add();
  addFullBleedBackground(slide, "tech", "#EAF2FA/74", "Blue digital network background, Unsplash photo-1644088379091-d574269d422f");
  addHeader(slide, 10, "AI 的位置：加速理解和检查，不替代研究判断", "我把 AI 当成辅助工具，最终数字、来源和边界都回到材料与数据");
  addShape(slide, "roundRect", 72, 220, 470, 290, C.mint, { borderRadius: "rounded-2xl", line: line("#B9D8CE", 1) });
  addText(slide, "AI 帮我加速的部分", 110, 254, 300, 34, { fontSize: 23, color: C.teal, bold: true });
  addText(slide, "• 把 plan.docx 的重点整理成学习路径\n• 解释 PCA、UMAP、过拟合等概念\n• 帮我检查代码和数字是否能对上\n• 帮我把分析结果压缩成汇报语言", 110, 312, 360, 130, { fontSize: 18, color: C.ink });
  addText(slide, "工具价值：减少查找和改写成本", 110, 470, 350, 26, { fontSize: 16, color: C.slate, italic: true });
  addShape(slide, "line", 590, 365, 100, 0, line(C.teal, 2));
  addShape(slide, "rightArrow", 690, 340, 60, 50, C.teal, { line: line(C.teal, 0) });
  addShape(slide, "roundRect", 790, 220, 390, 290, C.peach, { borderRadius: "rounded-2xl", line: line("#E7BDB5", 1) });
  addText(slide, "我必须自己完成的部分", 828, 254, 320, 34, { fontSize: 23, color: C.coral, bold: true });
  addText(slide, "• 回到原始表格核对 22.76%、40.25%、85.0%\n• 区分训练、未见患者和外部验证\n• 不把 loading 当因果，不把簇当亚型\n• 明确什么时候应停止解释", 828, 312, 310, 132, { fontSize: 18, color: C.ink });
  addText(slide, "研究判断：数字和边界必须能被我解释", 828, 470, 310, 26, { fontSize: 16, color: C.slate, italic: true });
  addText(slide, "我最后提交的不是“AI 生成的答案”，而是一套能回到证据、能说明限制、也能被追问的判断。", 72, 572, 1080, 38, { fontSize: 21, color: C.ink, bold: true });
  addFooter(slide, "AI 使用原则来自 plan.docx；数字核对回到 worksheet/evidence pack 与本地数据副本");
  notes[10] = `这一页是我想主动说明的 AI 使用边界。AI 对我有帮助的地方主要是加速：把课程大纲整理成学习路线，帮助我用更容易理解的语言解释距离、PCA、UMAP 和过拟合，也帮助我检查答案是否覆盖了题目要求。但 AI 不能替我决定一张图是否可靠。真正关键的地方，我都回到了材料和证据：比如 22.76% 来自表 C2 的标准化 PCA 二维累计方差，40.25% 来自表 C3 的 UMAP 基准近邻保留率，85.0% 来自固定 23,146 基因排列下未见患者的 34/40。还要自己区分三个经常混在一起的概念：见过患者表现、未见患者表现和外部验证；loading 和样本 score；重建误差和去噪效果。AI 可以帮我把这些概念说得更顺，但如果我没有亲自核对，表达得再漂亮也不能成为证据。我的底线是：所有数字要能回到表格，所有来源要能找到原文，所有结论都要带着适用条件。这样使用 AI，工具是加速器，但判断仍然是我的。`;
  addNotes(slide, notes[10]);
}

// Slide 11: conclusion
{
  const slide = presentation.slides.add();
  addFullBleedBackground(
    slide,
    "tech",
    "#071A38/44",
    "Blue digital network background, Unsplash photo-1644088379091-d574269d422f",
  );
  addShape(slide, "rect", 0, 0, 12, 720, C.teal);
  addText(slide, "11  /  学会的不是“选一个算法”", 72, 62, 580, 24, { fontSize: 15, color: "#BDE8D9", bold: true });
  addText(slide, "我学会的是：\n先定义问题，再相信结果。", 72, 132, 760, 130, { fontSize: 48, color: C.white, bold: true });
  const takeaways = [
    ["01", "先定义数据对象", "120 位患者和 23,146 个基因，决定了我必须认真处理 p≫n。"],
    ["02", "再拆开证据用途", "原始空间、PCA 和 UMAP 分别服务于核查、压缩和探索。"],
    ["03", "最后给出边界", "见过患者的高分不能替代外部验证，漂亮图也不能替代独立证据。"],
  ];
  let x = 72;
  takeaways.forEach((t, i) => {
    addText(slide, t[0], x, 360, 64, 24, { fontSize: 16, color: i === 1 ? "#FFC0B4" : "#BDE8D9", bold: true });
    addText(slide, t[1], x, 400, 300, 34, { fontSize: 23, color: C.white, bold: true });
    addText(slide, t[2], x, 448, 300, 72, { fontSize: 17, color: "#D7E7E2" });
    if (i < 2) addShape(slide, "line", x + 330, 370, 0, 172, line("#8FB9AE", 1));
    x += 390;
  });
  addText(slide, "最终建议：A 为主展示，PCA/UMAP 做有边界的辅助；模型结果等待新患者和新中心验证。", 72, 632, 1080, 28, { fontSize: 18, color: "#BDE8D9", bold: true });
  addText(slide, "11", 1158, 668, 52, 22, { fontSize: 12, color: "#C7D9D4", alignment: "right" });
  notes[11] = `最后我想把这次作业的收获压缩成三个层次。第一，我学会先定义数据对象。120 位患者、23,146 个基因和 p/n 约 192.9，不是背景数字，而是决定分析策略的约束。第二，我学会拆开不同用途的证据。原始特征适合做距离和近邻核查，PCA 适合做线性压缩和解释方差，UMAP 适合做局部结构探索，它们没有一张图能同时承担全部任务。第三，我学会给结论加边界。见过患者 100% 正确并不能替代未见患者和外部验证，UMAP 上的簇也不能自动变成新亚型。对我来说，最重要的改变不是学会了一个新算法，而是知道什么时候不应该过度解释。我的最终建议因此很保守但也更可执行：用 A，也就是标准化原始特征的距离、相关性和近邻稳定性做主要展示，PCA 和 UMAP 分别承担压缩和探索的辅助角色；如果要进入预测建模，就把预处理放进训练折内，并把新患者、新中心、批次和校准列为恢复可信度的必要证据。背景图：Actin Filaments in a Neuroblastoma cell Line，Howard Vindin，CC BY 4.0，Wikimedia Commons。`;
  addNotes(slide, notes[11]);
}

for (const slideNumber of Object.keys(notes)) {
  const current = notes[slideNumber];
  notes[slideNumber] = current.includes("视觉素材：") || current.includes("背景图：")
    ? current.replace(/(?:视觉素材|背景图)：[\s\S]*$/, TECH_BG_NOTE)
    : `${current}\n${TECH_BG_NOTE}`;
}

await fs.writeFile(NOTES_JSON, JSON.stringify(notes, null, 2), "utf8");
const artifact = await PresentationFile.exportPptx(presentation);
await artifact.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, notes: NOTES_JSON, slides: presentation.slides.items.length }));
