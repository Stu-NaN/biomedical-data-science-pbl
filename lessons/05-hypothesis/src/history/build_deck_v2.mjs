import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";
import JSZip from "jszip";

const ROOT = path.resolve(".");
const SKILL_DIR =
  "C:\\Users\\A\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1007.11041\\skills\\presentations";
const NODE_MODULES =
  "C:\\Users\\A\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules";
const PYTHON =
  "C:\\Users\\A\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const FONT = "Microsoft YaHei";
const MONO = "Consolas";
const WIDTH = 1280;
const HEIGHT = 720;
const FINAL_PPTX = path.join(
  ROOT,
  "deliverables",
  "统计检验与假设设计_演讲PPT_v2.pptx",
);
const PREVIEW_DIR = path.join(ROOT, "previews-v2");
const CONTENT_PATH = path.join(ROOT, "presenter-content-v2.json");
const DATA_ZIP = "C:\\Users\\A\\Downloads\\hypothesis.zip";
const COVER_IMAGE = path.join(ROOT, "assets-medical", "ecg-monitor.jpg");
const CONTENT = JSON.parse(await fs.readFile(CONTENT_PATH, "utf8"));

const COLORS = {
  ink: "#0E1620",
  red: "#C0392B",
  teal: "#0F6E7F",
  gray: "#8A949E",
  muted: "#59636C",
  rule: "#D9DEE2",
  paper: "#F6F7F9",
  white: "#FFFFFF",
  pale: "#ECEFF2",
  darkMuted: "#C7CFD6",
};

const presentation = Presentation.create({
  slideSize: { width: WIDTH, height: HEIGHT },
});

function rect(slide, name, x, y, w, h, fill = "none", line = "none", lineWidth = 0) {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: line, width: lineWidth },
  });
}

function rule(slide, name, x1, y1, x2, y2, color = COLORS.rule, width = 1, style = "solid") {
  const deltaX = x2 - x1;
  const deltaY = y2 - y1;
  return slide.shapes.add({
    geometry: "line",
    name,
    position: {
      left: Math.min(x1, x2),
      top: Math.min(y1, y2),
      width: Math.abs(deltaX),
      height: Math.abs(deltaY),
      verticalFlip: deltaX * deltaY < 0,
    },
    fill: "none",
    line: { style, fill: color, width },
  });
}

function hatchedRect(slide, name, x, y, width, height) {
  rect(slide, `${name}-fill`, x, y, width, height, "#EEF0F2");
  const step = 22;
  let index = 0;
  for (let offset = -height; offset <= width; offset += step) {
    const points = [
      [0, -offset],
      [width, width - offset],
      [offset, 0],
      [height + offset, height],
    ].filter(([px, py]) => px >= 0 && px <= width && py >= 0 && py <= height);
    const unique = points.filter(
      (point, pointIndex) =>
        points.findIndex(
          ([px, py]) =>
            Math.abs(px - point[0]) < 0.01 && Math.abs(py - point[1]) < 0.01,
        ) === pointIndex,
    );
    if (unique.length >= 2) {
      rule(
        slide,
        `${name}-hatch-${index}`,
        x + unique[0][0],
        y + unique[0][1],
        x + unique[1][0],
        y + unique[1][1],
        "#DFE3E6",
        1,
      );
      index += 1;
    }
  }
}

function text(
  slide,
  name,
  value,
  x,
  y,
  w,
  h,
  {
    size = 22,
    color = COLORS.ink,
    bold = false,
    font = FONT,
    align = "left",
    valign = "top",
    insets = 0,
  } = {},
) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: font,
    fontSize: size,
    color,
    bold,
    alignment: align,
    verticalAlignment: valign,
    autoFit: "none",
    wrap: "square",
    insets,
  };
  return shape;
}

function addBase(slide, title, index) {
  slide.background.fill = COLORS.paper;
  text(slide, "slide-title", title, 72, 45, 1080, 58, {
    size: 38,
    bold: true,
  });
  rule(slide, "title-rule", 72, 120, 1208, 120, COLORS.rule, 1);
  text(slide, "slide-index", `${String(index).padStart(2, "0")} / 13`, 1112, 667, 96, 22, {
    size: 15,
    color: COLORS.gray,
    font: MONO,
    align: "right",
  });
}

function addNotes(slide, item) {
  const sections = [
    `本页要讲清楚：${item.goal}`,
    `一句话记住：${item.key}`,
    `口语稿：\n${item.talk}`,
    `分析步骤：\n${item.process.map((line, index) => `${index + 1}. ${line}`).join("\n")}`,
    `代码或计算：\n${item.code}`,
    `代码解释：${item.codeExplanation}`,
    `互动提问：${item.ask}`,
    `接话提示：${item.guide}`,
    `过渡句：${item.transition}`,
    `容易说错的地方：${item.pitfall}`,
    "资料来源：用户提供的 hypothesis.zip/data/heart-disease-hypothesis.csv、data/README.md、teacher-reference.md、review/numeric_answers.md；代码实现依据 build_docs.py。统计结果为未调整的观察性完整病例比较。",
    ...(item.index === 1 || item.index === 13 ? [CONTENT.photoCredit] : []),
  ];
  slide.speakerNotes.text = sections.join("\n\n");
  slide.speakerNotes.setVisible(true);
}

async function addCover(slide, item) {
  slide.background.fill = COLORS.ink;
  const image = new Uint8Array(await fs.readFile(COVER_IMAGE));
  slide.images.add({
    blob: image,
    contentType: "image/jpeg",
    alt: "暗场下显示心电波形的临床监护仪，作为封面氛围图",
    fit: "cover",
    position: { left: 758, top: 0, width: 522, height: HEIGHT },
  });
  rect(slide, "cover-image-shade", 758, 0, 522, HEIGHT, "#0E1620/38");
  text(slide, "cover-kicker", "PBL  /  统计实践", 84, 105, 520, 30, {
    size: 18,
    color: COLORS.darkMuted,
    bold: true,
  });
  text(slide, "cover-title", item.title, 80, 192, 700, 82, {
    size: 58,
    color: COLORS.white,
    bold: true,
  });
  text(
    slide,
    "cover-subtitle",
    "两组均值差了约 20 次/分，是真实差距，还是抽样巧合？",
    84,
    305,
    640,
    96,
    { size: 30, color: "#E6EAED" },
  );
  rect(slide, "group-red-mark", 84, 453, 94, 3, COLORS.red);
  rect(slide, "group-teal-mark", 178, 453, 94, 3, COLORS.teal);
  text(slide, "cover-context", "四中心教学数据   /   原始 920 人   /   完整病例 865 人", 84, 496, 630, 42, {
    size: 19,
    color: COLORS.darkMuted,
  });
}

function addQuestion(slide) {
  addBase(slide, "今天只回答一个问题", 2);
  text(slide, "delta-symbol", "Δ", 100, 205, 90, 100, {
    size: 72,
    bold: true,
    font: MONO,
  });
  text(slide, "mu-one", "μ₁", 222, 210, 108, 78, {
    size: 54,
    color: COLORS.red,
    bold: true,
    font: MONO,
  });
  text(slide, "minus-symbol", "−", 337, 220, 70, 62, {
    size: 46,
    color: COLORS.gray,
    font: MONO,
    align: "center",
  });
  text(slide, "mu-zero", "μ₀", 417, 210, 108, 78, {
    size: 54,
    color: COLORS.teal,
    bold: true,
    font: MONO,
  });
  text(slide, "equals-question", "是否等于 0？", 556, 222, 490, 64, {
    size: 42,
    bold: true,
  });
  rule(slide, "equation-frame-top", 205, 198, 538, 198, COLORS.rule, 2);
  rule(slide, "equation-frame-bottom", 205, 292, 538, 292, COLORS.rule, 2);
  text(slide, "term-one", "有疾病证据组平均值", 181, 313, 190, 30, {
    size: 18,
    color: COLORS.red,
    align: "center",
  });
  text(slide, "term-zero", "无疾病证据组平均值", 376, 350, 190, 30, {
    size: 18,
    color: COLORS.teal,
    align: "center",
  });
  const facts = [
    ["主要终点", "thalach，最大运动心率（次/分）"],
    ["分析单位", "每位受试者，每行一人"],
    ["终点口径", "thalach 完整的受试者"],
  ];
  facts.forEach(([label, value], i) => {
    const y = 438 + i * 46;
    text(slide, `fact-label-${i}`, label, 104, y, 180, 28, {
      size: 19,
      color: COLORS.gray,
    });
    text(slide, `fact-value-${i}`, value, 300, y, 770, 30, {
      size: 20,
    });
  });
  text(slide, "study-design", "观察性资料，不是随机分组试验", 812, 580, 350, 28, {
    size: 18,
    color: COLORS.muted,
    align: "right",
  });
}

function addGroups(slide) {
  addBase(slide, "参赛的两组是谁", 3);
  const top = 181;
  const left = 88;
  const right = 666;
  rule(slide, "group-divider", 640, top, 640, 563, COLORS.rule, 2);
  rect(slide, "group-red-topline", left, top, 478, 8, COLORS.red);
  rect(slide, "group-teal-topline", right, top, 478, 8, COLORS.teal);
  text(slide, "group-one-code", "disease = 1", left, top + 42, 478, 48, {
    size: 34,
    color: COLORS.red,
    bold: true,
    font: MONO,
  });
  text(slide, "group-one-name", "有疾病证据组", left, top + 101, 478, 40, {
    size: 27,
    bold: true,
  });
  text(slide, "group-one-definition", "原始疾病等级 num > 0", left, top + 156, 478, 34, {
    size: 21,
    color: COLORS.muted,
  });
  text(slide, "group-one-raw", "原始人数", left, top + 230, 190, 30, {
    size: 18,
    color: COLORS.gray,
  });
  text(slide, "group-one-raw-n", "509", left + 196, top + 218, 230, 46, {
    size: 32,
    font: MONO,
    bold: true,
  });
  text(slide, "group-one-complete", "完整病例", left, top + 288, 190, 30, {
    size: 18,
    color: COLORS.gray,
  });
  text(slide, "group-one-complete-n", "474", left + 196, top + 276, 230, 46, {
    size: 32,
    font: MONO,
    bold: true,
  });
  text(slide, "group-zero-code", "disease = 0", right, top + 42, 478, 48, {
    size: 34,
    color: COLORS.teal,
    bold: true,
    font: MONO,
  });
  text(slide, "group-zero-name", "无疾病证据组", right, top + 101, 478, 40, {
    size: 27,
    bold: true,
  });
  text(slide, "group-zero-definition", "原始疾病等级 num = 0", right, top + 156, 478, 34, {
    size: 21,
    color: COLORS.muted,
  });
  text(slide, "group-zero-raw", "原始人数", right, top + 230, 190, 30, {
    size: 18,
    color: COLORS.gray,
  });
  text(slide, "group-zero-raw-n", "411", right + 196, top + 218, 230, 46, {
    size: 32,
    font: MONO,
    bold: true,
  });
  text(slide, "group-zero-complete", "完整病例", right, top + 288, 190, 30, {
    size: 18,
    color: COLORS.gray,
  });
  text(slide, "group-zero-complete-n", "391", right + 196, top + 276, 230, 46, {
    size: 32,
    font: MONO,
    bold: true,
  });
  rule(slide, "group-bottom-rule", 88, 568, 1160, 568);
  text(slide, "group-caveat", "编码只说明数据定义，不等于完整的临床健康状态。", 90, 594, 1050, 34, {
    size: 21,
    color: COLORS.muted,
  });
}

function addHypotheses(slide) {
  addBase(slide, "规则先定好，再看数据", 4);
  const y = 221;
  const xs = [90, 466, 842];
  const widths = [348, 348, 348];
  rule(slide, "hypothesis-column-1", 452, 196, 452, 448, COLORS.rule, 1);
  rule(slide, "hypothesis-column-2", 828, 196, 828, 448, COLORS.rule, 1);
  text(slide, "h0-label", "H0", xs[0], y, widths[0], 38, {
    size: 26,
    color: COLORS.ink,
    bold: true,
    font: MONO,
  });
  text(slide, "h0-formula", "Δ = 0", xs[0], y + 52, widths[0], 68, {
    size: 48,
    bold: true,
    font: MONO,
  });
  text(slide, "h0-detail", "两组总体均值相同", xs[0], y + 133, widths[0], 36, {
    size: 20,
    color: COLORS.muted,
  });
  text(slide, "h1-label", "H1", xs[1], y, widths[1], 38, {
    size: 26,
    color: COLORS.red,
    bold: true,
    font: MONO,
  });
  text(slide, "h1-formula", "Δ ≠ 0", xs[1], y + 52, widths[1], 68, {
    size: 48,
    color: COLORS.red,
    bold: true,
    font: MONO,
  });
  text(slide, "h1-detail", "两组总体均值不同", xs[1], y + 133, widths[1], 36, {
    size: 20,
    color: COLORS.muted,
  });
  text(slide, "alpha-label", "α", xs[2], y, widths[2], 38, {
    size: 26,
    color: COLORS.gray,
    bold: true,
    font: MONO,
  });
  text(slide, "alpha-value", "0.05", xs[2], y + 52, widths[2], 68, {
    size: 48,
    bold: true,
    font: MONO,
  });
  text(slide, "alpha-detail", "查看结果前预先设定", xs[2], y + 133, widths[2], 36, {
    size: 20,
    color: COLORS.muted,
  });
  rule(slide, "alpha-reference", 90, 494, 1164, 494, COLORS.gray, 2, "dashed");
  text(slide, "two-sided-reason", "题目问“是否不同”，方向未预先指定，所以采用双侧检验。", 90, 525, 1060, 35, {
    size: 21,
  });
  text(slide, "prediction-prompt", "在看数据之前，你猜哪一组更低？", 90, 581, 1060, 35, {
    size: 22,
    bold: true,
  });
}

function addFunnel(slide) {
  addBase(slide, "数据从哪来：920 → 865", 5);
  const barY1 = 271;
  const barY2 = 472;
  const barX = 110;
  const barW = 1030;
  const rawRed = barW * (509 / 920);
  const rawTeal = barW - rawRed;
  const analysisW = barW * (865 / 920);
  const finalRed = barW * (474 / 920);
  const finalTeal = barW * (391 / 920);
  const excludedW = barW - analysisW;
  text(slide, "raw-total", "920 人原始样本", 110, 177, 470, 42, {
    size: 30,
    bold: true,
    font: MONO,
  });
  rect(slide, "raw-disease", barX, barY1, rawRed, 66, COLORS.red);
  rect(slide, "raw-no-evidence", barX + rawRed, barY1, rawTeal, 66, COLORS.teal);
  text(slide, "raw-disease-label", "509", barX + 18, barY1 + 12, rawRed - 30, 40, {
    size: 25,
    color: COLORS.white,
    bold: true,
    font: MONO,
  });
  text(slide, "raw-no-evidence-label", "411", barX + rawRed + 18, barY1 + 12, rawTeal - 30, 40, {
    size: 25,
    color: COLORS.white,
    bold: true,
    font: MONO,
  });
  text(slide, "missing-total", "thalach 缺失 55 人", 392, 359, 480, 34, {
    size: 22,
    color: COLORS.gray,
    bold: true,
    align: "center",
  });
  text(slide, "missing-split", "有疾病证据组 35  人        无疾病证据组 20", 235, 399, 790, 30, {
    size: 18,
    color: COLORS.muted,
    font: MONO,
    align: "center",
  });
  text(slide, "complete-total", "865 人进入分析", 110, 429, 470, 38, {
    size: 28,
    bold: true,
    font: MONO,
  });
  rect(slide, "complete-disease", barX, barY2, finalRed, 66, COLORS.red);
  rect(slide, "complete-no-evidence", barX + finalRed, barY2, finalTeal, 66, COLORS.teal);
  hatchedRect(slide, "excluded-missing", barX + analysisW, barY2, excludedW, 66);
  text(slide, "complete-disease-label", "474", barX + 18, barY2 + 12, finalRed - 30, 40, {
    size: 25,
    color: COLORS.white,
    bold: true,
    font: MONO,
  });
  text(slide, "complete-no-evidence-label", "391", barX + finalRed + 18, barY2 + 12, finalTeal - 30, 40, {
    size: 25,
    color: COLORS.white,
    bold: true,
    font: MONO,
  });
  text(
    slide,
    "complete-excluded-label",
    "55",
    barX + analysisW,
    barY2 + 19,
    excludedW,
    28,
    { size: 15, color: COLORS.ink, bold: true, font: MONO, align: "center" },
  );
  text(
    slide,
    "excluded-explanation",
    "末端斜纹 = 55 人未进入分析",
    747,
    548,
    390,
    28,
    { size: 16, color: COLORS.gray, align: "right" },
  );
  text(slide, "funnel-legend", "朱红  有疾病证据组   深青  无疾病证据组", 110, 553, 610, 30, {
    size: 16,
    color: COLORS.muted,
  });
  text(slide, "funnel-note", "完整病例分析，没有插补", 110, 602, 470, 28, {
    size: 19,
    bold: true,
  });
  text(slide, "missing-warning", "两组缺失人数不同，可能带来选择偏倚", 612, 602, 548, 28, {
    size: 18,
    color: COLORS.gray,
    align: "right",
  });
}

function readCsvRows() {
  const py = [
    "import csv, io, json, zipfile",
    `with zipfile.ZipFile(r"${DATA_ZIP}") as z:`,
    "    raw = z.read('data/heart-disease-hypothesis.csv').decode('utf-8-sig')",
    "rows = list(csv.DictReader(io.StringIO(raw)))",
    "print(json.dumps([{'disease': int(float(r['disease'])), 'thalach': (r.get('thalach') or '').strip()} for r in rows], ensure_ascii=False))",
  ].join("\n");
  const result = spawnSync(PYTHON, ["-c", py], {
    encoding: "utf8",
    maxBuffer: 8 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(`Could not read source CSV: ${result.stderr}`);
  }
  return JSON.parse(result.stdout);
}

function addDotPlot(slide, rows) {
  addBase(slide, "差距画在哪", 6);
  const plot = { left: 132, right: 1160, top: 190, bottom: 535 };
  const minX = 70;
  const maxX = 210;
  const scaleX = (value) =>
    plot.left + ((value - minX) / (maxX - minX)) * (plot.right - plot.left);
  const lanes = [
    { group: 1, y: 315, color: COLORS.red, label: "有疾病证据组", mean: 128.26 },
    { group: 0, y: 443, color: COLORS.teal, label: "无疾病证据组", mean: 148.8 },
  ];
  for (const lane of lanes) {
    text(slide, `lane-label-${lane.group}`, lane.label, 86, lane.y - 61, 360, 32, {
      size: 21,
      color: lane.color,
      bold: true,
    });
    rows.forEach((row, index) => {
      if (row.disease !== lane.group || !row.thalach) return;
      const value = Number(row.thalach);
      const stable = ((index * 37 + lane.group * 17) % 101) / 100;
      const jitter = (stable - 0.5) * 74;
      const dot = slide.shapes.add({
        geometry: "ellipse",
        name: `observation-${lane.group}-${index}`,
        position: {
          left: scaleX(value) - 3,
          top: lane.y + jitter - 3,
          width: 6,
          height: 6,
        },
        fill: `${lane.color}/72`,
        line: { style: "solid", fill: "none", width: 0 },
      });
      dot.altText = `${lane.label}，thalach=${value} 次/分`;
    });
    rule(slide, `mean-line-${lane.group}`, scaleX(lane.mean), lane.y - 47, scaleX(lane.mean), lane.y + 47, lane.color, 5);
    text(
      slide,
      `mean-label-${lane.group}`,
      `${lane.mean.toFixed(2)}`,
      scaleX(lane.mean) - 48,
      lane.y + 49,
      120,
      28,
      { size: 18, color: lane.color, bold: true, font: MONO },
    );
  }
  rule(slide, "dotplot-axis", plot.left, 550, plot.right, 550, COLORS.gray, 1.5);
  for (let tick = 80; tick <= 200; tick += 20) {
    const x = scaleX(tick);
    rule(slide, `tick-${tick}`, x, 544, x, 557, COLORS.gray, 1);
    text(slide, `tick-label-${tick}`, String(tick), x - 28, 561, 56, 22, {
      size: 15,
      color: COLORS.gray,
      font: MONO,
      align: "center",
    });
  }
  text(slide, "axis-label", "最大运动心率 thalach（次/分）", 433, 598, 430, 28, {
    size: 18,
    color: COLORS.muted,
    align: "center",
  });
  const meanOne = scaleX(128.2616);
  const meanZero = scaleX(148.8005);
  rule(slide, "delta-bracket-left", meanOne, 226, meanOne, 207, COLORS.ink, 1.5);
  rule(slide, "delta-bracket", meanOne, 207, meanZero, 207, COLORS.ink, 1.5);
  rule(slide, "delta-bracket-right", meanZero, 207, meanZero, 226, COLORS.ink, 1.5);
  text(slide, "delta-label", "Δ = −20.54 次/分", meanOne - 24, 164, meanZero - meanOne + 80, 30, {
    size: 19,
    font: MONO,
    bold: true,
    align: "center",
  });
  text(slide, "dotplot-aside", "每个点 = 1 人    /    点的纵向位置只为减少重叠", 430, 632, 700, 22, {
    size: 15,
    color: COLORS.gray,
    align: "right",
  });
}

function addPipeline(slide) {
  addBase(slide, "差距来自哪里：差值 ÷ 标准误", 7);
  const columns = [
    { x: 76, title: "01  均值差", value: "−20.54", unit: "次/分", color: COLORS.ink },
    { x: 359, title: "02  标准误", value: "1.6258", unit: "差值估计的不确定度", color: COLORS.gray },
    { x: 642, title: "03  t 值", value: "−12.633", unit: "差值 ÷ 标准误", color: COLORS.ink },
    { x: 925, title: "04  判定", value: "p≈1.30×10⁻³³", unit: "df≈837.18  /  双侧", color: COLORS.ink },
  ];
  columns.forEach((item, index) => {
    text(slide, `pipeline-heading-${index}`, item.title, item.x, 151, 255, 34, {
      size: 19,
      color: item.color,
      bold: true,
    });
    text(slide, `pipeline-value-${index}`, item.value, item.x, 195, 260, 53, {
      size: index === 3 ? 27 : 36,
      color: item.color,
      bold: true,
      font: MONO,
    });
    text(slide, `pipeline-unit-${index}`, item.unit, item.x, 250, 260, 42, {
      size: 16,
      color: COLORS.muted,
    });
    if (index < 3) {
      rule(slide, `pipeline-arrow-${index}`, item.x + 258, 221, item.x + 280, 221, COLORS.gray, 2);
      rule(slide, `pipeline-arrow-head-${index}`, item.x + 272, 214, item.x + 280, 221, COLORS.gray, 2);
      rule(slide, `pipeline-arrow-foot-${index}`, item.x + 272, 228, item.x + 280, 221, COLORS.gray, 2);
    }
  });
  rule(slide, "pipeline-divider", 76, 294, 1190, 294, COLORS.rule, 1);
  text(slide, "pipeline-code-title", "build_docs.py 中实际执行的核心代码", 82, 305, 760, 26, {
    size: 17,
    color: COLORS.gray,
    bold: true,
  });
  rect(slide, "code-strip", 76, 335, 1128, 282, "#ECEFF2");
  const codeLines = [
    "standard_error = math.sqrt(",
    "    variances[1] / len(groups[1]) +",
    "    variances[0] / len(groups[0])",
    ")",
    "difference = means[1] - means[0]",
    "t_value = difference / standard_error",
    "df = (",
    "    (variances[1] / len(groups[1]) + variances[0] / len(groups[0])) ** 2",
    "    / ((variances[1] / len(groups[1])) ** 2 / (len(groups[1]) - 1)",
    "       + (variances[0] / len(groups[0])) ** 2 / (len(groups[0]) - 1))",
    ")",
    "p_value = t_two_sided_p(t_value, df)",
  ];
  codeLines.forEach((line, index) => {
    text(slide, `pipeline-code-${index}`, line, 94, 338 + index * 22, 1090, 23, {
      size: 20,
      font: MONO,
    });
  });
  text(
    slide,
    "pipeline-p-method",
    "p 值由正则化不完全 Beta 函数的连分数计算；95% 临界值用二分搜索；仅使用 Python 标准库。",
    82,
    632,
    1090,
    22,
    { size: 15, color: COLORS.muted },
  );
}

function addInterval(slide) {
  addBase(slide, "差距有多大把握", 8);
  text(slide, "mean-summary-title", "两组样本均值（次/分）", 94, 137, 300, 23, {
    size: 17,
    color: COLORS.muted,
  });
  const meanAxis = { left: 390, right: 890, min: 110, max: 170 };
  const meanX = (value) =>
    meanAxis.left +
    ((value - meanAxis.min) / (meanAxis.max - meanAxis.min)) *
      (meanAxis.right - meanAxis.left);
  const means = [
    { label: "有疾病证据组", value: 128.26, y: 178, color: COLORS.red },
    { label: "无疾病证据组", value: 148.8, y: 205, color: COLORS.teal },
  ];
  means.forEach((item, index) => {
    text(slide, `mean-group-${index}`, item.label, 96, item.y - 13, 235, 24, {
      size: 17,
      color: item.color,
      bold: true,
    });
    slide.shapes.add({
      geometry: "ellipse",
      name: `mean-point-${index}`,
      position: {
        left: meanX(item.value) - 7,
        top: item.y - 7,
        width: 14,
        height: 14,
      },
      fill: item.color,
      line: { style: "solid", fill: COLORS.white, width: 1 },
    });
    text(
      slide,
      `mean-value-${index}`,
      item.value.toFixed(2),
      meanX(item.value) + 12,
      item.y - 13,
      115,
      25,
      { size: 17, color: item.color, bold: true, font: MONO },
    );
  });
  rule(slide, "mean-axis", meanAxis.left, 232, meanAxis.right, 232, COLORS.gray, 1);
  [110, 140, 170].forEach((tick) => {
    const x = meanX(tick);
    rule(slide, `mean-tick-${tick}`, x, 228, x, 236, COLORS.gray, 1);
    text(slide, `mean-tick-label-${tick}`, String(tick), x - 28, 238, 56, 18, {
      size: 15,
      color: COLORS.gray,
      font: MONO,
      align: "center",
    });
  });
  text(slide, "ci-chart-subtitle", "均值差 Δ（次/分）及其 95% 置信区间", 94, 260, 650, 22, {
    size: 17,
    color: COLORS.muted,
  });
  slide.charts.add("scatter", {
    position: { left: 110, top: 284, width: 1040, height: 200 },
    series: [
      {
        name: "95% 置信区间",
        xValues: [-23.73, -17.35],
        values: [0, 0],
        line: { style: "solid", fill: COLORS.ink, width: 5 },
        marker: { symbol: "none", size: 2 },
      },
      {
        name: "均值差",
        xValues: [-20.54],
        values: [0],
        line: { style: "solid", fill: COLORS.ink, width: 0 },
        marker: { symbol: "diamond", size: 13, fill: COLORS.ink },
      },
      {
        name: "零差异参考线",
        xValues: [0, 0],
        values: [-0.8, 0.8],
        line: { style: "dashed", fill: COLORS.gray, width: 2 },
        marker: { symbol: "none", size: 2 },
      },
    ],
    scatterOptions: { style: "lineWithMarkers" },
    hasLegend: false,
    chartFill: COLORS.white,
    chartLine: { style: "solid", fill: COLORS.rule, width: 1 },
    plotAreaFill: COLORS.white,
    plotAreaLine: { style: "solid", fill: COLORS.white, width: 0 },
    xAxis: {
      visible: true,
      min: -30,
      max: 5,
      majorUnit: 5,
      numberFormatCode: "0",
      textStyle: { fill: COLORS.muted, fontSize: 18 },
      line: { style: "solid", fill: COLORS.gray, width: 1 },
      majorGridlines: { style: "solid", fill: "#E9ECEF", width: 1 },
    },
    yAxis: {
      visible: false,
      min: -1,
      max: 1,
      majorGridlines: null,
      line: { style: "solid", fill: COLORS.white, width: 1 },
    },
  });
  hatchedRect(slide, "zero-compatible-band", 987, 309, 141, 111);
  rule(slide, "zero-reference-overlay", 987, 309, 987, 420, COLORS.gray, 2, "dashed");
  text(slide, "zero-label", "0  均值相同", 981, 286, 152, 22, {
    size: 16,
    color: COLORS.gray,
    bold: true,
    font: MONO,
    align: "center",
  });
  text(slide, "zero-band-label", "与 0 相容的范围", 990, 425, 138, 21, {
    size: 15,
    color: COLORS.gray,
    align: "center",
  });
  text(slide, "ci-left-value", "−23.73", 242, 319, 150, 25, {
    size: 18,
    color: COLORS.ink,
    bold: true,
    font: MONO,
    align: "center",
  });
  text(slide, "ci-center-value", "Δ = −20.54", 290, 415, 220, 26, {
    size: 18,
    color: COLORS.ink,
    bold: true,
    font: MONO,
    align: "center",
  });
  text(slide, "ci-right-value", "−17.35", 425, 319, 150, 25, {
    size: 18,
    color: COLORS.ink,
    bold: true,
    font: MONO,
    align: "center",
  });
  rule(slide, "ci-chart-baseline", 94, 484, 1185, 484, COLORS.rule, 1);
  const stats = [
    ["有疾病证据组", "n=474   128.26 ± 24.02   中位数 128", COLORS.red],
    ["无疾病证据组", "n=391   148.80 ± 23.61   中位数 151", COLORS.teal],
  ];
  stats.forEach(([label, value, color], index) => {
    const y = 493 + index * 39;
    text(slide, `ci-stat-label-${index}`, label, 98, y, 210, 27, {
      size: 16,
      color,
      bold: true,
    });
    text(slide, `ci-stat-value-${index}`, value, 306, y, 760, 27, {
      size: 17,
      font: MONO,
    });
  });
  text(
    slide,
    "ci-result",
    "Δ=−20.54   95% CI [−23.73, −17.35]   t=−12.633   df≈837.18   p≈1.30×10⁻³³",
    98,
    576,
    1080,
    27,
    { size: 17, font: MONO, bold: true },
  );
  text(slide, "ci-interpretation", "区间整个落在 0 左侧，数据不支持“没有差别”。", 98, 620, 660, 25, {
    size: 16,
    color: COLORS.muted,
  });
  text(slide, "ci-decision", "α = 0.05  /  拒绝 H0", 842, 635, 330, 24, {
    size: 17,
    color: COLORS.ink,
    bold: true,
    align: "right",
  });
}

function addScope(slide) {
  addBase(slide, "我们做的是关联", 9);
  const columns = [
    {
      x: 82,
      title: "关联",
      question: "两组平均值是否不同？",
      status: "本次问题",
      detail: "比较 disease 组与 thalach",
      color: COLORS.ink,
    },
    {
      x: 365,
      title: "预测",
      question: "能判断新个体属于哪组吗？",
      status: "本次未运行",
      detail: "没有训练模型或测试性能",
      color: COLORS.gray,
    },
    {
      x: 648,
      title: "因果",
      question: "疾病状态造成心率改变吗？",
      status: "本次未运行",
      detail: "观察性比较不能识别因果",
      color: COLORS.gray,
    },
    {
      x: 931,
      title: "本次范围",
      question: "未调整的组间均值比较",
      status: "已完成",
      detail: "Welch 独立样本 t 检验",
      color: COLORS.ink,
    },
  ];
  hatchedRect(slide, "prediction-unrun", 360, 190, 270, 337);
  hatchedRect(slide, "causal-unrun", 643, 190, 270, 337);
  rule(slide, "scope-top-rule", 82, 226, 1193, 226, COLORS.rule, 1);
  rule(slide, "scope-bottom-rule", 82, 492, 1193, 492, COLORS.rule, 1);
  [351, 634, 917].forEach((x, i) =>
    rule(slide, `scope-divider-${i}`, x, 190, x, 527, COLORS.rule, 1),
  );
  columns.forEach((item, index) => {
    text(slide, `scope-title-${index}`, item.title, item.x, 263, 248, 48, {
      size: 29,
      color: item.color,
      bold: true,
    });
    text(slide, `scope-question-${index}`, item.question, item.x, 334, 244, 68, {
      size: 20,
    });
    text(slide, `scope-status-${index}`, item.status, item.x, 428, 248, 28, {
      size: 16,
      color: item.color,
      bold: true,
    });
    text(slide, `scope-detail-${index}`, item.detail, item.x, 547, 248, 54, {
      size: 17,
      color: COLORS.muted,
    });
  });
  text(slide, "scope-bottom", "关联 ≠ 预测能力    /    预测能力 ≠ 因果效应", 82, 620, 1070, 27, {
    size: 19,
    color: COLORS.muted,
    bold: true,
  });
}

function addClaims(slide) {
  addBase(slide, "能做 / 不能做", 10);
  hatchedRect(slide, "claims-unsupported-band", 650, 208, 528, 244);
  text(slide, "claims-left-title", "这份数据支持我们说", 88, 164, 500, 40, {
    size: 25,
    bold: true,
  });
  text(slide, "claims-right-title", "这份数据不支持我们说", 660, 164, 520, 40, {
    size: 25,
    color: COLORS.gray,
    bold: true,
  });
  rule(slide, "claims-mid-divider", 632, 164, 632, 450, COLORS.rule, 1);
  const supported = [
    "两组最大运动心率的样本均值不同",
    "差值约 20.5 次/分，疾病证据组更低",
    "差值的 95% 区间没有包含 0",
    "这是观察性资料中的统计关联",
  ];
  const unsupported = [
    "疾病导致最大运动心率下降",
    "thalach 单项可以诊断某个个体",
    "差异一定具有临床重要性",
    "某种治疗已经证明有效",
  ];
  supported.forEach((value, index) => {
    const y = 224 + index * 58;
    text(slide, `supported-mark-${index}`, "可说", 92, y, 72, 30, {
      size: 17,
      bold: true,
    });
    text(slide, `supported-text-${index}`, value, 177, y, 430, 36, {
      size: 18,
    });
    text(slide, `unsupported-mark-${index}`, "越界", 660, y, 72, 30, {
      size: 17,
      color: COLORS.gray,
      bold: true,
    });
    text(slide, `unsupported-text-${index}`, value === "疾病导致最大运动心率下降" ? value : unsupported[index], 744, y, 430, 36, {
      size: 18,
      color: COLORS.muted,
    });
  });
  rule(slide, "claim-interaction-rule", 88, 478, 1178, 478, COLORS.rule, 1);
  text(slide, "claim-interaction-title", "现场判断", 92, 499, 148, 27, {
    size: 18,
    color: COLORS.ink,
    bold: true,
  });
  [
    "① “有疾病证据的人平均更低” → 可说",
    "② “疾病导致心率下降” → 不能说",
    "③ “单看 thalach 能诊断个人” → 不能说",
  ].forEach((value, index) => {
    text(slide, `claim-interaction-${index}`, value, 92, 534 + index * 27, 1070, 24, {
      size: 17,
    });
  });
  text(slide, "claim-stat-clinical", "p 很小 ≠ 临床意义确定    /    区间不含 0 ≠ 因果成立", 92, 624, 1070, 28, {
    size: 18,
    color: COLORS.muted,
    bold: true,
  });
}

function addNext(slide) {
  addBase(slide, "如果再做一步", 11);
  text(slide, "done-title", "本次已完成", 88, 164, 300, 34, {
    size: 23,
    bold: true,
  });
  text(
    slide,
    "done-summary",
    "分组描述统计、thalach 缺失计数、完整病例筛选、未调整 Welch 均值比较与 95% 区间",
    88,
    211,
    1080,
    52,
    { size: 21, color: COLORS.ink },
  );
  rule(slide, "next-stage-rule", 88, 286, 1188, 286, COLORS.rule, 1);
  text(slide, "proposed-title", "下一步研究建议", 88, 307, 390, 36, {
    size: 23,
    color: COLORS.gray,
    bold: true,
  });
  hatchedRect(slide, "next-unrun-band", 88, 352, 1100, 228);
  text(slide, "not-run-tag", "尚未运行", 106, 367, 138, 26, {
    size: 16,
    color: COLORS.muted,
    bold: true,
  });
  text(slide, "proposed-model", "thalach ~ disease + age + sex + C(site)", 106, 405, 1030, 35, {
    size: 23,
    font: MONO,
  });
  text(slide, "proposed-model-explain", "含义：比较疾病组时，同时把年龄、性别与来源中心纳入模型。", 106, 452, 1020, 30, {
    size: 18,
    color: COLORS.muted,
  });
  text(slide, "proposed-sensitivity", "另可比较有无 thalach 者的已观测特征，评估缺失筛选是否影响样本。", 106, 497, 1020, 30, {
    size: 18,
    color: COLORS.muted,
  });
  text(slide, "proposed-multiple", "多项确认性检验可考虑 Holm / Bonferroni；探索性筛查可控制 FDR。", 106, 537, 1020, 27, {
    size: 16,
    color: COLORS.gray,
    bold: true,
  });
  text(slide, "next-caveat", "模型、敏感性分析、多重校正、预测、因果分析和插补均尚未运行。", 88, 610, 1080, 30, {
    size: 18,
    color: COLORS.gray,
    bold: true,
  });
}

function addLadder(slide) {
  addBase(slide, "主张 · 证据 · 推理 · 边界", 12);
  const rows = [
    {
      label: "主张",
      color: COLORS.ink,
      text: "在 thalach 完整的四中心样本中，有疾病证据组平均最大运动心率较低。",
    },
    {
      label: "证据",
      color: COLORS.red,
      text: "n=474 对 n=391；Δ=−20.54 次/分；95% CI [−23.73, −17.35]；p≈1.30×10⁻³³。",
    },
    {
      label: "推理",
      color: COLORS.ink,
      text: "区间不含 0 且方向为负；α=0.05 下拒绝均值相同的 H0。",
    },
    {
      label: "边界",
      color: COLORS.gray,
      text: "未调整的观察性完整病例比较，混杂、缺失与中心差异仍可能影响估计。",
    },
  ];
  rows.forEach((item, index) => {
    const y = 159 + index * 111;
    const lineWidth = 1110 - index * 84;
    text(slide, `ladder-label-${index}`, item.label, 90, y, 148, 36, {
      size: 24,
      color: item.color,
      bold: true,
    });
    text(slide, `ladder-text-${index}`, item.text, 260, y, 886, 61, {
      size: 20,
      color: COLORS.ink,
    });
    if (index < rows.length - 1) {
      rule(slide, `ladder-rule-${index}`, 90, y + 78, 90 + lineWidth, y + 78, COLORS.rule, 1);
    }
  });
  text(slide, "ladder-final-boundary", "不能据此推断因果，也不能推断个体诊断能力。", 260, 579, 886, 34, {
    size: 19,
    color: COLORS.gray,
    bold: true,
  });
}

async function addClosing(slide, item) {
  const image = new Uint8Array(await fs.readFile(COVER_IMAGE));
  slide.images.add({
    blob: image,
    contentType: "image/jpeg",
    alt: "暗场下显示心电波形的临床监护仪，与封面使用同一氛围图",
    fit: "cover",
    position: { left: 0, top: 0, width: WIDTH, height: HEIGHT },
  });
  rect(slide, "closing-image-shade", 0, 0, WIDTH, HEIGHT, "#0E1620/72");
  text(slide, "closing-kicker", "统计检验与假设设计", 82, 196, 1116, 32, {
    size: 18,
    color: COLORS.darkMuted,
    bold: true,
    align: "center",
  });
  text(
    slide,
    "closing-title",
    item.title,
    48,
    310,
    1184,
    76,
    { size: 31, color: COLORS.white, bold: true, align: "center", valign: "middle" },
  );
  rule(slide, "closing-rule", 380, 420, 900, 420, "#68727B", 1);
  text(
    slide,
    "closing-source",
    "四中心心脏病教学数据 · 数值依据原始 CSV 复算",
    180,
    460,
    920,
    34,
    { size: 17, color: COLORS.darkMuted, align: "center" },
  );
  text(slide, "closing-number", "13 / 13", 1095, 667, 100, 22, {
    size: 15,
    color: COLORS.darkMuted,
    font: MONO,
    align: "right",
  });
}

async function setChartAxisFonts(pptxPath) {
  const zip = await JSZip.loadAsync(await fs.readFile(pptxPath));
  const chartFile = zip.file("ppt/slides/charts/chart1.xml");
  if (!chartFile) {
    throw new Error("The exported presentation is missing chart1.xml.");
  }

  let axisCount = 0;
  let chartXml = await chartFile.async("string");
  chartXml = chartXml.replace(/<c:valAx>[\s\S]*?<\/c:valAx>/g, (axisXml) => {
    axisCount += 1;
    let textPropertyCount = 0;
    const patchedAxis = axisXml.replace(
      /<c:txPr>[\s\S]*?<\/c:txPr>/g,
      (textProperties) => {
        textPropertyCount += 1;
        const defaultRunCount = (
          textProperties.match(/<a:defRPr\b[\s\S]*?<\/a:defRPr>/g) || []
        ).length;
        if (defaultRunCount !== 1) {
          throw new Error(
            `Expected one default run style in chart axis ${axisCount}, found ${defaultRunCount}.`,
          );
        }

        return textProperties.replace(
          /<a:defRPr\b[\s\S]*?<\/a:defRPr>/,
          (defaultRunStyle) => {
            let patchedRunStyle = defaultRunStyle;
            for (const script of ["latin", "ea"]) {
              const scriptFont = new RegExp(`<a:${script}\\b[^>]*/>`);
              const declaration = `<a:${script} typeface="${FONT}" />`;
              if (scriptFont.test(patchedRunStyle)) {
                patchedRunStyle = patchedRunStyle.replace(scriptFont, declaration);
              } else {
                patchedRunStyle = patchedRunStyle.replace(
                  /<\/a:defRPr>$/,
                  `${declaration}</a:defRPr>`,
                );
              }
            }
            return patchedRunStyle;
          },
        );
      },
    );
    if (textPropertyCount !== 1) {
      throw new Error(
        `Expected one text property in chart axis ${axisCount}, found ${textPropertyCount}.`,
      );
    }
    return patchedAxis;
  });
  if (axisCount !== 2) {
    throw new Error(`Expected 2 value axes, found ${axisCount}.`);
  }

  zip.file("ppt/slides/charts/chart1.xml", chartXml);
  await fs.writeFile(
    pptxPath,
    await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    }),
  );
}

const csvRows = readCsvRows();
if (csvRows.length !== 920) {
  throw new Error(`Expected 920 source rows, got ${csvRows.length}`);
}
const completeCount = csvRows.filter((row) => row.thalach).length;
if (completeCount !== 865) {
  throw new Error(`Expected 865 complete cases, got ${completeCount}`);
}

for (const item of CONTENT.slides) {
  const slide = presentation.slides.add();
  switch (item.layout) {
    case "cover":
      await addCover(slide, item);
      break;
    case "question":
      addQuestion(slide);
      break;
    case "groups":
      addGroups(slide);
      break;
    case "hypotheses":
      addHypotheses(slide);
      break;
    case "funnel":
      addFunnel(slide);
      break;
    case "dotplot":
      addDotPlot(slide, csvRows);
      break;
    case "pipeline":
      addPipeline(slide);
      break;
    case "interval":
      addInterval(slide);
      break;
    case "scope":
      addScope(slide);
      break;
    case "claims":
      addClaims(slide);
      break;
    case "next":
      addNext(slide);
      break;
    case "ladder":
      addLadder(slide);
      break;
    case "closing":
      await addClosing(slide, item);
      break;
    default:
      throw new Error(`Unknown layout: ${item.layout}`);
  }
  addNotes(slide, item);
}

if (presentation.slides.items.length !== 13) {
  throw new Error(`Expected 13 slides, got ${presentation.slides.items.length}`);
}
if (await fs.stat(FINAL_PPTX).catch(() => null)) {
  throw new Error(`Refusing to overwrite ${FINAL_PPTX}`);
}

await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
await fs.mkdir(PREVIEW_DIR, { recursive: true });
for (const [index, slide] of presentation.slides.items.entries()) {
  const png = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(
    path.join(PREVIEW_DIR, `slide-${String(index + 1).padStart(2, "0")}.png`),
    new Uint8Array(await png.arrayBuffer()),
  );
}
const montage = await presentation.export({ format: "webp", montage: true });
await fs.writeFile(
  path.join(PREVIEW_DIR, "deck-montage.webp"),
  new Uint8Array(await montage.arrayBuffer()),
);

const stagingDir = path.join(ROOT, ".codex-finalizer-v2");
await fs.mkdir(stagingDir, { recursive: true });
const candidatePath = path.join(stagingDir, "candidate-v2.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
await setChartAxisFonts(candidatePath);
const { finalizePresentation } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href
);
const result = await finalizePresentation({
  workspaceDir: ROOT,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: PYTHON,
  integrityValidatorPath: path.join(
    SKILL_DIR,
    "container_tools/inspect_presentation_package_integrity.py",
  ),
  layoutValidatorPath: path.join(
    SKILL_DIR,
    "container_tools/inspect_presentation_layout_geometry.py",
  ),
  layoutArgs: [
    "--expected-slide-size-emu",
    "12192000,6858000",
    "--validate-bullet-geometry",
    "--validate-heading-fit",
  ],
  explicitTotalSlideCount: 13,
  requiredNativeChartOwnerSlides: [8],
  materializeLiteralChartWorkbooks: true,
  fontPolicy: {
    basis: "design",
    families: [FONT, MONO],
  },
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, "statistics-hypothesis-v2.validation.json"),
});
console.log(JSON.stringify({ finalPath: FINAL_PPTX, validation: result }, null, 2));
