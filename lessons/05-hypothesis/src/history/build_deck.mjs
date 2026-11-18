import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = path.resolve(".");
const SKILL_DIR =
  "C:\\Users\\A\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1007.11041\\skills\\presentations";
const FONT = "Microsoft YaHei";
const WIDTH = 1280;
const HEIGHT = 720;
const FINAL_PPTX = path.join(
  ROOT,
  "deliverables",
  "统计检验与假设设计_演讲PPT.pptx",
);
const CONTENT = JSON.parse(
  await fs.readFile(path.join(ROOT, "presentation-content.json"), "utf8"),
);
const BACKGROUND = new Uint8Array(
  await fs.readFile(
    "C:\\Users\\A\\ppt_high_dimension_reflection\\.codex-build\\assets-v8\\tech-network.jpg",
  ),
);
const { applyPresentationChartFont } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href
);

const COLORS = {
  ink: "#142B40",
  blue: "#087FA3",
  cyan: "#13A7C4",
  coral: "#FF735D",
  amber: "#C57B16",
  gray: "#536576",
  lightGray: "#D8E0E8",
  white: "#FFFFFF",
};

const presentation = Presentation.create({
  slideSize: { width: WIDTH, height: HEIGHT },
});

function addShape(slide, name, position, fill = "none", lineFill = "none") {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position,
    fill,
    line: { style: "solid", fill: lineFill, width: lineFill === "none" ? 0 : 1 },
  });
}

function addText(
  slide,
  name,
  text,
  position,
  {
    fontSize = 24,
    color = COLORS.ink,
    bold = false,
    alignment = "left",
    verticalAlignment = "top",
    fill = "none",
    insets = 0,
  } = {},
) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: FONT,
    fontSize,
    color,
    bold,
    alignment,
    verticalAlignment,
    autoFit: "shrinkText",
    wrap: "square",
    insets,
  };
  return shape;
}

function addBackground(slide, cover = false) {
  slide.background.fill = "#09182A";
  slide.images.add({
    blob: BACKGROUND,
    contentType: "image/jpeg",
    alt: "Blue digital network background, used as a decorative visual",
    fit: "cover",
    position: { left: 0, top: 0, width: WIDTH, height: HEIGHT },
  });
  addShape(
    slide,
    cover ? "cover-dark-overlay" : "content-light-overlay",
    { left: 0, top: 0, width: WIDTH, height: HEIGHT },
    cover ? "#06152A/86" : "#F2F4F8/86",
  );
  addShape(slide, "coral-left-rule", { left: 0, top: 0, width: 8, height: HEIGHT }, COLORS.coral);
}

function addContentBase(slide, title, index) {
  addBackground(slide, false);
  addText(
    slide,
    "section-label",
    "PBL  /  统计检验",
    { left: 76, top: 34, width: 280, height: 26 },
    { fontSize: 15, color: COLORS.blue, bold: true },
  );
  addText(
    slide,
    "slide-title",
    title,
    { left: 76, top: 72, width: 1125, height: 62 },
    { fontSize: 36, color: COLORS.ink, bold: true },
  );
  addShape(
    slide,
    "heading-rule",
    { left: 76, top: 144, width: 1125, height: 2 },
    "#087FA3/52",
  );
  addText(
    slide,
    "page-number",
    String(index).padStart(2, "0"),
    { left: 1150, top: 674, width: 50, height: 22 },
    { fontSize: 14, color: "#627485", alignment: "right" },
  );
}

function addLabelValue(slide, y, label, value, index) {
  addText(
    slide,
    `fact-label-${index}`,
    label,
    { left: 92, top: y, width: 220, height: 38 },
    { fontSize: 21, color: COLORS.blue, bold: true },
  );
  addText(
    slide,
    `fact-value-${index}`,
    value,
    { left: 340, top: y, width: 820, height: 48 },
    { fontSize: 24, color: COLORS.ink },
  );
  if (index < 4) {
    addShape(
      slide,
      `fact-rule-${index}`,
      { left: 92, top: y + 53, width: 1068, height: 1 },
      "#AEBBC8/65",
    );
  }
}

function addCover(slide, content) {
  addBackground(slide, true);
  addText(
    slide,
    "cover-kicker",
    "第五章  /  PBL 统计实践",
    { left: 82, top: 70, width: 440, height: 32 },
    { fontSize: 18, color: "#A9E9F1", bold: true },
  );
  addText(
    slide,
    "cover-title",
    content.title,
    { left: 82, top: 176, width: 960, height: 96 },
    { fontSize: 58, color: COLORS.white, bold: true },
  );
  addText(
    slide,
    "cover-subtitle",
    content.subtitle,
    { left: 86, top: 298, width: 1050, height: 88 },
    { fontSize: 29, color: "#E6F2F5" },
  );
  addText(
    slide,
    "cover-topics",
    "研究问题  ·  检验方法  ·  结果边界",
    { left: 88, top: 428, width: 750, height: 44 },
    { fontSize: 22, color: COLORS.coral, bold: true },
  );
}

function addQuestionSlide(slide, content) {
  addContentBase(slide, content.title, 2);
  addText(
    slide,
    "research-question",
    content.headline,
    { left: 90, top: 171, width: 1080, height: 66 },
    { fontSize: 31, color: COLORS.coral, bold: true },
  );
  const yPositions = [264, 334, 404, 474, 544];
  content.facts.forEach(([label, value], index) => {
    addLabelValue(slide, yPositions[index], label, value, index);
  });
}

function addHypothesisSlide(slide, content) {
  addContentBase(slide, content.title, 3);
  const blocks = [
    { x: 92, label: "H0", formula: "Δ = 0", detail: "两组总体均值相同", color: COLORS.blue },
    { x: 660, label: "H1", formula: "Δ ≠ 0", detail: "两组总体均值不同", color: COLORS.coral },
  ];
  for (const [index, block] of blocks.entries()) {
    addText(
      slide,
      `hypothesis-label-${index}`,
      block.label,
      { left: block.x, top: 205, width: 180, height: 42 },
      { fontSize: 27, color: block.color, bold: true },
    );
    addText(
      slide,
      `hypothesis-formula-${index}`,
      block.formula,
      { left: block.x, top: 270, width: 460, height: 68 },
      { fontSize: 48, color: COLORS.ink, bold: true },
    );
    addText(
      slide,
      `hypothesis-detail-${index}`,
      block.detail,
      { left: block.x, top: 360, width: 460, height: 48 },
      { fontSize: 25, color: COLORS.gray },
    );
  }
  addShape(slide, "hypothesis-divider", { left: 625, top: 205, width: 2, height: 210 }, "#98AABA/65");
  addText(
    slide,
    "alpha",
    content.alpha,
    { left: 92, top: 474, width: 260, height: 56 },
    { fontSize: 34, color: COLORS.coral, bold: true },
  );
  addText(
    slide,
    "test-direction-reason",
    content.reason,
    { left: 360, top: 474, width: 800, height: 72 },
    { fontSize: 23, color: COLORS.ink },
  );
  addText(
    slide,
    "pre-specification",
    "方向与 α 应在查看结果前确定",
    { left: 92, top: 583, width: 800, height: 34 },
    { fontSize: 19, color: COLORS.gray },
  );
}

function addPowerSlide(slide, content) {
  addContentBase(slide, content.title, 4);
  const columns = [
    { x: 92, label: content.errors[0][0], definition: content.errors[0][1], example: content.errors[0][2], color: COLORS.coral },
    { x: 660, label: content.errors[1][0], definition: content.errors[1][1], example: content.errors[1][2], color: COLORS.blue },
  ];
  for (const [index, item] of columns.entries()) {
    addText(
      slide,
      `error-label-${index}`,
      item.label,
      { left: item.x, top: 186, width: 460, height: 44 },
      { fontSize: 25, color: item.color, bold: true },
    );
    addText(
      slide,
      `error-definition-${index}`,
      item.definition,
      { left: item.x, top: 246, width: 480, height: 65 },
      { fontSize: 24, color: COLORS.ink },
    );
    addText(
      slide,
      `error-example-${index}`,
      item.example,
      { left: item.x, top: 324, width: 480, height: 52 },
      { fontSize: 21, color: COLORS.gray },
    );
  }
  addShape(slide, "error-divider", { left: 625, top: 190, width: 2, height: 190 }, "#98AABA/65");
  addShape(slide, "power-rule", { left: 92, top: 403, width: 1068, height: 1 }, "#AEBBC8/65");
  addText(
    slide,
    "power-label",
    "Power",
    { left: 92, top: 433, width: 160, height: 42 },
    { fontSize: 25, color: COLORS.blue, bold: true },
  );
  content.power.forEach(([label, value], index) => {
    const x = 260 + index * 300;
    addText(
      slide,
      `power-factor-${index}`,
      `${label}\n${value}`,
      { left: x, top: 431, width: 270, height: 88 },
      { fontSize: 21, color: COLORS.ink },
    );
  });
  addText(
    slide,
    "power-boundary",
    content.boundary,
    { left: 92, top: 574, width: 700, height: 50 },
    { fontSize: 26, color: COLORS.coral, bold: true },
  );
}

function addMethodsSlide(slide, content) {
  addContentBase(slide, content.title, 5);
  const table = slide.tables.add({
    rows: content.rows.length,
    columns: content.rows[0].length,
    left: 76,
    top: 174,
    width: 1128,
    height: 365,
    columnWidths: [326, 300, 502],
    values: content.rows,
  });
  table.borders.assign({ style: "solid", fill: "#CAD4DE", width: 1 });
  for (let row = 0; row < content.rows.length; row += 1) {
    for (let column = 0; column < content.rows[0].length; column += 1) {
      const cell = table.getCell(row, column);
      cell.fill = row === 0 ? "#173B57" : row % 2 ? "#FFFFFF" : "#E9F0F4";
      cell.text.style = {
        typeface: FONT,
        fontSize: row === 0 ? 20 : 19,
        color: row === 0 ? "#FFFFFF" : COLORS.ink,
        bold: row === 0 || column === 0,
        verticalAlignment: "middle",
        alignment: "left",
        autoFit: "shrinkText",
        wrap: "square",
      };
      cell.text.insets = { top: 10, right: 12, bottom: 10, left: 12 };
    }
  }
  addText(
    slide,
    "method-selection-note",
    "先看研究设计与变量类型，再看检验假设；备选方法改变时，结论解释也要相应调整。",
    { left: 80, top: 578, width: 1100, height: 56 },
    { fontSize: 20, color: COLORS.gray },
  );
}

function addAssociationSlide(slide, content) {
  addContentBase(slide, content.title, 6);
  const columns = [
    { x: 92, item: content.columns[0], color: COLORS.blue },
    { x: 660, item: content.columns[1], color: COLORS.coral },
  ];
  for (const [index, { x, item, color }] of columns.entries()) {
    addText(
      slide,
      `question-heading-${index}`,
      item.heading,
      { left: x, top: 180, width: 480, height: 52 },
      { fontSize: 25, color, bold: true },
    );
    addText(
      slide,
      `question-method-${index}`,
      item.method,
      { left: x, top: 246, width: 480, height: 42 },
      { fontSize: 22, color: COLORS.ink, bold: true },
    );
    addText(
      slide,
      `question-detail-${index}`,
      item.detail,
      { left: x, top: 311, width: 478, height: 108 },
      { fontSize: 22, color: COLORS.gray },
    );
    addShape(
      slide,
      `question-divider-${index}`,
      { left: x, top: 437, width: 478, height: 1 },
      "#AEBBC8/65",
    );
    addText(
      slide,
      `question-alternative-${index}`,
      `备选｜${item.alternative}`,
      { left: x, top: 461, width: 480, height: 88 },
      { fontSize: 20, color: COLORS.ink },
    );
  }
  addShape(slide, "association-divider", { left: 625, top: 185, width: 2, height: 364 }, "#98AABA/65");
  addText(
    slide,
    "association-distinction",
    content.distinction,
    { left: 92, top: 584, width: 1080, height: 46 },
    { fontSize: 23, color: COLORS.coral, bold: true },
  );
}

function addSampleSlide(slide, content) {
  addContentBase(slide, content.title, 7);
  content.metrics.forEach(([label, number, description], index) => {
    const x = 90 + index * 380;
    addText(
      slide,
      `sample-label-${index}`,
      label,
      { left: x, top: 190, width: 340, height: 40 },
      { fontSize: 21, color: COLORS.blue, bold: true },
    );
    addText(
      slide,
      `sample-number-${index}`,
      number,
      { left: x, top: 246, width: 340, height: 86 },
      { fontSize: 58, color: index === 1 ? COLORS.coral : COLORS.ink, bold: true },
    );
    addText(
      slide,
      `sample-description-${index}`,
      description,
      { left: x, top: 340, width: 340, height: 34 },
      { fontSize: 21, color: COLORS.gray },
    );
    if (index < 2) {
      addShape(slide, `sample-divider-${index}`, { left: x + 350, top: 194, width: 1, height: 178 }, "#98AABA/65");
    }
  });
  addShape(slide, "sample-groups-rule", { left: 92, top: 411, width: 1068, height: 1 }, "#AEBBC8/65");
  addText(
    slide,
    "sample-group-disease",
    content.groups[0][0],
    { left: 92, top: 444, width: 235, height: 38 },
    { fontSize: 21, color: COLORS.coral, bold: true },
  );
  addText(
    slide,
    "sample-group-disease-detail",
    content.groups[0][1],
    { left: 334, top: 444, width: 340, height: 44 },
    { fontSize: 21, color: COLORS.ink },
  );
  addText(
    slide,
    "sample-group-no-disease",
    content.groups[1][0],
    { left: 684, top: 444, width: 250, height: 38 },
    { fontSize: 21, color: COLORS.blue, bold: true },
  );
  addText(
    slide,
    "sample-group-no-disease-detail",
    content.groups[1][1],
    { left: 924, top: 444, width: 280, height: 60 },
    { fontSize: 20, color: COLORS.ink },
  );
  addText(
    slide,
    "sample-site-note",
    content.site,
    { left: 92, top: 555, width: 1070, height: 44 },
    { fontSize: 18, color: COLORS.gray },
  );
}

function addResultSlide(slide, content) {
  addContentBase(slide, content.title, 8);
  addText(
    slide,
    "result-chart-title",
    content.chartTitle,
    { left: 94, top: 170, width: 600, height: 36 },
    { fontSize: 21, color: COLORS.ink, bold: true },
  );
  const chart = slide.charts.add("bar", {
    position: { left: 82, top: 212, width: 650, height: 365 },
    categories: content.chartLabels,
    series: [
      {
        name: "平均 thalach",
        values: content.chartValues,
        fill: COLORS.cyan,
        points: [
          { idx: 0, fill: COLORS.coral },
          { idx: 1, fill: COLORS.cyan },
        ],
      },
    ],
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 105 },
    hasLegend: false,
    chartFill: "#FFFFFF",
    chartLine: { style: "solid", fill: "#CFD8E2", width: 1 },
    plotAreaFill: "#FFFFFF",
    plotAreaLine: { style: "solid", fill: "#FFFFFF", width: 0 },
    xAxis: {
      visible: true,
      textStyle: { fill: COLORS.ink, fontSize: 17 },
      line: { style: "solid", fill: "#B7C3CD", width: 1 },
      majorGridlines: null,
    },
    yAxis: {
      visible: true,
      title: { text: "次/分", textStyle: { fill: COLORS.gray, fontSize: 15 } },
      min: 0,
      max: 180,
      majorUnit: 30,
      textStyle: { fill: COLORS.gray, fontSize: 15 },
      line: { style: "solid", fill: "#B7C3CD", width: 1 },
      majorGridlines: { style: "solid", fill: "#E1E7EC", width: 1 },
    },
    dataLabels: {
      showValue: true,
      position: "outEnd",
      numberFormatCode: "0.00",
      textStyle: { fill: COLORS.ink, fontSize: 18, bold: true },
    },
  });
  applyPresentationChartFont(chart, { fontFamily: FONT });
  content.stats.forEach(([label, value], index) => {
    const y = 187 + index * 78;
    addText(
      slide,
      `result-label-${index}`,
      label,
      { left: 774, top: y, width: 390, height: 30 },
      { fontSize: 19, color: index === 2 ? COLORS.coral : COLORS.blue, bold: true },
    );
    addText(
      slide,
      `result-value-${index}`,
      value,
      { left: 774, top: y + 34, width: 410, height: 44 },
      { fontSize: 20, color: COLORS.ink },
    );
  });
  addText(
    slide,
    "result-decision",
    content.decision,
    { left: 774, top: 535, width: 390, height: 44 },
    { fontSize: 24, color: COLORS.coral, bold: true },
  );
}

function addInterpretationSlide(slide, content) {
  addContentBase(slide, content.title, 9);
  const yPositions = [181, 280, 379, 478];
  content.points.forEach(([label, text], index) => {
    addText(
      slide,
      `interpret-label-${index}`,
      label,
      { left: 92, top: yPositions[index], width: 185, height: 43 },
      { fontSize: 22, color: index === 3 ? COLORS.coral : COLORS.blue, bold: true },
    );
    addText(
      slide,
      `interpret-text-${index}`,
      text,
      { left: 286, top: yPositions[index], width: 865, height: 73 },
      { fontSize: 21, color: COLORS.ink },
    );
    if (index < 3) {
      addShape(
        slide,
        `interpret-rule-${index}`,
        { left: 92, top: yPositions[index] + 80, width: 1060, height: 1 },
        "#AEBBC8/65",
      );
    }
  });
}

function addRisksSlide(slide, content) {
  addContentBase(slide, content.title, 10);
  const leftX = 92;
  const rightX = 670;
  addText(
    slide,
    "confounder-heading",
    "混杂与缺失",
    { left: leftX, top: 175, width: 480, height: 42 },
    { fontSize: 24, color: COLORS.blue, bold: true },
  );
  content.confounders.forEach(([label, detail], index) => {
    const y = 231 + index * 78;
    addText(
      slide,
      `confounder-label-${index}`,
      label,
      { left: leftX, top: y, width: 150, height: 34 },
      { fontSize: 19, color: index === 2 ? COLORS.coral : COLORS.ink, bold: true },
    );
    addText(
      slide,
      `confounder-detail-${index}`,
      detail,
      { left: leftX + 155, top: y, width: 390, height: 66 },
      { fontSize: 18, color: COLORS.gray },
    );
  });
  addText(
    slide,
    "adjusted-model",
    content.model,
    { left: leftX, top: 478, width: 530, height: 57 },
    { fontSize: 18, color: COLORS.ink, bold: true },
  );
  addShape(slide, "risks-divider", { left: 635, top: 180, width: 1, height: 385 }, "#98AABA/65");
  addText(
    slide,
    "multiple-heading",
    "多重比较",
    { left: rightX, top: 175, width: 470, height: 42 },
    { fontSize: 24, color: COLORS.coral, bold: true },
  );
  addText(
    slide,
    "multiple-risk",
    content.multiple,
    { left: rightX, top: 232, width: 470, height: 98 },
    { fontSize: 20, color: COLORS.ink },
  );
  addText(
    slide,
    "multiple-control",
    content.control,
    { left: rightX, top: 349, width: 470, height: 150 },
    { fontSize: 19, color: COLORS.gray },
  );
  addText(
    slide,
    "risk-boundary",
    "调整能处理已测量因素，不能消除未测量混杂。",
    { left: rightX, top: 520, width: 470, height: 50 },
    { fontSize: 18, color: COLORS.coral, bold: true },
  );
}

function addExitSlide(slide, content) {
  addContentBase(slide, content.title, 11);
  const rows = [
    ["主张", content.claim],
    ["证据", content.evidence],
    ["推理", content.reasoning],
    ["边界", content.boundary],
  ];
  rows.forEach(([label, text], index) => {
    const y = 184 + index * 105;
    addText(
      slide,
      `exit-label-${index}`,
      label,
      { left: 92, top: y, width: 145, height: 44 },
      { fontSize: 23, color: index === 3 ? COLORS.coral : COLORS.blue, bold: true },
    );
    addText(
      slide,
      `exit-text-${index}`,
      text.slice(text.indexOf("｜") + 1).trim(),
      { left: 258, top: y, width: 900, height: 78 },
      { fontSize: 21, color: COLORS.ink },
    );
    if (index < 3) {
      addShape(
        slide,
        `exit-rule-${index}`,
        { left: 92, top: y + 83, width: 1065, height: 1 },
        "#AEBBC8/65",
      );
    }
  });
}

const builders = {
  cover: addCover,
  question: addQuestionSlide,
  hypothesis: addHypothesisSlide,
  power: addPowerSlide,
  methods: addMethodsSlide,
  association: addAssociationSlide,
  sample: addSampleSlide,
  result: addResultSlide,
  interpretation: addInterpretationSlide,
  risks: addRisksSlide,
  exit: addExitSlide,
};

for (const [index, content] of CONTENT.slides.entries()) {
  const slide = presentation.slides.add();
  if (content.kind === "cover") {
    builders.cover(slide, content);
  } else {
    builders[content.kind](slide, content);
  }
  const sourceNote =
    "\n\n资料来源：用户提供的 hypothesis.zip/data/heart-disease-hypothesis.csv、data/README.md 与 lesson-plan.md，以及课程统计检验课件。视觉背景：Unsplash photo-1644088379091-d574269d422f；仅作装饰，不作为统计证据。";
  slide.speakerNotes.text = `${content.notes}${sourceNote}`;
  slide.speakerNotes.setVisible(true);
}

await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
if (await fs.stat(FINAL_PPTX).catch(() => null)) {
  throw new Error(`Refusing to overwrite ${FINAL_PPTX}`);
}

const previewDir = path.join(ROOT, "previews");
await fs.mkdir(previewDir, { recursive: true });
for (const [index, slide] of presentation.slides.items.entries()) {
  const png = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(
    path.join(previewDir, `slide-${String(index + 1).padStart(2, "0")}.png`),
    new Uint8Array(await png.arrayBuffer()),
  );
}
const montage = await presentation.export({ format: "webp", montage: true });
await fs.writeFile(
  path.join(previewDir, "deck-montage.webp"),
  new Uint8Array(await montage.arrayBuffer()),
);

const candidatePath = path.join(ROOT, "candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);

const { finalizePresentation } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href
);
const finalizerDir = path.join(ROOT, "finalizer");
await fs.mkdir(finalizerDir, { recursive: true });
const result = await finalizePresentation({
  workspaceDir: ROOT,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: process.env.RUNTIME_PYTHON,
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
    "--require-native-table-slide",
    "5",
  ],
  explicitTotalSlideCount: 11,
  requiredNativeTableOwnerSlides: [5],
  requiredNativeChartOwnerSlides: [8],
  materializeLiteralChartWorkbooks: true,
  fontPolicy: { basis: "design", families: [FONT] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(finalizerDir, "statistics-hypothesis.validation.json"),
});
console.log(JSON.stringify({ finalPath: FINAL_PPTX, validation: result }, null, 2));
