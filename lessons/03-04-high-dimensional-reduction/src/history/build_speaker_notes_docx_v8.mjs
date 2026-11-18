import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const runtimeDocx = "C:/Users/A/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/docx/dist/index.mjs";
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  PageOrientation,
  Packer,
  Paragraph,
  SectionType,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} = await import(pathToFileURL(runtimeDocx).href);

const WORKSPACE_DIR = "C:/Users/A/ppt_high_dimension_reflection";
const BUILD_DIR = path.join(WORKSPACE_DIR, ".codex-build");
const NOTES_PATH = path.join(BUILD_DIR, "speaker-notes.json");
const OUTPUT_PATH = "C:/Users/A/Downloads/student-package-class3-learning-reflection-speaker-notes-v8.docx";

const notes = JSON.parse(await fs.readFile(NOTES_PATH, "utf8"));

const FONT = "Microsoft YaHei";
const COLORS = {
  ink: "17363A",
  teal: "218E84",
  coral: "E27667",
  muted: "6C7C7C",
  line: "D9D9D9",
  pale: "F3F7F5",
  paleCoral: "F9ECE8",
  paleGold: "FBF6E8",
};

function run(text, options = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: options.size ?? 22,
    bold: options.bold ?? false,
    italics: options.italics ?? false,
    color: options.color ?? "000000",
  });
}

function para(children, options = {}) {
  return new Paragraph({
    children,
    alignment: options.alignment ?? AlignmentType.LEFT,
    style: options.style,
    spacing: options.spacing ?? { after: 140, line: 300 },
    keepNext: options.keepNext ?? false,
    pageBreakBefore: options.pageBreakBefore ?? false,
  });
}

function textPara(text, options = {}) {
  return para([run(text, options)], options);
}

function heading(text, level = HeadingLevel.HEADING_1, pageBreakBefore = false) {
  return new Paragraph({
    text,
    heading: level,
    pageBreakBefore,
    keepNext: true,
    spacing: { before: level === HeadingLevel.HEADING_1 ? 380 : 240, after: 150, line: 280 },
  });
}

function labelParagraph(label, text) {
  return para([run(label, { bold: true, color: COLORS.teal }), run(text)]);
}

function numberedItem(text) {
  return new Paragraph({
    numbering: { reference: "main-numbering", level: 0 },
    children: [run(text)],
    spacing: { after: 100, line: 290 },
  });
}

function simpleTable(rows) {
  const widths = [2300, 4300, 2760];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, rowIndex) => new TableRow({
      children: row.map((cellText, colIndex) => new TableCell({
        width: { size: widths[colIndex], type: WidthType.DXA },
        shading: {
          fill: rowIndex === 0 ? COLORS.ink : (rowIndex % 2 === 0 ? COLORS.pale : "FFFFFF"),
          type: ShadingType.CLEAR,
        },
        margins: { top: 110, bottom: 110, left: 140, right: 140 },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
          left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
          right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
        },
        children: [
          para([run(cellText, {
            size: 20,
            bold: rowIndex === 0 || colIndex === 0,
            color: rowIndex === 0 ? "FFFFFF" : "000000",
          })], { spacing: { after: 0, line: 260 } }),
        ],
      })),
    })),
  });
}

const slideTitles = [
  "从读讲义到有边界的判断",
  "我先把学习任务拆成一条证据链",
  "课程重点决定了我真正要回答的问题",
  "先把数据对象说清楚，很多误解会自动消失",
  "Q2 距离：维度越高，绝对距离变大，近远差别反而变小",
  "Q3 见过不等于学会：训练高分必须和未见患者一起看",
  "Q4 PCA：压缩的是方差，不是生物学真相",
  "Q5 UMAP：我把它当作探索图，而不是结论机器",
  "我的最终判断：A 做主展示，PCA 与 UMAP 各司其职",
  "AI 的位置：加速理解和检查，不替代研究判断",
  "我学会的是：先定义问题，再相信结果",
];

const slidePurposes = [
  "交代这次汇报的主线：我如何按课程重点学习，再把概念放回高维表达数据中验证。",
  "说明学习顺序和证据链，强调先读大纲与讲义，再做数据实践。",
  "把课程重点整理成数据对象、维度灾难、PCA、UMAP 和泛化五个问题。",
  "定义矩阵、患者、基因、标签和开发集/未见患者组，说明 p/n 约 192.9 的意义。",
  "用相对对比度和有效维数解释高维距离为什么不能只看绝对值。",
  "比较见过患者与未见患者表现，说明训练高分不能替代泛化验证。",
  "区分 PCA 的 loading、score、解释方差和重建误差，避免把数学方向说成机制。",
  "用近邻保留率、参数敏感性和批次竞争解释限定 UMAP 的用途。",
  "给出方法选择：A 为主展示，PCA 做压缩辅助，UMAP 做探索辅助。",
  "明确 AI 的辅助边界，展示哪些判断必须回到材料、表格和数据。",
  "总结这次作业带来的方法变化，以及后续验证和停止解释的条件。",
];

const cleanedNotes = Object.fromEntries(
  Object.entries(notes).map(([key, value]) => [
    key,
    value.replace(/视觉素材：.*$/, "").replace(/背景图：.*$/, "").trim(),
  ]),
);

const children = [];

children.push(
  new Paragraph({
    text: "高维数据学习复盘讲稿",
    style: "Title",
    alignment: AlignmentType.LEFT,
    spacing: { after: 120, line: 300 },
  }),
  textPara("课程主题：高维表达数据、PCA 与 UMAP 的学习过程复盘", { size: 26, color: COLORS.teal, bold: true }),
  textPara("配套文件：student-package-class3-learning-reflection-v8.pptx", { size: 20, color: COLORS.muted }),
  textPara("使用方式：这份讲稿不是 worksheet 的逐格朗读，而是按课程大纲组织的一次学习与分析复盘。汇报时可以按页讲，也可以把每页的“本页作用”和“讲稿”压缩成自己的口语。", { size: 22 }),

  heading("一 我这次具体做了什么"),
  numberedItem("先回到 plan.docx，确认课程要求我学习的重点不是某一个算法，而是数据对象、高维困难、降维、过拟合和验证边界。"),
  numberedItem("再读 readings 文件夹里的讲义，把距离、标准化、PCA、UMAP、有效维数和泛化这些概念弄清楚，尤其区分“图看起来有结构”和“证据足以支持结论”。"),
  numberedItem("最后进入 worksheet 的数据实践，用固定的 120 位患者和 23,146 个基因检验这些概念，把 worksheet 的问题重新组织成一条可以追踪的论证链。"),
  textPara("这次作业的重点不在于把每个格子都填满，而在于把每个数字放回它适用的概念和边界中。"),

  heading("二 核心数据与最终判断"),
  simpleTable([
    ["对象", "本次分析使用的事实", "我如何使用它"],
    ["数据规模", "120 位患者，23,146 个基因，p/n 约 192.9", "把它作为小样本高维问题处理"],
    ["泛化观察", "开发集 80 人，未见患者 40 人；全部基因未见患者 34/40 = 85.0%", "把见过患者与未见患者表现分开"],
    ["距离与降维", "相对对比度从 2.3028 降到 0.8544；标准化 PCA 二维累计方差 22.76%", "报告结构变化，同时写出压缩损失"],
    ["UMAP", "基准全体近邻保留率 40.25%", "作为探索图，不直接命名新亚型"],
  ]),
  textPara("我的最终判断是：用标准化原始特征的距离、相关性和近邻稳定性做主要展示；PCA 负责压缩和解释方差；UMAP 负责局部结构探索。任何预测性结论都要等新患者、新中心、批次信息和外部验证。"),

  heading("三 逐页讲稿"),
);

for (let i = 1; i <= 11; i += 1) {
  children.push(
    heading(`第 ${String(i).padStart(2, "0")} 页  ${slideTitles[i - 1]}`, HeadingLevel.HEADING_2, i > 1),
    labelParagraph("本页作用：", slidePurposes[i - 1]),
    labelParagraph("讲稿：", cleanedNotes[String(i)]),
  );
}

children.push(
  heading("四 课堂追问准备"),
  labelParagraph("为什么没有直接选 UMAP？", "因为 UMAP 改变原空间中的邻域关系，基准全体近邻保留率只有 40.25%。它很适合提出结构问题，但不能独立证明新亚型或批次因果。"),
  labelParagraph("85.0% 的未见患者准确率能不能说明模型很好？", "不能。它说明在这一次预先固定的 40 人观察中，34 人被正确分类。还需要外部队列、独立中心、完整批次信息、校准和更严格的交叉验证。"),
  labelParagraph("PCA 二维只解释 22.76% 是不是说明 PCA 没有用？", "不是。PCA 仍然可以帮助压缩、展示线性方向和比较重建损失，只是二维图不能被当成原空间的无损替代。"),
  labelParagraph("为什么 p/n 约 192.9 会改变分析策略？", "因为特征数远高于患者数，模型有更大的自由度去记忆训练样本，距离也会受相关特征和尺度影响，所以必须优先检查标准化、近邻稳定性和未见患者表现。"),
  labelParagraph("AI 在这次作业里具体做了什么？", "AI 帮我整理学习路径、解释概念、检查表达和协助排版，但我自己回到表格核对了 22.76%、40.25% 和 85.0%，并自己决定哪些结论应该停止在探索层面。"),

  heading("五 结束语"),
  textPara("这次作业让我真正学会的不是“选一个最强的降维算法”，而是先定义数据对象，再判断证据用途，最后给结论加边界。对我来说，AI 的价值是减少查找和改写成本，但研究判断仍然需要由我自己承担。"),
  textPara("汇报结束时可以用一句话收束：在高维数据里，真正重要的不是图能不能把样本分开，而是我能不能解释这张图保留了什么、丢失了什么，以及它还需要什么证据才能被相信。"),

  heading("六 视觉素材说明"),
  textPara("整套 PPT 使用一张蓝色数字网络背景图铺满页面，来源为 Unsplash photo-1644088379091-d574269d422f。图片只作为科技感的背景视觉，不作为本次数据分析的证据。"),
);

const doc = new Document({
  creator: "Student",
  title: "高维数据学习复盘讲稿",
  subject: "Course 03 high-dimensional data learning reflection speaker notes",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 22, color: "000000" },
        paragraph: { spacing: { after: 140, line: 300 } },
      },
    },
    paragraphStyles: [
      {
        id: "Title",
        name: "Title",
        basedOn: "Normal",
        next: "Normal",
        run: { font: FONT, size: 38, bold: true, color: "000000" },
        paragraph: { spacing: { before: 0, after: 160, line: 320 } },
      },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        run: { font: FONT, size: 30, bold: true, color: "000000" },
        paragraph: { spacing: { before: 380, after: 160, line: 300 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        run: { font: FONT, size: 25, bold: true, color: COLORS.ink },
        paragraph: { spacing: { before: 260, after: 130, line: 290 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "main-numbering",
        levels: [{
          level: 0,
          format: "decimal",
          text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      type: SectionType.CONTINUOUS,
      page: {
        size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
        margin: { top: 1080, right: 1260, bottom: 1080, left: 1260 },
      },
    },
    headers: {},
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { before: 80, after: 0 },
            children: [
              run("Course 03 学习复盘讲稿", { size: 16, color: COLORS.muted }),
            ],
          }),
        ],
      }),
    },
    children,
  }],
});

const buffer = await Packer.toBuffer(doc);
await fs.writeFile(OUTPUT_PATH, buffer);
console.log(JSON.stringify({ output: OUTPUT_PATH, bytes: buffer.length, slideNoteCount: Object.keys(notes).length }));
