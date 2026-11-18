# 第 3 至 4 讲 高维数据与降维

课纲编号：3a / 3b。课堂课件编号优先，完整对应表见根目录 `course/lesson-map.csv`。

研究主题：距离集中、PCA、UMAP 与信息泄漏。

状态：已归档历史补充作业。基础归档检查入口已在本机运行通过。

## 作业与展示

- [student-package-class3-completed-worksheet.docx](homework/supplement/student-package-class3-completed-worksheet.docx)
- [student-package-class3-learning-reflection-speaker-notes-v8.docx](homework/supplement/student-package-class3-learning-reflection-speaker-notes-v8.docx)
- [student-package-class3-learning-reflection-v8.pptx](homework/supplement/student-package-class3-learning-reflection-v8.pptx)

## 代码与对话

- [本次新增的可移植分析入口](src/analyze.py)
- [历史文档与图表生成代码](src/history/README.md)
- [可见 AI 对话](ai-dialogues/conversation.md)，共 45 条消息
- [本次运行结果](outputs/reproduced/results.json)
- [反思与复现边界](reflection.md)
- [报告文档索引](report/README.md)

## 运行

在仓库根目录安装 `requirements-lock.txt` 后运行：

```console
python lessons/03-04-high-dimensional-reduction/src/analyze.py
```

输出位于 `outputs/reproduced/`，不会覆盖 `homework/`、`outputs/history/` 或原始课件。
`src/history` 是原任务生成代码的存档，不承诺直接运行；部分依赖原本机路径、
Office 渲染环境、Node 工具或教师标签。新增入口不依赖这些条件。

数据为 120 人、23146 基因，使用固定 80 人开发集作 PCA 和距离集中检查。此入口不重新运行 UMAP，也不重做历史完整预测实验；相应 Notebook 和文档另行保留。NPZ 的对象数组字符串用 pickle 操作码解析，不执行随包 pickle。

## 证据链

课程输入数据 -> `src/analyze.py` -> 根目录 `tools/run_lesson.py` -> `outputs/reproduced/`。
历史作业来源与源文件哈希见根目录 `provenance/imported-files.json`；
逐文件代码和对话关联见 `provenance/file-index.csv`。
逐文件关联用于导航，不表示某次聊天逐一生成了所有教师课件。
