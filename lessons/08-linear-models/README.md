# 第 8 讲 线性模型 LDA 与 Logistic Regression

课纲编号：7。课堂课件编号优先，完整对应表见根目录 `course/lesson-map.csv`。

研究主题：线性概率模型、共线性、阈值与评价协议。

状态：用户本次确认的作业已归档。基础归档检查入口已在本机运行通过。

## 作业与展示

- [Linear-models-student-worksheet-completed.docx](homework/Linear-models-student-worksheet-completed.docx)

## 代码与对话

- [本次新增的可移植分析入口](src/analyze.py)
- [历史文档与图表生成代码](src/history/README.md)
- [可见 AI 对话](ai-dialogues/conversation.md)，共 18 条消息
- [本次运行结果](outputs/reproduced/results.json)
- [反思与复现边界](reflection.md)
- [报告文档索引](report/README.md)

## 运行

在仓库根目录安装 `requirements-lock.txt` 后运行：

```console
python lessons/08-linear-models/src/analyze.py
```

输出位于 `outputs/reproduced/`，不会覆盖 `homework/`、`outputs/history/` 或原始课件。
`src/history` 是原任务生成代码的存档，不承诺直接运行；部分依赖原本机路径、
Office 渲染环境、Node 工具或教师标签。新增入口不依赖这些条件。

本包 `data/heart.csv` 与 `assets/.../statlog-heart.csv` 并非同一份数据。原 data/README 声称无缺失，但实际发现 557 个缺失单元格。本次入口在训练阶段插补，结果只记为新基线，不冒充原工作纸结果的精确复现。历史工作纸与 Notebook 已暴露测试标签的限制继续保留。

## 证据链

课程输入数据 -> `src/analyze.py` -> 根目录 `tools/run_lesson.py` -> `outputs/reproduced/`。
历史作业来源与源文件哈希见根目录 `provenance/imported-files.json`；
逐文件代码和对话关联见 `provenance/file-index.csv`。
逐文件关联用于导航，不表示某次聊天逐一生成了所有教师课件。
