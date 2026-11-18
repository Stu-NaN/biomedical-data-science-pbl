# 第 9 讲 SVM 与 Margin

课纲编号：8。课堂课件编号优先，完整对应表见根目录 `course/lesson-map.csv`。

研究主题：间隔、支持向量、C、核函数与尺度。

状态：用户本次确认的作业已归档。基础归档检查入口已在本机运行通过。

## 作业与展示

- [SVM-student-worksheet-completed(1).docx](homework/SVM-student-worksheet-completed%281%29.docx)
- [SVM-student-worksheet-completed.docx](homework/SVM-student-worksheet-completed.docx)
- [第9讲作业_SVM与Margin_南铂山.docx](homework/%E7%AC%AC9%E8%AE%B2%E4%BD%9C%E4%B8%9A_SVM%E4%B8%8EMargin_%E5%8D%97%E9%93%82%E5%B1%B1.docx)

## 代码与对话

- [本次新增的可移植分析入口](src/analyze.py)
- [历史文档与图表生成代码](src/history/README.md)
- [可见 AI 对话](ai-dialogues/conversation.md)，共 25 条消息
- [本次运行结果](outputs/reproduced/results.json)
- [反思与复现边界](reflection.md)
- [报告文档索引](report/README.md)

## 运行

在仓库根目录安装 `requirements-lock.txt` 后运行：

```console
python lessons/09-svm/src/analyze.py
```

输出位于 `outputs/reproduced/`，不会覆盖 `homework/`、`outputs/history/` 或原始课件。
`src/history` 是原任务生成代码的存档，不承诺直接运行；部分依赖原本机路径、
Office 渲染环境、Node 工具或教师标签。新增入口不依赖这些条件。

正式主版本为 `homework/第9讲作业_SVM与Margin_南铂山.docx`。另两份 `SVM-student-worksheet-completed*.docx` 是较早版本，二者 SHA-256 相同，为保留用户提供的文件清单而都保留。新增运行入口仅使用训练和验证数据。

## 证据链

课程输入数据 -> `src/analyze.py` -> 根目录 `tools/run_lesson.py` -> `outputs/reproduced/`。
历史作业来源与源文件哈希见根目录 `provenance/imported-files.json`；
逐文件代码和对话关联见 `provenance/file-index.csv`。
逐文件关联用于导航，不表示某次聊天逐一生成了所有教师课件。
