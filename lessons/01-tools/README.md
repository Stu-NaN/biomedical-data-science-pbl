# 第 1 讲 课程介绍与 AI 工具

课纲编号：1。课堂课件编号优先，完整对应表见根目录 `course/lesson-map.csv`。

研究主题：工具环境、版本管理与 AI 使用边界。

状态：仅课程资料与新增工具演示。基础归档检查入口已在本机运行通过。

## 作业与展示

未提供本课正式作业；`outputs/reproduced` 仅为本次整理时新增的工具演示。

## 代码与对话

- [本次新增的可移植分析入口](src/analyze.py)
- [历史文档与图表生成代码](src/history/README.md)
- [可见 AI 对话](ai-dialogues/conversation.md)，共 0 条消息
- [本次运行结果](outputs/reproduced/results.json)
- [反思与复现边界](reflection.md)
- [报告文档索引](report/README.md)

## 运行

在仓库根目录安装 `requirements-lock.txt` 后运行：

```console
python lessons/01-tools/src/analyze.py
```

输出位于 `outputs/reproduced/`，不会覆盖 `homework/`、`outputs/history/` 或原始课件。
`src/history` 是原任务生成代码的存档，不承诺直接运行；部分依赖原本机路径、
Office 渲染环境、Node 工具或教师标签。新增入口不依赖这些条件。

## 证据链

课程输入数据 -> `src/analyze.py` -> 根目录 `tools/run_lesson.py` -> `outputs/reproduced/`。
历史作业来源与源文件哈希见根目录 `provenance/imported-files.json`；
逐文件代码和对话关联见 `provenance/file-index.csv`。
逐文件关联用于导航，不表示某次聊天逐一生成了所有教师课件。
