# 生物医学数据科学 PBL 课程仓库

南铂山的课程资料、作业、代码、结果与 AI 协作记录。整理日期：2026-11-18。
本仓库按用户要求公开访问，老师可以直接通过链接查看，无需添加协作者。
公开访问不改变课程课件、教材和数据原有的版权或许可条件，不作额外再授权。

## 课程索引

| 课程主题 | 归档状态 | 历史 AI 消息数 |
| --- | --- | ---: |
| [第 1 讲 课程介绍与 AI 工具](lessons/01-tools/README.md) | 仅课程资料与新增工具演示 | 0 |
| [第 2 讲 变量与数据类型](lessons/02-data-types/README.md) | 用户本次确认的作业已归档 | 11 |
| [第 3 至 4 讲 高维数据与降维](lessons/03-04-high-dimensional-reduction/README.md) | 已归档历史补充作业 | 45 |
| [第 5 讲 统计检验与假设设计](lessons/05-hypothesis/README.md) | 用户本次确认的作业已归档 | 85 |
| [第 6 讲 多重检验校正](lessons/06-multiple-testing/README.md) | 用户本次确认的作业已归档 | 34 |
| [第 7 讲 机器学习基础与朴素贝叶斯](lessons/07-naive-bayes/README.md) | 用户本次确认的作业已归档 | 35 |
| [第 8 讲 线性模型 LDA 与 Logistic Regression](lessons/08-linear-models/README.md) | 用户本次确认的作业已归档 | 18 |
| [第 9 讲 SVM 与 Margin](lessons/09-svm/README.md) | 用户本次确认的作业已归档 | 25 |
| [第 10 讲 集成学习与 Bootstrapping](lessons/10-ensemble/README.md) | 已归档历史补充作业 | 12 |

课纲与课堂编号不完全一致，本仓库按实际课件编号归档，并合并高维与降维主题。
原文件名中的编号不静默改写，见 [课次对应表](course/lesson-map.csv)。
本次用户新提供的 8 份 Word 作业全部保留原字节；SVM 的两个同内容旧版本也保留。

## 文件结构

```text
biomedical-data-science-pbl/
|-- course/                 大纲、评价标准与仓库提交要求
|-- lessons/
|   `-- <编号>-<主题>/
|       |-- README.md       主题、作业、代码与运行入口
|       |-- slides/         教师课件，部分参考阅读保留 readings/
|       |-- pbl/            工作纸与任务材料，教师参考标明来源
|       |-- data/           数据与原始来源说明
|       |-- assets/         课程包原有证据表、图和部分输入数据
|       |-- src/
|       |   |-- analyze.py  本次新增的可移植复查入口
|       |   `-- history/    之前任务的实际生成源码
|       |-- notebook(s)/   原课程 Notebook
|       |-- homework/       用户确认的作业原件及历史补充材料
|       |-- outputs/        reproduced 新结果与 history 历史结果
|       |-- report/         报告与展示文档索引
|       |-- ai-dialogues/   真实可见对话脱敏导出
|       `-- reflection.md   复现限制与下一步
|-- provenance/             逐文件来源、哈希、代码及对话关联
|-- tools/                  统一运行和仓库校验
`-- tests/                  归档与统计校正的回归测试
```

课件包的 `scripts/`、`readings/`、`notebook/` 等内部名称保留以便查对来源。
教师专用标签、缓存、临时渲染和规划草稿不提交；原 ZIP 在用户本机保持不变。

## 复查运行

需要 Python 3.12。在仓库根目录执行：

```console
python -m venv .venv
# Windows
.venv\Scripts\python -m pip install -r requirements-lock.txt
.venv\Scripts\python tools/run_lesson.py all
.venv\Scripts\python tools/validate_archive.py
.venv\Scripts\python -m unittest discover -s tests -v
```

macOS 或 Linux 使用 `.venv/bin/python`。也可以只运行某一课，例如：

```console
python lessons/09-svm/src/analyze.py
```

有意修改代码或重新生成结果后，运行 `python tools/update_file_index.py` 更新文件索引，
再运行归档校验。导入原件的哈希清单不会因此改写，原件变化仍会报错。
Git 不自动转换换行符，确保跨系统克隆后源文件哈希保持一致。

基础分析入口已在本机运行通过；这不等于全部历史作业、PPT、Word 和 Notebook
都已重建。高维入口只检查开发集 PCA 和距离集中，不重跑 UMAP。
线性模型数据存在版本差异，新运行明确标为归档基线，不声称重现旧工作纸的每个数字。
SVM 在当前 scikit-learn 下有来自原教师 helper 的弃用警告，不影响本次运行完成。

## 证据与诚实边界

数据文件 -> 本课 `src/analyze.py` -> `tools/run_lesson.py` -> `outputs/reproduced/`。
旧作业和旧代码由 [来源清单](provenance/imported-files.json) 追溯，
每个文件的课程代码、聊天和哈希关联见 [文件索引](provenance/file-index.csv)。
该关联是导航，不代表每份教师课件都由对应 AI 聊天生成。

对话只包含本课程相关的可见用户消息和助手回复，不导出系统提示、开发者指令、
隐藏推理、工具输出或本机凭据。本机路径、微信标识和疑似密钥已脱敏。
AI 历史回复中的完成宣称不自动成为本次验证结论。

未读取教师专用锁定标签文件，不把开发集结果写成外部验证。
部分公开原始数据和旧材料本来包含标签，不能据此声称历史上从未接触测试标签。
没有伪造同伴复现记录；[同伴复现模板](provenance/peer-reproduction.md) 尚待真人填写。

## 版本与许可

本次提交历史记录真实的归档、运行入口和校验阶段，不回填或伪造过去的 commit。
课程课件与教材版权归原作者，数据许可见各课程原始说明，未统一加 MIT 等授权。
本仓库已公开，可直接分享仓库链接；引用和再分发材料时仍应遵守原始许可条件。
