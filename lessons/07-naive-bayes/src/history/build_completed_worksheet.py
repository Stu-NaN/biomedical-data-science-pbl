from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(r"C:\Users\A\Downloads\Naive-Bayes-work\7. Naive-Bayes")
SOURCE = ROOT / "student-worksheet.md"
OUTPUT = Path(r"C:\Users\A\Downloads\Naive-Bayes-student-worksheet-completed.docx")


REPLACEMENTS = {
    "# 第 7 讲 机器学习基础与朴素贝叶斯 学生工作单":
        "# 第 7 讲 机器学习基础与朴素贝叶斯 学生工作单 完成版",
    "姓名：________　学号：________　日期：________":
        "姓名：________　学号：________　日期：2026-10-28",
    "我的办法：________________________________________________________":
        "我的办法：**作答：**先把训练记录按六项症状找相似病例；若有完全相同的组合，就统计其中膀胱炎症阳性的比例作为预测概率，再按0.5阈值给出类别。",
    "它需要从已有记录中统计什么？______________________________________":
        "它需要从已有记录中统计什么？**作答：**每种症状组合出现的次数，以及这些记录中阳性和阴性的数量。",
    "如果找不到完全相同的症状组合，我的办法会怎样处理？__________________":
        "如果找不到完全相同的症状组合，我的办法会怎样处理？**作答：**改用与输入最接近的训练记录，或分别统计每项症状在两类中的出现频率；不能把“没有完全相同记录”直接当作阴性。",
    "一条记录（sample）是____；一项特征（feature）是____；本次标签（label）是____。":
        "一条记录（sample）是**一行专家设计的病例描述**；一项特征（feature）是**六项二元症状中的一项**；本次标签（label）是**膀胱炎症是否为阳性（0/1）**。",
    "模型输入有____个数；标签能否同时作为输入？____，理由是________________。":
        "模型输入有**6**个数；标签能否同时作为输入？**不能**，理由是**标签是要预测的答案，作为输入会造成目标泄漏，模型等于提前看到了答案**。",
    "- `p=0.8` 与 `ŷ=1` 表达的内容有什么不同？____________________________":
        "- `p=0.8` 与 `ŷ=1` 表达的内容有什么不同？**作答：**`p=0.8`是模型给出的阳性概率；`ŷ=1`只是按阈值把概率转成阳性类别。概率不等于已经验证过的个体患病风险。",
    "- 如果只列出症状清单，却没有说明症状与疾病如何联系，是否已经定义了模型？还缺什么？________________________________________________":
        "- 如果只列出症状清单，却没有说明症状与疾病如何联系，是否已经定义了模型？还缺什么？**作答：**还没有；还需要规定输入如何映射到输出的规则或函数、由训练数据估计的参数，以及把概率转成类别的阈值。",
    "- 用自己的话解释“数据模型”，不只写一个算法名称：__________________":
        "- 用自己的话解释“数据模型”，不只写一个算法名称：**作答：**数据模型是一套从输入到输出的明确关系。本题先用训练记录估计疾病先验和症状概率，再把六项症状变成阳性概率，最后按阈值输出类别。",
    "**小结**：画出你的“输入 → 关系或规则 → 输出”，标出哪些部分需要由数据确定。":
        "**小结：**六项症状 `x` → 训练数据估计先验与类内症状概率（Bernoulli NB）→ 阳性概率 `p` → 阈值 `t=0.5` → 类别 `ŷ`。先验和症状概率由训练数据确定；阈值由验证集选择。",
    "训练得到：`P̂(y=1|x=0)=____/____=____`；`P̂(y=1|x=1)=____/____=____`。":
        "训练得到：`P̂(y=1|x=0)=**1/3**=**0.333**`；`P̂(y=1|x=1)=**4/5**=**0.800**`。",
    "对于新输入 `x=1`，概率预测为____；若 `p≥0.5` 判为阳性，类别为____。预测时需要提前知道新输入的 y 吗？____。":
        "对于新输入 `x=1`，概率预测为**0.800**；若 `p≥0.5` 判为阳性，类别为**1（阳性）**。预测时需要提前知道新输入的 y 吗？**不需要**，真实标签只在事后评价时使用。",
    "若再加入一条 `(x=1,y=0)` 重新估计，概率变为____。若只是把新输入送入**固定的模型**，概率表会不会自动改变？____。":
        "若再加入一条 `(x=1,y=0)` 重新估计，概率变为**4/6=0.667**。若只是把新输入送入**固定的模型**，概率表会不会自动改变？**不会**；只有重新训练才会更新参数。",
    "- **训练**时已知什么，改变什么？__________________________________":
        "- **训练**时已知什么，改变什么？**作答：**已知训练记录的症状 `x` 和标签 `y`，据此估计阳性频率等模型参数。",
    "- **预测**时输入什么，哪些量保持固定？______________________________":
        "- **预测**时输入什么，哪些量保持固定？**作答：**输入新记录的症状 `x`；已训练的参数和当前阈值保持固定，模型输出概率及类别。",
    "- 模型在这 8 条训练记录上表现好，能否说明它能预测其他记录？你需要怎样检验？________________________________________________________":
        "- 模型在这 8 条训练记录上表现好，能否说明它能预测其他记录？你需要怎样检验？**作答：**不能。训练表现只说明模型适配了已见记录；应在未参与估参和选择、且符合目标使用场景的留出数据上评价，并说明样本来源限制。",
    "用自己的话解释**泛化**：__________________________________________":
        "用自己的话解释**泛化**：**模型参数固定后，面对未参与训练或模型选择的新记录，仍能保持有用预测表现的能力。**",
    "**小结**：在“数据 → 参数 → 新输入的预测 → 未见记录上的误差”中，标出训练、预测和泛化评价的位置。":
        "**小结：**`带标签训练数据 → 训练估计参数`是训练；`固定参数+新症状 → 概率/类别`是预测；把预测与未见记录的真实标签比较并汇总误差，才是在评价泛化。",
    "- 用训练记录估计参数，再用同一批记录选阈值并报告成绩，会有什么问题？________________________________________________________":
        "- 用训练记录估计参数，再用同一批记录选阈值并报告成绩，会有什么问题？**作答：**模型和阈值都根据这批记录调整过，报告成绩会偏乐观，不能代表对新记录的表现。",
    "- 若先看测试结果，再挑成绩最好的阈值，测试数据承担了什么职责？______":
        "- 若先看测试结果，再挑成绩最好的阈值，测试数据承担了什么职责？**它已经参与模型选择，变成验证/开发数据，不再是独立的最终测试。**",
    "- 怎样分配数据，既能作选择，又保留一次未参与选择的评价？____________":
        "- 怎样分配数据，既能作选择，又保留一次未参与选择的评价？**训练集估计参数，验证集选阈值，全部选择固定后只用测试集作最终评价。**",
    "| 估计类别比例和类内症状概率 | ____ | ____ |":
        "| 估计类别比例和类内症状概率 | **train** | 用训练标签估计先验和症状条件概率。 |",
    "| 比较阈值，决定采用哪一个 | ____ | ____ |":
        "| 比较阈值，决定采用哪一个 | **validation** | 用验证表现按事先规则选择，不能用测试集挑选。 |",
    "| 固定全部选择后报告留出表现 | ____ | ____ |":
        "| 固定全部选择后报告留出表现 | **test** | 配置已冻结后作最终留出评价，不再据此修改模型。 |",
    "| 看完 test 后再换输入特征 | 这会怎样改变 test 的用途？____ | ____ |":
        "| 看完 test 后再换输入特征 | **test已参与开发选择，不再是独立最终评价。** | 之后若要报告无偏的最终表现，需要另一份此前未查看的新测试数据。 |",
    "从 Notebook A 记录实际划分：train____条，validation____条；test 预设为 24 条，暂不读取其标签。训练与验证的 ID 是否相交？____。":
        "从 Notebook A 记录实际划分：train**72**条，validation**24**条；test 预设为24条，暂不读取其标签。训练与验证的 ID 是否相交？**不相交**。",
    "validation 有____条记录的六输入组合在 train 出现过。“新的一行”“未见过的症状组合”和“另一家医院的患者”是同一种泛化问题吗？举例说明：______":
        "validation 有**24**条记录的六输入组合在 train 出现过。“新的一行”“未见过的症状组合”和“另一家医院的患者”是同一种泛化问题吗？举例说明：**不是。新行可能只是已见组合的另一条记录；新组合要求预测训练中未出现的症状型；新医院还可能带来人群、采集流程或测量分布变化。**",
    "**小结**：画出数据流程，标出“允许估参”“允许选择”“只能最终评价”的位置。后面照此执行。":
        "**小结：**`train（72条）→估计参数`；`validation（24条）→按预先规则选择阈值`；`冻结方案→test（24条）最终评价`。三份记录的ID不重叠，但相同症状组合可以跨集合出现。",
    "`P̂(Y=1)=____`；`P̂(x=1|Y=1)=____`；`P̂(x=1|Y=0)=____`。":
        "`P̂(Y=1)=**5/8**`；`P̂(x=1|Y=1)=**4/5**`；`P̂(x=1|Y=0)=**1/3**`。",
    "`P̂(Y=1|x=1) = P̂(x=1|Y=1)P̂(Y=1) / [P̂(x=1|Y=1)P̂(Y=1) + P̂(x=1|Y=0)P̂(Y=0)] = ____`。":
        "`P̂(Y=1|x=1) = P̂(x=1|Y=1)P̂(Y=1) / [P̂(x=1|Y=1)P̂(Y=1) + P̂(x=1|Y=0)P̂(Y=0)] = **(4/5×5/8)/[(4/5×5/8)+(1/3×3/8)]=0.5/0.625=0.800**`。",
    "它与任务 2 直接计算的阳性比例是否相同？____。为什么 `P(x=1|Y=1)` 不能直接当作 `P(Y=1|x=1)`？____________________________________________":
        "它与任务2直接计算的阳性比例是否相同？**相同，都是0.800**。为什么 `P(x=1|Y=1)` 不能直接当作 `P(Y=1|x=1)`？**两者条件方向不同、参照的记录集合和分母也不同，必须结合先验和 `P(x=1|Y=0)` 用 Bayes 定理转换。**",
    "若直接统计 `P(X=x|Y=c)`，六个二元症状共有____种可能组合。训练数据可能没有覆盖所有组合，怎样用每个症状的计数近似联合概率？":
        "若直接统计 `P(X=x|Y=c)`，六个二元症状共有**2⁶=64**种可能组合。训练数据可能没有覆盖所有组合，怎样用每个症状的计数近似联合概率？**在给定类别后假设特征条件独立，把六项单症状条件概率相乘近似联合似然。**",
    "- 若 `xj=1`，乘____；若 `xj=0`，乘____。“无症状”为何也有用？________":
        "- 若 `xj=1`，乘**θcj=P(Xj=1|Y=c)**；若 `xj=0`，乘**1−θcj=P(Xj=0|Y=c)**。“无症状”为何也有用？**它同样提供与类别相符或不符的似然证据。**",
    "- 假设的是所有记录中的症状互相独立，还是同一类别内部独立？__________":
        "- 假设的是所有记录中的症状互相独立，还是同一类别内部独立？**同一类别内部的条件独立，即给定Y后作独立近似。**",
    "记录 `n0=____，n1=____`。对固定输入 `x=(0,0,0,1,1,1)` 填表；输入为 0 时，似然项必须取补概率。":
        "记录 `n0=**37**，n1=**35**`。对固定输入 `x=(0,0,0,1,1,1)` 填表；输入为0时，似然项必须取补概率。",
    "| fever_ge_38 | 0 | ____ | ____ | ____ | ____ |":
        "| 发热（≥38℃） | 0 | 25 | 13 | 13/39 = 0.333333 | 23/37 = 0.621622 |",
    "| nausea | 0 | ____ | ____ | ____ | ____ |":
        "| 恶心 | 0 | 5 | 13 | 33/39 = 0.846154 | 23/37 = 0.621622 |",
    "| lumbar_pain | 0 | ____ | ____ | ____ | ____ |":
        "| 腰痛 | 0 | 30 | 13 | 8/39 = 0.205128 | 23/37 = 0.621622 |",
    "| urine_pushing | 1 | ____ | ____ | ____ | ____ |":
        "| 尿急 | 1 | 13 | 35 | 14/39 = 0.358974 | 36/37 = 0.972973 |",
    "| micturition_pain | 1 | ____ | ____ | ____ | ____ |":
        "| 排尿痛 | 1 | 5 | 28 | 6/39 = 0.153846 | 29/37 = 0.783784 |",
    "| burning_urethra | 1 | ____ | ____ | ____ | ____ |":
        "| 尿道灼痛 | 1 | 13 | 17 | 14/39 = 0.358974 | 18/37 = 0.486486 |",
    "`π0=____，π1=____`；`w0=π0×六项乘积=____`；`w1=π1×六项乘积=____`。":
        "`π0=**37/72=0.513889**，π1=**35/72=0.486111**`；`w0=π0×六项乘积=**0.0005894354**`；`w1=π1×六项乘积=**0.0433193284**`。",
    "`p=P̂(Y=1|x)=w1/(w0+w1)=____`。阈值为 0.5 时，预测类别____。":
        "`p=P̂(Y=1|x)=w1/(w0+w1)=**0.9865759063**`。阈值为0.5时，预测类别**1（阳性）**。",
    "完成后运行 Notebook A2 核对：程序概率____，与手算差____。若不同，定位到计数、补概率、先验或归一化的哪一步？________________________________":
        "完成后运行 Notebook A2 核对：程序概率**0.986576**，与手算差**小于0.000001（四舍五入差异）**。若不同，定位到计数、补概率、先验或归一化的哪一步？**依次检查特征顺序和类内计数、0取补概率、先验分母，以及最后是否除以`w0+w1`。**",
    "为什么零计数可能让整个乘积变成零？平滑如何改变这种情况？____________":
        "为什么零计数可能让整个乘积变成零？平滑如何改变这种情况？**任一条件概率为0会使相乘后的类别似然为0；α=1给每个二元结果各加一个伪计数，使估计概率不落在0或1，从而避免单个零计数抹掉整类概率，但不会消除特征相关或抽样偏差。**",
    "先想：复制一次腰痛列，没有增加新测量，模型概率会不会变化？为什么？____":
        "先想：复制一次腰痛列，没有增加新测量，模型概率会不会变化？**会。NB会把复制列当成额外的独立证据，再乘一次同一项似然，因此后验概率可能改变。**",
    "运行 Notebook C，选一个记录比较不复制与额外复制五次：ID____，原概率____，复制后____，类别是否改变____。哪项似然被重复使用了？____。这说明条件独立假设怎样影响预测？______________________________________________":
        "运行 Notebook C，选一个记录比较不复制与额外复制五次：ID：**A081**，原概率**0.010210**，复制后**0.724999**，类别是否改变**是，从0变为1**。哪项似然被重复使用了？**腰痛=0对应的`P(X_lumbar=0|Y=c)=1−θc`**。这说明条件独立假设怎样影响预测？**模型把同一测量的五个副本误当成新且独立的信息，重复放大其证据，后验可大幅偏移甚至跨过分类阈值。**",
    "**小结**：按“训练计数 → 参数 → 类别权重 → 归一化概率 → 类别”解释一次完整预测。":
        "**小结：**训练集计数经加一平滑得到先验`πc`和症状概率`θcj`；按输入0/1取相应似然并相乘得到各类权重`wc`；用`w1/(w0+w1)`归一化为后验，再按0.5阈值输出类别。本例后验为0.986576，预测阳性。",
    "| a | 1 | 0.6 | 0.9 | ____ | ____ |":
        "| a | 1 | 0.6 | 0.9 | 1 | 1 |",
    "| b | 0 | 0.4 | 0.1 | ____ | ____ |":
        "| b | 0 | 0.4 | 0.1 | 0 | 0 |",
    "两者 accuracy 各为____、____。两组概率预测一样好吗？先解释理由：________":
        "两者accuracy各为**1.000**、**1.000**。两组概率预测一样好吗？先解释理由：**不一样。两者类别都正确，但模型B给真实类别的概率更高（0.9而非0.6），因此概率质量更好。**",
    "计算 `LA=____`、`LB=____`。数值越大还是越小越好？____。它揭示了 accuracy 没显示的什么差别？________________________________________________":
        "计算 `LA=**0.510826**`、`LB=**0.105361**`。数值越大还是越小越好？**越小越好**。它揭示了accuracy没显示的什么差别？**模型给真实类别分配的概率有多高；accuracy只保留阈值后的对错，不体现置信程度。**",
    "再比较真实 `y=0` 的一个样本：预测 `p=0.6` 与 `p=0.99` 都分错，两者损失分别为____、____。为什么“非常自信地犯错”受到更大惩罚？________________":
        "再比较真实 `y=0` 的一个样本：预测 `p=0.6` 与 `p=0.99` 都分错，两者损失分别为**−ln(0.4)=0.916291**、**−ln(0.01)=4.605170**。为什么“非常自信地犯错”受到更大惩罚？**因为真实类别是0，模型给它的概率分别只有0.4和0.01；log loss会强烈惩罚把真实结果判得极不可能的预测。**",
    "- 任务 4 真正从训练记录学到哪些参数？______________________________":
        "- 任务4真正从训练记录学到哪些参数？**阳性先验`πc`以及每一类中六项症状为1的条件概率`θcj`；本课用α=1的加一平滑估计。**",
    "- 能否用 validation 标签一起估参，再把 validation 当未见数据评价？为什么？__________________________________________________________":
        "- 能否用validation标签一起估参，再把validation当未见数据评价？为什么？**不能。标签一旦参与参数估计，validation就不再是未见数据；它可用于比较/选择，但不能再声称是独立评价。**",
    "| NB train | ____ | ____ |":
        "| NB train | 1.000 | 0.098840 |",
    "| NB validation | ____ | ____ |":
        "| NB validation | 1.000 | 0.087214 |",
    "| Prior-only validation | ____ | ____ |":
        "| Prior-only validation | 0.500 | 0.693533 |",
    "用 Notebook B 的 **validation** 输出填表。若 precision 分母为零，写“未定义”；程序的 NaN 不等于观察到零精确率。":
        "用Notebook B的**validation**输出填表。若precision分母为零，写“未定义”；程序的NaN不等于观察到零精确率。表中sens.=sensitivity，spec.=specificity。",
    "Prior-only 对所有输入都输出训练集阳性比例。NB 是否比它更有用？从类别与概率两个角度回答；不预设本次一定存在很大的训练／验证差距：____________":
        "Prior-only对所有输入都输出训练集阳性比例。NB是否比它更有用？从类别与概率两个角度回答；不预设本次一定存在很大的训练／验证差距：**在这24条validation记录上，NB分类准确率为1.000，高于先验基线的0.500；平均log loss为0.087214，低于0.693533，说明本次既改善类别判断，也给真实标签分配了更合适的概率。但这只是小型同源留出结果，不能据此保证其他来源也有同样表现。**",
    "**小结**：说明“单样本损失”“平均损失”“按阈值计算的错误数”的联系和区别。":
        "**小结：**单样本loss衡量一条预测与真实标签的差距；平均loss汇总一批记录的概率质量；按阈值得到的混淆矩阵则统计分类错误及其类型。三者评价对象不同，不能互相替代。",
    "阈值 0.5 时：TP 对应____，FN 对应____，FP 对应____，TN 对应____。":
        "阈值0.5时：TP对应**a、b**，FN对应**c、d**，FP对应**e**，TN对应**f、g、h**。混淆矩阵（行=真实0/1，列=预测0/1）：`[[TN=3, FP=1], [FN=2, TP=2]]`。",
    "- accuracy=`(TP+TN)/n`=____；sensitivity/recall=`TP/(TP+FN)`=____。":
        "- accuracy=`(TP+TN)/n`=**5/8=0.625**；sensitivity/recall=`TP/(TP+FN)`=**2/4=0.500**。",
    "- specificity=`TN/(TN+FP)`=____；precision=`TP/(TP+FP)`=____。":
        "- specificity=`TN/(TN+FP)`=**3/4=0.750**；precision=`TP/(TP+FP)`=**2/3=0.667**。",
    "- F1=`2TP/(2TP+FP+FN)`=____。":
        "- F1=`2TP/(2TP+FP+FN)`=**4/7=0.571**。",
    "“真实阳性中检出多少”和“预测阳性中多少是真的”分别是哪项指标？圈出各自分母对应的记录：__________________________________________________":
        "“真实阳性中检出多少”和“预测阳性中多少是真的”分别是哪项指标？圈出各自分母对应的记录：**前者是sensitivity/recall，分母为真实阳性a、b、c、d；后者是precision，分母为预测阳性a、b、e。**",
    "先预测：阈值从 0.5 降到 0.3，漏诊与误报各会怎样变化？________________":
        "先预测：阈值从0.5降到0.3，漏诊与误报各会怎样变化？**阳性判定会更宽松；本例漏诊从2条降到1条，误报从1条升到3条。一般而言FN不会增加、FP不会减少，但具体变化取决于概率分布。**",
    "手算阈值 0.3：TP____、FN____、FP____、TN____；sensitivity____、specificity____。概率有没有变？____；为何类别和指标会变？________________________":
        "手算阈值0.3：TP**3**、FN**1**、FP**3**、TN**1**；sensitivity**3/4=0.750**、specificity**1/4=0.250**。概率有没有变？**没有**；为何类别和指标会变？**阈值降低后，更多既有概率跨过判阳性界线，类别改变，混淆矩阵随之改变。**",
    "| Prior-only | ____ | ____ | ____ | ____ | ____ | ____ | ____ | ____ | ____ |":
        "| Prior-only | 12 | 0 | 12 | 0 | 0.500 | 0 | 1.000 | 未定义 | 0 |",
    "| Bernoulli NB | ____ | ____ | ____ | ____ | ____ | ____ | ____ | ____ | ____ |":
        "| Bernoulli NB | 12 | 0 | 0 | 12 | 1.000 | 1.000 | 1.000 | 1.000 | 1.000 |",
    "| 模型（阈值0.5） | TN | FP | FN | TP | accuracy | sensitivity | specificity | precision | F1 |":
        "| 模型（阈值0.5） | TN | FP | FN | TP | accuracy | sens. | spec. | precision | F1 |",
    "| 0.1 | ____ | ____ | ____ | ____ |":
        "| 0.1 | 5 | 0 | 1.000 | 7/12 = 0.583 |",
    "| 0.5 | ____ | ____ | ____ | ____ |":
        "| 0.5 | 0 | 0 | 1.000 | 1.000 |",
    "| 0.9 | ____ | ____ | ____ | ____ |":
        "| 0.9 | 0 | 2 | 10/12 = 0.833 | 1.000 |",
    "选中阈值____，依据____。如果更看重减少漏诊，为什么不能只比 accuracy？____":
        "选中阈值**0.5**，依据**validation sensitivity≥0.90的可行阈值中，0.5的specificity最高（为1.000）；0.9的sensitivity只有10/12=0.833，不符合规则。**如果更看重减少漏诊，为什么不能只比accuracy？**accuracy把FP与FN合并计数，不能显示漏诊数量及其用途代价；应查看sensitivity和混淆矩阵。**",
    "**小结**：用两句完整的话解释当前阈值下的一项优点和一种可能错误，不仅列指标名称。":
        "**小结：**在当前validation结果中，NB于阈值0.5检出12/12条阳性且未误报，敏感度和特异度均为1.000。这个结果只基于24条同源记录；换到新来源或新症状组合时仍可能出现漏诊或误报，不能把本次零错误当成保证。",
    "accuracy____，sensitivity____，specificity____，precision____，检出阳性数____。与任务 6 的平衡例子相比，为什么 accuracy 容易掩盖问题？____________":
        "accuracy**990/1000=0.990**，sensitivity**0/10=0**，specificity**990/990=1.000**，precision**未定义（没有预测阳性）**，检出阳性数**0**。与任务6的平衡例子相比，为什么accuracy容易掩盖问题？**因为990条阴性占绝大多数，全判阴性仍有99% accuracy，却漏掉全部10条阳性。**",
    "漏诊与误报各意味着什么？在你熟悉的一种医学用途里，哪种错误更需要关注，为什么？________________________________________________________":
        "漏诊与误报各意味着什么？在你熟悉的一种医学用途里，哪种错误更需要关注，为什么？**漏诊（FN）是实际有病却判阴性，可能延误进一步检查；误报（FP）是实际无病却判阳性，可能带来额外检查、焦虑和资源负担。若是初筛，我会优先避免漏诊，但仍需报告误报及其后续代价，不能只优化一个指标。**",
    "从任务 5 的全对模型和任务 4 的复制症状实验中各举一例，解释为什么还要检查概率预测：______________________________________________________":
        "从任务5的全对模型和任务4的复制症状实验中各举一例，解释为什么还要检查概率预测：**任务5中模型A和B的accuracy都为1，但平均log loss分别为0.510826和0.105361，说明分类相同而概率质量不同；任务4中A081复制腰痛五次后概率从0.010210升至0.724999，类别翻转，说明重复信息和条件独立假设会明显改变概率。**",
    "Notebook C 提供可靠性表。任选非空箱：样本数____，平均预测概率____，观察阳性比例____。这个箱在比较什么？为什么不能将小样本观察直接推广到所有新记录？__________________________________________________________":
        "Notebook C提供可靠性表。任选非空箱：样本数**12**，平均预测概率**0.089002**，观察阳性比例**0/12=0**。这个箱在比较什么？**比较该预测概率箱内的平均预测值与实际阳性频率。**为什么不能将小样本观察直接推广到所有新记录？**箱内只有12条，比例有抽样波动；而且资料来自专家设计记录，不代表所有新患者或新医院。**",
    "现在才运行 Notebook D。先记录固定方案：目标____；六输入清单____；α____；阈值____；训练 ID 文件____；提交日期____。维持原 train 模型，不合并 validation 重训，也不根据 test 重新选阈值。":
        "现在才运行Notebook D。先记录固定方案：目标**bladder_inflammation（膀胱炎症，0/1）**；六输入清单**fever_ge_38、nausea、lumbar_pain、urine_pushing、micturition_pain、burning_urethra**；α**1**；阈值**0.5**；训练ID文件**data/student/train.csv（72条，sample_id列）**；提交日期**2026-10-28（本完成稿日期）**。维持原train模型，不合并validation重训，也不根据test重新选阈值。",
    "| TN / FP / FN / TP | ____ / ____ / ____ / ____ | 总记录数____ |":
        "| TN / FP / FN / TP | 12 / 0 / 0 / 12 | 总记录数24（阴性12、阳性12） |",
    "| accuracy | ____ | ____ |":
        "| accuracy | 24/24 = 1.000 | 24 |",
    "| sensitivity | ____ | ____ |":
        "| sensitivity | 12/12 = 1.000 | 真实阳性12 |",
    "| specificity | ____ | ____ |":
        "| specificity | 12/12 = 1.000 | 真实阴性12 |",
    "| log loss | ____ | 参与求平均的记录数____ |":
        "| log loss | 0.074492 | 24 |",
    "对比 validation，哪些数值相同，哪些不同？若恰好全对，为什么也不能推断未来永远不犯错？____________________________________________________":
        "对比validation，哪些数值相同，哪些不同？若恰好全对，为什么也不能推断未来永远不犯错？**两者混淆计数均为TN=12、FP=0、FN=0、TP=12，因此accuracy、sensitivity、specificity都为1.000；log loss不同，validation为0.087214、test为0.074492，表示概率置信程度不同。test只有24条，且记录来自同一来源和规则，有限样本全对不能证明未来必然全对。资料还说明编写者在早期探索中看过测试结果，所以这不是新的盲测或外部验证。**",
    "测试记录中____条六输入组合曾在 train 出现。这次测试回答了哪种“未见数据”问题？没回答哪一种？________________________________________________":
        "测试记录中**24**条六输入组合曾在train出现。这次测试回答了哪种“未见数据”问题？没回答哪一种？**它评价固定模型对同一来源中未参与拟合的记录行能否复现已见症状模式；它没有检验训练未见组合、独立患者抽样或跨医院/人群泛化。**",
    "1. 数据模型接收什么、输出什么，哪些参数来自数据。":
        "1. 数据模型接收什么、输出什么，哪些参数来自数据。**作答：**输入是发热、恶心、腰痛、尿急、排尿痛和尿道灼痛六个0/1值；训练数据估计类别先验`πc`及类内症状概率`θcj`，模型输出阳性概率`p`，再由阈值转为类别。",
    "2. 训练改变什么，预测时做什么，泛化需要怎样评价。":
        "2. 训练改变什么，预测时做什么，泛化需要怎样评价。**作答：**训练用带标签记录估计参数；预测时固定参数，把新症状映射为概率和类别；泛化要在未参与估参和选择、且来源符合目标场景的数据上评价。",
    "3. train、validation、test 各有什么职责，不能怎样混用。":
        "3. train、validation、test各有什么职责，不能怎样混用。**作答：**train估计参数，validation选择阈值，方案冻结后test只作最终评价；不能用test调阈值或改特征后仍把它称为独立测试。",
    "4. log loss 怎样衡量概率误差，为什么不等价于 accuracy。":
        "4. log loss怎样衡量概率误差，为什么不等价于accuracy。**作答：**log loss惩罚模型给真实类别的概率过低，越小越好；accuracy只看阈值后的类别。本例两模型accuracy均为1，但平均log loss为0.510826和0.105361。",
    "5. 混淆矩阵和至少两项指标怎样计算，分母是什么。":
        "5. 混淆矩阵和至少两项指标怎样计算，分母是什么。**作答：**阈值0.5的教学例子TP=2、FN=2、FP=1、TN=3；sensitivity=TP/(TP+FN)=2/4，分母是真实阳性；precision=TP/(TP+FP)=2/3，分母是预测阳性。",
    "6. prior、likelihood、posterior 怎样通过 Bayes theorem 联系，条件独立假设用在哪里。":
        "6. prior、likelihood、posterior怎样通过Bayes theorem联系，条件独立假设用在哪里。**作答：**先验`πc`乘以给定类别下症状的似然得到类别权重，归一化为后验。本例预测概率0.986576；朴素假设用于把六个给定类别后的特征条件概率相乘。",
    "7. 只看 accuracy 会遗漏哪些医学上需要关注的信息。":
        "7. 只看accuracy会遗漏哪些医学上需要关注的信息。**作答：**类别不平衡时全判阴性也可有99% accuracy，却漏掉全部阳性；accuracy不显示sensitivity、precision、错误代价、概率质量或适用人群边界。",
    "最后补一句：我原以为____；通过____的计算或实验，现在理解为____。还有一个没有弄懂的问题是____。":
        "最后补一句：我原以为**只要accuracy高，模型的预测概率就可靠**；通过**比较log loss并观察A081复制腰痛后的概率变化**，现在理解为**分类正确、概率质量和假设稳健性是不同问题，不能只报告accuracy**。还有一个没有弄懂的问题是**怎样用更大且真正独立的外部样本验证概率校准**。",
}


def set_east_asia_font(run, name: str) -> None:
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), name)


def set_cell_margins(cell, top=75, start=90, bottom=75, end=90) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "D9E2E0")


def format_document(path: Path) -> None:
    doc = Document(path)

    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.65)
        section.bottom_margin = Cm(1.65)
        section.left_margin = Cm(1.75)
        section.right_margin = Cm(1.75)
        header = section.header.paragraphs[0]
        header.text = "第7讲 机器学习基础与朴素贝叶斯 | 学生工作单完成版"
        header.alignment = 1
        for run in header.runs:
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(96, 110, 108)
            set_east_asia_font(run, "Microsoft YaHei")
        footer = section.footer.paragraphs[0]
        footer.text = "完成版工作单 | 2026-10-28"
        footer.alignment = 1
        for run in footer.runs:
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(96, 110, 108)
            set_east_asia_font(run, "Microsoft YaHei")

    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.08
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

    for style_name, size, color in (
        ("Title", 19, "173F3A"),
        ("Heading 1", 15, "173F3A"),
        ("Heading 2", 12.5, "275D55"),
        ("Heading 3", 11, "275D55"),
    ):
        style = doc.styles[style_name]
        style.font.name = "Aptos Display"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.paragraph_format.keep_with_next = True

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            set_east_asia_font(run, "Microsoft YaHei")
        if paragraph.style.name.startswith("Heading"):
            paragraph.paragraph_format.keep_with_next = True

    for table in doc.tables:
        table.autofit = True
        table.style = "Table"
        set_table_borders(table)
        for row_index, row in enumerate(table.rows):
            if row_index == 0:
                tr_pr = row._tr.get_or_add_trPr()
                header_marker = OxmlElement("w:tblHeader")
                header_marker.set(qn("w:val"), "true")
                tr_pr.append(header_marker)
            for cell in row.cells:
                cell.vertical_alignment = 1
                set_cell_margins(cell)
                if row_index == 0:
                    shade_cell(cell, "DDEBE8")
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(1.5)
                    paragraph.paragraph_format.line_spacing = 1.0
                    if row_index == 0:
                        paragraph.alignment = 1
                    for run in paragraph.runs:
                        set_east_asia_font(run, "Microsoft YaHei")
                        run.font.size = Pt(8.8)
                        if row_index == 0:
                            run.font.bold = True

    doc.save(path)


def main() -> None:
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    if OUTPUT.exists():
        existing = Document(OUTPUT)
        existing_text = "\n".join(p.text for p in existing.paragraphs)
        if "A081" not in existing_text or "0.9865759063" not in existing_text:
            raise FileExistsError(f"Refusing to overwrite existing output: {OUTPUT}")

    text = SOURCE.read_text(encoding="utf-8")
    for original, completed in REPLACEMENTS.items():
        count = text.count(original)
        if count != 1:
            raise ValueError(f"Expected one occurrence, found {count}: {original[:120]}")
        text = text.replace(original, completed, 1)

    unresolved = [
        (line_number, line)
        for line_number, line in enumerate(text.splitlines(), start=1)
        if re.search(r"_{3,}", line) and not line.startswith("姓名：")
    ]
    if unresolved:
        raise ValueError(f"Unfilled response blanks remain: {unresolved[:12]}")

    text = re.sub(
        r"`([^`]+)`",
        lambda match: "`" + match.group(1).replace("**", "") + "`",
        text,
    )

    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise FileNotFoundError("pandoc was not found on PATH")

    result = subprocess.run(
        [pandoc, "--from=markdown", "--to=docx", "--output", str(OUTPUT)],
        input=text,
        text=True,
        encoding="utf-8",
        capture_output=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr or "pandoc failed")

    format_document(OUTPUT)

    check = Document(OUTPUT)
    body_text = "\n".join(p.text for p in check.paragraphs)
    table_text = "\n".join(cell.text for table in check.tables for row in table.rows for cell in row.cells)
    all_text = body_text + "\n" + table_text
    if "0.9865759063" not in all_text or "0.074492" not in all_text:
        raise ValueError("Key answer values are missing from the generated document")
    print(f"Created {OUTPUT}")
    print(f"Paragraphs: {len(check.paragraphs)}; tables: {len(check.tables)}")
    print(f"Bytes: {OUTPUT.stat().st_size}")


if __name__ == "__main__":
    main()
