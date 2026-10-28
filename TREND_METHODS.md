# 批量结果趋势分析的方法依据

检索与整理日期：2026-10-28。此文说明本程序采用的比较方法和解释范围，文献中的具体金属、介质和温度不能直接代替当前实验条件。

## 先确定电流密度的物理含义

原程序计算的是两段人工指定 Tafel 直线的交点电流密度。该交点如何命名，取决于测试体系：原有电池反应动力学分析沿用交换电流密度 j0；腐蚀极化实验可在适用的阴、阳极 Tafel 区间外推得到腐蚀电流密度 jcorr。不能仅凭计算式相同就把两者等同，也不能统一认为数值越小越好。

IUPAC 的 exchange current 定义针对给定电极反应的平衡状态；mixed potential 定义涉及不同氧化还原反应在净电流为零时的共同电位。来源：[exchange current](https://goldbook.iupac.org/terms/view/E02238)、[mixed potential](https://goldbook.iupac.org/terms/view/M03944)。趋势窗口的模式选择只改变解释和比较指标，不会重新拟合，也不会验证反应机理。

## 文献中常见的比较方式

| 方法 | 用途 | 文献依据 |
| --- | --- | --- |
| 对比 jcorr、Ecorr 和阴/阳极 Tafel 斜率 | 比较不同样品、浓度、温度或处理条件 | Canto 等，2023，[Corrosion Inhibition in CO2-Saturated Brine by Nd3+ Ions](https://pmc.ncbi.nlm.nih.gov/articles/PMC10537565/)，表 1–2、图 3–4 |
| 保留重复实验，展示离散程度 | 判断样品间差异是否大于实验散布，避免只看一次拟合 | Özkan 等，2024，[Laying the experimental foundation for corrosion inhibitor discovery through machine learning](https://www.nature.com/articles/s41529-024-00435-z)，每种缓蚀剂至少三次实验；其误差棒通常为标准误，程序采用的标准差会单独标明 |
| 相对空白条件计算缓蚀效率 IE，同时保留密度和比值 | 比较抑制程度；较高 IE 会压缩实际性能差异 | 同上，式 4–6，讨论 IE 与 inhibition power 的区别 |
| 比较浓度响应及重复测量 | 观察随添加浓度变化的电化学参数 | 2025，[Influence of praseodymium(III) chloride on corrosion resistance of AS1020 steel in an environment containing chloride ions and carbon dioxide](https://pubs.rsc.org/en/content/articlehtml/2025/ra/d5ra05060j)，公开检索文本列出密度、Ecorr、Tafel 斜率及重复测量；全文访问受限 |

2026 年的研究 [How do system-specific conditions limit the validity of linear polarization resistance and potentiodynamic polarization in determining corrosion rates?](https://www.sciencedirect.com/science/article/pii/S0010938X26000259) 的公开摘要也强调 LPR/PDP 在不同表面及保护条件下的适用性问题；本次无法读取其全文，因此未据此引入额外计算模型。

## 本程序采用的统计与单位约定

以下为程序的分析约定，不能解读为所有文献统一采用的规则。

- 默认按样品/处理条件作分类比较；不根据文件名猜测时间、浓度或重复实验。只有用户明确填写相同分组时才合并重复实验。数值横轴需要用户填写值及单位。
- 原始电流密度统一为 mA/cm²，来自 log10(abs(I/A) × 1000 / S/cm²) 的反对数。每个样品仍使用自己的面积和人工拟合范围。
- 组均值采用算术平均；误差棒采用样本标准差 SD，分母 n−1。n=1 时 SD 留空，不表示误差为零。SD 表示组内散布，不能替代置信区间或显著性检验；同一样品的连续扫描也不自动构成独立重复实验。
- 平均 log10(j) 与 log10(平均 j) 是不同统计量。按 log10(j) 分析时须明确前者，不能把两种结果混用。
- 原程序拟合斜率 a 的单位是 decade/V；文献常见的 Tafel 斜率幅值 β 为 mV/decade，因此 β=1000/abs(a)，不能直接把 a 当成 β。
- 基准比较使用用户选择的组/条件的算术平均密度 jref。比值=j/jref；相对变化=100×(j/jref−1)%。腐蚀模式下 IE=100×(1−j/jref)%；IP=10×log10(jref/j) dB。负 IE 保留，代表比基准更高的腐蚀电流密度。jref 应对应可比条件下的空白/对照；程序不自动判断它是否合适。
- 上述基准指标是组均值的点估计，不传播对照组误差。单次实验的拟合 R² 高，不能证明样品间差异具有统计显著性，也不能证明所选区域符合 Tafel 动力学。
- 数值横轴上的线性趋势仅是描述性拟合，用户显式开启后才使用；不能从拟合斜率、R² 或腐蚀电位偏移自动推断机理。

## 本阶段范围

支持批量拟合结果的比较、分组统计、基准比较及分析数据导出。保留原始拟合参数与来源，失败、过期或非有限结果不得作为有效趋势点。用户的分组、横轴及选择变化后必须重新生成分析，避免导出旧图与新表混杂的结果。

不根据电流密度自动换算 mm/year。ASTM [G102 的公开范围说明](https://store.astm.org/standards/g102) 将电化学结果换算限定于均匀腐蚀的适用范围；实际计算还需材料等效重量、密度等信息。当前工具没有这些实验输入，也不由趋势图自动推断局部腐蚀、吸附等温线、活化能或显著性。
