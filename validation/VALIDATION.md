# Tafel 分析器当前验证记录

更新日期：2026-11-25。Windows 11 / Python 3.12.10。

## 当前维护内容

- 唯一源码工程：当前目录。
- 唯一 Windows 程序：`dist/TafelAnalyzer.exe`。
- 唯一桌面入口：“Tafel 分析器”，指向 `dist/TafelAnalyzer.exe`。
- 当前应用版本：1.0.0。后续需求继续更新这套源码、程序和桌面入口。

## 回归验证

- `python -m unittest discover -s tests -v`：68/68 通过。
- Windows EXE 使用 UPX 5.3.0 压缩；构建脚本可通过 `-UPXDirectory` 指定其目录。
- 覆盖 CHI TXT/CSV 解析、面积归一化、手动 Tafel 拟合与交点、区间绘图数据 Excel、批量计算及结果管理、图上选择拟合区间、拟合质量提示、趋势统计与基准比较、项目保存恢复、参数模板、实验信息、600 dpi 图片导出、离线帮助，以及滚轮不会更改面积或拟合/导出区间。
- 包括主窗口创建后首帧绘制前立即关闭的测试；Matplotlib 不再对已销毁的 Qt 绘图控件报错。
- 单文件、批量、趋势和项目的原生 Qt 检查脚本位于 `validation/`，使用固定文件名，不保存多套成品版本。

## 科学与真实数据验证边界

- 合成案例 1 和案例 2 的拟合、交点和电流密度均通过已知公式核对。
- 用户目前没有 Origin 对照结果，因此尚未完成实际 CHI604E 文件与 Origin 的逐例对照。所需输入和核对步骤见 [Origin 对照计划](ORIGIN_COMPARISON.md)；获得原始 CHI 文件和对应 Origin 拟合数值后再补做。
- 趋势统计方法和适用范围见 [趋势方法说明](../TREND_METHODS.md)。趋势图不自动判断机理、显著性或腐蚀速率。
- 当前不支持 XLSX 原始数据导入或自动挑选 Tafel 区域；区间仍由用户判断并手动设置。

## 当前成品

- 文件：`dist/TafelAnalyzer.exe`
- 大小：65,426,907 bytes（比原 81,308,707 bytes 减少约 19.5%）
- SHA256：`8FA78F1558CE8B44D3065FEA464B666E5B2CFB42225B86E3F1FF12B2EAA85A70`
- 单文件 EXE 自检：成功，退出码 0；涵盖 Excel、批量、趋势、PNG/TIFF/PDF/SVG 图片及内置帮助。
- 桌面快捷方式：“Tafel 分析器”，目标为当前唯一 `dist/TafelAnalyzer.exe`。
- `dist/` 中只有这一个 EXE；替换时程序未运行。
