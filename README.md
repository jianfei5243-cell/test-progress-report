# 测试进度报告生成器

OS4.0 第一阶段全功能测试进度汇总与报告生成工具。

📄 单页功能描述：[功能描述.md](功能描述.md)

## 功能

- 读取测试管理平台导出的 `Test导出_*.xlsx`
- 自动汇总：总体进度、功能模块(Folder)、测试批次(Run Name)、优先级、Top 缺陷、人力投入
- 一键输出：网页版测试报告(HTML)、数据汇总表(Excel)、模块执行率柱状图(PNG)

## 使用方式

### 方式一：可视化界面（推荐，无需终端、无需联网）

1. 双击打开 `测试报告生成器.html`
2. 拖入或点击选择 `Test导出_*.xlsx`
3. 查看报告，点击按钮下载 HTML / Excel / PNG

### 方式二：命令行脚本

```bash
pip install openpyxl pillow
python3 generate_test_report.py                      # 自动发现最新导出文件
python3 generate_test_report.py -i /path/导出.xlsx    # 指定文件
python3 generate_test_report.py -o ./report          # 指定输出目录
```

## 目录结构

- `测试报告生成器.html` — 可视化操作界面（单文件、纯本地运行）
- `generate_test_report.py` — 命令行自动化脚本
- 说明：本仓库不含任何测试数据（xlsx / 报告 / 图表已在 `.gitignore` 中忽略）

## 输出文件

- `YYYYMMDD_测试报告.html`
- `YYYYMMDD_测试进度汇总.xlsx`
- `YYYYMMDD_测试进度_模块执行率.png`
