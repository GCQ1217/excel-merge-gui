# Excel 合并工具

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

[English](README.md) | 简体中文

将多个 Excel / CSV 文件按行合并为一个文件的桌面工具。

## 功能特性

- 合并 .xlsx、.xls、.csv 文件，按行追加
- 为每个 Excel 文件选择指定 sheet
- 全局起始行设置（跳过表头）
- 单个文件行范围设置
- CSV 自动编码检测（UTF-8 / GBK / GB18030 等）
- CSV 自动分隔符检测
- CSV 转 Excel 格式

## 支持格式

| 格式 | 扩展名 | 作为输入 | 作为输出 |
|------|--------|----------|----------|
| Excel 新版 | .xlsx | ✅ | ✅ |
| Excel 旧版 | .xls | ✅ | — |
| CSV | .csv | ✅ | ✅ |

> .xls 文件可以作为输入参与合并，但输出仅支持 .xlsx 或 .csv。

---

## 快速开始

### 从源码运行

```bash
pip install -r requirements.txt
python excel_merge_gui.py
```

### 直接使用 exe

双击 excel_merge_gui.exe 即可运行。

也可以自行打包：

```bash
scripts\build_pyinstaller.bat   # 速度快，体积大
scripts\build_nuitka.bat        # 速度慢，体积小
```

---

## 使用说明

### 基本合并流程

1. 点击「添加文件」，选择一个或多个文件
2. （可选）点击「选择」设置输出文件路径
3. 点击「合并并保存」

合并结果按文件添加顺序，依次追加每个文件的所有行。

### 界面说明

| 按钮 | 功能 |
|------|------|
| **添加文件** | 打开文件选择对话框，支持多选 |
| **移除选中** | 从列表中移除当前选中的文件 |
| **清空列表** | 清除所有已添加的文件 |
| **为选中文件选择 sheet** | 为选中的 Excel 文件逐个选择 sheet |
| **为选中文件设置行范围** | 为选中文件设置要复制的行范围 |
| **清除所选文件行范围** | 将选中文件的行范围恢复为默认 |
| **CSV -> Excel** | 将选中的 CSV 文件转换为 .xlsx |
| **合并并保存** | 执行合并 |

### 选择 Sheet

适用于 .xlsx / .xls 文件（CSV 会自动跳过）。

1. 在列表中选中文件
2. 点击「为选中文件选择 sheet」
3. 在弹出的下拉列表中选择，点击「确定」

未设置的文件默认读取第一个 sheet。

### 设置行范围

可以为每个文件单独指定要复制的行范围（第 1 行 = 文件中第一行）。

| 设置 | 效果 |
|------|------|
| 仅设置全局起始行 = N | 所有文件从第 N 行开始复制到末尾 |
| 仅设置单文件行范围 S-E | 取 max(全局起始行, S) 到 E |
| 两者都设置 | 取更大的起始行，结束行以单文件设置为准 |

**举例**：全局起始行 = 2（跳过表头），某文件另设 5-20 → 实际复制第 5~20 行。

### CSV 转 Excel

1. 选中 CSV 文件，点击「CSV -> Excel」
2. 转换后的 .xlsx 保存在原文件同目录下
3. 同名文件已存在时自动追加 _converted 后缀

### 输出格式

| 扩展名 | 格式 | 说明 |
|--------|------|------|
| .xlsx | Excel | 默认格式 |
| .csv | CSV | UTF-8-BOM 编码，兼容 Excel 直接打开 |

### 合并规则

- **合并模式**：按行追加，将每个文件的数据纵向拼接
- **表头处理**：不自动识别表头，跳过表头请将起始行设为 2
- **列数不一致**：取所有列的并集，缺失位置填充空值

---

## 常见问题

**Q: CSV 文件乱码？**
工具会自动尝试多种编码（UTF-8/GBK/GB18030 等）。若仍有问题，建议先用记事本将文件另存为 UTF-8 编码。

**Q: 合并后第一行出现数字（0, 1, 2...）？**
将全局起始行改为 2 可跳过表头。

**Q: 输出文件路径为空？**
默认保存到 exe 所在目录下的 merged.xlsx。

**Q: 部分文件读取失败？**
合并仍会继续，结果对话框会列出失败的文件及原因。

---

## 项目结构

```
├── README.md
├── README.zh-CN.md
├── requirements.txt
├── excel_merge_gui.py
├── testdata/
│   ├── employees_1.xlsx
│   ├── employees_2.xlsx
│   ├── multi_sheet.xlsx
│   ├── products.xls
│   ├── cities_utf8.csv
│   └── cities_gbk.csv
└── scripts/
    ├── build_nuitka.bat
    └── build_pyinstaller.bat
```

## 技术栈

- **GUI**: tkinter
- **数据处理**: pandas
- **Excel 读写**: openpyxl (.xlsx), xlrd (.xls)

## License

[MIT](LICENSE)
