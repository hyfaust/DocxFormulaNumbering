# DocxFormulaNumbering (DFN)

> Word 文档公式自动编号工具

[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![Word 2016+](https://img.shields.io/badge/Word-2016+-green.svg)](https://products.office.com/word)
[![License: GPL v3](https://img.shields.io/badge/license-GPLv3-blue.svg)](LICENSE)

[English](README.md) | [简体中文](README_zh.md)

---

## 简介

DocxFormulaNumbering (DFN) 是一个 Python 工具，可自动为 Microsoft Word 文档中的单行公式添加公式编号和书签。它支持基于章节的编号格式，并可创建书签用于交叉引用。

## 功能特性

- ✅ 自动识别 Word 文档中的单行公式
- ✅ 章节号编号格式：`(1-1)`, `(1-2)`, `(2-1)`...
- ✅ 简单连续编号格式：`(1)`, `(2)`, `(3)`...
- ✅ 创建书签用于交叉引用
- ✅ 支持 eqArr 格式（专业的公式排版）
- ✅ 通过 Word COM 自动化实现 100% Word 兼容

## 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 基本用法

```bash
# 默认模式：章节号 - 序号格式
python dfn.py input.docx output.docx

# 简单模式：连续编号
python dfn.py input.docx output.docx --simple

# eqArr 格式（需要 t1.docx 模板）
python dfn.py input.docx output.docx --t1

# 覆盖原文件
python dfn.py document.docx
```

## 命令行选项

| 选项 | 说明 |
|------|------|
| `input` | 输入 docx 文件路径（必需） |
| `output` | 输出 docx 文件路径（可选，默认覆盖原文件） |
| `--simple` | 使用简单连续编号模式（不含章节号） |
| `--t1` | 使用 t1 格式（eqArr 环境，需要 t1.docx 模板） |

## 前提条件

### 默认模式要求

文档需要配置多级列表：

1. 章节标题使用"标题 1"样式
2. 通过"开始"→"段落"→"多级列表"应用编号
3. 标题样式与多级列表级别关联

### eqArr 格式要求

使用 `--t1` 参数时，确保 `t1.docx` 模板文件存在于同一目录。

## 输出格式

### 默认格式（t2 格式）

```
E = mc²#(1-1)
```

编号为普通文本，与公式在同一行。

### eqArr 格式（t1 格式）

```xml
<m:eqArr>
  <m:t>E = mc²#</m:t>
  <m:d>
    <m:t>1</m:t>
  </m:d>
</m:eqArr>
```

编号在 delimiter 元素中，显示更美观，与公式融为一体。

## 书签命名

- **默认模式**: `eq{章}_{序号}` 如 `eq1_1`, `eq1_2`, `eq2_1`
- **简单模式**: `eq{序号}` 如 `eq1`, `eq2`, `eq3`

可通过 Word 的"插入"→"交叉引用"→"书签"进行引用。

## 项目结构

```
DocxFormulaNumbering/
├── dfn.py              # 主脚本（入口文件）
├── README.md           # 英文文档
├── README_zh.md        # 中文文档
├── USAGE.md            # 详细使用指南
├── requirements.txt    # Python 依赖
├── LICENSE             # GPL v3 许可证
├── .gitignore          # Git 忽略规则
├── t1.docx             # eqArr 格式模板
└── example.docx        # 测试文档
```

## 技术细节

- **语言**: Python 3.6+
- **依赖**: pywin32 (Word COM 自动化)
- **核心**: Word 对象模型 + OMML (Office Math Markup Language)

## 示例

### 输入文档

```
第 1 章 引言

E = mc²

F = ma

第 2 章 相对论

p = mv
```

### 输出（默认模式）

```
第 1 章 引言

E = mc²#(1-1)  [书签：eq1_1]

F = ma#(1-2)   [书签：eq1_2]

第 2 章 相对论

p = mv#(2-1)   [书签：eq2_1]
```

## 常见问题

### Q: 提示"文档未配置多级列表"错误

A: 您的章节标题没有应用多级列表。使用 `--simple` 参数跳过章节检测。

### Q: 公式编号显示为乱码

A: 按 `F9` 更新所有域，或右键点击编号选择"更新域"。

### Q: 如何引用公式编号？

A: 在 Word 中使用"插入"→"交叉引用"→"书签"，选择对应的 `eq*` 书签。

## 许可证

本项目采用 GNU 通用公共许可证 v3.0 授权 - 详见 [LICENSE](LICENSE) 文件。

## 作者

FaustSherpad

## 版本

1.0.0
