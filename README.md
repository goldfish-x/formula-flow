# 公式美化大师 (formula-flow)

> LaTeX ⇄ Word 原生公式（OMML）的双向通道。生成的公式**双击可编辑**，不是图片、不是文本模拟。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## 这是什么

在 Word 里写数学公式有两个常见的坑：

| 坑 | 表现 | 本项目的做法 |
|---|---|---|
| **公式变成图片** | 看着对，但不能编辑、改字号，打印发虚 | 生成 Word **原生 oMath 对象**，双击可编辑 |
| **公式退化成纯文本** | `x^2` 直接显示成 `x^2`、`√(ab)` 是字符串 | LaTeX 解析后转成真正的上标/根号结构 |

核心链路：

```
LaTeX ──latex2mathml──▶ MathML ──MML2OMML.XSL──▶ OMML ──▶ Word 原生公式
```

支持 WPS / Microsoft Word 打开（公式格式为 OMML 开放标准，ECMA-376）。

## 功能一览

### 1. Markdown → Word（公式可编辑）

```bash
python scripts/md2docx.py input.md -o output.docx --title "高中数学 必修一"
```

Markdown 里直接写 `$A\subsetneqq B$`、`$$\sum_{i=1}^{n}i=\frac{n(n+1)}{2}$$`，
输出为 Word 原生公式。

### 2. 伪公式批量升级

很多老讲义的公式是 Unicode 字符硬拼的（`a²+b²≥2ab`、`A⊆B`、`√(ab)`）——
`upgrade_math.py` 把它们批量升级成真正的可编辑公式：

```bash
python scripts/upgrade_math.py 讲义.docx --report     # 先出报告（只读）
python scripts/upgrade_math.py 讲义.docx --apply      # 确认后应用
```

### 3. MathType 公式抢救（MTEF 解析）

docx 里 MathType 公式 = 一张 WMF 呈现图 + 一个 OLE 二进制（`Equation Native` 流）。
`mtef.py` 直接解析 **MTEF v5 二进制**还原完整公式结构，不靠 OCR：

```bash
python scripts/mtef.py 讲义.docx --json eqs.json       # 提取为 LaTeX
python scripts/mtef2docx.py 讲义.docx --apply          # 原位替换为原生公式
```

全库实测：214 份文档 / 8521 个 MTEF 对象，解析成功率 **100%**（v5 部分）。

### 4. WMF 图元符号提取

`wmf_symbols.py` 从 MathType 的 WMF 图元里精确提取符号（ExtTextOut 字符 +
PolyDraw 路径），交叉验证 MTEF 解析结果。

### 5. 矢量文本型 PDF → 可编辑 Word（无需 OCR）

`pdf2docx_math.py`：PDF 是矢量文本型时（非扫描件），不走 OCR，
直接由**字符几何**反推公式结构：

```
行聚类(垂直中心) → 上下标(字号+基线) → 分数(分数线) → 根号(描边路径)
→ 可伸缩定界符(纵向拼接件) → 表格(网格线) → LaTeX → OMML
```

```bash
python scripts/pdf2docx_math.py input.pdf -o output.docx --dump
```

支持：上下标、分数、根号、定界符、方程组（原生 `m:eqArr`）、表格、
填空横线（下划线空格格式）、WPS 的 SymbolMT 私有区编码自动解码。

### 6. 九大数学域符号全覆盖

生成侧 + 识别侧双向支持，近 300 条测试表达式全部通过：

| 域 | 代表符号 |
|---|---|
| 代数/集合 | `⊆ ⊂ ∪ ∩ ∅ ∈ ∉ ⊊` |
| 线性代数 | 行列式 `vmatrix`、转置 `A^T`、秩、⊗ ⊕、范数 |
| 希腊字母 | 全部大小写 + 变体两形（ε θ φ ρ σ） |
| 统计/组合 | `binom{n}{k}` 真堆叠、`x̄`、卡方、正态、条件概率 |
| 数论 | `pmod`、`gcd`、上下取整、数集 ℤℕℚℝℂ |
| 平面几何 | `∠ △ ≅ ∼ ∥ ⊥ ⊙ ⌒AB` |
| 微积分/场论 | `∬ ∭ ∮ ∯ ∰`、`∂`、`∇`、lim 正下方、牛顿点 |
| 化学 | 反应条件上标（点燃/高温/催化剂）、`⇌`、离子、同位素 |
| 生物/地理 | `♀ ♂` 杂交图解、中心法则、经纬度度分秒、‰ 盐度 |

## 安装

依赖：Python 3.9+（Windows）

```cmd
install.cmd
```

一条命令完成：建独立虚拟环境 `.venv` → 装依赖 → 跑自检。

或手动：

```bash
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe scripts\latex2omml.py    # 自检
```

### 在 WorkBuddy 中使用

本 skill 兼容 WorkBuddy 技能格式（SKILL.md frontmatter）：
技能市场 → 上传技能 → 选择本目录打包的 ZIP；或
「从 Git 仓库导入」粘贴本仓库地址。

## 目录结构

```
latex-omml/
├── SKILL.md              # AI agent 使用说明（能力地图 + 踩坑记录）
├── INSTALL.md            # 安装说明
├── install.cmd           # Windows 一键安装
├── requirements.txt
├── references/
│   └── symbols.md        # 符号速查表（LaTeX ↔ Unicode ↔ OMML 结构）
└── scripts/
    ├── latex2omml.py     # 核心：LaTeX → MathML → OMML（含预映射层）
    ├── md2docx.py        # Markdown → docx（公式可编辑）
    ├── upgrade_math.py   # 伪公式批量升级
    ├── mtef.py           # MTEF v3/v5 二进制解析器（MathType OLE）
    ├── mtef2docx.py      # MathType OLE → 原生公式批量替换
    ├── wmf_symbols.py    # WMF 图元符号提取
    ├── pdf2docx_math.py  # 矢量文本型 PDF → docx（无需 OCR）
    ├── verify_docx.py    # 生成结果校验
    ├── audit_math.py     # 讲义公式质量审计
    └── MML2OMML.XSL      # 微软官方转换表（见下方版权说明）
```

## 版权说明

`scripts/MML2OMML.XSL` 来自 Microsoft Office 安装目录，仅为本项目与
Word 公式格式互通之目的包含，版权归 Microsoft Corporation 所有。
公式输出格式为 OMML（ECMA-376 开放标准），由 WPS / Microsoft Word
正常渲染与编辑，与 Office 版本无关。如果你是权利人并反对收录，
请提 issue，我们会移除。

## License

[MIT](LICENSE)
