---
name: latex-omml
version: 1.1.1
description: 数学公式与 Word 文档的双向解决方案。当需要在 Word/docx 中生成可编辑的数学公式（分数、根号、上下标、希腊字母、集合符号、矩阵、积分求和），从 PPT/PDF/图片中提取数学符号并还原为 LaTeX，解析 MathType OLE 对象里的 MTEF 二进制还原完整公式结构，或把已有 docx 里 Unicode 硬拼的伪公式（a²+b²≥2ab、A⊆B、√(ab)）批量升级为真正的可编辑公式时使用。触发词：数学公式、LaTeX、OMML、公式转Word、可编辑公式、集合符号、⊆ ∅ ∈、生成讲义、试题排版、公式识别、MathType 替代、讲义公式升级、伪公式、MTEF、Equation Native、MathType 转 LaTeX、OLE 公式提取。
category: 内容创作
platforms: [windows, macos, linux]
license: MIT
author: goldfish-x
homepage: https://github.com/goldfish-x/formula-flow
keywords: [latex, omml, mathml, word, docx, equation, math-formulas, education, wps, mathtype]
tags: [latex, omml, word, education, equations]
agent_created: true
---

# latex-omml — 数学公式与 Word 的双向通道

## 这个 skill 解决什么

Word 里写数学公式有两个常见坑，本 skill 全部绕开：

| 坑 | 表现 | 本 skill 的做法 |
|---|---|---|
| **公式变成图片** | 看着对，但不能编辑、不能改字号、打印发虚 | 生成 Word **原生 oMath 对象**，双击可编辑 |
| **公式退化成纯文本** | `x^2` 直接显示成 `x^2` | LaTeX 解析后转成真正的上标结构 |

### 技术链路

```
LaTeX ──latex2mathml──▶ MathML ──MML2OMML.XSL──▶ OMML ──▶ Word 原生公式
```

`MML2OMML.XSL` 是微软 Office 自带的官方转换表，已内置在 `scripts/` 下，
skill 自包含、可移植到没装 Office 的机器。

---

## 快速开始

### 方式一：Markdown → Word（推荐）

写一个 `.md` 文件，公式用 `$...$`（行内）或 `$$...$$`（独立成行）：

```markdown
# 1.2 集合间的基本关系

若 $A\subsetneqq B$，$B\subsetneqq C$，则 $A\subsetneqq C$.

空集 $\varnothing$ 是任何非空集合的真子集.

$$x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}$$
```

执行：

```bash
python scripts/md2docx.py input.md -o output.docx --title "高中数学 必修一"
```

### 方式二：Python API

```python
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else '.', 'scripts'))
# 若在别处运行，把下面路径换成 skill 的 scripts 目录
# sys.path.insert(0, '<skill目录>/scripts')

from docx import Document
from latex2omml import add_paragraph, latex_to_omml

doc = Document()
add_paragraph(doc, r'真子集：$A\subsetneqq B$', style='Heading 2')
add_paragraph(doc, r'$$\sum_{i=1}^{n}i=\frac{n(n+1)}{2}$$')
doc.save('out.docx')
```

### 方式三：升级已有 docx 里的「伪公式」

很多讲义的公式是 **Unicode 字符硬拼**的（`a²+b² ≥ 2ab`、`A⊆B`、`√(ab)`、
`(a+b)/2`）—— 看着像公式，实际是纯文本，不能编辑、排版不专业。

`upgrade_math.py` 把它们批量升级成真正的 oMath 对象：

```bash
# 1. 先出报告，不动原文件（默认只读）
python scripts/upgrade_math.py "第4讲.docx"

# 2. 看转换抽样，确认质量
python scripts/upgrade_math.py "第4讲.docx" --sample 15

# 3. 确认后应用，输出到独立目录
python scripts/upgrade_math.py "C:\讲义目录" --apply --outdir "C:\升级输出"

# 4. 保守模式：只转高置信度
python scripts/upgrade_math.py "第4讲.docx" --apply --min-level high
```

**安全设计**：

- 默认**只读**，加 `--apply` 才写入
- **不覆盖原文件**，输出到 `_公式升级.docx` 或 `--outdir`
- 逐 run 操作并**复制原 rPr**，加粗/颜色/字号全部保留
- 每段 LaTeX 生成后**回读校验**，校验不通过就保持原文（计入"校验跳过"）
- 可选 `--no-backup`；默认也会备份为 `.bak.docx`

**已实测**：16 讲高中数学讲义，866 处伪公式 → 真公式，
中文内容 100% 一致、加粗零丢失、0 失败。

### 方式四：MathType 文档 → LaTeX（免 OCR，带完整结构）

讲义里的公式如果是 **MathType OLE 对象**（不是图片、不是 Unicode 硬拼），
可以直接从 `word/embeddings/*.bin` 里解出 MTEF 二进制，**还原成完整结构的
LaTeX** —— 不用 OCR、不用读图，分数/根号/向量/绝对值/上下标的结构全部保留。

```bash
# 列出全部公式的 LaTeX
python scripts/mtef.py "6.2.6n.docx"

# 导出成 JSON（part / latex / error 三个字段）
python scripts/mtef.py "6.2.6n.docx" --json eqs.json

# 验证解析器本身没坏（官方 SDK 示例流）
python scripts/mtef.py --selftest
```

**一键端到端**（直接把整份讲义的 MathType 对象换成原生公式）：

```bash
# 1. 先体检，不动文件（默认只读）
python scripts/mtef2docx.py "6.2.6n.docx" --sample 12

# 2. 确认后生成，输出到独立文件（原文件不动）
python scripts/mtef2docx.py "6.2.6n.docx" --apply

# 3. 整个目录批量
python scripts/mtef2docx.py "C:\讲义目录" --apply --outdir "C:\输出"
```

详见 [MathType 文档抢救（MTEF 解析）](#mathtype-文档抢救mtef-解析)。

### 方式五：自检

```bash
python scripts/latex2omml.py        # 跑 12 项核心符号测试
python scripts/md2docx.py --selftest
python scripts/mtef.py --selftest   # MTEF 解析器自检
```

---

## 环境

依赖（见 `requirements.txt`，可用 `install.cmd` 一键装到本目录 `.venv`）：

- `lxml` 5.4+（XSLT 转换）
- `python-docx` 1.2+（docx 读写）
- `latex2mathml` 3.81+（LaTeX → MathML）

如果报 `ModuleNotFoundError`，用虚拟环境里的解释器跑：

```
.venv\Scripts\python.exe scripts\md2docx.py input.md -o output.docx
```

---

## 支持的语法

### 结构类

| LaTeX | 效果 | OMML 元素 |
|---|---|---|
| `\frac{a}{b}` | 分数 | `m:f` |
| `x^{2}` | 上标 | `m:sSup` |
| `a_{i}` | 下标 | `m:sSub` |
| `a_{i}^{j}` | 上下标 | `m:sSubSup` |
| `\sqrt{x}` / `\sqrt[3]{x}` | 根号 | `m:rad` |
| `\sum_{i=1}^{n}` | 求和 | `m:nary` |
| `\int_{0}^{1}` | 积分 | `m:nary` |
| `\begin{pmatrix}a & b \\ c & d\end{pmatrix}` | 矩阵 | `m:m` |
| `\overline{AB}` | 上划线 | `m:bar` |

### 符号类（常用，完整表见 `references/symbols.md`）

**集合与逻辑**：`\subseteq ⊆` `\subsetneqq ⫋` `\supseteq ⊇` `\cap ∩` `\cup ∪`
`\varnothing ∅` `\in ∈` `\notin ∉` `\mid ∣` `\{ \}`

**数集**：`\mathbb{N} ℕ` `\mathbb{Z} ℤ` `\mathbb{Q} ℚ` `\mathbb{R} ℝ` `\mathbb{C} ℂ`

**关系与运算**：`\leq ≤` `\geq ≥` `\neq ≠` `\approx ≈` `\equiv ≡` `\pm ±` `\times ×` `\div ÷` `\cdot ⋅`

**希腊字母**：`\alpha β` … `\omega ω`，大写 `\Gamma \Delta \Theta \Pi \Sigma \Omega`

**箭头**：`\to →` `\Rightarrow ⇒` `\leftrightarrow ↔` `\mapsto ↦`

**几何**：`\angle ∠` `\perp ⊥` `\parallel ∥` `\triangle △` `\cong ≅` `\sim ∼`

### 九大域扩展（审计：84/84 + 微积分 47/47 + 化学 45/45 + 生物地理 56/54+）

**生物·地理**：`\female ♀` `\male ♂`（U+2640/2642 原生支持）杂交图解 `♀AA\times♂aa\rightarrow Aa`、
`P\times P\rightarrow F_{1}\rightarrow F_{2}`；遗传平衡 `p^{2}+2pq+q^{2}=1`、J/S 型增长、
标记重捕 `p=Mn/N`；中心法则 `\xrightarrow{\text{转录}}`（m:limUpp）、
光合作用双向条件 `\underset{\text{叶绿体}}{\overset{\text{光能}}{\rightarrow}}`（limUpp+limLow 同框）；
激素缩写 `GA_{3}` `\alpha-\text{淀粉酶}`、肽键 `-\mathrm{CO}-\mathrm{NH}-`、米氏方程、
`P<0.05` `\bar{x}\pm s`、呼吸商；
地理：正午太阳高度 `H=90^{\circ}-\left|\varphi-\delta\right|`、经纬度度分秒
`116^{\circ}23^{\prime}29^{\prime\prime}\mathrm{E}`、N/S/E/W 标注、hPa/mm/m³/s、
比例尺 `1:100000`、盐度 `\text{‰}`、`hm^{2}` 公顷、`\mathrm{Ma}` 地质年代。
识别侧补 ♀♂‱≈∝（UNI2TEX 814 条）。

**测试脚本的坑**：python 脚本里写 `r'\u2640'` 是**字面 6 字符**不转 Unicode，
测特殊字符要直接把 ♀ 写进脚本（文件存 UTF-8）。

**微积分·场论**：`\iint ∬` `\iiint ∭` `\oint ∮` `\oiint ∯` `\oiiint ∰` 全部输出
正确的 U+222C–2230 码位（`m:nary` 原生结构）；`\varoiint`/`\varoiiint` 会**静默泄漏**，
premap 已重写为 `\oiint`/`\oiiint`。四重积分无原生命令，用 `\int\!\!\!\!\int`。
`\lim_{}` 走 m:sSub（右下角），`\lim\limits_{}` 走 `m:limLow`（正下方）——试卷排版用后者。
`\frac{\partial^{n}z}{\partial x^{n}}` 偏导、`\nabla` 梯度/散度/旋度、`\dot{x}` `\ddot{x}`
牛顿记法（m:acc）、`\Gamma(n+1)=n!`。

**化学公式**（45/45 通过，全程零泄漏零 premap）：
反应条件 `\overset{\text{点燃}}{\rightarrow}`（m:limUpp，中文条件原生支持）、
`\xrightarrow{\text{高温}}`、等号上写催化剂 `\overset{\mathrm{MnO_{2}}}{=}`；
可逆 `\rightleftharpoons ⇌`（U+21CC），双向条件 `\overset{}{\underset{}{\rightleftharpoons}}`；
离子 `SO_{4}^{2-}`（m:sSubSup）、同位素 `^{23}_{11}Na`、结晶水 `\cdot`；
气体 `\uparrow ↑`、沉淀 `\downarrow ↓`（U+2191/2193）；
热化学 `\Delta H=-92.4\,\mathrm{kJ/mol}`、`\lg`、`\varphi^{\ominus}`（U+2296）、
`100\%`。识别侧补 ↑↓·⟶⇀↽← 六码位（UNI2TEX 811 条）。

**线性代数**：`\begin{vmatrix}...\end{vmatrix}` 行列式（`m:m`+`m:d`） `\begin{bmatrix}` 方阵
`A^{T}` 转置 `\det` `\operatorname{r}(A)` 秩 `\otimes ⊗` `\oplus ⊕` `\Vert ‖` 范数 `\left\|\vec{a}\right\|`

**统计与排列组合**：`\binom{n}{k}` 二项式（`m:d`+`m:f` 真堆叠，非文本模拟） `A_n^k` `C_n^k`
`\bar{x}` `\chi^{2}` `X\sim N(\mu,\sigma^{2})` `P(A\mid B)`

**数论**：`a\mid b` `a\nmid b` `a\equiv b\pmod{m}` `2^{100}\bmod 7` `\gcd(a,b)`
`\lfloor x\rfloor` `\lceil x\rceil`

**平面几何**：`\odot O` 圆 `\widehat{AB}` 弧（m:acc） `\overset{\frown}{AB}` 弧（⌒形）
`\nparallel ∦` `\triangle ABC\cong\triangle DEF`

**单位符号——预映射命令**（latex2mathml 对下列命令**静默失败**：不报错、把命令名
原样吐进公式。`premap_latex()` 在转换前自动重写，已在 `latex_to_mathml/latex_to_omml` 入口生效）：

| 扩展命令 | 重写为 | 用途 |
|---|---|---|
| `\celsius` / `\degreeCelsius` | `^{\circ}\mathrm{C}` | 摄氏度 |
| `\degree` | `^{\circ}` | 度 |
| `\arcmin` / `\arcsec` | `^{\prime}` / `^{\prime\prime}` | 角分/角秒 |
| `\ohm` | `\Omega` | 欧姆 |
| `\micro` | `\mu` | 微（μF 等） |
| `\permil` | `\text{‰}` | 千分号 |
| `\AA` | `\text{Å}` | 埃 |
| `\varoiint` / `\varoiiint` | `\oiint` / `\oiiint` | 曲面/体积分（预映射） |
| `\unit{m/s^2}` / `\si{...}` | `\mathrm{...}` | 单位正文化（支持一层嵌套花括号） |

**premap 正则的坑**：模式里的 `\b` 在 `\varoiint_` 处失配（`_` 也是词字符），
负向断言用 `(?![A-Za-z])` 代替 `\b`。

识别侧（`pdf2docx_math.py` 的 `UNI2TEX`，现 805 条）同步扩充：⊗⊕⊖⊘‖∣∤∦△∆⊙⌒∽∡、
数集 ℂℍℕℙℚℝℤ、℃℉‰Å、Ohm 兼容码位 U+2126、微符号 U+00B5、
微积分 ∂∇∬∭∮∯∰⨌⇌⇋⋁⋀⋂⋃，
以及 **U+1D400–1D7FF 数学字母数字变体**全块（𝑥 𝐴 𝒩 等斜体/粗体/花体还原为普通字母，
含 Unicode 标准洞位表——WPS 导出的 PDF 常用这些变体码位）。

新增命令时先跑 `pdfwork/symbol_audit.py` 风格的审计：`latex2mathml` 不抛异常 ≠ 转换正确，
必须检查 OMML 的**结构标记**（m:m/m:d/m:f/m:acc）和**码位**（`\perp` 出过 ⟂→⊥ 映射错误），
静默失败要靠「命令名是否泄漏进 m:t」来抓。

### 填空下划线（**注意：`\underline{空格}` 不要用**）

`\underline{\qquad}` 看上去是标准写法，实测**不能用**（2026-09-13 复测）：

| 写法 | OMML 结果 | 能否用 |
|---|---|---|
| `$\underline{\qquad}$`（单独成式） | **空公式**（`m:t` 数 = 0），Word 里一片空白 | ❌ 静默失败 |
| `$a\ \underline{\qquad}\ \{a,b,c\}$` | 有文本但**没有 `m:uLn` 元素** → 下划线丢失，只剩空格 | ❌ 效果错 |
| `\underline{AB}`（有实字符） | 正常输出 `AB`，但下划线仍不一定保留 | ⚠️ 仅当字符必填时可用 |

**填空占位该怎么做**：不要用 `\underline{}`，直接在**正文文本**里写下划线
（`______`，或按 坑#7 用「U+3000 空格 + underline 格式」的 run），
数学公式部分只放真正的式子：

```markdown
则 $A\cup B$ 的元素个数为 ______．
```

批量化自检同样适用坑#7：凡是 LaTeX 里含字母/数字、但 OMML 内 `m:t` 为空的，
一律报出来 —— `\underline{\qquad}` 正是靠这条才能被发现。

---

## 从 PPT / PDF 里抢救公式（识别侧）

**重要前提**：很多课件的公式不是 MathType 对象，而是被存成 **EMF/WMF 图元图片**。
这类公式**无法直接从 XML 提取**，必须先渲染成图片，再用多模态模型识别成 LaTeX。

### 判断公式类型

```bash
python -c "
import zipfile, re
z = zipfile.ZipFile('input.pptx')
for n in z.namelist():
    if re.match(r'ppt/slides/slide\d+\.xml$', n):
        x = z.read(n).decode('utf8','ignore')
        print(n, '| OMML原生公式:', 'YES' if 'oMath' in x else 'no')
"
```

- **OMML 原生**：直接读 `<m:oMath>` 拿结构，无需 OCR
- **EMF 图元**：走下面的图片路线

### EMF 图元 → 图片

用 `scripts/emf2png.py`（纯 GDI+ ctypes，不依赖 ImageMagick）：

```bash
python scripts/emf2png.py emf_dir/ png_dir/
```

### 缩略图（关键步骤：让模型能看清图）

EMF 图元转换出来的 PNG **尺寸夸张**（常上万像素），多模态模型读不动。
先用 `mkthumb.py` 缩到 1000 宽以内：

```python
# mkthumb.py —— 把上万像素的大 PNG 缩到模型能读的尺寸
import ctypes, os, sys
from ctypes import wintypes, byref, c_void_p, c_uint32, c_uint16, c_ubyte

# （完整源码见 scripts/mkthumb.py）
# 用法: python mkthumb.py png_dir/ thumb_dir/
# 自动缩到 max_w=1000, max_h=420
# 同时打印每张图的「宽高比」，>2.6 标为「可能跨行」
```

读图时建议直接看缩略图，能保留视觉特征但尺寸合理。

### 完整工作流（已用 3.1.2 + 4.1.1 验证，52/52 通过）

```
docx/PDF 
  ↓ emf2png.py 抽图元
EMF 矢量图 
  ↓ mkthumb.py 缩到 1000 宽
缩略图 
  ↓ 多模态模型读图
LaTeX 文本 → 写入 recognition.json
  ↓ verify_all.py
对照 docx（原图 + LaTeX + OMML 真公式）
```

**关键判断**：宽高比 > 2.6 的几乎都是跨行大括号或并列根号；
正方形 / 1.2~1.5 倍的通常是普通表达式；
高瘦（< 0.6）的可能是嵌套根号或大分式。

### 识别模板（参考本次真实样本）

| 类型 | 图元特征 | LaTeX 写法 |
|---|---|---|
| 集合构造 | 横向 4~5 倍 | `\{x\mid -2\leq x\leq -1\ \text{或}\ 0\leq x\leq 2\}` |
| 标准 cases | 横向 3 倍 | `\begin{cases}x^{2}-x, & x>1 \\ f(x+1)-1, & x\leq 1\end{cases}` |
| cases 嵌套方括号 | 横向 3 倍 | 在 cases 行内加 `f\!\left[f(x+6)\right]` |
| 嵌套根号（并列） | 高瘦 0.6 | `\sqrt{(a-1)^{2}+\sqrt[3]{a^{3}}}` |
| 嵌套根号（真套娃） | 横向 2.5 | `\sqrt[3]{a^{\frac{3}{2}}\sqrt{a^{-3}}}` |
| 根号套分数 | 高瘦 1.0 | `\sqrt{\frac{a}{b}}` |
| 分母有理化 | 横向 3 倍 | `\sqrt{3-2\sqrt{2}}=`（外根号套内根号） |
| 带分数 + 指数 | 横向 2.5 | `\left(5\frac{4}{9}\right)^{0.5}` |

### 识别侧的已知限制

- **当前 MiniMax-M3 模型支持看图**（2026-09 实测可用，9-01 那次被过滤是偶然）
  → 直接 `Read` 缩略图就能读公式
- 如果模型不支持读图，识别侧需要**切换到多模态模型**（Claude / GPT-4V / Gemini）
- 生成侧（LaTeX → Word）**不依赖任何视觉能力**，随时可用
- 第三方 OCR（如扫描全能王）需要云端 OAuth 授权，讲义内容会上传

## 矢量文本型 PDF → 可编辑 Word（无需 OCR）

上面那节是「公式被存成图片」的抢救方案，需要多模态模型。
但**绝大多数 WPS / Word 导出的 PDF 是矢量文本型的** —— 文字、符号全在 PDF 里，
根本不用 OCR，直接由**字符几何**反推公式结构即可，准确率高得多。

### 先判断是不是矢量文本型

```bash
python -c "
import pdfplumber
pdf = pdfplumber.open('input.pdf')
print('页数', len(pdf.pages), '| 字体',
      sorted({c['fontname'].split('+')[-1] for p in pdf.pages for c in p.chars}))
print('总字符', sum(len(p.chars) for p in pdf.pages))
"
```

- 字符数接近版面文字量、且能提取出中文 → **矢量文本型，走本节方案**
- 字符数接近 0、每页一张大图 → 扫描件，走上一节的多模态路线

### 一条命令转换

```bash
python scripts/pdf2docx_math.py input.pdf -o output.docx --dump
python scripts/verify_docx.py output.docx     # 校验：公式数/残留/zip 完整性
```

`--dump` 会把每页还原结果打印成 `正文$LaTeX$正文` 的形式，方便核对。

### 还原原理（六步）

| 步骤 | 做法 | 关键判据 |
|---|---|---|
| 1 行聚类 | 按**垂直中心**聚类 | 容差 5pt。**不能用 top** —— 同一行里 SymbolMT / TNR / 汉字 的 top 能差 1pt，按 top 会把一行切碎 |
| 2 上下标 | 小字（`< 主字号 × 0.75`）+ **底边 vs 基线** | 底边 ≤ 基线−1 → 上标；底边 ≥ 基线+0.5 → 下标。用**中心**判断会错（`x₀` 的中心反而略高于基线） |
| 3 分数 | 找短水平描边（3–60pt） | 上方 15pt 内、下方 15pt 内各有字符 → `\frac{}{}` |
| 4 根号 | 水平描边 + 左端有向下勾 | `∃s: s.x1 == 横线.x0 且 s.底 > 横线.y+3` → 横线下方、x 在横线范围内的字符包进 `\sqrt{}` |
| 5 定界符 | 同一 x 上纵向堆叠的**拼接件**（按连续段切分） | Symbol 字体 0xE6–0xF9 是可伸缩括号的零件；**raw 码位是 U+F0xx，须先减 0xF000 再查 `PIECE_CLASS`**（空 ToUnicode 的零件只能靠码位分清圆/方/花括号）。同一列可能有多个独立括号（如两个大括号组），要按纵向连续性**切段**再成堆叠。同型配对包 `\left( \right)` 等；落在定界符 x 范围外的下标要移到括号外（`[f(x)]_{min}` 不是 `[f(x)_{min}]`） |
| 6 表格 | 长度 > 60pt 的横/竖网格线 | 按 x 区间分组，行边界 = 横线 y 序列，列边界 = 竖线 x 序列 |
| 7 大括号分组 | 未配对大括号纵向覆盖 ≥2 行 → 手工构造 `m:d(begChr={) + m:eqArr` | latex2mathml 不支持 `cases`（报 regex 错），且 `latex_to_omml` 返回 **bytes** 要先 decode。成员行 = 紧贴括号右侧（x0 ∈ [括号x−2, 括号x+40]）；标签行（如"①恒成立问题："）整体在括号左侧（x1 ≤ 括号x+2），并入组段落作前缀文本 |
| 8 版式对齐 | 公式字号 = 行内数学字符字号众数（PDF 公式通常比正文大 1–2 号）；正文 = 非数学字符众数；标题判定**只看非数学字符**的最大字号（定义行里 16pt 的显示用 `∀` 不该把整行撑成标题）；页边距收窄到 ~80pt（PDF 版心 ~420pt > Word 默认 415.6pt，否则长公式跨行） | 中文 eastAsia 字体跟 PDF（微软雅黑），颜色显式黑色（防 Heading 样式主题色渗入） |

### 踩过的坑（按顺序，都实际发生过）

1. **SymbolMT 私有区编码**：WPS 导出的 PDF 里 SymbolMT 的 ToUnicode 把符号映射到
   `U+F0xx` 私有区。必须二次映射 `Symbol码位 = 私有区码位 - 0xF000`，再查 Symbol 字体表。
   已验证：`U+F022→∀`、`U+F024→∃`、`U+F061→α`、`U+F0A3→≤`、`U+F0B3→≥`、`U+F0CE→∈`、`U+F0D8→¬`。
   完整映射表在 `scripts/pdf2docx_math.py` 的 `SYM`，**含 0xE0–0xFF 可伸缩括号段**（漏了这段，
   `[` `]` `{` 会变成空字符串，进而无法识别定界符）。
2. **定界符误判**：用 `ord(c) >= 0x2500` 筛选「拼接件」会把**所有汉字**（U+4E00+）圈进来，
   而每行行首 x 坐标相同 → 被误判成「纵向堆叠」→ 满屏 `\left[]{}`。
   必须只认 `DELIM_PIECES` 白名单，并校验**纵向连续性**（相邻件空隙 ≤ 4pt）。
3. **上下标污染行锚点**：上下标字符比主行高/低，若参与行聚类会拉偏锚点，把一行劈成两行。
   做法：先用**非上下标字符**定行，再把上下标按基线就近归行。
4. **LaTeX 命令粘连**：`\alpha` 后面紧跟字母会被解析成 `\alphax`。
   凡命令名结尾是字母的，一律补一个空格。
5. **纯数字不该进公式**：`§1.5`、`（1）`、`能被2整除` 里的数字是正文。
   数学段里若既无字母、也无 `∀∃∈≤≥` 等强制符号 → 当正文处理。
6. **`{x | ...}` 的竖杠可能不是字符**：常被画成**矢量竖线**（高 14pt 的描边）而非 `|` 字形。
   要单独检测纵向描边（高 6–40pt、非表格竖线）并插入 `\mid`，否则竖杠直接丢失。
7. **填空横线**要按 x 坐标插回行内（`是____.`），不能一律追加到行尾 —— 行尾可能还有句号。
   呈现用「空格 + 下划线格式」（U+3000 全角空格 run 设 `underline=True`），不要用 `＿` 字符——
   下划线字符是实心黑线、字距发闷，下划线格式的空格才和试卷印刷效果一致。空格数按横线实测宽度换算。
8. **统计 `m:oMath` 数量**要用 `body.iter(...)`，`body.findall(...)` 只找直接子节点，会永远返回 0。
9. **不要按 PDF 页面插硬分页**（`doc.add_page_break()`）：PDF 分页位置 ≠ Word 自然分页位置，
   硬分页会把内容拦腰截断、留下半页空白。版面交给 Word 自动分页，段落间距用样式控制。
10. **表格之间的间隔段要压扁**（段前/段后 0 + 段落标记字号 3pt）：既防止相邻表格粘连，
    又不会因空段把后续内容挤出尴尬的分页。PDF 里无内容的空行一律跳过不写段。

### 已知边界

- 保存目标若在 Word 中打开会 PermissionError，脚本自动另存为 `*_new.docx`
- 跨页表格、旋转文本、真正的扫描件不在覆盖范围内

## MathType 文档抢救（MTEF 解析）

老讲义里的公式常常是 **MathType OLE 对象** —— 双击能编辑，但批量处理时
既不是图片也不是文本，常规手段碰不到。这一节讲怎么**无损还原成 LaTeX**，
再走主链路变成 Word 原生公式。

### 双层存储（关键定位，踩过坑）

MathType 对象在 docx 里**分两处存**，用途完全不同：

| 位置 | 内容 | 能拿什么 |
|---|---|---|
| `word/embeddings/*.bin` | OLE 复合文档 → `Equation Native` 流 → **MTEF 二进制** | ✅ **完整公式结构**（主路径） |
| `word/media/*.wmf` | 呈现图（矢量），文字原样在 `META_EXTTEXTOUT` | 符号清单（辅助校验） |

**不要**去 `word/embeddings/` 找 WMF —— 那里是 OLE 包，直接搜 WMF 头搜不到。

**判断是不是 MathType OLE**：

```bash
python -c "
import zipfile
z=zipfile.ZipFile('input.docx')
wmf=[n for n in z.namelist() if n.startswith('word/media/') and n.lower().endswith('.wmf')]
emb=[n for n in z.namelist() if n.startswith('word/embeddings/')]
x=z.read('word/document.xml').decode('utf8')
print('WMF呈现图:', len(wmf), '| OLE嵌入:', len(emb), '| OMML原生公式:', x.count('oMath'))
"
```

`OLE嵌入 > 0` 且 `OMML原生公式 == 0` 就是纯 MathType 文档，走本节流程。

### 主路径：MTEF → LaTeX

```bash
python scripts/mtef.py "6.2.6n.docx" --json eqs.json
```

输出示例（真实样本，高中数学平面向量章节）：

```
oleObject9.bin   a\cdot b=\left|a\right|cos\theta \cdot \left|b\right|
oleObject45.bin  \vec{a}\perp \vec{b}\Leftrightarrow \left|\vec{a}+\vec{b}\right|=\left|\vec{a}-\vec{b}\right|
oleObject50.bin  \left|\vec{a}\right|=\left|\vec{b}\right|=2,\left\langle \vec{a},\vec{b}\right\rangle =120^{\circ}
oleObject53.bin  \left|\vec{b}\right|=2
```

Python API：

```python
import sys
sys.path.insert(0, '<skill目录>/scripts')
from mtef import equations_from_docx

for r in equations_from_docx('6.2.6n.docx'):
    print(r['part'], '->', r['latex'] or r['error'])
```

### 端到端：整份讲义换成原生公式

```bash
python scripts/mtef2docx.py "6.2.6n.docx" --apply
```

做的事：找到 `<w:object>` 里的 `<o:OLEObject>`（按 ProgID 认 MathType）→ 顺着
r:id 取 embedding 字节 → MTEF 解出 LaTeX → 转 OMML → **原地替换 w:object**。

**安全设计**（跟 `upgrade_math.py` 一致）：

- 默认**只读体检**，加 `--apply` 才写文件
- **不覆盖原文件**，输出到 `原名_原生公式.docx` 或 `-o` / `--outdir`
- 解析或转换失败的公式**保留原 MathType 对象**，不会把公式弄丢
- 只动 `w:object` 元素，周围文字、格式、样式一概不动

**实测**（平面向量讲义 6.2.6n.docx）：

| 指标 | 结果 |
|---|---|
| MathType 对象 | 59 → 原生 `m:oMath` 59，残留 OLE 0 |
| 段落数 | 47 → 47 |
| 正文中文字符 | 451 → 451（**一字不差**） |
| 文件体积 | 显著缩小（59 个 WMF + 59 个 OLE 包全去掉了） |

### 辅助路径：WMF → 符号清单（用来校验）

```bash
python scripts/wmf_symbols.py "6.2.6n.docx" --json symbols.json
```

自动把 `word/media/*.wmf` 抽到 `6.2.6n_wmf/` 再解析，输出每个图元用了哪些字体
（Symbol / MT Extra / **Euclid Extra** / Cambria Math）以及对应符号。

**WMF 侧的硬限制**：MathType 输出文字记录时**按字体分组**
（如 `|a||b|cosθ` 会输出成 `coscos` / `ababba` / `qq` / `×=×=×` 四组），
而且**可伸缩的竖线 `|`、大括号是用绘图图元（LineTo）画的、不在文字记录里**。
所以它只能回答"用了哪些符号"，**不能还原结构** —— 只配当 MTEF 结果的校验器。

### 实测结果

**单份文档**（高中数学「平面向量」讲义 `6.2.6n.docx`）：

- **59 / 59** 个公式解析成功，结构全对（分数、根号、上下标、向量、绝对值 fence、
  点乘、垂直、等价、希腊字母、角度全部正确）
- 与 WMF 呈现图**交叉对账 61 项特征、0 缺检（100% 一致）**

**全库压测**（214 份讲义、8521 个 MathType 对象）：

| 类别 | 数量 | 说明 |
|---|---|---|
| MTEF v5（MathType） | 8387 | **解析失败 0，成功率 100%** |
| MTEF v3（公式编辑器 3.0） | 24 | 老格式，不支持（见下） |
| docx 内部 zip 损坏 | 122 | 文件本身坏了，跟解析器无关 |

官方 SDK 示例流 `-b±√(b²-4ac)/2a` 自检 PASS。

### 三种取 LaTeX 的来源（按可靠度排序）

1. **FUTURE 0x66「TeX Input Language」** —— MathType 6+ 把作者当初敲的
   **原始 TeX 源码**存在这里（`equations_from_docx` 返回的 `tex_input` 字段）。
   没有经过任何启发式猜测，**最准**。不是每个公式都有，手写方式建的就没有。
   `mtef2docx.py` 会优先用它，转不成 OMML 才回退。
2. **MTEF 结构翻译** —— 从对象树翻译回来，覆盖所有 v5 公式，少数地方有启发式。
3. **WMF 呈现图** —— 只能提符号清单，没有结构，仅作校验用。

### 已知限制：MTEF v3（微软公式编辑器 3.0）

v3 是 Office 自带的老公式编辑器格式，记录布局跟 MathType v5 不是一回事
（版本头没有 app key，记录字段也不同），本解析器**不支持**。

识别方法很简单 —— 看版本头第一个字节：

```
raw[28] == 5   -> MathType v5，本解析器能解
raw[28] == 3   -> 公式编辑器 3.0，不支持
```

**怎么救**：用 Word 打开该文档 → 双击公式 → 「公式 → 转换为新格式」
（或全选后「公式 → 全部转换」）→ 存盘。转换后就变成 MathType 兼容格式了。

实测全库 8521 个对象里只有 24 个是 v3（集中在 7 份文档），影响很小。

### MTEF 格式要点（全是踩过的坑）

实现在 `scripts/mtef.py`，改代码前务必先读：

1. `Equation Native` = 28 字节 OLE 头 + **12 字节 MathType 版本头** + MTEF 记录体。
   喂给解析器前必须**剥掉前 40 字节**（自检第一版就是忘了剥，报"未知记录类型 82"）。
2. 每条记录 = 1 字节 type + 1 字节 options，但**不是所有记录都有 options 字节**
   （见 `NO_OPTION` 集合）。
3. 整数编码三套：signed int = 单字节(`value+128`) 或 `FF`+2字节(`value+32768`)；
   unsigned int = 单字节 或 `FF`+2字节；simple 16-bit = 固定 2 字节小端。
4. NUDGE(0x08)：2 字节(`dx+128,dy+128`) 或 6 字节(`80 80` + 两个 16-bit)。
5. **`EQN_PREFS(18)` 不以 END 结尾**，长度由内部三个数组决定：
   sizes(dim array) + spaces(dim array) + styles(array)。
   dim array = count 字节 + **nibble 流**，每个 dimension = `[units nibble][数字nibble...][0xF 终止]`。
   这里错一个字节，后面全部错位。
6. TMPL selector 0–37；variation 支持 1/2 字节扩展（高位 `0x80` 是续读标志）。
7. **slot 顺序反直觉**：
   - Frac = `[分子, 分母]`
   - Root = `[被开方数, 根指数]`（根指数在后！）
   - Scr = `[下标, 上标]`（下标在前！）
   - BigOp = `[主体, 上限, 下限, 算子CHAR]`
8. **`LINE_NULL(0x01)` 的 LINE 连 END 对象列表都没有**，不能按常规读。
9. MTCode 本身就是 Unicode 码位，**优先于字体码位**。
10. MathType 私有区码位（`typeface=22 fnEXPAND`）是**排版部件**，不是字符：
    `0xEC07/0xEC08` 是可伸缩竖线的左右半边。已登记的在 `MTCODE_LATEX` 里，
    未登记的（`0xE000–0xF8FF`）直接丢弃，否则会输出这种乱码：`\ue007a`。

### OLE 解包：olefile 可选，装不上也能跑

`mtef.py` 解 OLE 复合文档时**优先用 `olefile`，没装就自动回退到内置的 `_CFBF` 类**
（纯标准库，约 100 行）。实测两条路径在同一份文档的 59 个公式上**逐字节一致**。

`requirements.txt` 里写了 `olefile`，但它是可选的 —— 换机器装不上依赖时，
MathType 功能不会全瘫。

自己实现 CFBF 时有两个必踩的坑：

1. **`_read_chain` 返回的是字节串，不是数组。** FAT / MINIFAT 必须 `struct.unpack`
   成 uint32 列表才能当索引用。直接拿 `bytes` 下标会取到单字节（0–255），
   既永远小于 `FREESECT(0xFFFFFFF0)`、又指向错误扇区 —— 症状是**死循环**，
   不报错，很难查（用 `faulthandler.dump_traceback_later()` 才抓到）。
2. **Equation Native 流一律走迷你流。** CFBF 里 <4096B 的流放 mini stream，
   靠 mini FAT（64B/扇区）串起来，不是普通 FAT（512B/扇区）。只实现普通 FAT 解不出来。
   实测 59 个流全是 254–427 字节，无一例外。

**最容易翻车的一处 —— 「向量的模」**：

MathType 把 $|\vec{b}|$ 存成 `VEC(BAR(b))`，按字面翻译会得到
`\overrightarrow{\left|b\right|}`（毫无数学意义）。`tmpl_latex()` 里有启发式修正：
当 `VEC` 内部是**单个 fence 模板**时，把箭头挪进 fence 里面，输出 `\left|\vec{b}\right|`。

判据是呈现图里箭头只有 2 段（只够盖住单个字符），说明渲染的确实是 $|\vec{b}|$。
实测这个启发式在 59 个公式里**只命中 1 个**（`oleObject53`），是安全的。

**另外两处翻车点**：

11. **FUTURE 记录（type ≥ 100）没有 options 字节**，格式是 `type + 长度 + 数据`。
    必须在 `read_record` 最前面就分流。放到通用逻辑里会把**长度字节当成 options**，
    而且长度字节只要 bit3 为 1（如 `0x1e`）就会被误判成 `OPT_NUDGE` 再吞 2 字节 ——
    从这条记录起后面全错位，报出来的错误（「未知类型 0x43」）离真正的病root 很远。
    **就是这一条，让全库成功率从 95.1% 掉到那里。**

12. **括号的 variation 位不可靠，直接读 slot 里的实际字符。**
    `INTERVAL`（selector 9）不在 `sel <= tmOBRACK(8)` 范围内，而且它的 variation
    编码（实测 `0x30` / `0x12`）跟 `BRACK`（`0x3`）**不是同一套位定义**。
    按位判断的结果是把 `(` 和 `]` 当成普通内容拼到末尾，输出成 `-\infty ,3(]`。
    正确做法：slot1/slot2 非空且字符在括号白名单里，就无条件当作开闭符号。

---

## 常见坑

1. **`\subsetneq` 和 `\subsetneqq` 不同**
   `\subsetneq` 是 ⫋ 的简写变体，`\subsetneqq` 是标准真子集 ⫋。教材通用后者。

2. **花括号必须转义**
   写 `\{x\mid x>0\}`，不是 `{x|x>0}`。后者在 LaTeX 里是分组语法，不显示。

3. **集合构造的竖线用 `\mid`**
   `\mid` 渲染成 ∣（U+2223），比直接打 `|` 间距更规范。

4. **中英文混排**
   公式里的中文用 `\text{是平行四边形}` 包起来，否则会被当成斜体变量。

5. **不要修改 Word 公式的字体**
   原生公式默认用 **Cambria Math**，改成其他字体会导致符号缺失变豆腐块。

6. **转换失败会自动降级**
   单个公式出错时保留 `[原始LaTeX]` 文本并在 stderr 报警，**不会中断整个文档**。
   看到方括号内容就说明那个公式没转成，去查 stderr。

7. **相邻的绝对值 / 竖线 fence 会静默变空公式（2026-09 实测）**
   `|a_{i}||a_{j}|`、`|a||b|` 这类**两段 `|...|` 直接相连**的写法，
   `latex_to_omml` **不报错**，但产出的 `m:oMath` 是空壳（没有任何 `m:t` 文本），
   在 Word 里就是一片空白 —— 降级机制完全抓不到它，只能靠回读校验发现。
   - **正确写法**：用 `\left|...\right|` 成对包裹，即
     `\left|a_{i}\right|\left|a_{j}\right|`（实测输出 `aiaj`，正常）。
   - **无效的绕过**：在中间插 `\,` `\;` `~` 都没用，一样是空的。
   - 单个 `|a|`、被运算符隔开的 `|x|+|y|`、`|a|<|b|` 都正常，只有**紧邻**才翻车。
   - **同类静默空公式还有 `\underline{\qquad}`**（2026-09-13 实测）：
     `$\underline{\qquad}$` 单独成式时 `m:t` 数为 0，Word 里直接是空白；
     写进句子里（`$a\ \underline{\qquad}\ b$`）虽不报错，但**不生成 `m:uLn`**、
     下划线丢失。填空横线请改在正文文本里写（见「填空下划线」一节）。
   - **批量自检建议**：转完后统一回读一遍，凡 LaTeX 里含字母/数字但
     OMML 内 `m:t` 文本为空的一律报出来（`verify_docx.py` 的「空公式对象」就是查这个）。

8. **分段函数/不等式组右边多出一个撑大的 `)` —— `m:endChr` 被删了（2026-09 实测，已修复）**

   症状：`\begin{cases}...\end{cases}` 在 Word 里左边是正常的大括号 `{`，
   **右边却凭空多出一个和公式一样高的右圆括号 `)`**。

   根因在 OMML 的默认值规则（ECMA-376）：

   | | m:begChr 默认 | m:endChr 默认 |
   |---|---|---|
   | | `(` | `)` |

   MML2OMML.XSL 的规矩是「**定界符等于默认值时才省略**」，所以三种输出的含义完全不同：

   ```
   \left( ... \right)   -> 不写 begChr/endChr         渲染 '(' ')'    正确
   A=[-2,4)             -> begChr='' + 不写 endChr    渲染 ''  ')'    正确（半开区间）
   \begin{cases}        -> begChr='{' + endChr=''     渲染 '{'  无     正确
   ```

   旧版 `latex2omml.postprocess_omml` 里有一段"修复"，为了「让 Word 不画右括号」
   把 cases 的 **`<m:endChr m:val=""/>` 删掉了** —— 删掉 == 没声明 == Word 套用默认 `)`。
   这个"修复"就是 bug 本身，已删除（2026-09-13）。**那个元素绝不能动。**

   - **识别它**：`python scripts/check_omml_delims.py <文件.docx>`，判据只有两条 ——
     `begChr 非空且 endChr 缺失`、`endChr 非空且 begChr 缺失`，都是必坏。
     （`begChr=''` + 无 endChr 是**正常**的半开区间写法，别误报。）
     `--selftest` 带负例测试：故意复现旧 bug，检测器必须抓到。
   - **修复**：改 `latex2omml.py` 的 `postprocess_omml`（别在生成后去删元素），
     改完必须重跑 `--selftest` + 重建文档 + 对照旧版 docx 确认可疑数归零。
   - **实测影响面**：本次一册汇编里 4 处中招（2015AB10、2006\*3 两组不等式组、
     2004\*三 的 f(n) 分段函数），另 176 个定界符组不受影响。
     所以**不要以为只有一处**，改完必须全库扫。

9. **目标 docx 被 Word/WPS 打开时保存会 `PermissionError`**
   Word 会生成 `~$xxx.docx` 锁文件并拒绝写入、**连改名替换也会被拒**。
   构建脚本要带兜底：写 `.new` → `os.replace` → 再把旧文件改名成 `.old` 顶上；
   全都失败就保留 `.new` 并明确告知用户「合上 Word 后改名回去」。
   **别默默失败**（那样用户看到的还是旧文件，会以为修复没生效）。

---

## 文件说明

```
latex-omml/
├── SKILL.md                   本文件
├── INSTALL.md                 安装说明（换机器看这个）
├── install.cmd                Windows 一键安装（建 .venv + 装依赖 + 自检）
├── requirements.txt           依赖清单：lxml / latex2mathml / python-docx
├── scripts/
│   ├── latex2omml.py          核心引擎（LaTeX→OMML，含自检）
│   ├── md2docx.py             CLI：Markdown+LaTeX → docx
│   ├── upgrade_math.py        CLI：已有 docx 的伪公式批量升级为真公式
│   ├── audit_math.py          CLI：体检 docx，统计真公式/图片公式/伪公式
│   ├── emf2png.py             EMF/WMF 图元 → PNG（识别侧用）
│   ├── mtef.py                CLI：MathType OLE → MTEF v5 解析 → LaTeX（含自检）
│   ├── mtef2docx.py           CLI：整份讲义的 MathType 对象批量换成原生公式
│   ├── wmf_symbols.py         CLI：从 MathType WMF 图元精确提取符号（不靠 OCR）
│   ├── pdf2docx_math.py       CLI：矢量文本型 PDF → docx，公式还原为原生 OMML
│   ├── verify_docx.py         校验生成结果：公式数 / 降级残留 / zip 完整性
│   ├── check_omml_delims.py   体检/回归自测：定界符组 m:d 的 begChr/endChr 默认值陷阱
│   └── MML2OMML.XSL           微软官方转换表（内置，194KB）
└── references/
    └── symbols.md             数学符号 LaTeX 速查表
```

**典型工作流**（已有讲义改造）：

```bash
# 1 体检：看清现状
python scripts/audit_math.py "C:\讲义目录"
# 2 试转：只读 + 抽样
python scripts/upgrade_math.py "C:\讲义目录" --sample 10
# 3 应用：输出到新目录，原文件不动
python scripts/upgrade_math.py "C:\讲义目录" --apply --outdir "C:\升级版"
# 4 复核：对比中文是否一致
python scripts/audit_math.py "C:\升级版"
```
