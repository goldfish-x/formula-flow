# 数学符号 LaTeX 速查表

配合 `latex-omml` skill 使用。所有符号均已实测可正确转换为 Word 原生公式。

---

## 一、集合与逻辑（高中数学高频）

| 符号 | LaTeX | Unicode | 说明 |
|---|---|---|---|
| ⊆ | `\subseteq` | U+2286 | 子集（包含于） |
| ⊇ | `\supseteq` | U+2287 | 超集（包含） |
| ⊂ | `\subset` | U+2282 | 子集（不强调相等） |
| ⊃ | `\supset` | U+2283 | 超集 |
| ⫋ | `\subsetneqq` | U+2ACB | **真子集**（教材标准写法） |
| ⫌ | `\supsetneqq` | U+2ACC | 真超集 |
| ⊊ | `\subsetneq` | U+228A | 真子集（变体，间距略窄） |
| ∅ | `\varnothing` | U+2205 | 空集（推荐） |
| ∅ | `\emptyset` | U+2205 | 空集（同义） |
| ∈ | `\in` | U+2208 | 属于 |
| ∉ | `\notin` | U+2209 | 不属于 |
| ∋ | `\ni` | U+220B | 包含（反向属于） |
| ∩ | `\cap` | U+2229 | 交集 |
| ∪ | `\cup` | U+222A | 并集 |
| ∁ | `\complement` | U+2201 | 补集 |
∣ | `\mid` | U+2223 | 集合构造分隔竖线 |
| ∅̸ | `\nexists` | U+2204 | 不存在 |
| ∀ | `\forall` | U+2200 | 任意 |
| ∃ | `\exists` | U+2203 | 存在 |
| ∴ | `\therefore` | U+2234 | 所以 |
| ∵ | `\because` | U+2235 | 因为 |
| ⇒ | `\Rightarrow` | U+21D2 | 推出 |
| ⇔ | `\Leftrightarrow` | U+21D4 | 等价 |
| → | `\to` | U+2192 | 趋于 / 映射 |

### 数集（黑板粗体）

| 符号 | LaTeX | Unicode |
|---|---|---|
| ℕ | `\mathbb{N}` | U+2115 |
| ℤ | `\mathbb{Z}` | U+2124 |
| ℚ | `\mathbb{Q}` | U+211A |
| ℝ | `\mathbb{R}` | U+211D |
| ℂ | `\mathbb{C}` | U+2102 |

### 花括号（必须转义）

```latex
\{x\mid x^{2}-1=0\}
\{1,2,3\}
\varnothing\neq\{0\}
```

> ⚠️ 直接写 `{x|x>0}` 不会显示花括号——`{}` 在 LaTeX 里是分组语法。

---

## 二、关系与运算

| 符号 | LaTeX | 符号 | LaTeX |
|---|---|---|---|
| = | `=` | ≠ | `\neq` |
| < | `<` | > | `>` |
| ≤ | `\leq` | ≥ | `\geq` |
| ⩽ | `\leqslant` | ⩾ | `\geqslant` |
| ≈ | `\approx` | ≡ | `\equiv` |
| ≅ | `\cong` | ∼ | `\sim` |
| ≪ | `\ll` | ≫ | `\gg` |
| ∝ | `\propto` | ≃ | `\simeq` |
| + | `+` | − | `-` |
| × | `\times` | ÷ | `\div` |
| ⋅ | `\cdot` | ± | `\pm` |
| ∓ | `\mp` | ∗ | `\ast` |
| ∘ | `\circ` | ⊕ | `\oplus` |

---

## 三、希腊字母

### 小写

| 符号 | LaTeX | 符号 | LaTeX | 符号 | LaTeX |
|---|---|---|---|---|---|
| α | `\alpha` | β | `\beta` | γ | `\gamma` |
| δ | `\delta` | ε | `\epsilon` | ε | `\varepsilon` |
| ζ | `\zeta` | η | `\eta` | θ | `\theta` |
| θ | `\vartheta` | ι | `\iota` | κ | `\kappa` |
| λ | `\lambda` | μ | `\mu` | ν | `\nu` |
| ξ | `\xi` | π | `\pi` | ρ | `\rho` |
| σ | `\sigma` | τ | `\tau` | υ | `\upsilon` |
| φ | `\phi` | φ | `\varphi` | χ | `\chi` |
| ψ | `\psi` | ω | `\omega` | | |

### 大写

| 符号 | LaTeX | 符号 | LaTeX | 符号 | LaTeX |
|---|---|---|---|---|---|
| Γ | `\Gamma` | Δ | `\Delta` | Θ | `\Theta` |
| Λ | `\Lambda` | Ξ | `\Xi` | Π | `\Pi` |
| Σ | `\Sigma` | Υ | `\Upsilon` | Φ | `\Phi` |
| Ψ | `\Psi` | Ω | `\Omega` | | |

---

## 四、结构类

### 分数与根式

```latex
\frac{a}{b}                 分数
\dfrac{a}{b}                大分数（独立显示用）
\tfrac{a}{b}                小分数（行内用）
\sqrt{x}                    平方根
\sqrt[3]{x}                 立方根
\sqrt[n]{x}                 n 次根
```

### 上下标

```latex
x^{2}                       上标
a_{i}                       下标
a_{i}^{j}                   上下标
x^{2n+1}                    多项式上标（记得加花括号）
\sum_{i=1}^{n}              求和上下限
\int_{0}^{1}                积分上下限
\lim_{x\to 0}               极限
```

> ⚠️ 多位上标必须加花括号：`x^2n` 会渲染成 `x²n`，`x^{2n}` 才是 `x²ⁿ`。

### 大型运算符

```latex
\sum_{i=1}^{n} a_{i}        求和
\prod_{i=1}^{n}             连乘
\int_{a}^{b} f(x)\,dx       定积分
\iint                       二重积分
\oint                       曲线积分
\bigcup_{i=1}^{n}           并集
\bigcap_{i=1}^{n}           交集
\lim_{x\to\infty}           极限
\max_{x\in D}               最大值
\min_{x\in D}               最小值
```

> 积分里的小间距用 `\,`，例如 `\int_{0}^{1}x^{2}\,dx`

### 矩阵与行列式

```latex
\begin{matrix}a & b \\ c & d\end{matrix}           无括号
\begin{pmatrix}a & b \\ c & d\end{pmatrix}         圆括号
\begin{bmatrix}a & b \\ c & d\end{bmatrix}         方括号
\begin{vmatrix}a & b \\ c & d\end{vmatrix}         行列式
\begin{Bmatrix}a & b \\ c & d\end{Bmatrix}         花括号
```

`&` 分列，`\\` 分行。

### 多行公式与方程组

```latex
\begin{cases}
x+y=1 \\
x-y=3
\end{cases}
```

### 装饰

```latex
\overline{AB}               上划线
\underline{x}               下划线
\overrightarrow{AB}         向量箭头
\widehat{x}                 上尖括号
\widetilde{x}               上波浪
\hat{a}                     上帽
\vec{a}                     向量
\dot{x}                     点（导数）
\ddot{x}                    双点（二阶导）
```

---

## 五、几何

| 符号 | LaTeX | 符号 | LaTeX |
|---|---|---|---|
| ∠ | `\angle` | ⊥ | `\perp` |
| ∥ | `\parallel` | △ | `\triangle` |
| ⊙ | `\odot` | ⌢ | `\frown` |
| ∘ | `^\circ` | ′ | `'` |
| ⌣ | `\smile` | ⊚ | `\circledcirc` |

**角度写法**：`30^\circ`，`90^\circ`，`\angle ABC = 60^\circ`

⚠️ 两个易混点（2026-11 实测校正）：

1. **度数不是 `°`**。`^\circ` 生成的是 `m:sSup`（底数 + 上标的 `∘` U+2218 RING
   OPERATOR）；直接写 `°` U+00B0 只落在基线上，是个普通度号，不是上标。两者
   Cambria Math 都有字形，但只有 `^\circ` 的结构是对的。
2. **`\frown`/`\smile` 曾在本表里写反**：`\frown` = `⌢` U+2322，`\smile` = `⌣` U+2323。
   `⌒` U+2312 是 ARC，跟这两个都不同；要画弧线用 `\overparen{AB}`（U+23DC `⏜`）。

---

## 六、中学常见完整公式模板

### 一元二次

```latex
x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}
\Delta=b^{2}-4ac
x_{1}+x_{2}=-\frac{b}{a},\quad x_{1}x_{2}=\frac{c}{a}
```

### 数列

```latex
a_{n}=a_{1}+(n-1)d
S_{n}=\frac{n(a_{1}+a_{n})}{2}=na_{1}+\frac{n(n-1)}{2}d
a_{n}=a_{1}q^{n-1}
S_{n}=\frac{a_{1}(1-q^{n})}{1-q}\quad(q\neq 1)
\sum_{i=1}^{n}i=\frac{n(n+1)}{2}
\sum_{i=1}^{n}i^{2}=\frac{n(n+1)(2n+1)}{6}
```

### 三角

```latex
\sin^{2}\alpha+\cos^{2}\alpha=1
\sin(\alpha+\beta)=\sin\alpha\cos\beta+\cos\alpha\sin\beta
\frac{a}{\sin A}=\frac{b}{\sin B}=\frac{c}{\sin C}=2R
c^{2}=a^{2}+b^{2}-2ab\cos C
```

### 集合运算

```latex
A\cap B=\{x\mid x\in A\text{ 且 }x\in B\}
A\cup B=\{x\mid x\in A\text{ 或 }x\in B\}
\complement_{U}A=\{x\mid x\in U\text{ 且 }x\notin A\}
\varnothing\subseteq A
A\cap\varnothing=\varnothing
```

### 函数与导数

```latex
f'(x)=\lim_{\Delta x\to 0}\frac{f(x+\Delta x)-f(x)}{\Delta x}
(x^{n})'=nx^{n-1}
(\sin x)'=\cos x
(\ln x)'=\frac{1}{x}
```

### 概率统计

```latex
P(A\cup B)=P(A)+P(B)-P(A\cap B)
\bar{x}=\frac{x_{1}+x_{2}+\cdots+x_{n}}{n}
s^{2}=\frac{1}{n}\sum_{i=1}^{n}(x_{i}-\bar{x})^{2}
```

---

## 七、中文混排

公式里出现中文要用 `\text{}` 包裹，否则会被渲染成斜体变量：

```latex
\{x\mid x\text{是平行四边形}\}
\{x\mid x\text{是 8 的约数}\}
A\cup B=\{x\mid x\in A\text{ 或 }x\in B\}
```

---

## 八、填空题下划线

讲义、试卷高频用法，实测转换正常：

```latex
a\ \underline{\qquad}\ \{a,b,c\}
\varnothing\ \underline{\qquad}\ \{x\in\mathbb{R}\mid x^{2}+1=0\}
```

`\underline{\qquad}` 生成一段可书写的空白下划线，`\ ` 是手动空格。

---

## 九、n 次方根与嵌套根号

已在必修一 4.1.1 真实讲义上验证通过（OMML 生成 `m:rad` 对象，Word 里可编辑）。

| 需求 | LaTeX | 说明 |
|---|---|---|
| 平方根 | `\sqrt{a}` | |
| n 次方根 | `\sqrt[n]{a}` | `∛a`、`⁵√a` 都用它 |
| 根号套括号 | `\sqrt{(a-1)^{2}}` | `\left(` `\right)` 让括号随内容变高 |
| **嵌套根号（并列）** | `\sqrt{(a-1)^{2}+\sqrt[3]{a^{3}}}` | 实测生成 rad ×2 ✓ |
| **嵌套根号（真套娃）** | `\sqrt[3]{\sqrt[4]{a}}` | 实测生成 rad ×2 嵌套 ✓ |
| 根号套分数 | `\sqrt{\frac{a}{b}}` | |
| 分数指数幂 | `a^{\frac{m}{n}}=\sqrt[n]{a^{m}}` | rad + f + sSup 组合 |
| 负分数指数 | `\left(\frac{8}{27}\right)^{-\frac{2}{3}}` | |

> **要点**：根号内的括号必须用 `\left(...\right)`，否则括号不会随根号高度自动拉伸，
> 排出来就是 `√((a-1)²+∛a³)` 里括号偏小的样子。

---

## 十、跨行大括号（分段函数 / 方程组）

用 `\begin{cases}...\end{cases}`。转换后 OMML 结构是
`m:d`（分隔符，`begChr='{'`、`endChr=''`）包一个 `m:m`（矩阵，每行一个 `m:mr`）——
**这正是 Word / MathType 里分段函数的标准内部表示**，所以排版效果和 MathType 一致。

### 两行

```latex
f(x)=\begin{cases}x^{2}, & x<4 \\ f(x-1), & x\geq 4\end{cases}
```

### 三行

```latex
f(x)=\begin{cases}-x+1, & x<1 \\ x^{2}-2x, & x\geq 1 \\ 0, & x=1\end{cases}
```

### 行内嵌方括号（已验证，外层 `{}` 内层 `[]` 互不干扰）

```latex
f(x)=\begin{cases}x-2, & x\geq 10 \\ f\!\left[f(x+6)\right], & x<10\end{cases}
```

> 这里 `m:d` 会出现**两次**：外层 `{}` 包矩阵，内层 `[]` 包 `f(x+6)`。

### 语法要点（错了会整段降级）

1. 换行用 `\\`，列分隔用 `&`
2. Markdown 文件里写 `\\`；Python 字符串里要写 `\\\\`（或直接用 `r'...'`）
3. 只写 `\begin{cases}` 不写 `\end{cases}` → 转换失败
4. 条件里的中文用 `\text{或}`、`\text{且}`，不要直接写中文
5. `f\!\left[...\right]` 里的 `\!` 是负空格，让 `f` 和 `[` 贴紧一点，更像教科书排版

### 集合构造式（不是 cases，别混）

```latex
\{x\mid -2\leq x\leq -1\ \text{或}\ 0\leq x\leq 2\}
\{x\mid x>-1\ \text{且}\ x\neq 2\}
```

用 `\{` `\}` + `\mid`，生成的是**单个** `m:d`（左右都有分隔符），结构比 cases 简单。

---

## 十一、排查：符号没出来怎么办

1. **检查是否拼写错误** — `\subseteq` 不是 `\subeateq`
2. **检查花括号转义** — `\{` `\}` 而不是 `{` `}`
3. **看 stderr 有没有降级警告** — 出现 `[原始LaTeX]` 说明转换失败
4. **跑自检定位** — `python scripts/latex2omml.py`
5. **单独测这个符号** —
   ```bash
   python -c "
   import sys; sys.path.insert(0, '<skill目录>/scripts')
   from latex2omml import latex_to_omml
   print(latex_to_omml(r'\subseteq').decode('utf-8'))
   "
   ```

---

## 十二、向量 / 平行垂直 / 角度符号

平面向量章节的三大类符号，已全部验证可生成。

### LaTeX 写法与对应 OMML 结构

| LaTeX | 视觉 | OMML 结构 | 备注 |
|---|---|---|---|
| `\vec{a}` | a⃗ | `m:acc` + `chr="⃗"` | 单字母向量，短箭头 |
| `\overrightarrow{AB}` | AB⃗ | `m:acc` + `chr="⃗"` | 多字母向量，长箭头 |
| `\overleftarrow{AB}` | AB⃖ | `m:acc` | 反向 |
| `\overleftrightarrow{AB}` | AB↔ | `m:acc` | 双向 |
| `\parallel` | ∥ | 普通字符 U+2225 | — |
| `\perp` | ⊥ | 普通字符 **U+22A5** | 见下方"坑" |
| `\angle` | ∠ | 普通字符 U+2220 | — |
| `\theta` | θ | 普通字符 U+03B8 | 夹角 |
| `\cdot` | ⋅ | 普通字符 U+22C5 | 数量积 |

### 常用句式

```latex
\vec{a}\cdot\vec{b}=|\vec{a}||\vec{b}|\cos\theta          % 数量积定义
\vec{a}\parallel\vec{b}\Leftrightarrow\vec{a}=\lambda\vec{b}   % 平行充要条件
\vec{a}\perp\vec{b}\Leftrightarrow\vec{a}\cdot\vec{b}=0        % 垂直充要条件
\vec{a}\cdot\vec{a}=|\vec{a}|^{2}                        % 模方
```

### 坑 1：`\perp` 的 Unicode 码位

latex2mathml 把 `\perp` 映射成 **U+27C2（⟂）**，这是错的。
标准的垂直符号是 **U+22A5（⊥ UP TACK）**。
`latex2omml.py` 里的 `preprocess_mathml()` 已自动修正。

### 坑 2：accent 符号退化成上极限（重要）

latex2mathml 输出 `<mover>` 时**不带 `accent="true"`**，
MML2OMML.XSL 因此把 `\vec{a}` 转成 `m:limUpp`（上极限），
视觉上箭头悬空在字母正上方、间距过大，完全不是教科书的样子。

受影响的命令（已全部修复）：

```
\vec  \overrightarrow  \overleftarrow  \overleftrightarrow
\hat  \widehat  \tilde  \widetilde
\dot  \ddot  \check  \acute  \grave  \breve  \mathring
```

修复方式：`preprocess_mathml()` 检测 `<mover>`/`<munder>` 上方字符，
属于 accent 字符集就补上 `accent="true"`，XSL 才会输出 `m:acc`。

> 判断是否 accent 的依据是**上方字符本身**（箭头/点/帽子/波浪…），
> 不是命令名。所以 `\xrightarrow{f}` 这种上方是文字的，仍然走 `m:limUpp`，正确。

### 坑 3：`\underline` 与 XSL 硬限制

`MML2OMML.XSL` **不支持** `munder -> m:bar`，加任何属性都只会得到 `m:limLow`。
`postprocess_omml_root()` 在 OMML 层面把「`m:limLow` + 横线字符」
改写成 `m:bar` + `pos="bot"`，才是紧贴主体的下划线。

---

## 十三、从 MathType 图元里识别符号（不靠 OCR）

老讲义里的公式多是 **MathType 7.0 的 OLE 对象**，在 docx 中以 WMF 呈现图片保存。
`python-docx` 读不到任何文字，但 WMF 是**矢量**格式，文字原样存在
`META_EXTTEXTOUT` 记录里，配合当时的字体就能精确还原。

```bash
python scripts/wmf_symbols.py <目录或单个.wmf> [--json out.json]
```

### 字体与符号的对应

| 字体 | 用途 | 关键码位 |
|---|---|---|
| **MT Extra** | 向量箭头 | `0x72` = 箭头头，`0x75` = 可重复线段 |
| **Euclid Extra** | 同上（MathType 6+ 新公式改用这字体） | 码位**跟 MT Extra 一致** |
| **Symbol** | 希腊字母 + 运算符 | `0x71`=θ `0x70`=π `0x5E`=⊥ `0xD7`=⋅ `0xD0`=∠ |
| Times New Roman | 普通字母数字 | ASCII |

**向量箭头的判定**：`0x75` 出现 ≥3 次再跟 `0x72` → 长箭头（`\overrightarrow`，覆盖多字母）；
只有 `0x72` → 短箭头（`\vec`，单字母）。

**平行符号**：MathType 用 Times New Roman 的两个斜杠 `//` 画 ∥。

> ⚠️ **同一份文档里 MT Extra 和 Euclid Extra 会混用**。只认 MT Extra 会漏掉一半向量符号
> （实测 59 个公式里漏了 1 个）。脚本里的 `ARROW_FONTS` 常量把四种名称都列了。

### 四个必踩的坑

1. **WMF 头是 40 字节** = placeable(22) + standard(18)。只跳 22 会解析出 0 条记录。
2. **LOGFONT 是 Win16 版**：5 个 SHORT(10B) + 8 个 BYTE(8B)，
   facename 在 **offset 18**，32 字节**单字节 ANSI**，不是 UTF-16LE。
3. **文字记录按字体分组输出，不是视觉顺序**。
   例如 `|a||b|cosθ` 会输出成 `coscos` / `ababba` / `qq` / `×=×=×` 四段。
4. **可伸缩的竖线 `|`、大括号根本不在文字记录里** —— MathType 用绘图图元
   （`LineTo` / `Rectangle`）画它们，字号变了要跟着拉伸。
   所以 WMF 侧**看不到绝对值符号**，别拿它当"这个公式没有竖线"的证据。

前三条决定了本脚本只能还原「用了哪些符号」，**不能还原公式结构**。
要结构请走 MTEF 解析（`scripts/mtef.py`，见 SKILL.md 的
[MathType 文档抢救](#mathtype-文档抢救mtef-解析) 一节）—— 那份数据里有
完整对象树，分数、根号、上下标、fence 的嵌套关系一个不漏。

### 提取 WMF

```python
import zipfile
z = zipfile.ZipFile('讲义.docx')
for n in z.namelist():
    if n.startswith('word/media/') and n.lower().endswith('.wmf'):
        open(os.path.basename(n), 'wb').write(z.read(n))
```

---

## 附：符号库体检与上游勘误（2026-11）

`scripts/scan_symbol_table.py` 会一次性做三件事：扫上游 `SYMBOLS` 表里退化为
ASCII 的条目、逐条核对本文件与引擎的实际输出、统计 Cambria Math 无字形的码点。

### 已修正的上游笔误（引擎导入时自动打补丁）

| 命令 | 上游错值 | 正确码点 | 说明 |
|---|---|---|---|
| `\sim` | `0007E` `~` | U+223C `∼` | `\thicksim` 用的才是正确码点；ASCII `~` 是 `asciitilde` 字形 |
| `\ast` | `0002A` `*` | U+2217 `∗` | ASCII `*` 在数学字体里是偏上的 `asterisk`，不是 `asteriskmath` |
| `\cdot` | `000B7` `·` | U+22C5 `⋅` | `periodcentered` 是文本中点；Word 公式用 `dotmath` |

这三条在 Cambria Math 里是**三个不同的字形**，所以修正后肉眼可见，不只是码点变好看。
补丁写在 `latex2omml._patch_upstream_symbols()`：改的是 `symbols_parser.SYMBOLS`
字典本身（保留 `<mo>` 语义），上游哪天修好了会自动跳过。

`\colon`→`:`、`\vert`→`|`、`\less`→`<` 这类**不是笔误**，是 LaTeX 的标准语义，别改。

### 箭头 / 关系符 / 圈码：会不会在 Word 里丢字符？

实测 `C:\Windows\Fonts\cambria.ttc` 的第 1 号字体（Cambria Math）：

| 字符 | 码点 | LaTeX | Cambria Math |
|---|---|---|---|
| ⟹ | U+27F9 | `\Longrightarrow` | 有字形 |
| ⇒ | U+21D2 | `\Rightarrow` | 有字形 |
| ∽ | U+223D | `\backsim` | 有字形 |
| ∼ | U+223C | `\sim`（已修正） | 有字形 |
| ① ② … ⑩ | U+2460…U+2469 | 直接写字符 | 有字形 |
| ∘ / ° | U+2218 / U+00B0 | `^\circ` / 直接写 | 都有字形 |

结论：**这些符号都不会丢**，"未映射"只说明转换脚本没有对应的语义标签，
不等于 Word 显示不出来。转换时把它们原样放进 `m:t` 即可；只有圈码 ① 这类
**文本符号**建议放在正文 run 里，别塞进公式（字形宽度按全角算，基线会不齐）。

### 常用命令的字形覆盖（2026-11 实测）

上游 4054 条里有 179 个码点在 Cambria Math 中无字形，但逐条筛过
`\Box` `\Diamond` `\lll` `\ggg` `\lessdot` `\nleq` `\ngeq` `\nsim` `\ncong`
`\thicksim` `\thickapprox` `\overparen` `\dashrightarrow` `\boxplus` 等常用命令，
**全部有字形**。缺的都是 `\enleadertwodots`、`\caretinsert`、`\Question`、
`\enclosetriangle` 这类生僻排版符号 —— 用到才会丢字，属于信息项不是故障。
`scan_symbol_table.py --all` 可以列出全部。

### 旧文档里已经固化的错误字符

引擎补丁只管新生成的文档。此前转换出来的成品 docx，`·` `~` 已经写进 `<m:t>`，
用 `fix_docx_symbols.py` 单独修：

```bash
python scripts/fix_docx_symbols.py "C:\讲义目录" -r        # 只读体检
python scripts/fix_docx_symbols.py 文档.docx --fix          # 修到 _符号修正.docx
```

默认修 `·`→U+22C5、`~`→U+223C；`*` 默认不动（多为 `n∈N*` 原稿写法），
确需修加 `--stars`。2026-11 实测桌面 3198 份 docx：3168 份可解析、
**40 份命中**；本项目两份成品各修 3 处与 11 处，公式数与插图数守恒。

### 体检判据的两条边界

- A 类（必坏）看的是**最终 OMML 输出**，不是上游表原始值。`\perp` 上游写
  U+27C2，引擎的字符纠错表已补成 U+22A5，只算"已补偿"不算故障。
- `\degree` / `\degreeCelsius` / `\celsius` 经 premap 改写成 `^{\circ}` 这类
  **结构型表达式**，不适用"最终文本 == 单个字符"的判据，单列跳过。
  直接写 `°` U+00B0 只落在基线上，结构反而是错的。

### 度数上标的唯一正确写法

```
30^\circ        ✅  m:sSup（底数 30 + 上标 ∘ U+2218）
30^{\circ}      ✅  同上
30\degree       ✅  premap 改写为 ^{\circ}
30°             ❌  U+00B0 落在基线，不是上标
30^{\mathrm{o}} ❌  字母 o，不是数学度号
```

---

## 附：跨格式公式字符体检（2026-11）

`fix_docx_symbols.py` 只看 `<m:t>`，覆盖不到 MathType OLE、图片公式和 PDF 文字层。
`audit_formula_chars.py` 把 docx + PDF 一起盘，并**按处置级别分开报**：

```bash
python scripts/audit_formula_chars.py "C:\资料目录" "试卷.pdf" -r
python scripts/audit_formula_chars.py "C:\资料目录" -r --fix   # 只修 ERROR，输出到副本
```

| 级别 | 命中内容 | 处置 |
|---|---|---|
| ERROR | 公式内 `·` U+00B7、`~` U+007E、U+FFFD、私有区码点 | 修，改到 `<m:t>` 内 |
| WARN | 公式内全角标点、Unicode 伪上标 `²³ⁿ`、SYM 未覆盖私有码点 | 只报告 |
| INFO | `*` 星号、`□` 填空框、正文题源 `·`、OLE/图片公式数、扫描件 | 明确不动 |

### 三条容易改坏的规则

1. **正文 `·` 不是点乘**。实测某讲义 265 个 `·` 全是 `2021·北京人大附中模拟`
   这类题源标注；三个目标合计 1366 个正文 `·`，改成 `⋅` 会把题源全毁掉。
   修复正则必须严格限定 `<m:t>...</m:t>`。
2. **`*` 不修**。抽样看几乎都是 `n∈N*`、星级标记等原稿写法（用户 2026-11-11 确认），
   换成二元运算符 `∗` 会多出间距。
3. **`□` 不修**。组合数学原题里是合法的填空占位符，不是缺字方框。
   真正的缺字是 U+FFFD 替换符或私有区乱码。

### PDF：Symbol 编码与 Euclid 字体的边界

`pdf2docx_math.SYM` 是 Adobe Symbol 0x00–0xFF 编码表（191 条）。
PDF 文字层里的 U+F0XX 由 `remap_symbol()` **减 0xF000** 后查表：

| 码位 | 值 | 说明 |
|---|---|---|
| `0x7E` | `∼` U+223C | Adobe Symbol 的 similar。此前误写 `'~'`，2026-11 修正 |
| `0xA3` | `≤` | 常见 |
| `0xD7` | `⋅` U+22C5 | 数学点乘 |
| `0xB7` | `•` U+2022 | Symbol 的 bullet，**不是**中文中点 `·` |
| `0xFF` | `''` | 可伸缩括号片，有意丢弃 |

实测 239 份 PDF（含 396 页讲义）：197 份带私有区码点，110 种 / 32809 次，
其中 **3 种（U+F084 / F085 / F08C，10 次）不在 SYM 表内**。它们来自
**EuclidMathOne / EuclidMathTwo**（MathType 专用字体，U+F08C 字形盒仅 3.3pt 宽，
是整除竖线），不属于 Adobe Symbol 编码，**不能照 Symbol 表硬猜**。
现在 `UNMAPPED_PRIVATE` 会记录被静默丢弃的码位，`unmapped_private_report()`
可查；体检报告会列出未覆盖明细，提示人工核对原页。

### 覆盖边界

- **11 份纯扫描件 PDF** 没有文字层，字符检查无从谈起，要走多模态识别路线。
- **`.doc` 老格式**（2 份）需先转成 docx。
- **MathType OLE 公式**：字符在 MTEF 二进制里，`m:t` 扫不到；抽查 112 个
  公式成功解析 106 个，`\cdot` 都是结构化 LaTeX，没有码点问题。
  要查走 `mtef.py --json`。
