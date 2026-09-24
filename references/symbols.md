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
| ⊙ | `\odot` | ⌒ | `\frown` |
| ° | `^\circ` | ′ | `'` |
| ⌢ | `\smile` | ⊚ | `\circledcirc` |

**角度写法**：`30^\circ`，`90^\circ`，`\angle ABC = 60^\circ`

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
