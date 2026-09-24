# latex-omml 安装说明

skill 本体只有 **300 KB**，自包含（微软的 `MML2OMML.XSL` 已内置），**不需要装 Office**。
只需要 Python 3.9+ 和三个第三方库。

---

## Windows（推荐，一键）

双击 `install.cmd`。

它会：

1. 在 skill 目录下建独立虚拟环境 `.venv`（**不污染系统 Python**）
2. 装 `requirements.txt` 里的依赖（优先清华源，失败自动切官方源）
3. 跑一遍内置自检，打印 12 项 PASS 列表

装完以后用 `.venv` 里的 python 跑脚本：

```bat
.venv\Scripts\python.exe scripts\md2docx.py input.md -o output.docx
.venv\Scripts\python.exe scripts\audit_math.py "C:\讲义目录"
.venv\Scripts\python.exe scripts\wmf_symbols.py "讲义.docx" --json sym.json
```

> **注意**：虚拟环境**不能**整个文件夹拷到别的机器用（里面的路径是写死的）。
> 换机器要把 `latex-omml` 文件夹拷过去后**重新跑一次 install.cmd**。

---

## macOS / Linux（手动）

```bash
cd latex-omml
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/latex2omml.py      # 自检，应看到 12/12
```

> macOS 上的 `emf2png.py` 依赖 GDI+，**用不了**（WMF/EMF 转 PNG 仅限 Windows）。
> 其余功能（生成 OMML、符号提取、docx 升级）跨平台正常。

---

## 装到 WorkBuddy

整个 `latex-omml` 文件夹放到：

| 系统 | 路径 |
|---|---|
| Windows | `C:\Users\<你的用户名>\.workbuddy\skills\latex-omml\` |
| macOS / Linux | `~/.workbuddy/skills/latex-omml/` |

放好后重启/新开会话，WorkBuddy 会自动识别。识别不到的话检查 `SKILL.md`
顶部的 frontmatter 是否完好（`name` / `description` 两个字段必须有）。

---

## 依赖说明

```
lxml>=5.0          XML 变换（跑 MML2OMML.XSL）
latex2mathml>=3.80 LaTeX → MathML
python-docx>=1.1   读写 docx
```

就这三个。WMF/EMF 解析、缩略图生成、docx 拆包全用标准库手写，不额外依赖。

---

## 验证装好了没

```bash
python scripts/latex2omml.py
```

看到 `通过 12 / 12` 就 OK。

再跑个端到端：

```bash
echo '求根公式 $$x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}$$' > t.md
python scripts/md2docx.py t.md -o t.docx
```

打开 `t.docx`，公式应该能**双击编辑**，而不是图片。
