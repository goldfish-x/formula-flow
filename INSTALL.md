# latex-omml 安装说明

skill 本体只有 **300 KB**，自包含（微软的 `MML2OMML.XSL` 已内置），**不需要装 Office**。
只需要 Python 3.9+ 和三个第三方库。

> **装完不确定好不好使？** 一句命令体检，缺什么它直接告诉你装什么：
>
> ```bash
> python scripts/doctor.py        # 也可以加 --fix 直接把缺的装上
> ```
>
> 看到「结论：全部就绪，可以直接用」就成了。可选组件（PDF 转 Word、扫描件导出、
> 图元转图）缺失**不影响**公式转换主功能。

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

## macOS / Linux（一键）

```bash
bash install.sh
```

它和 `install.cmd` 做同样的事：建 `.venv` → 装依赖（先清华源，失败自动切官方源）
→ 跑 `scripts/doctor.py` 体检。装完用 `.venv/bin/python` 跑脚本。

### 或者手动来（想自己控制环境的话）

```bash
cd latex-omml
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/doctor.py          # 体检，应看到「全部就绪」
```

> macOS 上的 `emf2png.py` 默认用不了 GDI+，但**不会报错崩掉**：
> 它会自动去找 LibreOffice(`soffice`) 或 Inkscape，找到就用它们转；
> 都没有则明确告诉你装哪个（`brew install --cask libreoffice`）。
> 其余功能（生成 OMML、符号提取、docx 升级）跨平台正常。
> 老讲义的 WMF 公式其实**不用转图**——`wmf_symbols.py` 纯 Python 就能读出里面的字符。

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
