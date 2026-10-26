#!/usr/bin/env bash
# latex-omml 一键安装（macOS / Linux）
#
#   bash install.sh
#
# 干三件事：建独立虚拟环境 .venv（不污染系统 Python）→ 装依赖 → 跑体检。
set -u

cd "$(dirname "$0")" || exit 1

echo
echo "=== latex-omml 安装 ==="
echo
echo "本脚本会："
echo "  1. 在当前目录建独立虚拟环境 .venv（不污染系统 Python）"
echo "  2. 安装 requirements.txt 里的依赖"
echo "  3. 跑一遍体检，确认真的能用"
echo

# ---------- 1. 找 Python ----------
PY=""
for cand in python3.12 python3.11 python3.10 python3.9 python3 python; do
    if command -v "$cand" >/dev/null 2>&1; then
        if "$cand" -c 'import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)' 2>/dev/null; then
            PY="$cand"; break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo "[X] 没找到 Python 3.9+"
    echo
    echo "    macOS:  brew install python@3.12"
    echo "    Ubuntu: sudo apt install python3 python3-venv"
    echo "    CentOS: sudo yum install python3"
    echo
    exit 1
fi

echo "[1/3] 找到 Python: $("$PY" -c 'import sys; print(sys.version.split()[0])')"

# ---------- 2. 建虚拟环境 ----------
if [ -x ".venv/bin/python" ]; then
    echo "[2/3] .venv 已存在，复用"
else
    echo "[2/3] 正在创建虚拟环境 .venv ..."
    if ! "$PY" -m venv .venv; then
        echo "[X] 虚拟环境创建失败。"
        echo "    Ubuntu/Debian 可能需要先装：sudo apt install python3-venv"
        exit 1
    fi
fi

VPY=".venv/bin/python"

# ---------- 3. 装依赖 ----------
echo "[3/3] 正在安装依赖（首次约 1-2 分钟）..."
"$VPY" -m pip install --upgrade pip -q
if ! "$VPY" -m pip install -r requirements.txt -q \
        -i https://pypi.tuna.tsinghua.edu.cn/simple; then
    echo "      清华源失败，改用官方源重试..."
    if ! "$VPY" -m pip install -r requirements.txt -q; then
        echo "[X] 依赖安装失败，检查网络后重试。"
        exit 1
    fi
fi

# ---------- 4. 体检 ----------
echo
echo "=== 体检 ==="
"$VPY" scripts/doctor.py
STATUS=$?

echo
if [ $STATUS -eq 0 ]; then
    echo "=== 装好了 ==="
else
    echo "=== 有必需组件没装上，看上面的 → 提示 ==="
fi
echo
echo "以后用 .venv 里的 python 跑脚本："
echo
echo "  .venv/bin/python scripts/md2docx.py input.md -o output.docx"
echo "  .venv/bin/python scripts/pdf2docx_math.py input.pdf"
echo "  .venv/bin/python scripts/audit_math.py \"<讲义目录>\""
echo "  .venv/bin/python scripts/doctor.py          # 任何时候不确定就先跑这个"
echo
echo "macOS 提示：EMF/WMF 转 PNG 需要 LibreOffice 或 Inkscape，"
echo "           装不上也不影响公式转换（详见 SKILL.md）。"
echo
echo "换机器：整个文件夹拷过去后，重新跑一次 bash install.sh。"
echo

exit $STATUS
