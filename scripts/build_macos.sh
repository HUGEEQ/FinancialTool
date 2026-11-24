#!/usr/bin/env bash
# ============================================================
#  FinSummaryTool macOS 客户端构建脚本（在 macOS 上运行）
#  产物：dist/FinSummaryTool.app（发给用户，未签名需右键打开）
#        dist/FinSummaryTool-macos.zip
# ============================================================
set -e
cd "$(dirname "$0")/.."

# 1) 依赖（含打包工具；如无 .venv 会自动创建）
if [ ! -d .venv ]; then
    python3 -m venv .venv
fi
source .venv/bin/activate
python -m pip install -r requirements.txt -r requirements-build.txt

# 2) 打包
python -m PyInstaller --noconfirm --clean FinSummaryTool.spec

# 3) 自检（验证客户端内资源可解析）
"./dist/FinSummaryTool.app/Contents/MacOS/FinSummaryTool" --selftest

# 4) 压缩分发包
cd dist
rm -f FinSummaryTool-macos.zip
zip -q -r FinSummaryTool-macos.zip FinSummaryTool.app
cd ..

echo ""
echo "============================================================"
echo " 构建完成："
echo "   应用包: dist/FinSummaryTool.app（未签名，首次需右键->打开）"
echo "   分发压缩包: dist/FinSummaryTool-macos.zip"
echo "============================================================"
