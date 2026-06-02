#!/bin/bash
# 评估明细表 Skill — GitHub 同步安装脚本
# 用途: Mac Mini / 新机器首次安装 / 已有仓库更新
# 用法: bash install_from_github.sh [--verify]

set -e

REPO_URL="https://github.com/zer0-lyz/valuation-detail-table.git"
SKILL_DIR="$HOME/.codex/skills/valuation-detail-table"
EXPECTED_VERSION="v0.2.2"

# 1) 准备目录
mkdir -p "$(dirname "$SKILL_DIR")"

# 2) 克隆或拉取
if [ -d "$SKILL_DIR/.git" ]; then
    echo "→ 仓库已存在,执行 git pull"
    cd "$SKILL_DIR"
    git fetch origin
    LOCAL=$(git rev-parse --short HEAD)
    REMOTE=$(git rev-parse --short origin/main)
    if [ "$LOCAL" = "$REMOTE" ]; then
        echo "  已是最新 ($LOCAL)"
    else
        echo "  本地 $LOCAL → 远端 $REMOTE,拉取中..."
        git pull --rebase origin main
    fi
else
    echo "→ 首次安装,克隆 $REPO_URL"
    git clone "$REPO_URL" "$SKILL_DIR"
    cd "$SKILL_DIR"
fi

# 3) 校验
ACTUAL_VERSION=$(grep -A3 "^## 版本" SKILL.md | grep -E "^v[0-9]" | head -1 | awk '{print $1}')
if [ "$ACTUAL_VERSION" = "$EXPECTED_VERSION" ]; then
    echo "✓ SKILL.md 版本匹配: $ACTUAL_VERSION"
else
    echo "⚠️  版本不匹配: 期望 $EXPECTED_VERSION, 实际 $ACTUAL_VERSION"
    echo "   请在终端打开 SKILL.md 确认"
fi

# 4) 可选自检
if [ "$1" = "--verify" ]; then
    echo ""
    echo "→ 运行 validate_skill.py 自检"
    cd "$SKILL_DIR"
    python3 scripts/validate_skill.py 2>&1 | tail -15
fi

echo ""
echo "✓ 安装完成"
echo "  路径: $SKILL_DIR"
echo "  版本: $ACTUAL_VERSION"
echo ""
echo "  使用: 任何项目目录下,执行"
echo "    python3 ~/.codex/skills/valuation-detail-table/valuation-detail-table/scripts/dt_runner.py --phase all --project <项目路径>"
