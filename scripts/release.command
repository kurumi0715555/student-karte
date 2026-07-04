#!/bin/bash
# 生徒カルテ — リリース作成スクリプト
# Finderでこのファイルをダブルクリックすると実行されます。
set -e
cd "$(dirname "$0")/.."

echo "=================================="
echo " 生徒カルテ — 新しいバージョンを公開"
echo "=================================="
echo ""
echo "現在の最新タグ:"
git tag --sort=-v:refname | head -3 || echo "  （まだタグがありません）"
echo ""
read -p "新しいバージョン番号を入力してください（例: 1.3.0）: " VERSION

if [ -z "$VERSION" ]; then
  echo "何も入力されなかったので中止しました。"
  read -p "Enterキーで閉じます..."
  exit 1
fi

TAG="v$VERSION"

if git rev-parse "$TAG" >/dev/null 2>&1; then
  echo "⚠ タグ $TAG はすでに存在します。中止しました。"
  read -p "Enterキーで閉じます..."
  exit 1
fi

echo ""
echo "未pushの変更がないか確認します..."
git fetch origin main --quiet
LOCAL=$(git rev-parse main)
REMOTE=$(git rev-parse origin/main)
if [ "$LOCAL" != "$REMOTE" ]; then
  echo "⚠ ローカルの main と GitHub の main がずれています。"
  echo "  先に「git push」でmainを最新にしてから、もう一度実行してください。"
  read -p "Enterキーで閉じます..."
  exit 1
fi

echo "タグ $TAG を作成してGitHubに送ります..."
git tag "$TAG"
git push origin "$TAG"

echo ""
echo "✅ 完了しました。数分でGitHub上に新しいリリースが自動生成されます。"
echo "   （配布ZIPの作成も自動で行われます）"
echo ""
echo "確認はこちら:"
echo "https://github.com/kurumi0715555/student-karte/releases"
echo ""
read -p "Enterキーで閉じます..."
