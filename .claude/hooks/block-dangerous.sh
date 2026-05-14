#!/bin/bash
# 危険なコマンドをブロックするhook
# Claude Codeがbashコマンドを実行する前に自動チェック

# 標準入力からJSONを読み取り、コマンドを抽出
COMMAND=$(cat /dev/stdin | jq -r '.tool_input.command // empty')

# 危険なパターンのリスト
DANGEROUS_PATTERNS=(
  "rm -rf"
  "git push --force"
  "git push -f"
  "format "
  "del /f /s /q"
  "Remove-Item -Recurse -Force"
)

for pattern in "${DANGEROUS_PATTERNS[@]}"; do
  if echo "$COMMAND" | grep -qF "$pattern"; then
    echo "{\"block\": true, \"message\": \"⛔ 危険なコマンドをブロックしました: '$pattern' が含まれています。本当に必要な場合は手動で実行してください。\"}"
    exit 2
  fi
done

exit 0
