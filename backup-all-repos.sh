#!/usr/bin/env bash
# نسخة كاملة (mirror) من كل مستودعات Mohamedabdelguom1
# الاستخدام: GITHUB_TOKEN=ghp_xxx ./backup-all-repos.sh
set -euo pipefail

USER_NAME="Mohamedabdelguom1"
DEST="${HOME}/github-backup"
: "${GITHUB_TOKEN:?ضع التوكن: export GITHUB_TOKEN=ghp_xxx (صلاحية repo)}"

mkdir -p "$DEST" && cd "$DEST"

REPOS=(
  AI-SEALER-PROJECT queue-system sinan-dental-system clinic-project-no-2
  mosque-prayer-display wa-booking-agent hermes-n8n-mcp
  n8n-whatsapp-dental-clinic-mvp clinic-whatsapp-bot
  maxxautomation-facebook-content-engine rawaq-cafe-bakery fb-ig-automation
  ai-seller-crm-os my-first-android-app cultural-games-platform rgp Mohamedabdo
)

for r in "${REPOS[@]}"; do
  url="https://${USER_NAME}:${GITHUB_TOKEN}@github.com/${USER_NAME}/${r}.git"
  if [ -d "${r}.git" ]; then
    echo "تحديث: $r"
    git -C "${r}.git" remote update --prune
  else
    echo "نسخ: $r"
    git clone --mirror "$url" "${r}.git"
  fi
done

echo "تمّ. النسخ في: $DEST"
