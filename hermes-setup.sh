#!/usr/bin/env bash
# تثبيت Hermes Agent محلياً داخل WSL2 (Ubuntu) أو ماك/لينكس
# التشغيل: bash hermes-setup.sh
set -e

echo "==> 1/4 تحديث الحزم الأساسية"
if command -v apt-get >/dev/null 2>&1; then
  sudo apt-get update -y && sudo apt-get install -y curl git
fi

echo "==> 2/4 تثبيت Hermes Agent"
curl -fsSL https://hermes-agent.nousresearch.com/install.sh | bash
# تحميل المسار الجديد في نفس الجلسة
[ -f ~/.bashrc ] && source ~/.bashrc || true

echo "==> 3/4 إنشاء فولدر مادة الكورس"
mkdir -p ~/course
echo "حط ملفات الدروس (.md أو .txt) في: ~/course"

echo "==> 4/4 تثبيت Claude Code (اختياري - لاستخدام اشتراك Claude)"
read -r -p "تثبّت Claude Code؟ (y/n): " ans
if [ "$ans" = "y" ]; then
  curl -fsSL https://claude.ai/install.sh | bash
  [ -f ~/.bashrc ] && source ~/.bashrc || true
  claude auth login
fi

echo ""
echo "تم. الخطوة الجاية:"
echo "  - للموديلات المجانية:  hermes setup --portal"
echo "  - لاختيار موديل/إضافة Claude:  hermes model"
echo "  - للتشغيل:  hermes"
