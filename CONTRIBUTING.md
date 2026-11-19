<div dir="rtl">

# دليل المساهمة وإدارة الإصدارات

## نموذج الفروع (GitHub Flow)

| الفرع | الغرض |
|---|---|
| `main` | مستقر دائمًا وقابل للإصدار. لا يُدفع إليه مباشرة. |
| `feature/<وصف>` | ميزة جديدة، مثل `feature/timezone-rules` |
| `fix/<وصف>` | إصلاح خطأ، مثل `fix/r09-midnight` |
| `docs/<وصف>` | توثيق فقط |

خطوات العمل:

</div>

```bash
git switch main && git pull
git switch -c feature/timezone-rules
# ... التعديل ...
make all                       # تأكد أن البناء يعمل
git add -A && git commit -m "feat(rules): add UTC conversion check R27"
git push -u origin feature/timezone-rules
# ثم افتح Pull Request إلى main
```

<div dir="rtl">

## رسائل الـ commits

نتبع [Conventional Commits](https://www.conventionalcommits.org/): `type(scope): وصف مختصر بصيغة الأمر`.

| النوع | متى يُستخدم | أثره على الإصدار |
|---|---|---|
| `feat` | قاعدة أو مخرج أو ميزة جديدة | MINOR |
| `fix` | إصلاح خطأ في قاعدة أو تقرير | PATCH |
| `docs` | توثيق فقط | — |
| `refactor` | إعادة هيكلة دون تغيير السلوك | — |
| `chore` | إعدادات وأدوات | — |
| `feat!` أو `BREAKING CHANGE:` | تغيير غير متوافق (مثل تغيير مخطط البيانات) | MAJOR |

النطاقات المقترحة: `rules` و `content` و `reports` و `data` و `docs`.

## إصدار نسخة جديدة

1. انقل البنود من `[Unreleased]` إلى قسم جديد `[X.Y.Z] - YYYY-MM-DD` في `CHANGELOG.md`.
2. أعد توليد التقارير على الملف الكامل: `make all`.
3. نفّذ commit بالرسالة `chore(release): vX.Y.Z`.
4. أنشئ وسمًا موثقًا وادفعه:

</div>

```bash
git tag -a vX.Y.Z -m "vX.Y.Z: ملخص الإصدار"
git push origin main --tags
```

<div dir="rtl">

5. أنشئ GitHub Release من الوسم، وأرفق `flights.csv` والتقارير.

</div>
