<div dir="rtl">

# مشروع جودة بيانات الرحلات الجوية ✈️

**Flights Data Quality Project**

[![Release](https://img.shields.io/badge/release-v1.0.0-blue)](CHANGELOG.md)
[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB)](requirements.txt)
[![Program](https://img.shields.io/badge/SDAIA%20Academy-Data%20Quality%20Management-0A7E8C)](https://github.com/SDAIAAcademy)

مشروع تطبيقي متكامل لتقييم وإدارة جودة بيانات رحلات المغادرة من مطارات نيويورك الثلاثة (EWR و JFK و LGA) لعام 2013، يضم **336,776** سجلًا و**21** حقلًا. يقيس المشروع جودة البيانات على **ستة أبعاد**، ويحصر مشاكلها بعينات حقيقية، ويضع خطة جودة كاملة بنطاقها وفريقها وميزانيتها، ويربط قواعد الأعمال بقواعد الجودة ويوزع مهام الإصلاح على الأدوار.

> 🎓 نُفّذ هذا المشروع ضمن **برنامج إدارة جودة البيانات** المقدَّم من [أكاديمية سدايا (SDAIA Academy)](https://github.com/SDAIAAcademy).

---

## المحتويات

- [فكرة المشروع](#فكرة-المشروع)
- [أبرز النتائج](#أبرز-النتائج)
- [المخرجات](#المخرجات)
- [هيكل المستودع](#هيكل-المستودع)
- [التشغيل](#التشغيل)
- [طريقة الاستخدام](#طريقة-الاستخدام)
- [التوثيق الفني](#التوثيق-الفني)
- [إدارة الإصدارات](#إدارة-الإصدارات)
- [البرنامج التدريبي](#البرنامج-التدريبي)
- [English summary](#english-summary)

## فكرة المشروع

البيانات التشغيلية للرحلات (المواعيد والتأخير والطائرات والمسارات) هي أساس مؤشرات الانضباط في المواعيد وتخطيط الأساطيل وتقارير الجهات التنظيمية. أي خلل فيها ينعكس مباشرة على القرارات. يجيب المشروع عن ثلاثة أسئلة:

1. **ما مستوى جودة البيانات الحالي؟** قياس آلي لـ 26 قاعدة جودة (18 قاعدة على مستوى السجل و8 ضوابط على مستوى الملف) موزعة على الأبعاد الستة: الاكتمال، والدقة، والحداثة، والاتساق، والصلاحية، والتوفر.
2. **ما المشاكل وما أثرها؟** سجل بـ 14 مشكلة مصنّفة حسب الخطورة، لكل منها وصف وأثر على الأعمال وسبب جذري وعينات حقيقية من الملف.
3. **كيف نعالجها؟** خطة جودة بيانات تشمل النطاق، وأولويات عناصر البيانات، والأنشطة والمراحل (12 أسبوعًا)، والميزانية، والفريق (11 دورًا بينها مسؤول حوكمة البيانات)، ومصفوفة RACI، و14 قاعدة أعمال مرتبطة بقواعد الجودة، و16 مهمة إصلاح موزعة على الأدوار.

## أبرز النتائج

| البعد | النتيجة | البعد | النتيجة |
|---|---|---|---|
| الدقة (Accuracy) | **99.96%** | الاتساق (Consistency) | **94.38%** |
| الاكتمال (Completeness) | **99.11%** | التوفر (Availability) | **40.00%** |
| الصلاحية (Validity) | **98.30%** | الحداثة (Timeliness) | **33.33%** |

- **المؤشر العام للجودة: 77.5%**، واكتمال الخلايا 99.34%.
- **14 مشكلة جودة**: 4 عالية الخطورة، و7 متوسطة، و3 منخفضة. أبرزها: أوقات محلية دون منطقة زمنية (130,280 سجلًا)، وأرقام تسجيل طائرات غير مطابقة لصيغة FAA (22,754)، ورحلات ملغاة دون حقل حالة صريح (8,255).
- **ميزانية الخطة**: 535,480 ريال سعودي على 12 أسبوعًا.

## المخرجات

| الملف | الوصف |
|---|---|
| [`reports/flights_dq_project_ar.docx`](reports/flights_dq_project_ar.docx) | وثيقة المشروع (نحو 20 صفحة، عربي RTL): الملخص التنفيذي، وخطة الجودة، ودراسة الأبعاد الستة، والتقرير التنفيذي للمشاكل، وقواعد الأعمال، ومهام الإصلاح، والتوصيات. |
| [`reports/flights_dq_analysis_ar.xlsx`](reports/flights_dq_analysis_ar.xlsx) | مصنف التحليل (12 ورقة): الملخص، والتحليل الوصفي، وقواعد الجودة ونتائجها، وسجل المشاكل وعيناتها، والسجلات المخالفة، وأولويات العناصر، وقواعد الأعمال، والفريق و RACI، والمراحل، والميزانية، ومهام الإصلاح. |

## هيكل المستودع

</div>

```text
flights-data-quality/
├── src/
│   ├── analysis.py      # محرك قواعد الجودة R01–R18 على مستوى السجل
│   ├── content.py       # المحتوى: القواعد، الضوابط، المشاكل، الخطة، الفريق، الميزانية
│   ├── build_docx.py    # يبني وثيقة Word
│   └── build_xlsx.py    # يبني مصنف Excel
├── data/
│   ├── README.md        # مصدر البيانات وطريقة الحصول على الملف الكامل
│   ├── sample/          # عينة 11,036 سجلًا (اليوم الأول من كل شهر)
│   └── raw/             # يوضع هنا flights.csv الكامل (مستبعد من git)
├── reports/             # المخرجات النهائية للإصدار
├── docs/
│   └── TECHNICAL.md     # التوثيق الفني
├── Makefile
├── requirements.txt
├── CHANGELOG.md
└── CONTRIBUTING.md
```

<div dir="rtl">

## التشغيل

### المتطلبات

- Python 3.10 أو أحدث
- الحزم: pandas و numpy و python-docx و openpyxl

### الخطوات

</div>

```bash
# 1. نسخ المستودع
git clone https://github.com/hhummadi/flights-data-quality.git
cd flights-data-quality

# 2. إنشاء بيئة افتراضية وتثبيت الحزم
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 3. (اختياري) تنزيل الملف الكامل flights.csv من صفحة الإصدار v1.0.0
#    ووضعه في data/raw/flights.csv  — راجع data/README.md

# 4. تشغيل الفحوص وبناء التقارير
make all
```

<div dir="rtl">

إن لم يتوفر `make` (مثل Windows)، شغّل الأوامر مباشرة:

</div>

```bash
cd src
python analysis.py      # يطبع نتيجة كل قاعدة
python build_docx.py    # ينشئ reports/flights_dq_project_ar.docx
python build_xlsx.py    # ينشئ reports/flights_dq_analysis_ar.xlsx
```

<div dir="rtl">

> **ملاحظة:** إذا لم يوجد `data/raw/flights.csv` فإن الأدوات تعمل تلقائيًا على العينة في `data/sample/`، وهي مناسبة للتجربة السريعة. الأرقام المنشورة أعلاه وفي `reports/` ناتجة عن الملف الكامل. بعض النصوص الوصفية في الوثيقة مكتوبة لأرقام الملف الكامل.

## طريقة الاستخدام

| الهدف | الأمر |
|---|---|
| فحص الجودة فقط وعرض نتيجة كل قاعدة | `make check` |
| بناء وثيقة Word | `make docx` |
| بناء مصنف Excel | `make xlsx` |
| استخدام ملف بيانات آخر بالمخطط نفسه | `make all DQ_DATA=/path/to/flights.csv` |
| توجيه المخرجات إلى مجلد آخر | `make all DQ_REPORTS=/path/to/out` |
| حذف المخرجات المولدة | `make clean` |

لتعديل محتوى الخطة (الفريق، الميزانية، قواعد الأعمال، المشاكل) عدّل `src/content.py` ثم أعد البناء. ولإضافة قاعدة جودة جديدة اتبع الخطوات في [التوثيق الفني](docs/TECHNICAL.md#إضافة-قاعدة-جودة-جديدة).

## التوثيق الفني

التفاصيل الكاملة في [`docs/TECHNICAL.md`](docs/TECHNICAL.md): المعمارية وتدفق البيانات، وقاموس البيانات، وتعريف كل قاعدة ومنطق فحصها، وطريقة حساب درجات الأبعاد والمؤشر العام، وأولويات العناصر، ومتغيرات البيئة، وطريقة التوسعة، والقيود المعروفة.

## إدارة الإصدارات

- **ترقيم دلالي** [SemVer](https://semver.org/lang/ar/): `MAJOR.MINOR.PATCH`، والإصدار الحالي `v1.0.0`.
- **وسوم git موقعة بالتعليق** (annotated tags) لكل إصدار، مع إصدار GitHub مرفق به الملف الكامل والتقارير.
- **سجل تغييرات** وفق [Keep a Changelog](https://keepachangelog.com/ar/1.1.0/) في [`CHANGELOG.md`](CHANGELOG.md).
- **رسائل commits** وفق [Conventional Commits](https://www.conventionalcommits.org/ar/v1.0.0/) (`feat:` و `fix:` و `docs:` و `chore:` ...).
- **نموذج الفروع**: الفرع `main` مستقر دائمًا، وكل تغيير في فرع قصير العمر (`feature/*` أو `fix/*` أو `docs/*`) يُدمج عبر Pull Request. التفاصيل في [`CONTRIBUTING.md`](CONTRIBUTING.md).

## البرنامج التدريبي

نُفّذ هذا المشروع تطبيقًا عمليًا ضمن **برنامج إدارة جودة البيانات** الذي تقدمه **أكاديمية سدايا** التابعة للهيئة السعودية للبيانات والذكاء الاصطناعي (SDAIA).

- حساب أكاديمية سدايا على GitHub: **<https://github.com/SDAIAAcademy>**

## المؤلف

**هشام** — المتدرب في برنامج إدارة جودة البيانات.

## مصدر البيانات

مجموعة بيانات [nycflights13](https://github.com/tidyverse/nycflights13) العامة (مصدرها مكتب إحصاءات النقل الأمريكي BTS)، مع إضافة عمودي `id` و `name`. الافتراضات الواردة في الخطة (مثل مالك البيانات "إدارة عمليات الرحلات" والميزانية بالريال السعودي) افتراضات تعليمية.

</div>

---

## English summary

An end-to-end **data quality assessment and management project** on 336,776 departure records from New York's three airports (2013, 21 columns), built as a capstone for the **Data Quality Management Program** at [SDAIA Academy](https://github.com/SDAIAAcademy).

**What it does**

- Runs 26 automated data-quality rules (18 record-level, 8 file-level) across six dimensions: completeness, accuracy, timeliness, consistency, validity and availability.
- Produces an issue log of 14 issues with severity, business impact, root cause and real sample records.
- Lays out a full data quality plan: scope, data element priorities, a 12-week roadmap, a SAR 535,480 budget, an 11-role team (including a Data Governance Officer) with a RACI matrix, 14 business rules mapped to DQ rules, and 16 remediation tasks.
- Generates an Arabic (RTL) Word report and a 12-sheet Excel workbook.

**Key results:** overall score 77.5% · accuracy 99.96% · completeness 99.11% · validity 98.30% · consistency 94.38% · availability 40% · timeliness 33.3%.

**Quick start**

```bash
pip install -r requirements.txt
make all            # uses data/raw/flights.csv if present, otherwise the bundled sample
```

See [`docs/TECHNICAL.md`](docs/TECHNICAL.md) for the architecture, data dictionary and rule catalogue, and [`CHANGELOG.md`](CHANGELOG.md) for release history.
