# PSA Eligibility Checklist — Patient Web App

Web App مخصص (ليس Google Form) لتحويل استمارة **"PSA Eligibility Checklist – Prostate Awareness / Augusta Victoria Hospital"** الورقية إلى تجربة رقمية سريعة للمريض على الموبايل، مبنية بالكامل على **Google Apps Script**.

- **Web App URL (الإنتاج):**
  `https://script.google.com/macros/s/AKfycby9HptliHahlzHKlhWxqWKZAg-JfKtZsBrlgsAVTm6AxR9gpRBhryEYXiLpP5DzIBk3pA/exec`
- **Deployment ID:** `AKfycby9HptliHahlzHKlhWxqWKZAg-JfKtZsBrlgsAVTm6AxR9gpRBhryEYXiLpP5DzIBk3pA`
- **Script ID:** `1Gxk0elewDm8K-iqDQ3r-8VlDODStQrBxCi60E3DdZi7gGnriLYQtlNmP`
- **Apps Script Editor:** https://script.google.com/d/1Gxk0elewDm8K-iqDQ3r-8VlDODStQrBxCi60E3DdZi7gGnriLYQtlNmP/edit
- **Google Sheet (قاعدة البيانات):** `PSA_RESPONSES_DB — Prostate Awareness Campaign` (الرابط داخل حساب Google الذي شغّل `setupProject()`)

## الغرض

جمع بيانات المريض ومعايير الأهلية للفحص المبكر لسرطان البروستاتا (PSA) بنموذج Wizard خطوة-بخطوة، عربي RTL، Mobile-first، مع حفظ كل استمارة كسجل مستقل وآمن في Google Sheet — دون أن يتصل المتصفح بالشيت مباشرة أبدًا.

## البنية المعمارية (باختصار — التفاصيل في [ARCHITECTURE.md](ARCHITECTURE.md))

```
متصفح المريض (Wizard SPA)
        ↓ google.script.run
Code.gs — submitPsaForm()
        ↓ تحقق صارم (Server-side validation)
        ↓ LockService.getScriptLock()
        ↓ فحص Client_Submission_ID (منع التكرار)
        ↓ Utilities.getUuid() + Submitted_At من السيرفر
Google Sheet (PSA_RESPONSES)
```

## ملفات المشروع (Apps Script)

كل ملفات Apps Script داخل [src/](src):

| ملف | الدور |
|---|---|
| [src/appsscript.json](src/appsscript.json) | Manifest: V8 runtime، إعدادات Web App |
| [src/Code.gs](src/Code.gs) | كل المنطق السيرفري (doGet, submitPsaForm, validation, Lock, setupProject) |
| [src/Index.html](src/Index.html) | هيكل الصفحة الأساسي (يجمع Styles + JavaScript) |
| [src/Styles.html](src/Styles.html) | كل CSS (Mobile-first, RTL, Healthcare theme) |
| [src/JavaScript.html](src/JavaScript.html) | محرّك الـ Wizard بالكامل (Vanilla JS، بدون مكتبات خارجية) |

## هيكل Google Sheet

اسم الشيت: **`PSA_RESPONSES`** داخل Spreadsheet باسم **`PSA_RESPONSES_DB — Prostate Awareness Campaign`**.

Header ثابت في الصف الأول (34 عمودًا)، أهمها:

```
Submission_ID | Submitted_At | Client_Submission_ID | Region | Full_Name | Address | Age | Phone
| Smoker | Previous_PSA | Previous_PSA_Result
| Q1_Age50_GoodHealth | Q3_Age40_FamilyHistory_EarlyOnset | Q4_FamilyHistory_Breast_Ovarian_Pancreatic_BRCA
| Q5_Urinary_Symptoms | Back_Pain_Pelvic_Numbness | Unexplained_Weight_Loss | Q6_No_PSA_Last_2Years
| Q7_Chronic_Disease | Chronic_BPH | Chronic_UTI | Chronic_Hypertension | Chronic_Diabetes
| Chronic_Other | Chronic_Other_Text
| Q8_No_Previous_Prostate_Cancer | Previous_Cancer_Following_Specialist | BPH_Medications | Q9_Wants_PSA
| Age_Above_75 (محسوب تلقائيًا من العمر)
| Staff_Recommendation | Staff_Notes | Staff_Reviewed_By | Staff_Reviewed_At (فارغة — محجوزة لـ Staff View مستقبلي)
```

القائمة الكاملة والترتيب الدقيق موجودان في `FIELD_MAP` بأعلى [src/Code.gs](src/Code.gs) — هذا هو المصدر الوحيد للحقيقة (Single Source of Truth)، ويُستخدم لكل من إنشاء الـ Header وكتابة الصفوف، فلا يمكن أن يختلفا.

> ⚠️ رقم السؤال **2 غير موجود** في الاستمارة الأصلية (فجوة ترقيم حقيقية في مصدر Word) — لم يُخترع أي سؤال بديل له، والترقيم في الأعمدة يعكس ذلك بأمانة (Q1, Q3, Q4...).

## كيفية تغيير Spreadsheet ID (أو نقل المشروع لحساب/بيئة أخرى)

الـ `SPREADSHEET_ID` **غير موجود في أي كود HTML/JS** — محفوظ فقط في **Script Properties** (لا يظهر للمتصفح إطلاقًا):

1. افتح [محرر Apps Script](https://script.google.com/d/1Gxk0elewDm8K-iqDQ3r-8VlDODStQrBxCi60E3DdZi7gGnriLYQtlNmP/edit).
2. من الترس ⚙️ (Project Settings) → **Script Properties**.
3. عدّل قيمة `SPREADSHEET_ID` يدويًا، أو ببساطة احذفها ثم شغّل الدالة `setupProject()` مرة أخرى (من القائمة المنسدلة أعلى المحرر ثم Run) — ستُنشئ Spreadsheet جديدًا تلقائيًا وتحفظ الـ ID الجديد.

## كيفية تشغيل clasp (PowerShell)

المشروع مربوط بالفعل عبر `.clasp.json` (في جذر المجلد، `rootDir: "src"`).

```bash
clasp show-authorized-user   # التحقق من تسجيل الدخول
clasp login                  # إذا لزم تسجيل الدخول (يفتح المتصفح)
clasp show-file-status       # معاينة الملفات التي سترفع (dry run)
clasp push --force           # رفع الكود إلى Apps Script
clasp list-deployments       # عرض الـ deployments الحالية
```

> ملاحظة إصدار: هذا المشروع استُخدم مع **clasp 3.4.0**، وأوامره تختلف عن clasp 2.x — مثلاً `clasp create-script` بدل `clasp create`، و`clasp create-deployment` / `clasp update-deployment` بدل `clasp deploy`. تحقق دائمًا بـ `clasp --version` و`clasp <command> --help` قبل التنفيذ إن تغيّر الإصدار.

## كيفية تحديث الـ Web App بعد أي تعديل بالكود

**لا تُنشئ Deployment جديدًا في كل مرة** — حدّث نفس الـ deployment الموجود حتى يبقى الرابط ثابتًا للمرضى:

```bash
clasp push --force
clasp update-deployment AKfycby9HptliHahlzHKlhWxqWKZAg-JfKtZsBrlgsAVTm6AxR9gpRBhryEYXiLpP5DzIBk3pA --description "وصف التحديث"
```

رابط الـ Web App **يبقى كما هو** لأنه مرتبط بمعرّف الـ deployment وليس برقم الإصدار.

## كيفية اختبار المشروع

- **اختبار وظيفي سريع:** افتح رابط الـ Web App على الموبايل وجرّب المسار كاملًا (منطقة → بيانات → أسئلة طبية → مراجعة → إرسال).
- **اختبار الخادم (Server-side logic):** يمكن إعادة تفعيل دالة `doPost` مؤقتًا (محذوفة حاليًا من نسخة الإنتاج عمدًا) لإرسال طلبات HTTP مباشرة عبر PowerShell/`Invoke-RestMethod` ومحاكاة حالات التزامن (راجع تقرير الاختبار الكامل في محادثة التسليم، أو [ARCHITECTURE.md](ARCHITECTURE.md) لمنهجية الاختبار).
- **إعداد أولي/بيئة جديدة:** شغّل `setupProject()` من المحرر مرة واحدة لإنشاء الشيت والـ Header.

## الحالة الحالية

الشيت فارغ من بيانات الاختبار (تم تنظيفها بعد التحقق) وجاهز لاستقبال بيانات المرضى الحقيقية.
