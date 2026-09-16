<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — حزمة المكتب المفتوحة المصدر المدعومة بالذكاء الاصطناعي: Docs و Sheets و Slides و PDF و Markdown و HTML مع لوحة ذكاء اصطناعي مدمجة" width="100%">
    </picture>
  </a>
</p>

<h1 align="center" dir="rtl">GenOffice</h1>

<p align="center" dir="rtl"><b>أول حزمة Office متكاملة المزايا ومفتوحة المصدر تعمل بالذكاء الاصطناعي في العالم.</b><br>
ملفات Word و Excel و PowerPoint و PDF، يحرّرها أنت والذكاء الاصطناعي الخاص بك، وتُحفظ بالصيغ الأصلية نفسها.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="الترخيص: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="أحدث إصدار"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="عدد التنزيلات"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="نجوم GitHub"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <b>العربية</b> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center" dir="rtl">
  <a href="#download"><b>التنزيل</b></a> ·
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
  <a href="https://genoffice.ai/"><b>الموقع الإلكتروني</b></a> ·
  <a href="https://genoffice.ai/join"><b>المجتمع</b></a> ·
  <a href="../../PRIVACY.md"><b>الخصوصية</b></a>
</p>

GenOffice بديل مجاني ومفتوح المصدر لـ Microsoft Office على macOS و Windows
و Linux. يفتح البرنامج ويحفظ مباشرة ملفات `.docx` و `.xlsx` و `.pptx` الأصلية،
ويحرّر ملفات PDF و Markdown و HTML، ويضع وكيل ذكاء اصطناعي إلى جانب كل مستند —
لا كمربع محادثة مُلحق بشكل عرضي، بل كمحرّر يقرأ الملف، وينفّذ التغيير، ويوضّح
لك تحديدًا ما لمسه.

- **صِيَغ أصلية، حفظ يحافظ على البايتات.** يُعاد كتابة الجزء الذي تحرّره فقط،
  وتبقى بقية الملف كما هي حرفيًا بايتًا ببايت، فتستمر المستندات في العمل
  بسلاسة في Word و Excel و PowerPoint.
- **ذكاء اصطناعي يمكنك مراجعته.** تصل التعديلات على شكل تتبّع للتغييرات
  وفروقات (diffs) قابلة للتراجع بضغطة واحدة. تحصل جداول البيانات على صِيَغ
  حيّة، لا أرقام ملصقة؛ وتُولَّد الشرائح والصفحات مباشرة على لوحة الرسم وتبقى
  قابلة للتحرير بالكامل.
- **محلي بالتصميم.** تُفتح الملفات وتُحرَّر وتُحفَظ وتُحوَّل على جهازك مباشرة.
  جميع عمليات التحويل — PDF إلى Word / Excel / PowerPoint، وMarkdown إلى Word،
  وHTML إلى Word — تعمل على الجهاز نفسه. لا يغادر جهازك شيء سوى طلبات الذكاء
  الاصطناعي، وتُرسَل فقط إلى مزوّد الخدمة الذي تختاره.
- **مفاتيحك أو بلا مفاتيح.** سجّل الدخول بواسطة Genspark لتجنّب إدارة المفاتيح،
  أو استخدم مفتاحك الخاص مع Claude أو OpenAI أو Gemini أو DeepSeek أو Kimi أو
  GLM أو Qwen أو Doubao أو MiniMax أو Grok أو Mistral أو OpenRouter أو Requesty، أو أي
  نقطة نهاية متوافقة مع OpenAI، بما في ذلك الخوادم المحلية.
- **قابل للبرمجة وجاهز للوكلاء.** يأتي التطبيق مع سطر أوامر `genoffice` ومهارة
  لكل من Claude Code وCodex وCursor وGemini CLI وGitHub Copilot وOpenCode
  وWindsurf، ليتمكّن وكيل البرمجة من إنشاء ملفات Office حقيقية وتحويلها
  وقراءتها وتحريرها على جهازك دون فتح أي نافذة.

**الحصول عليه:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon و Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 و Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb، rpm، AppImage) —
التفاصيل والمتطلبات في [التنزيل](#download).

## العرض التوضيحي

ست تطبيقات، ولوحة ذكاء اصطناعي واحدة، وسطر أوامر لوكيل البرمجة الخاص بك. كل
لقطة شاشة هي من التطبيق الفعلي على macOS، والذكاء الاصطناعي فيها ناتج عن
التوجيه (prompt) الذي يمكنك قراءته في اللوحة.

### 1 · Docs — افتح وحرّر ملفات `.docx` بمساعدة ذكاء اصطناعي يمكنك مراجعته

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="تطبيق GenOffice Docs يعرض صفحة تقرير سنوي بعمودين مع صورة غلاف بعرض الصفحة كاملة، وجدول مؤشرات أداء مظلّل، وترويسة وتذييل، بتكبير 80% مع لوحة الذكاء الاصطناعي مطوية"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: نظرة عامة عن الشركة مع صورة بانر؛ أعاد الذكاء الاصطناعي صياغة قسم النظرة العامة بإحكام وأضاف قسمًا جديدًا بنقاط، وتوفّر اللوحة تراجعًا بضغطة واحدة"></td>
</tr>
<tr>
<td dir="rtl"><b>يفتح الملف كما يعرضه Word تمامًا</b> — أقسام بعمودين، صور بعرض الصفحة كاملة، جداول مظلّلة، ترويسات وتذييلات، وتقسيم إلى صفحات بحسب مقاييس أسطر Word. الأنماط والتعليقات وتتبّع التغييرات والمعادلات والكتابة بالحبر تبقى سليمة دون تغيير عند الحفظ.</td>
<td dir="rtl"><b>اطلب التعديل</b> — يقرأ الذكاء الاصطناعي الكتل النصية التي يحتاجها، ويعيد كتابة قسم النظرة العامة ويضيف قسمًا جديدًا بنقاط. كل خطوة ذكاء اصطناعي هي لقطة يمكنك التراجع عنها؛ ومع تفعيل <b>تتبّع التغييرات</b>، تصل التعديلات كمراجعات على طراز Word.</td>
</tr>
</table>

### 2 · Sheets — ملفات `.xlsx` بصِيَغ ورسوم بيانية حيّة، لا أرقام ملصقة

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: أضاف الذكاء الاصطناعي ورقة ملخّص (Summary) تعرض الإيرادات بحسب المنطقة والفئة باستخدام صِيَغ SUMIF، إضافة إلى رسم بياني بالأعمدة، مع تقرير بـ 43 تغييرًا مُطبَّقًا وزر تراجع"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: عند سؤاله عن المنطقة الأعلى إيرادًا في الربع الثاني، يجيب الذكاء الاصطناعي بأوروبا مع تفصيل حسب الفئة، ويستشهد بالخلايا التي استخدمها كروابط، إلى جانب ورقة Orders"></td>
</tr>
<tr>
<td dir="rtl"><b>ابنِه</b> — من جملة واحدة، يضيف الوكيل ورقة ملخّص بصِيَغ <code>SUMIF</code> فعلية حسب المنطقة والفئة، ويُدرج رسمًا بيانيًا بالأعمدة، ويطبّق الـ43 تغييرًا كدفعة واحدة قابلة للتراجع.</td>
<td dir="rtl"><b>اسأله</b> — تُجاب الأسئلة عن المصنّف مع شرح المنطق والخلايا الدقيقة المستخدمة كاستشهادات قابلة للنقر. ويعمل خلف الكواليس: محرّك <code>.xlsx</code> داخلي مطوَّر بلغة Rust، وجداول محورية، ومرشِّحات تقطيع (slicers)، وتنسيق شرطي، وتتبّع للصِّيَغ.</td>
</tr>
</table>

### 3 · Slides — من توجيه (prompt) إلى عرض تقديمي `.pptx` كامل

<img src="../assets/readme/slides-generate.webp" alt="لقطات متسارعة لـ GenOffice Slides وهو يُنشئ عرض المستثمرين Aurora Home: يخطط الذكاء الاصطناعي القصة في اللوحة، وتظهر الشرائح على القماش واحدة تلو الأخرى، وينتهي العرض المكتمل بطلب الإغلاق (closing ask)" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: شريحة الغلاف لعرض تقديمي لمستثمري Aurora Home من إنشاء الذكاء الاصطناعي على لوحة الرسم، مع التوجيه الأصلي المكوَّن من سطر واحد وملخّص الذكاء الاصطناعي لما بناه في اللوحة"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: الشريحة الختامية المصمَّمة لنفس العرض المكوَّن من 11 شريحة، مع شريط المصغّرات على اليسار ولوحة الذكاء الاصطناعي تلخّص السرد القصصي"></td>
</tr>
<tr>
<td dir="rtl"><b>سطر واحد كمُدخل</b> — "أنشئ عرضًا تقديميًا من 10 شرائح لعرضه على المستثمرين لصالح Aurora Home…". يخطّط GenOffice للسرد القصصي، ويبحث عن الأرقام، ويُعِدّ مسودة كل شريحة على لوحة الرسم كملف <code>.pptx</code> فعلي.</td>
<td dir="rtl"><b>عرض تقديمي جاهز كمُخرَج</b> — إحدى عشرة شريحة مصمَّمة بخطوط وصور متّسقة ودعوة ختامية لاتخاذ إجراء؛ استمر في التحرير عبر الشرائح الرئيسية والتخطيطات والأدلة الذكية والقصّ غير المدمِّر، أو اطلب من اللوحة إعادة التصميم أو الصياغة أو إعادة الترتيب.</td>
</tr>
</table>

### 4 · PDF — حرّر نص PDF في مكانه، وحوّل PDF إلى Word على الجهاز مباشرة

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: يحدّد وضع تحرير النص كل كتلة نصية في الصفحة للتحرير في مكانها، بينما تجيب لوحة الذكاء الاصطناعي عن سؤال حول التقرير مع استشهادات بأرقام الصفحات"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="تطبيق GenOffice Docs يعرض مستند Word محوَّلًا محليًا من ملف PDF الخاص بالمراجعة الفصلية لشركة Helios، مفتوحًا في تبويب ثانٍ إلى جانب ملف PDF الأصلي"></td>
</tr>
<tr>
<td dir="rtl"><b>حرّر داخل الصفحة</b> — يحدّد وضع تحرير النص كل كتلة نصية لإعادة الكتابة في مكانها؛ ويُعاد كتابة تدفّق المحتوى عبر PDFium بالخطوط الأصلية نفسها، لا بتعليق تغطية فوق النص. اسأل الذكاء الاصطناعي عن تقرير طويل واحصل على إجابات مع استشهادات بأرقام الصفحات.</td>
<td dir="rtl"><b>حوّل على الجهاز مباشرة</b> — تُنتج ميزة <b>PDF Converter → PDF to Word</b> ملف <code>.docx</code> قابلًا للتحرير يُفتح في Docs إلى جانب المصدر، مع بقاء العناوين وصفوف الإحصاءات والفقرات سليمة. تعمل صيغتا الوجهة Excel و PowerPoint بالطريقة نفسها؛ وتمرّ الصفحات الممسوحة عبر تقنية OCR الخاصة بنظام التشغيل.</td>
</tr>
</table>

### 5 · HTML — أداة بناء صفحات وواجهات بالذكاء الاصطناعي، تبدأ بموجز تصميم

أخبر الذكاء الاصطناعي بالغرض من الصفحة والجهة المستهدفة بها. يقترح الذكاء
الاصطناعي أولًا **موجز تصميم (design brief)** — يشمل الفكرة الجاذبة، ولوحة
الألوان، والخطوط، وتوجهات الأسلوب — ثم يبني ملف `.html` واحدًا مستقلًا بالكامل
استنادًا إلى تلك العناصر (tokens).

<img src="../assets/readme/html-restyle-motion.webp" alt="لقطات متسارعة لـ GenOffice HTML وهو يعيد تصميم صفحة الوصول الخاصة بـ Lumen: طلب Restyle واحد في اللوحة يحوّل صفحة Midnight Studio الداكنة إلى نسخة Solar Daybreak الدافئة، بينما يبقى كل قسم وكل المحتوى في مكانه" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: صفحة هبوط تم إنشاؤها لمصباح مكتبي يعمل بالطاقة الشمسية باتجاه Midnight Studio الداكن، معروضة في المعاينة المباشرة مع لوحة الذكاء الاصطناعي تلخّص الصفحة التي بنتها للتو"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="نفس صفحة هبوط Lumen بعد إعادة تصميمها بواسطة الذكاء الاصطناعي إلى اتجاه Solar Daybreak الدافئ: خلفية بلون الورق، وعناوين بخط Serif، ولمسة برتقالية، مع الحفاظ على جميع الأقسام والنصوص"></td>
</tr>
<tr>
<td dir="rtl"><b>يُنشأ من توجيه واحد</b> — قسم رئيسي جريء، وبطاقات ميزات، وتسعير، ونموذج قائمة انتظار لصفحة Lumen، مبنيّة باتجاه Midnight Studio. اضغط على أي عنصر لإعادة تصميمه، أو اضغط ضغطًا مزدوجًا لتحرير النص، أو انتقل إلى عرض الكود المصدري في CodeMirror.</td>
<td dir="rtl"><b>نفس التصميم، اتجاه جديد</b> — طلب <b>إعادة تصميم (Restyle)</b> واحد يستبدل عناصر الموجز فتتبعه الصفحة تلقائيًا: خلفية ورقية دافئة، وخط Serif تحريري، ولمسة برتقالية شمسية، دون إعادة كتابة أي نص. اعرضها بملء الشاشة، أو صدّرها كملف PDF أو مستند Word أصلي قابل للتحرير.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: واجهة لوحة تحكم شخصية تم إنشاؤها لمصمِّم مستقل بأسلوب كتّاني دافئ، مع شريط جانبي أيسر، وترحيب بخط Serif، وأربع بطاقات مقاييس"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: تقرير بيانات عن سوق السيارات الكهربائية تم إنشاؤه بأسلوب الصحيفة الكبيرة (broadsheet)، مع ترويسة بخط Serif، ورقم عنوان رئيسي بقيمة 17.3 مليون، وصفّ إحصاءات"></td>
</tr>
<tr>
<td dir="rtl"><b>نماذج واجهات (UI mockups)</b> — تحوّل قالب البداية "لوحة التحكم الشخصية" شخصية افتراضية إلى تخطيط عملي: شريط جانبي، ترحيب، رسم Sparkline لساعات العمل القابلة للفوترة، وبطاقات فواتير ونسب استغلال، وكل ذلك HTML فعلي يمكنك تسليمه لمطوِّر.</td>
<td dir="rtl"><b>قصص البيانات</b> — يبني قالب البداية "تقرير البيانات" صحيفة تحريرية: ترويسة بخط Serif، ورقم عنوان رئيسي واحد، وصفّ إحصاءات مفصول بخطوط، ورسوم بيانية SVG مضمّنة، وملاحظة عن المنهجية.</td>
</tr>
</table>

### 6 · Markdown — محرِّر كُتل فوق ملفات `.md` النصية الصِّرفة، مع ميزة Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: فقرة محدَّدة تعرض نافذة منبثقة Ask AI مع تعليمات مكتوبة وشرائح اقتراحات مثل Polish وMake more concise وExpand وFix grammar، إضافة إلى زرَّي Send now وAdd to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="تطبيق GenOffice Markdown يعرض مستند ملاحظات إطلاق يحتوي على جدول، ومخطط تدفّق Mermaid، وقائمة مهام، مع توجيهات البداية الخاصة بلوحة الذكاء الاصطناعي على اليسار"></td>
</tr>
<tr>
<td dir="rtl"><b>اسأل الذكاء الاصطناعي عن نص محدَّد</b> — حدِّد أي فقرة فتظهر شريحة <b>Ask AI</b>: اكتب تعليمًا أو اختر اقتراحًا، وأرسله فورًا، أو أضِف عدة تعديلات مرتبطة بمواضع محدَّدة إلى قائمة انتظار ونفِّذها في مرة واحدة. نفس المدخل موجود في كل التطبيقات.</td>
<td dir="rtl"><b>يُعرض بصريًا، ويُحفَظ كملف Markdown نصي صِرف</b> — عناوين، وقوائم، وجداول، وصور، وكتل كود، ومخططات Mermaid داخل محرِّر كُتل من Tiptap، تُكتب مرة أخرى كملف <code>.md</code> نصي صِرف، مع تصدير <b>Markdown → Word</b> يعمل محليًا بالكامل.</td>
</tr>
</table>

### 7 · CLI — وكيل البرمجة الخاص بك يقود GenOffice، على جهازك

يأتي GenOffice مع سطر أوامر `genoffice` ومهارة وكيل. ثبّت المهارة، وسيتمكّن
Claude Code أو Codex أو Cursor أو Gemini CLI أو GitHub Copilot أو OpenCode أو
Windsurf من إنشاء ملفات Office حقيقية وتحويلها وقراءتها وتحريرها عبر المحرّكات
نفسها التي تستخدمها التطبيقات، دون فتح أي نافذة.

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides يعرض عرضًا تقديميًا من ثماني شرائح عن المجموعة الشمسية بناه وكيل برمجة عبر سطر أوامر genoffice: شريحة الغلاف على لوحة الرسم، وثماني مصغّرات على اليسار، ولوحة الذكاء الاصطناعي مفتوحة" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="الشرائح الثماني المُصيَّرة لعرض المجموعة الشمسية جنبًا إلى جنب: الغلاف، والجدول الزمني للاستكشاف، وأربعة أرقام رئيسية، ورسم بياني بالأعمدة لأقطار الكواكب، والعوالم الصخرية مقابل العمالقة، ورقم الشمس البارز 99.8%، وشبكة العمالقة الأربعة، والخلاصات"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="إعدادات GenOffice، صفحة Integrations: مهارة genoffice مثبَّتة في Claude Code، مع أزرار Install بجوار Codex وCursor"></td>
</tr>
<tr>
<td dir="rtl"><b>توجيه واحد إلى وكيلك</b> — "أنشئ عرضًا تقديميًا من ثماني شرائح عن المجموعة الشمسية." يقرأ الوكيل المهارة، ويكتب ورقة أنماط ومخططًا ومواصفة صفحة واحدة لكل شريحة، ويولّد الصورتين بواسطة <code>genoffice image</code>، ويترك <code>genoffice slides check</code> يرفض أي عنصر يتجاوز الحدود أو يتداخل قبل أن يجمّع <code>genoffice create</code> ملف <code>.pptx</code> ويُعيد <code>slides render</code> صورة PNG لكل شريحة لمعاينتها.</td>
<td dir="rtl"><b>ثبّت مرة واحدة، من الإعدادات → التكاملات</b> — يسرد GenOffice وكلاء البرمجة الذين يجدهم على هذا الجهاز ويكتب المهارة في كل وكيل تختاره. أو نزّل المهارة كملف zip، أو شغّل <code>npx skills add genspark-ai/genoffice</code>. الأوامر وسير العمل الكامل في <a href="#command-line-and-agent-skill">سطر الأوامر ومهارة الوكيل</a>.</td>
</tr>
</table>

### 8 · MCP — نفس الأدوات عبر Model Context Protocol

كل أمر من أوامر `genoffice` هو أيضًا أداة MCP. يمكن لـ Claude Code وClaude
Desktop وCursor وأي عميل MCP آخر تشغيل `genoffice mcp` بأنفسهم، دون الحاجة
إلى تثبيت مهارة أو فتح نافذة، والحصول على 29 أداة بالإضافة إلى مراجع
العمليات كموارد (resources). كما يتيح خادم HTTP ثانٍ داخل التطبيق للوكيل
بناء مستند Word في تبويب محرِّر مرئي بينما تشاهد.

<img src="../assets/readme/mcp-deck-motion.webp" alt="لقطات متسارعة لـ Claude Code وهو يبني عرضًا استثماريًا من ثماني شرائح عن الطاقة المتجدّدة عبر خادم MCP الخاص بـ genoffice: يبحث عن الرسوم البيانية والصور، ويتحقق من كل صورة مرشّحة عبر media، وdeck_start يكتب ورقة الأنماط والمخطط، وdeck_page يضيف صفحة واحدة مُتحقَّقة في كل مرة، وdeck_build يجمّع ملف .pptx، وslides_render يعيد صورة كل شريحة؛ ثم يُفتح العرض النهائي في GenOffice Slides" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides يعرض عرض Renewable Energy 2026 من ثماني شرائح الذي بناه Claude Code عبر خادم MCP الخاص بـ genoffice: شريحة الغلاف مع صورة فوتوغرافية لمزرعة رياح على لوحة الرسم وثماني مصغّرات على اليسار"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="إعدادات GenOffice، صفحة Integrations، جزء MCP: أمر claude mcp add من سطر واحد لـ Claude Code، وكتلة JSON لـ Cursor وClaude Desktop وعملاء MCP الآخرين، وخيار خادم HTTP المحلي أسفل ذلك"></td>
</tr>
<tr>
<td dir="rtl"><b>توجيه واحد، 38 استدعاءً للأدوات، بلا طرفية</b> — "ابنِ عرضًا استثماريًا من ثماني شرائح عن الطاقة المتجدّدة في 2026، بصورة حقيقية على الغلاف وحيثما تفيد صورة." يسحب الوكيل الرسوم البيانية والصور عبر <code>search</code>، ويسأل <code>media</code> عمّا إذا كانت كل صورة مرشّحة صورة فوتوغرافية حقيقية، ثم يستدعي <code>deck_start</code> بورقة أنماط ومخطط، ثم <code>deck_page</code> لكل شريحة؛ تُفحص كل صفحة مقابل المخطط والباليت قبل الاحتفاظ بها، ويجمّع <code>deck_build</code> ملف <code>.pptx</code>، ويبحث <code>slides_audit</code> عن التجاوزات، ويعيد <code>slides_render</code> صورة PNG لكل شريحة كمحتوى صوري يمكن للنموذج النظر إليه، ويُصلح <code>deck_replace</code> الصفحات الثلاث التي لم تعجبه.</td>
<td dir="rtl"><b>اتصل مرة واحدة، من الإعدادات → التكاملات</b> — انسخ سطر <code>claude mcp add</code> لـ Claude Code، أو كتلة JSON إلى Cursor أو Claude Desktop أو أي عميل MCP آخر. يفعِّل الخيار B خادم HTTP المحلي لمحرِّر Word المرئي. كلاهما موضّح في <a href="#mcp-server">خادم MCP</a>.</td>
</tr>
</table>

## لماذا GenOffice

- **مفتوح المصدر**، برخصة Apache-2.0، ويُطوَّر علنًا على GitHub.
- **لك لتشغيله.** تطبيقات أصلية لأنظمة macOS و Windows و Linux؛ تبقى الملفات
  على قرصك، وكل عملية تحرير وحفظ وتحويل تتم على جهازك.
- **ملفات Office حقيقية.** دعم أصلي لـ `.docx` و `.xlsx` و `.pptx`، مع الحفاظ
  على البايتات: الأجزاء التي لم تلمسها من الملف تُنسخ كما هي تمامًا.
- **ذكاء اصطناعي يحرِّر المستند نفسه.** تتبّع التغييرات في Docs، وصِيَغ ورسوم
  بيانية حيّة في Sheets، وشرائح تُرسم مباشرة على لوحة الرسم، وكل جولة AI هي لقطة
  يمكنك التراجع إليها.
- **نموذجك، مفتاحك.** سجّل الدخول عبر Genspark، أو استخدم مفتاحك الخاص مع
  Claude أو OpenAI أو Gemini أو DeepSeek وغيرها، مع دعم الخوادم المحلية وأي نقطة
  نهاية متوافقة مع OpenAI.
- **PDF يُعالَج كما ينبغي.** حرِّر النص داخل الصفحة مباشرة، وحوِّل PDF إلى Word
  أو Excel أو PowerPoint على جهازك، مع OCR للنظام لملفات المسح الممسوحة.
- **Markdown و HTML أيضًا**، بنفس لوحة الذكاء الاصطناعي وتصدير محلي إلى Word.
- **قابل للبرمجة.** سطر أوامر `genoffice`، ومهارة وكيل، وخادم MCP يضعون كل
  محرّك في خدمة Claude Code وClaude Desktop وCodex وCursor وغيرها من الوكلاء،
  مع بقاء كل شيء على الجهاز.
- **مجاني**، للأفراد والفرق على حد سواء.

## خلفيات الذكاء الاصطناعي

**سجّل الدخول بواسطة Genspark** فلا تحتاج إلى أي إعداد: تُوجَّه طلبات النماذج
عبر بروكسي Genspark (عائلات Claude و GPT و Gemini)، ويحصل الوكلاء على البحث
عبر الويب والصور، وتوليد الصور، وتحليل الصور والصوت والفيديو.

**أو استخدم مفتاحك الخاص.** تسرد صفحة Settings → AI مزوِّدين مثل Claude
وOpenAI وGemini وDeepSeek وKimi وGLM وQwen وDoubao وMiniMax وGrok وMistral
وOpenRouter وRequesty وOpenCode Zen/Go، إضافة إلى خانة مخصَّصة لأي نقطة نهاية متوافقة مع
OpenAI (عنوان أساسي + مفتاح)، بما في ذلك خوادم النماذج المحلية. ولكل من
البحث والوسائط مزوِّدوه الخاصون بحسب القدرة تحت **AI Media & Search**: Serper
أو Tavily للبحث عبر الويب، وOpenAI أو Gemini أو Doubao/Seedream أو GLM أو
Grok أو Qwen أو MiniMax أو أي نقطة نهاية للصور متوافقة مع OpenAI لتوليد
الصور وتحليل الصور والفيديو.

تأتي الحزمة كاملةً بثلاثة مظاهر: فاتح وداكن ومظهر النظام. تغيّر المظاهر فقط ما
يظهر على الشاشة: تحتفظ ملفات التصدير والطباعة والحفظ دائمًا بألوان المستند
نفسه.

<a id="command-line-and-agent-skill"></a>

## سطر الأوامر ومهارة الوكيل

كل ما تستطيع التطبيقات فعله بملف، يستطيع سطر أوامر `genoffice` فعله من
الطرفية: فحص ملفات Word وExcel وPowerPoint وPDF وMarkdown وHTML وتحويلها
وإنشاؤها وقراءتها وتحريرها على المحرّكات نفسها، دون واجهة رسومية. يُثبَّت مع
GenOffice، ولا يحتاج إلى بيئة تشغيل خاصة به، ولا يرسل أي مستند إلى أي مكان.
وبالاقتران مع **مهارة الوكيل** المرفقة، يحوّل وكيل البرمجة إلى عامل مستندات
يُنتج ملفات Office حقيقية بدلًا من تقريبات بصيغة Markdown.

**يعمل مع:** Claude Code وCodex وCursor وGemini CLI وGitHub Copilot وOpenCode
وWindsurf مباشرة دون أي إعداد، ومع أي وكيل آخر يقرأ المهارات، وكذلك، عبر
[خادم MCP](#mcp-server)، مع Claude Desktop وأي عميل MCP.

### تثبيت المهارة

| الطريقة                                | ما يحدث                                                                                                                                            |
| -------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| **الإعدادات → التكاملات** في التطبيق   | يسرد الوكلاء الموجودين على هذا الجهاز؛ وبنقرة واحدة يكتب المهارة في كل وكيل تختاره. ويظهر زر **Update** عندما يصدر إصدار من GenOffice بمهارة أحدث. |
| **التنزيل كملف zip** من الصفحة نفسها   | التنسيق الذي يقبله claude.ai وتطبيقات Claude لسطح المكتب والمساعدون الآخرون كمهارة مرفوعة.                                                         |
| `npx skills add genspark-ai/genoffice` | يثبّت من هذا المستودع في أي وكيل متوافق مع المهارات.                                                                                               |

ثم ابدأ محادثة جديدة واطلب مستندًا. تعلّم المهارة الوكيل متى يلجأ إلى
`genoffice`، وكيف يقرأ الملف قبل تحريره، وكيف يتحقّق من عمله بنفسه.

### بداية سريعة من الطرفية

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

يطبع كل أمر ملخّصًا من سطر واحد، أو كائن JSON واحدًا مع `--json`. التعديلات
ذرّية: العملية المرفوضة تترك الملف كما هو وتعود برسالة خطأ إرشادية. يسرد
`genoffice help` مجموعة الأوامر الحالية؛ والمرجع الكامل في
[packages/cli/README.md](../../packages/cli/README.md).

### ما الذي يشغّله الوكيل فعليًا

عرض المجموعة الشمسية في العرض التوضيحي أعلاه استغرق توجيهًا واحدًا في Claude Code.
وخلف الكواليس، اتّبع الوكيل سير العمل المرحلي للمهارة، وتحقّق سطر الأوامر من
كل مرحلة قبل أن تبدأ التالية:

```bash
genoffice capabilities --json                        # which cloud tools GenOffice has configured
genoffice guide slides design                        # the deck workflow and layout library
genoffice image "the eight planets in a row …" --aspect 16:9 --out deck/assets/cover.jpg
genoffice slides check deck/outline.json --json      # 8 pages, no findings
genoffice slides check deck/pages/01.json --json     # builds one slide, audits overflow and overlap
…                                                    # one page file per slide, fixed until each check is clean
genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/solar-system.pptx --json
genoffice slides render deck/solar-system.pptx --out deck/shots --json
genoffice slides audit deck/solar-system.pptx --json    # 8 slides, no layout issues
genoffice slides replace deck/solar-system.pptx --slide 4 --spec deck/pages/05.json --json
genoffice open deck/solar-system.pptx
```

لا يحدث أي استدعاء لنموذج داخل `genoffice`: الوكيل هو من يفكّر، وسطر الأوامر
هو من يبني ويتحقّق، وتُفتح النتيجة في GenOffice أو PowerPoint كملف `.pptx` عادي.

<a id="mcp-server"></a>

### خادم MCP

الأوامر نفسها متاحة كأدوات [Model Context Protocol](https://modelcontextprotocol.io)
للمساعدين الذين لا يمكنهم تشغيل طرفية، أو الذين تفضّل ألا تمنحهم واحدة. هناك
طريقتان للدخول، وكلتاهما موضَّحتان بمقتطفات جاهزة للنسخ في **الإعدادات →
التكاملات → MCP**:

| الطريقة                            | ما هي                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| ---------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **A · `genoffice mcp`** (موصى بها) | خادم stdio يشغّله المساعد بنفسه؛ لا حاجة لفتح GenOffice. أداة واحدة لكل أمر (`info`، `convert`، `create_docx`، `create_xlsx`، `create_pptx`، `create_pdf`، `docs_read` / `docs_apply` / `docs_check`، `sheet_*`، `slides_*`، `render`، `guide`، `search`، `image`، `media`، `open`) بالإضافة إلى تدفّق العرض المرحلي `deck_start` → `deck_page` → `deck_build` → `deck_replace`. تُمرَّر العمليات والمواصفات وMarkdown مباشرة ضمن الطلب، لذا يعمل حتى العميل بلا نظام ملفات. |
| **B · خادم HTTP محلي**             | يعمل داخل تطبيق GenOffice على `http://127.0.0.1:3093/mcp` (Streamable HTTP، مع دعم SSE القديم). تُشغِّل أدواته تبويب محرِّر Word المرئي: `create_session`، `insert_content`، `replace_blocks`، `apply_ops`، `read_document`، `save_session`، وتشاهد المستند وهو يتشكّل. مُعطَّل افتراضيًا؛ فعِّله من لوحة الإعدادات نفسها.                                                                                                                                                   |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor أو Claude Desktop أو أي عميل MCP آخر
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

`genoffice` هنا هو سطر الأوامر المرفق داخل التطبيق (على macOS في
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`؛ تعرض لوحة
الإعدادات المسار الدقيق لتثبيتك). يحمل الخادم تعليمات سير عمله الخاصة
ويعرض مراجع العمليات كموارد `genoffice://guide/*`، لذا لا حاجة لأي مهارة؛
يمكن للمهارة وخادم MCP التعايش معًا، ويختار المساعد أحدهما. الميزات
السحابية (`search`، `image`، `media`) لا تزال تمرّ عبر المزوِّد المُعدّ في
GenOffice؛ أما كل شيء آخر فيعمل محليًا، ويحصر `GENOFFICE_ALLOWED_ROOTS` كل
أداة ضمن المجلدات التي تحدّدها.

عرض الطاقة المتجدّدة في العرض التوضيحي أعلاه هو ما يبدو عليه توجيه واحد في
Claude Code، مع تفعيل خادم MCP الخاص بـ `genoffice` فقط، من جهة البروتوكول:

```text
capabilities · guide(slides, spec) · guide(slides, design)
search(query) ×4                         → IEA, BNEF and IRENA figures for the slides
search(query, images) ×7 · media(url, ask) ×7
                                         → candidate photos, each one checked to be a real photograph
deck_start(dir, style, outline)          → outline checked: 8 pages to write
deck_page(dir, 0, page) … deck_page(dir, 7, page)
                                         → each page checked against the outline and the palette; one page sent again
deck_build(dir, out)                     → renewables-2026.pptx, no image failures
slides_audit(file) · slides_render(file, out)
                                         → no layout findings; 8 PNGs come back as image content
deck_replace(dir, n, page) ×3 · slides_render(file, out)
                                         → three pages fixed after looking at the renders
```

ثمانية وثلاثون استدعاءً، في نحو ثلاث عشرة دقيقة، دون أن يلمس المساعد أي طرفية:
الرسوم البيانية والصور وكل الأدلة والفحوصات والعروض المصيَّرة انتقلت كنتائج
لاستدعاءات أدوات MCP. غادر الجهاز فقط `search` و`media`، إلى المزوّد المهيَّأ
في GenOffice.

<a id="download"></a>

## التنزيل

| المنصة                               | المتطلبات                                             | التنزيل                                                                              |
| ------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------ |
| **macOS** — شريحة Apple (arm64)      | macOS 11+                                             | [أحدث نسخة `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest) |
| **macOS** — Intel (x64)              | macOS 11+                                             | [أحدث نسخة `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** (x64، لمعظم الأجهزة)     | Windows 10+، معالجات Intel/AMD                        | [أحدث مثبِّت `-x64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** على Arm (ARM64)          | Windows 11 على Arm (معالج Snapdragon X وما شابهه)     | [أحدث مثبِّت `-arm64.exe`](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64، glibc 2.34+ (Ubuntu 22.04 أو أحدث)            | [أحدث نسخة `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **Linux** — Fedora / RHEL / openSUSE | x86_64، glibc 2.34+ (Fedora 35+، RHEL 9+، Leap 15.6+) | [أحدث نسخة `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **Linux** — توزيعات أخرى             | x86_64، glibc 2.34+، وFUSE 2                          | [أحدث نسخة `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)    |

جميع الإصدارات مبنية من الفرع `main`؛ ومثبِّتات macOS و Windows موقَّعة رقميًا.
الإصدارات الأقدم متوفرة في صفحة [Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>التثبيت على Linux</b></summary>

يتم تثبيت حزمة deb باستخدام apt — فيسحب التبعيات المطلوبة ويضيف GenOffice إلى
قائمة التطبيقات:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

على Fedora / عائلة RHEL / openSUSE، ثبِّت حزمة rpm بدلًا من ذلك:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

يعمل ملف AppImage مباشرة من مكانه: ثبِّت بيئة تشغيل FUSE 2
(`sudo apt install libfuse2`؛ وعلى Ubuntu 24.04 اسم الحزمة هو `libfuse2t64`)،
ثم اجعل الملف قابلاً للتنفيذ، وشغِّله:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## كيف يعمل

تتشارك سبعة تطبيقات Electron — Docs و Sheets و Slides و PDF و Markdown و
HTML بالإضافة إلى الغلاف (shell) ذي التبويبات — طبقة محرِّك واحدة من حزم
TypeScript خالصة، إلى جانب عملية Rust مساعدة (sidecar) لملفات `.xlsx`. الملف
الأصلي هو المرجع دائمًا: تُطبَّق التعديلات كرقعات (patches) ضيّقة النطاق، وكل
ما لم يلمسه المحرِّر يبقى سليمًا دون تغيير عند إعادة الحفظ.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

الجولة التفصيلية حزمة بحزمة (محرِّكات docx/pptx، و`pdf2docx`، و`html2docx`،
ونواة الوكيل ومزوِّدوه) موجودة في [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## التطوير

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

يحتاج تطبيق Sheets بالإضافة إلى ذلك إلى سلسلة أدوات Rust لعمليته المساعدة
الخاصة بـ xlsx (يجب أن يكون `cargo` متاحًا في PATH)؛ ويقوم أمر
`npm run build -w @genoffice/sheets` بتجميعه تلقائيًا. راجع
[CONTRIBUTING.md](../../CONTRIBUTING.md) لمعرفة الفحوصات التي يجب أن يمر بها
كل تغيير وكيفية دمج طلبات السحب (pull requests).

## المجتمع

GenOffice قيد تطوير نشط، وملاحظاتك هي ما يشكِّل مستقبله.

- **أبلغ عن خطأ أو اطلب ميزة جديدة** عبر
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **انضم إلى محادثة مجموعة GenOffice** على
  [GenTeam](https://genoffice.ai/join) للتواصل مع الفريق والمستخدمين الآخرين.
- **أضِف نجمة (Star) للمستودع** إذا كان GenOffice مفيدًا لك — فهذه أفضل طريقة
  لدعم المشروع.

## الأسئلة الشائعة

<details>
<summary><b>هل GenOffice مجاني؟</b></summary>

نعم. GenOffice مجاني ومفتوح المصدر بموجب ترخيص Apache-2.0 — لا نسخة تجريبية
ولا باقة مدفوعة للتطبيقات نفسها.

</details>

<details>
<summary><b>هل يمكن لـ GenOffice فتح ملفات Microsoft Word و Excel و PowerPoint؟</b></summary>

نعم. يفتح GenOffice ملفات `.docx` و `.xlsx` و `.pptx` الأصلية ويحفظها.
والحفظ يحافظ على البايتات: تُكتب الأجزاء التي لم تلمسها من الملف كما هي
حرفيًا بايتًا ببايت، فتستمر المستندات في العمل في Microsoft Office.

</details>

<details>
<summary><b>هل يعمل GenOffice دون اتصال بالإنترنت؟</b></summary>

تحرير المستندات محلي بالكامل — لا تغادر الملفات جهازك أبدًا عند الفتح أو
التحرير أو الحفظ أو التحويل. تحتاج ميزات الذكاء الاصطناعي (الوكلاء، والبحث،
وأدوات الصور) إلى اتصال بالشبكة، سواء عبر تسجيل الدخول بواسطة Genspark أو
باستخدام مفتاح API لنموذجك الخاص.

</details>

<details>
<summary><b>هل يمكن لـ GenOffice تحرير ملفات PDF؟</b></summary>

نعم — تحرير حقيقي لنص وصور PDF يعيد كتابة تدفّق محتوى الصفحة مع الحفاظ على
الخطوط الأصلية، وليس بتعليقات تغطية فوق النص.

</details>

<details>
<summary><b>هل يمكن لـ GenOffice تحويل PDF إلى Word أو Excel أو PowerPoint؟</b></summary>

نعم — بالكامل على الجهاز: استخراج على مستوى الحروف عبر PDFium إلى جانب تحليل
للتخطيط يعتمد على الأشكال الهندسية، دون أي خدمة سحابية أو رفع للملفات. وتشمل
هذه الميزة أيضًا الصفحات الممسوحة: تقرؤها تقنية OCR الخاصة بنظام التشغيل على
macOS و Windows، فتُحوَّل إلى نص قابل للتحرير بدلًا من صورة للصفحة.

</details>

<details>
<summary><b>هل يمكنني استخدام نموذج ذكاء اصطناعي أو مفتاح API خاص بي؟</b></summary>

نعم. إلى جانب تسجيل الدخول بواسطة Genspark الذي لا يتطلب أي مفتاح، يدعم
GenOffice استخدام مفتاحك الخاص مع Claude وOpenAI وGemini وDeepSeek وKimi
وGLM وQwen وDoubao وMiniMax وGrok وMistral وOpenRouter وRequesty وOpenCode Zen/Go،
إضافة إلى أي نقطة نهاية متوافقة مع OpenAI — بما في ذلك خوادم النماذج المحلية.
ولكل من البحث وتوليد الصور وتحليل الصور والفيديو مفاتيحه الخاصة تحت
Settings → AI Media & Search.

</details>

<details>
<summary><b>هل يمكن لـ GenOffice تحويل HTML إلى Word؟</b></summary>

نعم — يُنتج خيار Export as Word في تطبيق HTML ملف `.docx` أصليًا قابلًا
للتحرير، بالكامل على الجهاز. تُعرَض الصفحة داخل محرك Chromium المدمج وتُحوَّل
إلى عناصر Word حقيقية: عناوين، وفقرات، وقوائم، وجداول، وبطاقات، وصفوف مؤشرات
أداء، وحقول نماذج، وخلفيات صفحات؛ ولا تُضمَّن كصور إلا العناصر البصرية التي لا
نظير لها في Word (كالرسوم البيانية والأيقونات والصناديق المزخرفة).

</details>

<details>
<summary><b>هل يمكنني تشغيل GenOffice من Claude Code أو Codex أو Cursor أو من سكربت؟</b></summary>

نعم. يثبّت GenOffice سطر أوامر `genoffice` يشغّل المحرّكات نفسها دون واجهة
رسومية: افحص المستندات وحوّلها وأنشئها واقرأها وحرّرها من الطرفية أو من
سكربت، مع مخرجات `--json` للبرامج. وتعلّم مهارة الوكيل المرفقة Claude Code
وCodex وCursor وGemini CLI وGitHub Copilot وOpenCode وWindsurf استخدامه؛ ثبّتها
من **الإعدادات → التكاملات**. راجع
[سطر الأوامر ومهارة الوكيل](#command-line-and-agent-skill).

</details>

<details>
<summary><b>هل يجمع GenOffice أي بيانات؟</b></summary>

ترسل الإصدارات الرسمية المعبَّأة بيانات استخدام محدودة بشكل افتراضي، ويمكنك
تعطيل هذا الإرسال في أي وقت من Settings → General. لا تُرسَل عبر هذه البيانات
أبدًا محتويات المستندات، أو أسماء الملفات، أو مساراتها، أو هوية الحساب، أو
عناوين البريد الإلكتروني. راجع [GenOffice Privacy](../../PRIVACY.md)
للاطلاع على الإفصاح الكامل عن الأحداث والبيانات المُجمَّعة.

</details>

## الأمان

راجع [SECURITY.md](../../SECURITY.md) للتعرّف على وضع الأمان الخاص بالعمليات
(عزل واجهة العرض renderer في صندوق رمل، والتحقق من رسائل IPC، والتحكم في
الروابط الخارجية) ونماذج التهديد الخاصة بالمحتوى المُولَّد بالذكاء الاصطناعي.

## شكر وتقدير

ما كان لـ GenOffice أن يكون ممكنًا لولا هذه المشاريع المفتوحة المصدر:

- [Electron](https://www.electronjs.org/) — بيئة تشغيل سطح المكتب لكل التطبيقات.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — نواة واجهة جدول
  البيانات التي يبني عليها تطبيق Sheets.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause، مُضمَّن عبر
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) — محرِّك
  تدفّق المحتوى الذي يقوم عليه التحرير الحقيقي لنص وصور PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) و
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — عرض ملفات PDF وتجميع
  المستندات.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  محرِّرا الكُتل في تطبيقَي Docs و Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — محرِّر الكود المصدري في تطبيق HTML.
- [Konva](https://konvajs.org/) — عرض لوحة الرسم لتطبيق Slides ورسوم Sheets
  البيانية.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — مقاييس تشكيل
  النص (text-shaping) للأنظمة الكتابية المعقّدة.
- [calamine](https://github.com/tafia/calamine) و
  [IronCalc](https://github.com/ironcalc/IronCalc) — طبقتا القراءة والحساب في
  عملية Rust المساعدة الخاصة بـ xlsx.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — أداة فكّ تشفير
  MicroType Express لخطوط PowerPoint المُضمَّنة، تم نقلها إلى TypeScript.
- [React](https://react.dev/) (MIT) — طبقة واجهة المستخدم لكل التطبيقات.
- [Mermaid](https://mermaid.js.org/) (MIT) و [KaTeX](https://katex.org/)
  (MIT) — الرسوم البيانية والمعادلات الرياضية في Markdown و Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — تحليل الخطوط لاستخراج
  المقاييس وأشكال الحروف (glyphs).
- [JSZip](https://stuk.github.io/jszip/) (MIT) و
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — طبقتا حاوية OOXML و XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — مجموعة الأيقونات المستخدَمة في كل أشرطة الأدوات.
- [electron-updater](https://www.electron.build/) (MIT) — التحديثات داخل
  التطبيق.
- خطوط Liberation و Carlito و Caladea و Noto CJK (OFL/Apache-2.0) — الخطوط
  المُضمَّنة في المستندات.

يُعيد أمر `npm run notices` توليد ملخّص ترخيص الجهات الخارجية المُضمَّن
(`tools/gen-third-party-notices.mjs`)؛ وجميع تبعيات وقت التشغيل مرخَّصة بموجب
MIT/Apache-2.0/BSD-3-Clause/OFL.

## الترخيص

يخضع GenOffice لترخيص [Apache License 2.0](../../LICENSE)، مع استثناء واحد:
دليل `ee/` محجوز لوحدات المؤسسات (enterprise) المستقبلية ويخضع لـ
[GenOffice Enterprise License](../../ee/LICENSE).

اسما وشعارا GenOffice و Genspark هما علامتان تجاريتان مملوكتان لشركة
Mainfunc, Inc. لا يمنح ترخيص Apache-2.0 إذنًا باستخدامهما (راجع البند 6)؛
ويجب أن تستخدم النسخ المتفرعة (forks) هويتها البصرية الخاصة.
