<div dir="rtl">

# نظام العروض التقديمية — PowerPoint بأسلوب PlacePlate

**عروض تتحرك (Morph)، مبنية من البيانات لا من الانطباعات.**

[English README](README.md) · [SKILL.md](SKILL.md) · [الموقع](site/index.html)

نظام متكامل لبناء عروض PowerPoint احترافية بحركة Morph: مكتبة مرجعية
مدروسة، مولّد JSON ← ‏`.pptx`، **قاعدة بيانات وحشية** قابلة للاستعلام
(238 سجلاً)، بوابة فحص جودة آلية، وأدوات صانع محتوى كاملة — بالإنجليزية
والعربية (RTL كامل)، لقاعة الاجتماعات ولمنصات التواصل (Reels وCarousels).

## البدء السريع

```bash
pip install -r requirements.txt

# ١) المعالج: من الهدف إلى عرض جاهز ومفحوص
python tools/wizard.py --goal pitch --topic "وصل" --claim "نحوّل بقايا القهوة إلى تغليف"

# ٢) أو من ملاحظات خام إلى عرض (النص الأصلي يصبح ملاحظات المتحدث)
python tools/draft.py my-notes.txt --build

# ٣) أو يدوياً: حرّر ملف JSON ثم ابنِ وافحص
python generator/generate_pptx.py spec.json deck.pptx
python tools/qa_deck.py deck.pptx        # يجب أن يطبع GREEN

# ٤) حوّل أي عرض إلى حزمة تواصل اجتماعي (Reel ‏9:16 + Carousel ‏4:5)
python tools/social.py spec.json --handle @you
```

## العربية أولاً (RTL)

- `"rtl": true` في المواصفة: انعكاس كامل للهندسة، اتجاه فقرات RTL،
  قلب الأسهم والشيفرونات، خطوط Cairo/Tajawal تلقائياً.
- نوع شريحة **`bilingual`**: العربية يميناً (تُقرأ أولاً) والإنجليزية
  يساراً على الشريحة نفسها، وكل لغة بخطوطها واتجاهها الصحيح.
- `"numerals": "arabic-indic"`: الأرقام تصبح ٠١٢٣٤٥٦٧٨٩ (دون المساس
  بالنصوص الإنجليزية).
- التاريخ الهجري: اكتب `{hijri}` في أي نص ويتوسع تلقائياً
  (مثال: ٧ جمادى الأولى ١٤٤٨هـ)، مع أداة `tools/hijri.py`.
- ثيمات سعودية جاهزة: `saudi-green` و`desert-sand`.
- التوجيه الكامل في [`reference/14-arabic-and-rtl-decks.md`](reference/14-arabic-and-rtl-decks.md)،
  وعرض تجريبي ثنائي اللغة في `examples/arabic-bilingual-demo.pptx`.

## قاعدة البيانات الوحشية 🐲

SQLite ‏+ بحث نصي FTS5 — مبدعون، قوالب، تقنيات، هياكل عروض، لوحات
ألوان، خطوط، وبرومبتات:

```bash
python database/query.py recommend pitch     # هيكل + ثيمات مقترحة
python database/query.py search morph        # بحث نصي كامل
python database/query.py palettes            # ٢١ لوحة ألوان موثقة التباين
python database/build_db.py                  # إعادة بناء من data/*.json
python tools/db_lint.py                      # فاحص سلامة البيانات
```

## الأدوات

| الأداة | وظيفتها |
|---|---|
| `qa_deck.py` | بوابة الجودة: Morph، فيض النص، تباين WCAG، ميزانية الكلمات |
| `wizard.py` / `draft.py` / `coach.py` | وضع الصانع: معالج، ملاحظات→عرض، مدرب متحدث |
| `palette_lab.py` / `brandkit.py` / `anatomy.py` | ثيمات من لون أو شعار أو عرض قائم |
| `social.py` (+ `portrait.py`) | Reel وCarousel + تعليق وسيناريو تعليق صوتي |
| `decompile.py` | إنقاذ العروض القديمة: أي ‏`.pptx` ← مواصفة قابلة للتحرير |
| `merge.py` | دمج شرائح مختارة من عروض متعددة (الرسوم البيانية تنجو) |
| `imgkit.py` | معالجة الصور لتوافق الثيم (duotone/قص/تظليل) |
| `render.py` / `htmldeck.py` / `handout.py` | معاينة، عرض HTML قابل للتشغيل، مذكرة مطبوعة |
| `datafeed.py` | تحديث الرسوم من CSV — عرض ربع سنوي يتجدد تلقائياً |
| `diff.py` | سجل تغييرات بين نسختين من عرض |

## القوانين العشر (ملخص SKILL.md)

Morph أولاً · عناوين تدّعي ولا تسمّي · لون تمييز واحد في كل شريحة ·
فكرة واحدة لكل شريحة · النص الكامل في ملاحظات المتحدث لا على الشريحة ·
تباين موثّق (حبر ١٠:١، نص ثانوي ٤.٥:١) · أرقام كبيرة تُقال ببطء ·
لا شيء أصغر من ١٢ نقطة · كل عرض يمر ببوابة QA خضراء قبل التقديم ·
استعلم قاعدة البيانات قبل أن تخترع هيكلاً.

## الجودة والاختبارات

```bash
python -m unittest discover -s tests    # ٢٧+ اختباراً (نحو ثانية)
python tools/db_lint.py                 # يجب أن يطبع CLEAN
```

تعمل GitHub Actions على كل دفع: اختبارات + فحص بيانات + بوابة QA على
جميع العروض التجريبية. الموقع التفاعلي: `python tools/site.py` ثم
تصفّح `site/index.html` (إحصاءات حية، معرض ثيمات، عرض تجريبي يعمل في
المتصفح).

</div>
