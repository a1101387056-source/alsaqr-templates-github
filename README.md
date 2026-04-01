# الصقر لأنظمة المحاسبة — قوالب الفاتورة والإيصال (رسمية)

هذه الحزمة تحتوي قوالب Word جاهزة للطباعة والاستخدام داخل نظام **الصقر لأنظمة المحاسبة**:

- **فاتورة A4 (رسمي)**: `templates/AlSaqr_Invoice_A4_Template_v3.docx`
- **إيصال POS 80mm (رسمي مبسط)**: `templates/AlSaqr_POS_Receipt_80mm_Template_v3.docx`
- **فاتورة A4 (أحادي اللون للطباعة الاقتصادية)**: `templates/AlSaqr_Invoice_A4_Template_MONO_v3.docx`
- **إيصال POS 80mm (أحادي اللون للطباعة الاقتصادية)**: `templates/AlSaqr_POS_Receipt_80mm_Template_MONO_v3.docx`

## الحقول الديناميكية (Placeholders)

تم وضع حقول بين أقواس مزدوجة لتسهيل ربطها ببيانات النظام أثناء الطباعة:

- `{{InvNo}}` رقم الفاتورة/الإيصال
- `{{Date}}` تاريخ الإصدار (A4)
- `{{DateTime}}` التاريخ والوقت (الإيصال)
- `{{StoreName}}` اسم الفرع/المخزن
- `{{CustomerName}}`, `{{CustomerID}}`, `{{CustomerVAT}}`, `{{CustomerMobile}}`
- `{{BaseTotal}}`, `{{TaxTotal}}`, `{{NetTotal}}`, `{{Discount}}`, `{{Paid}}`, `{{Balance}}`
- `{{PayType}}`, `{{RefNo}}`, `{{Notes}}`
- `{{CompanyVAT}}` الرقم الضريبي للمنشأة

### بيانات الرأس (مضافة حسب طلبكم)

- `{{CompanyAddress}}` عنوان المنشأة
- `{{CompanyPhone}}` رقم التواصل

## توليد القوالب

يمكن إعادة توليد القوالب من نسخ المصدر باستخدام:

```bash
pip install python-docx
python generate_templates.py
```

يقوم السكريبت بإضافة أسطر عنوان المنشأة ورقم التواصل تلقائياً وحفظ النسخ المحدثة في مجلد `templates/`.

## موضع QR

تم تخصيص موضع واضح لرمز **QR** داخل الفاتورة والإيصال، ويُستبدل في مرحلة الطباعة بقيمة QR الفعلية.

## ملاحظات

- نسخ **MONO** مناسبة للطباعة الاقتصادية (أسود فقط).
- يمكن تعديل التصميم بحرية مع الحفاظ على الصياغات الرسمية.