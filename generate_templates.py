"""
generate_templates.py
=====================
Generates v2 AlSaqr Word templates by inserting {{CompanyAddress}} and
{{CompanyPhone}} placeholders into the existing v1 source templates, and
produces monochrome (MONO) variants with all text forced to black.

Usage
-----
    python generate_templates.py

Expected source files (place next to this script or update SRC_* paths):
    AlSaqr_Invoice_A4_Template.docx
    AlSaqr_POS_Receipt_80mm_Template.docx
    AlSaqr_Invoice_A4_Template_MONO.docx
    AlSaqr_POS_Receipt_80mm_Template_MONO.docx

Output
------
    templates/AlSaqr_Invoice_A4_Template_v2.docx
    templates/AlSaqr_POS_Receipt_80mm_Template_v2.docx
    templates/AlSaqr_Invoice_A4_Template_MONO_v2.docx
    templates/AlSaqr_POS_Receipt_80mm_Template_MONO_v2.docx
"""

import os
import shutil
import textwrap
import zipfile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_paragraph_rtl(paragraph):
    """Add w:bidi element to a paragraph so it renders right-to-left."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    # avoid duplicates
    for el in pPr.findall(qn('w:bidi')):
        pPr.remove(el)
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)


def set_run_ar(run, size=11, bold=False, color=None):
    """Apply Arabic font settings to a run."""
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Noto Sans Arabic'
    if color is not None:
        run.font.color.rgb = color
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Noto Sans Arabic')
    rFonts.set(qn('w:hAnsi'), 'Noto Sans Arabic')
    rFonts.set(qn('w:cs'), 'Noto Sans Arabic')
    rPr.append(rFonts)


# ---------------------------------------------------------------------------
# Template modification functions
# ---------------------------------------------------------------------------

def insert_company_contact_lines_invoice(doc: Document):
    """Insert {{CompanyAddress}} and {{CompanyPhone}} lines after the
    'فاتورة ضريبية' subtitle paragraph in an A4 invoice template."""
    idx = None
    for i, p in enumerate(doc.paragraphs):
        if 'فاتورة ضريبية' in p.text:
            idx = i
            break
    if idx is None:
        idx = 0

    p_addr = doc.paragraphs[idx].insert_paragraph_after('')
    p_addr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_rtl(p_addr)
    r = p_addr.add_run('عنوان المنشأة: {{CompanyAddress}}')
    set_run_ar(r, size=10)

    p_phone = p_addr.insert_paragraph_after('')
    p_phone.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_rtl(p_phone)
    r2 = p_phone.add_run('رقم التواصل: {{CompanyPhone}}')
    set_run_ar(r2, size=10)


def insert_company_contact_lines_receipt(doc: Document):
    """Insert {{CompanyAddress}} and {{CompanyPhone}} lines after the
    'الرقم الضريبي' line in a POS receipt template."""
    insert_after_idx = None
    for i, p in enumerate(doc.paragraphs):
        if 'الرقم الضريبي' in p.text:
            insert_after_idx = i
            break
    if insert_after_idx is None:
        insert_after_idx = 0

    p_addr = doc.paragraphs[insert_after_idx].insert_paragraph_after('')
    p_addr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_rtl(p_addr)
    r = p_addr.add_run('عنوان المنشأة: {{CompanyAddress}}')
    set_run_ar(r, size=9)

    p_phone = p_addr.insert_paragraph_after('')
    p_phone.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_rtl(p_phone)
    r2 = p_phone.add_run('رقم التواصل: {{CompanyPhone}}')
    set_run_ar(r2, size=9)


def force_black(doc: Document):
    """Set all text colour to black (suitable for monochrome printing)."""
    black = RGBColor(0, 0, 0)
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.color.rgb = black
            run.font.highlight_color = None
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.color.rgb = black
                        run.font.highlight_color = None


# ---------------------------------------------------------------------------
# Main generation logic
# ---------------------------------------------------------------------------

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base_dir, 'templates')
    os.makedirs(out_dir, exist_ok=True)

    # Paths to v1 source templates (expected alongside this script)
    src_invoice = os.path.join(base_dir, 'AlSaqr_Invoice_A4_Template.docx')
    src_receipt = os.path.join(base_dir, 'AlSaqr_POS_Receipt_80mm_Template.docx')
    src_invoice_mono = os.path.join(base_dir, 'AlSaqr_Invoice_A4_Template_MONO.docx')
    src_receipt_mono = os.path.join(base_dir, 'AlSaqr_POS_Receipt_80mm_Template_MONO.docx')

    out_invoice_v2 = os.path.join(out_dir, 'AlSaqr_Invoice_A4_Template_v2.docx')
    out_receipt_v2 = os.path.join(out_dir, 'AlSaqr_POS_Receipt_80mm_Template_v2.docx')
    out_invoice_mono_v2 = os.path.join(out_dir, 'AlSaqr_Invoice_A4_Template_MONO_v2.docx')
    out_receipt_mono_v2 = os.path.join(out_dir, 'AlSaqr_POS_Receipt_80mm_Template_MONO_v2.docx')

    print('Generating colour invoice v2 …')
    inv = Document(src_invoice)
    insert_company_contact_lines_invoice(inv)
    inv.save(out_invoice_v2)

    print('Generating colour receipt v2 …')
    rec = Document(src_receipt)
    insert_company_contact_lines_receipt(rec)
    rec.save(out_receipt_v2)

    print('Generating MONO invoice v2 …')
    inv_m = Document(src_invoice_mono)
    insert_company_contact_lines_invoice(inv_m)
    force_black(inv_m)
    inv_m.save(out_invoice_mono_v2)

    print('Generating MONO receipt v2 …')
    rec_m = Document(src_receipt_mono)
    insert_company_contact_lines_receipt(rec_m)
    force_black(rec_m)
    rec_m.save(out_receipt_mono_v2)

    # ------------------------------------------------------------------
    # Create a GitHub-ready zip bundle
    # ------------------------------------------------------------------
    zip_path = os.path.join(base_dir, 'alsaqr-templates-github.zip')
    print(f'Creating zip bundle → {zip_path} …')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        readme_content = textwrap.dedent('''
            # الصقر لأنظمة المحاسبة — قوالب الفاتورة والإيصال (رسمية)

            هذه الحزمة تحتوي قوالب Word جاهزة للطباعة والاستخدام داخل نظام **الصقر لأنظمة المحاسبة**:

            - **فاتورة A4 (رسمي)**: `AlSaqr_Invoice_A4_Template_v2.docx`
            - **إيصال POS 80mm (رسمي مبسط)**: `AlSaqr_POS_Receipt_80mm_Template_v2.docx`
            - **فاتورة A4 (أحادي اللون للطباعة الاقتصادية)**: `AlSaqr_Invoice_A4_Template_MONO_v2.docx`
            - **إيصال POS 80mm (أحادي اللون للطباعة الاقتصادية)**: `AlSaqr_POS_Receipt_80mm_Template_MONO_v2.docx`

            ## الحقول الديناميكية (Placeholders)
            تم وضع حقول بين أقواس مزدوجة لتسهيل ربطها ببيانات النظام أثناء الطباعة:

            - `{{InvNo}}` رقم الفاتورة/الإيصال
            - `{{Date}}` تاريخ الإصدار (A4)
            - `{{DateTime}}` التاريخ والوقت (الإيصال)
            - `{{StoreName}}` اسم الفرع/المخزن
            - `{{CustomerName}}`, `{{CustomerID}}`, `{{CustomerVAT}}`, `{{CustomerMobile}}`
            - `{{BaseTotal}}`, `{{TaxTotal}}`, `{{NetTotal}}`, `{{Discount}}`, `{{Paid}}`, `{{Balance}}`
            - `{{PayType}}`, `{{RefNo}}`, `{{Notes}}`

            ### بيانات الرأس (مضافة حسب طلبكم)
            - `{{CompanyAddress}}` عنوان المنشأة
            - `{{CompanyPhone}}` رقم التواصل

            ## موضع QR
            تم تخصيص موضع واضح لرمز **QR** داخل الفاتورة والإيصال، ويُستبدل في مرحلة الطباعة بقيمة QR الفعلية.

            ## ملاحظات
            - نسخ **MONO** مناسبة للطباعة الاقتصادية (أسود فقط).
            - يمكن تعديل التصميم بحرية مع الحفاظ على الصياغات الرسمية.
        ''').strip() + '\n'
        z.writestr('alsaqr-templates/README.md', readme_content)
        z.writestr('alsaqr-templates/.gitignore', '*.tmp\n~$*.docx\n.DS_Store\n')
        for fname in [
            'AlSaqr_Invoice_A4_Template_v2.docx',
            'AlSaqr_POS_Receipt_80mm_Template_v2.docx',
            'AlSaqr_Invoice_A4_Template_MONO_v2.docx',
            'AlSaqr_POS_Receipt_80mm_Template_MONO_v2.docx',
        ]:
            fpath = os.path.join(out_dir, fname)
            if os.path.exists(fpath):
                z.write(fpath, arcname=f'alsaqr-templates/templates/{fname}')

    print('Done.')


if __name__ == '__main__':
    main()
