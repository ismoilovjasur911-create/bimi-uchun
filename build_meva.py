# -*- coding: utf-8 -*-
"""
BMI (Bitiruv malakaviy ishi) docx quruvchi.
Mavzu: Kuniga 4 tonna quvvatga ega tuzlangan bodring ishlab chiqarish texnologiyasi.

Markdown-ga o'xshash 'bmi_content.md' faylini o'qib, formatlangan .docx yaratadi.

Belgilash (markup):
  # MATN        -> Bob / asosiy bo'lim sarlavhasi (yangi sahifadan, markazda, bold)
  ## MATN       -> Kichik bo'lim (2.1, 1.2 ...) bold, chapdan
  ### MATN      -> Ichki sarlavha (bold)
  | a | b | c | -> jadval qatori (birinchi qatordan keyin '|---|' bo'lsa sarlavha)
  [PB]          -> majburiy sahifa uzilishi
  [CENTER]matn  -> markazlashtirilgan oddiy abzats (masalan jadval nomi)
  bo'sh qator   -> abzaslar orasidagi ajratuvchi
  oddiy matn    -> justify abzas (birinchi qator chekinishi bilan)
"""
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Times New Roman"
SIZE = 14

def set_cell_font(cell, bold=False, size=12, align=None):
    for p in cell.paragraphs:
        if align is not None:
            p.alignment = align
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
        for r in p.runs:
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = bold
            r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

def add_page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
    run.font.name = FONT; run.font.size = Pt(12)

def style_doc(doc):
    style = doc.styles['Normal']
    style.font.name = FONT
    style.font.size = Pt(SIZE)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(3)
    sec.right_margin = Cm(1.5)

def add_para(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True, bold=False, size=SIZE,
             space_before=0, space_after=0, line=WD_LINE_SPACING.ONE_POINT_FIVE):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing_rule = line
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if indent:
        pf.first_line_indent = Cm(1.25)
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    return p

def add_heading1(doc, text):
    doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(12)
    run = p.add_run(text.upper())
    run.font.name = FONT; run.font.size = Pt(SIZE); run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

def add_heading2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(12); pf.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = FONT; run.font.size = Pt(SIZE); run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

def add_heading3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(6); pf.space_after = Pt(2)
    pf.first_line_indent = Cm(1.25)
    run = p.add_run(text)
    run.font.name = FONT; run.font.size = Pt(SIZE); run.font.bold = True; run.font.italic = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

def add_table(doc, rows):
    ncol = max(len(r) for r in rows)
    t = doc.add_table(rows=0, cols=ncol)
    t.style = 'Table Grid'
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j in range(ncol):
            val = row[j] if j < len(row) else ''
            cells[j].text = val
            set_cell_font(cells[j], bold=(i == 0), size=12,
                          align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT)
    # bo'shliq
    add_para(doc, '', indent=False, space_after=4)

def title_page(doc):
    def c(text, bold=False, size=14, before=0, after=0, caps=False, italic=False):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format; pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.space_before = Pt(before); pf.space_after = Pt(after)
        r = p.add_run(text.upper() if caps else text)
        r.font.name = FONT; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
        r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
        return p
    c("O‘ZBEKISTON RESPUBLIKASI", bold=True, size=14, after=0)
    c("OLIY TA’LIM, FAN VA INNOVATSIYALAR VAZIRLIGI", bold=True, size=14, after=10)
    c("_______________________________ UNIVERSITETI", bold=True, size=13, after=4)
    c("“OZIQ-OVQAT TEXNOLOGIYASI” KAFEDRASI", bold=True, size=13, after=40)
    c("Himoyaga ruxsat etildi", size=12, after=0)
    c("Kafedra mudiri ______________", size=12, after=0)
    c("“____” __________ 20___ yil", size=12, after=60)
    c("BITIRUV MALAKAVIY ISHI", bold=True, size=18, after=20)
    c("Mavzu: “Kunlik, quvvati 8 tonna bo‘lgan meva", bold=True, size=15, after=0)
    c("konservalarini ishlab chiqarish texnologiyasi”", bold=True, size=15, after=50)
    # bajardi/rahbar
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Bajardi: ___________________"); r.font.name = FONT; r.font.size = Pt(13)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Ilmiy rahbar: ___________________"); r.font.name = FONT; r.font.size = Pt(13)
    c("", after=80)
    c("Toshkent – 2026", bold=True, size=14, before=60)

def add_toc(doc):
    add_heading1(doc, "MUNDARIJA")
    p = doc.add_paragraph()
    run = p.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = "Mundarijani yangilash uchun: o‘ng tugma -> Update Field (F9)."
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar); run._r.append(instrText); run._r.append(fldChar2)
    run._r.append(t); run._r.append(fldChar3)

def enable_update_fields(doc):
    settings = doc.settings.element
    el = OxmlElement('w:updateFields'); el.set(qn('w:val'), 'true')
    settings.append(el)

def footer_page_numbers(doc):
    sec = doc.sections[0]
    footer = sec.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number_field(p)

def parse_and_build(doc, md_text):
    lines = md_text.split('\n')
    i = 0
    buffer = []  # paragraph text accumulation
    def flush():
        nonlocal buffer
        if buffer:
            text = ' '.join(x.strip() for x in buffer).strip()
            if text:
                add_para(doc, text)
            buffer = []
    while i < len(lines):
        line = lines[i].rstrip('\n')
        s = line.strip()
        if s == '':
            flush()
            i += 1
            continue
        if s == '[PB]':
            flush(); doc.add_page_break(); i += 1; continue
        if s.startswith('[CENTER]'):
            flush()
            add_para(doc, s[len('[CENTER]'):].strip(), align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, bold=True)
            i += 1; continue
        if s.startswith('### '):
            flush(); add_heading3(doc, s[4:].strip()); i += 1; continue
        if s.startswith('## '):
            flush(); add_heading2(doc, s[3:].strip()); i += 1; continue
        if s.startswith('# '):
            flush(); add_heading1(doc, s[2:].strip()); i += 1; continue
        if s.startswith('|'):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                row_line = lines[i].strip()
                cells = [c.strip() for c in row_line.strip('|').split('|')]
                if not re.match(r'^[\s\-:|]+$', row_line.strip('|')):
                    rows.append(cells)
                i += 1
            if rows:
                add_table(doc, rows)
            continue
        buffer.append(line)
        i += 1
    flush()

def main():
    doc = Document()
    style_doc(doc)
    title_page(doc)
    add_toc(doc)
    with open('meva_content.md', encoding='utf-8') as f:
        md = f.read()
    parse_and_build(doc, md)
    footer_page_numbers(doc)
    enable_update_fields(doc)
    out = "Meva_konservalari_BMI.docx"
    doc.save(out)
    # statistika
    words = sum(len(p.text.split()) for p in doc.paragraphs)
    print("Saqlandi:", out)
    print("Abzaslar:", len(doc.paragraphs), "Taxminiy so'zlar:", words)

if __name__ == '__main__':
    main()
