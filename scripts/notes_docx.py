"""Export slide speaker notes to a Word rehearsal script (one section per slide).

    python scripts/notes_docx.py --input deck.pptx --output 讲稿.docx [--rate 240]

rate = spoken Chinese characters per minute used for the time estimate.
"""
import argparse
import re
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pptx import Presentation

BLUE = RGBColor(0x1D, 0x50, 0xA2)


def set_font(run, size, bold=False, color=None):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def slide_title(slide):
    """页标题：info 风格按形状名取（页标题编号 + 页标题 / 封面题目 / 目录 / 结束页）；其他稿子取 y≈0.93 in 的粗体或最大字。"""
    named = {sh.name: sh.text_frame.text.strip().replace('\n', ' ') for sh in slide.shapes if sh.has_text_frame}
    if '页标题' in named:
        return (named.get('页标题编号', '') + ' ' + named['页标题']).strip()
    for key, fallback in (('封面题目', None), ('结束主句', None), ('目录项', '目录')):
        if key in named:
            return fallback or named[key]
    best, best_key = '', None
    for shape in slide.shapes:
        if not shape.has_text_frame or not shape.text_frame.text.strip():
            continue
        value = shape.text_frame.text.strip().replace('\n', ' ')
        top = shape.top / 914400
        sizes = [r.font.size.pt for p in shape.text_frame.paragraphs for r in p.runs if r.font.size]
        size = max(sizes, default=0)
        key = (0, abs(top - 0.93)) if 0.85 < top < 1.05 and size >= 17 else (1, -size)
        if best_key is None or key < best_key:
            best, best_key = value, key
    return best


def spoken_length(value):
    return len(re.sub(r'[\s，。、；：！？“”（）,.;:!?()\-]', '', value))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--rate', type=int, default=240)
    args = parser.parse_args()
    deck = Presentation(args.input)
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    style.font.size = Pt(12)
    entries = []
    for number, slide in enumerate(deck.slides, 1):
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ''
        entries.append((number, slide_title(slide), notes))
    total = sum(spoken_length(n) for _, _, n in entries)
    head = doc.add_paragraph()
    set_font(head.add_run(deck.core_properties.title or args.input.stem), 18, True, BLUE)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = doc.add_paragraph()
    set_font(info.add_run(f'讲稿 · 共 {len(entries)} 页 · 约 {total} 字 · 预计 {total / args.rate:.1f} 分钟（按每分钟 {args.rate} 字）'), 10.5)
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for number, title, notes in entries:
        heading = doc.add_paragraph()
        heading.paragraph_format.space_before = Pt(14)
        set_font(heading.add_run(f'第 {number} 页  {title}'), 13, True, BLUE)
        if not notes:
            set_font(doc.add_paragraph().add_run('（本页无讲稿）'), 11, color=RGBColor(0x7F, 0x7F, 0x7F))
            continue
        for line in [l for l in notes.split('\n') if l.strip()]:
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.space_after = Pt(4)
            set_font(paragraph.add_run(line.strip()), 12)
        seconds = round(spoken_length(notes) / args.rate * 60)
        meta = doc.add_paragraph()
        set_font(meta.add_run(f'约 {seconds} 秒'), 9.5, color=RGBColor(0x7F, 0x7F, 0x7F))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.output)
    print(f'Wrote rehearsal script for {len(entries)} slides (~{total / args.rate:.1f} min) -> {args.output}')


if __name__ == '__main__':
    main()
