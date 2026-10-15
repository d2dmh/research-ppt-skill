"""Normalize editable text and repair confirmed layout defects in the bundled reference."""
import argparse
from copy import deepcopy
import json
from pathlib import Path
import zipfile
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'a': A, 'p': P, 'm': M}
ROOT = Path(__file__).resolve().parents[1]


def child(parent, tag, **attributes):
    node = parent.find(f'{{{A}}}{tag}')
    if node is None:
        node = etree.SubElement(parent, f'{{{A}}}{tag}')
    for key, value in attributes.items():
        node.set(key, str(value))
    return node


def normalize(dom):
    count = 0
    for props in dom.xpath('//a:rPr | //a:defRPr | //a:endParaRPr', namespaces=NS):
        # OMML equations retain their math font, including shape-level defaults.
        if any(n.tag.startswith('{' + M + '}') for n in props.iterancestors()):
            continue
        containing = next((n for n in props.iterancestors() if n.tag == f'{{{P}}}sp'), None)
        if containing is not None and containing.xpath('.//m:oMath', namespaces=NS):
            continue
        for tag, face in [('latin', 'Arial'), ('ea', '微软雅黑'), ('cs', 'Arial')]:
            node = child(props, tag)
            node.attrib.clear()
            node.set('typeface', face)
        count += 1
    for font in dom.xpath('//a:majorFont | //a:minorFont', namespaces=NS):
        child(font, 'latin', typeface='Arial')
        child(font, 'ea', typeface='微软雅黑')
        for n in font.findall(f'{{{A}}}font'):
            if n.get('script') == 'Hans':
                n.set('typeface', '微软雅黑')
    for paragraph in dom.xpath('//a:p', namespaces=NS):
        if any(n.tag.startswith('{' + M + '}') for n in paragraph.iterancestors()):
            continue
        props = paragraph.find(f'{{{A}}}pPr')
        if props is None:
            props = etree.Element(f'{{{A}}}pPr')
            paragraph.insert(0, props)
        props.set('latinLnBrk', '0')
        props.set('eaLnBrk', '1')
        props.set('hangingPunct', '1')
    return count


def find_shape(dom, identifier):
    for shape in dom.xpath('//p:sp', namespaces=NS):
        props = shape.find('p:nvSpPr/p:cNvPr', NS)
        if props is not None and props.get('id') == str(identifier):
            return shape
    raise ValueError(f'Reference shape {identifier} not found')


def style_shape(dom, identifier, size=None, bold=None, align=None, spacing=None, **box):
    shape = find_shape(dom, identifier)
    if size is not None:
        for node in shape.xpath('.//a:rPr | .//a:defRPr | .//a:endParaRPr', namespaces=NS):
            node.set('sz', str(round(size * 100)))
    if bold is not None:
        for node in shape.xpath('.//a:rPr | .//a:defRPr', namespaces=NS):
            node.set('b', '1' if bold else '0')
    for paragraph in shape.xpath('./p:txBody/a:p', namespaces=NS):
        props = child(paragraph, 'pPr')
        if align:
            props.set('algn', align)
        if spacing:
            line = child(props, 'lnSpc')
            for old in list(line):
                line.remove(old)
            child(line, 'spcPct', val=round(spacing * 100000))
    transform = shape.find('p:spPr/a:xfrm', NS)
    for key, value in box.items():
        node = transform.find('a:off' if key in ('x', 'y') else 'a:ext', NS)
        node.set(key, str(round(value * 12700)))
    return shape


def chevron(identifier, x, text):
    shape = etree.Element(f'{{{P}}}sp')
    nonvisual = etree.SubElement(shape, f'{{{P}}}nvSpPr')
    etree.SubElement(nonvisual, f'{{{P}}}cNvPr', id=str(identifier), name=f'Editable process step {identifier}')
    etree.SubElement(nonvisual, f'{{{P}}}cNvSpPr')
    etree.SubElement(nonvisual, f'{{{P}}}nvPr')
    properties = etree.SubElement(shape, f'{{{P}}}spPr')
    transform = child(properties, 'xfrm')
    child(transform, 'off', x=round(x * 12700), y=round(179 * 12700))
    child(transform, 'ext', cx=round(163 * 12700), cy=round(72 * 12700))
    geometry = child(properties, 'prstGeom', prst='chevron')
    adjustment = child(geometry, 'avLst')
    child(adjustment, 'gd', name='adj', fmla='val 18000')
    child(child(properties, 'solidFill'), 'srgbClr', val='4472C4')
    child(child(child(properties, 'ln', w=12700), 'solidFill'), 'srgbClr', val='FFFFFF')
    body = etree.SubElement(shape, f'{{{P}}}txBody')
    child(body, 'bodyPr', wrap='square', anchor='ctr', lIns=round(15*12700),
          rIns=round(15*12700), tIns=round(4*12700), bIns=round(4*12700))
    child(body, 'lstStyle')
    for line in text.split('\n'):
        paragraph = child(body, 'p') if not body.findall(f'{{{A}}}p') else etree.SubElement(body, f'{{{A}}}p')
        props = child(paragraph, 'pPr', algn='ctr', latinLnBrk='0', eaLnBrk='1')
        child(child(props, 'lnSpc'), 'spcPct', val=115000)
        run = child(paragraph, 'r')
        rprops = child(run, 'rPr', sz=1300, b='0', lang='zh-CN')
        child(child(rprops, 'solidFill'), 'srgbClr', val='FFFFFF')
        child(rprops, 'latin', typeface='Arial')
        child(rprops, 'ea', typeface='微软雅黑')
        child(run, 't').text = line
    return shape


def break_before(shape, phrase):
    for text in shape.xpath('./p:txBody/a:p/a:r/a:t', namespaces=NS):
        value = text.text or ''
        position = value.find(phrase)
        if position < 0:
            continue
        run, paragraph = text.getparent(), text.getparent().getparent()
        if position == 0:
            paragraph.insert(paragraph.index(run), etree.Element(f'{{{A}}}br'))
        else:
            tail = deepcopy(run)
            text.text = value[:position]
            tail.find(f'{{{A}}}t').text = value[position:]
            index = paragraph.index(run)
            paragraph.insert(index + 1, etree.Element(f'{{{A}}}br'))
            paragraph.insert(index + 2, tail)
        return
    raise ValueError('Expected reference phrase missing')


def repair_reference(number, dom):
    if number == 1:
        style_shape(dom, 318, size=32, x=54, cx=852, align='ctr')
        for identifier in (321, 326, 54):
            style_shape(dom, identifier, size=16, bold=False)
    if number == 3:
        style_shape(dom, 49, cx=570)
        style_shape(dom, 50, cx=570)
        style_shape(dom, 10, size=13, spacing=1.12)
        style_shape(dom, 9, size=12, y=496)
    if number == 4:
        style_shape(dom, 30, size=16, spacing=1.3)
        style_shape(dom, 68, size=16, bold=False, spacing=1.2)
    if number == 7:
        for identifier in (38, 49, 51):
            style_shape(dom, identifier, size=16, spacing=1.25)
        break_before(find_shape(dom, 38), '连锁故障')
    if number == 10:
        frames = dom.xpath('//p:graphicFrame[a:graphic/a:graphicData[@uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"]]', namespaces=NS)
        if len(frames) != 1:
            raise ValueError('Expected one SmartArt process on reference slide 10')
        frame = frames[0]
        parent, index = frame.getparent(), frame.getparent().index(frame)
        parent.remove(frame)
        labels = ['随机初始故障，\n设备丧失完成\n规定功能能力',
                  '能流平衡被打破，\n系统负荷再分配',
                  '健全设备过负荷，\n保护装置动作导致\n连锁故障']
        for i, label in enumerate(labels):
            parent.insert(index + i, chevron(9001 + i, 50 + i * 165, label))
    if number == 13:
        style_shape(dom, 13, size=11)
        style_shape(dom, 22, size=10.5)
        style_shape(dom, 12, y=319)
        style_shape(dom, 19, size=11, y=346)
        style_shape(dom, 21, size=11, y=345, cy=85)
    if number == 14:
        for identifier in (128, 129):
            style_shape(dom, identifier, size=11)
    if number == 16:
        original = style_shape(dom, 10, size=13)
        equations = [p for p in original.xpath('./p:txBody/a:p', namespaces=NS) if p.xpath('.//m:oMath', namespaces=NS)]
        if len(equations) != 6:
            raise ValueError('Expected six reference reliability equations')
        parent, index = original.getparent(), original.getparent().index(original)
        parent.remove(original)
        for i, equation in enumerate(equations):
            shape = deepcopy(original)
            props = shape.find('p:nvSpPr/p:cNvPr', NS)
            props.set('id', str(10 if i == 0 else 9010 + i))
            props.set('name', f'Reliability equation {i+1}')
            body = shape.find('p:txBody', NS)
            for p in body.findall(f'{{{A}}}p'):
                body.remove(p)
            body.append(deepcopy(equation))
            transform = shape.find('p:spPr/a:xfrm', NS)
            transform.find('a:off', NS).set('y', str(round([218, 262, 318, 362, 414, 458][i] * 12700)))
            transform.find('a:ext', NS).set('cy', str(round(52 * 12700)))
            parent.insert(index + i, shape)
    if number == 15:
        break_before(find_shape(dom, 6), '属于')
    if number == 18:
        cell = dom.xpath('//a:tbl/a:tr[3]/a:tc[4]', namespaces=NS)[0]
        for props in cell.xpath('.//a:pPr', namespaces=NS):
            props.set('algn', 'l')
        for props in cell.xpath('.//a:rPr | .//a:defRPr | .//a:endParaRPr', namespaces=NS):
            props.set('sz', '1200')
        admittance = cell.xpath('./a:txBody/a:p[3]', namespaces=NS)[0]
        for props in admittance.xpath('.//a:rPr | .//a:endParaRPr', namespaces=NS):
            props.set('sz', '1100')
    if number == 19:
        style_shape(dom, 23, size=10, cx=60, align='ctr')
    if number == 21:
        for identifier in (6, 15):
            style_shape(dom, identifier, size=14, spacing=1.15)
    if number == 23:
        style_shape(dom, 10, size=15, spacing=1.18)
        style_shape(dom, 33, size=15, spacing=1.2)
    if number == 24:
        style_shape(dom, 20, size=15, spacing=1.18)
        style_shape(dom, 33, size=15, spacing=1.2)
    if number == 25:
        style_shape(dom, 20, size=16, spacing=1.25)
    if number == 26:
        style_shape(dom, 33, size=15, spacing=1.2)
        style_shape(dom, 20, size=13, spacing=1.25)
    if number == 28:
        for identifier in (11, 14, 24):
            shape = style_shape(dom, identifier, size=16, spacing=1.15)
            for paragraph in shape.xpath('./p:txBody/a:p', namespaces=NS):
                runs = paragraph.findall(f'{{{A}}}r')
                for previous, run in zip(runs, runs[1:]):
                    before = previous.find(f'{{{A}}}rPr')
                    current = run.find(f'{{{A}}}rPr')
                    if before is not None and before.get('b') == '1' and (current is None or current.get('b') != '1'):
                        paragraph.insert(paragraph.index(run), etree.Element(f'{{{A}}}br'))
                        break
    if number == 29:
        for identifier in (40, 43):
            style_shape(dom, identifier, size=16, spacing=1.3)
        break_before(find_shape(dom, 43), '探索')
    if number == 30:
        style_shape(dom, 3, size=16, align='l', spacing=1.25)
    if number == 31:
        for identifier in (27, 31):
            style_shape(dom, identifier, size=13, bold=False)


def fix(source, output, report, reference_layout):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output:
        raise ValueError('Use a new output path')
    changed = []
    total = 0
    with zipfile.ZipFile(source) as archive:
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, 'w') as dest:
            for info in archive.infolist():
                payload = archive.read(info.filename)
                name = info.filename
                prefixes = ('ppt/slides/', 'ppt/slideMasters/', 'ppt/slideLayouts/', 'ppt/theme/', 'ppt/diagrams/')
                if name.endswith('.xml') and name.startswith(prefixes):
                    dom = etree.fromstring(payload)
                    total += normalize(dom)
                    filename = Path(name).name
                    if reference_layout and name.startswith('ppt/slides/') and filename.startswith('slide'):
                        number = int(filename[5:-4])
                        repair_reference(number, dom)
                    updated = etree.tostring(dom, xml_declaration=True, encoding='UTF-8', standalone=True)
                    if updated != payload:
                        changed.append(name)
                        payload = updated
                dest.writestr(info, payload)
    result = {'mode': 'font-and-confirmed-layout-repair' if reference_layout else 'editable-font-normalization',
              'cjk_font': '微软雅黑', 'latin_font': 'Arial', 'math_font_preserved': True,
              'normalized_property_nodes': total, 'changed_parts': changed,
              'media_bytes_unchanged': True,
              'limitation': 'Raster and vector image labels retain original fonts; not native editable text.'}
    Path(report).parent.mkdir(parents=True, exist_ok=True)
    Path(report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Normalized {total} properties; repaired reference layout: {reference_layout}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, default=ROOT / 'assets' / 'reference.pptx')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--reference-layout', action='store_true',
                        help='Apply confirmed fixes only to the bundled 31-slide reference or its same-structure copies')
    args = parser.parse_args()
    fix(args.input, args.output, args.report, args.reference_layout)
