"""Fill a Word template (edited by hand in Word) with every person's duty rows.

Usage: python3 fill_appointment_template.py template.docx rows.json out.docx

The template holds one letter: {{ชื่อ-สกุล}} somewhere in the text, and a table whose
rows after the header contain {{วัน}} {{เวลา}} {{ห้องสอบ}} {{กระบวนวิชา/ตอน}} {{หน้าที่}}.
The first placeholder row is the row style; all placeholder rows are replaced with the
person's rows. Everything else (text, fonts, spacing, alignment) is kept as in the template.
rows.json comes from: node build_appointment_docx.js data.json rows.json --rows
"""
import copy, json, re, sys, zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
FIELDS = {'{{วัน}}': 'day', '{{เวลา}}': 'time', '{{ห้องสอบ}}': 'rooms',
          '{{กระบวนวิชา/ตอน}}': 'courses', '{{หน้าที่}}': 'role'}

tpl, rows_path, out = sys.argv[1:4]
people = json.load(open(rows_path, encoding='utf-8'))
zin = zipfile.ZipFile(tpl)
root = etree.fromstring(zin.read('word/document.xml'))
body = root.find(f'{W}body')
sect = body.find(f'{W}sectPr')
block = [el for el in body if el is not sect]


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(f'{W}t'))


def replace_text(el, old, new):
    """Replace a placeholder that may be split over several w:t in one paragraph."""
    for p in el.iter(f'{W}p'):
        ts = list(p.iter(f'{W}t'))
        full = ''.join(t.text or '' for t in ts)
        if old not in full:
            continue
        start = full.index(old)
        pos = 0
        for t in ts:
            s, e = pos, pos + len(t.text or '')
            pos = e
            if e <= start or s >= start + len(old):
                continue
            a, b = max(start, s) - s, min(start + len(old), e) - s
            t.text = (t.text or '')[:a] + (new if s <= start else '') + (t.text or '')[b:]
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def unhighlight(el):
    for h in list(el.iter(f'{W}highlight')):
        h.getparent().remove(h)


SMALL_PT = 10   # size for lines marked "small" (e.g. the long AUD room name)


def set_size(r, pt):
    rpr = r.find(f'{W}rPr')
    if rpr is None:
        rpr = etree.Element(f'{W}rPr')
        r.insert(0, rpr)
    for tag in ('sz', 'szCs'):
        el = rpr.find(f'{W}{tag}')
        if el is None:
            el = etree.SubElement(rpr, f'{W}{tag}')
        el.set(f'{W}val', str(pt * 2))
    # keep schema order: sz, szCs must precede highlight/lang etc.
    for tag in ('highlight', 'u', 'effect', 'shd', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'):
        for el in rpr.findall(f'{W}{tag}'):
            rpr.remove(el); rpr.append(el)


def set_cell(tc, value):
    lines = value if isinstance(value, list) else [value]
    paras = tc.findall(f'{W}p')
    proto = paras[0]
    for p in paras[1:]:
        tc.remove(p)
    for i, line in enumerate(lines):
        p = proto if i == 0 else copy.deepcopy(proto)
        if i:
            proto.addnext(p) if i == 1 else tc.findall(f'{W}p')[-1].addnext(p)
        small = isinstance(line, dict) and line.get('small')
        ts = list(p.iter(f'{W}t'))
        if ts:
            ts[0].text = line['text'] if isinstance(line, dict) else str(line)
            for t in ts[1:]:
                t.text = ''
            if small:
                set_size(ts[0].getparent(), SMALL_PT)


def fill_table(tbl, rows):
    trs = tbl.findall(f'{W}tr')
    ph = [tr for tr in trs if any(k in text_of(tr) for k in FIELDS)]
    proto = ph[0]
    anchor = ph[0].getprevious()
    for tr in ph:
        tbl.remove(tr)
    for n, r in enumerate(rows, 1):
        tr = copy.deepcopy(proto)
        tcs = tr.findall(f'{W}tc')
        set_cell(tcs[0], str(n))
        for tc in tcs[1:]:
            key = next((FIELDS[k] for k in FIELDS if k in text_of(tc)), None)
            if key:
                set_cell(tc, r[key])
        anchor.addnext(tr)
        anchor = tr


def page_break_before(p):
    ppr = p.find(f'{W}pPr')
    if ppr is None:
        ppr = etree.SubElement(p, f'{W}pPr')
        p.insert(0, ppr)
    el = etree.Element(f'{W}pageBreakBefore')
    # schema order: pStyle, keepNext, keepLines, pageBreakBefore, ...
    after = [c for c in ppr if c.tag in (f'{W}pStyle', f'{W}keepNext', f'{W}keepLines')]
    (after[-1].addnext(el) if after else ppr.insert(0, el))


for el in block:
    body.remove(el)
for i, person in enumerate(people):
    letter = [copy.deepcopy(el) for el in block]
    for el in letter:
        for node in el.iter():   # paragraph ids must stay unique across copies
            for a in (f'{W14}paraId', f'{W14}textId'):
                node.attrib.pop(a, None)
        replace_text(el, '{{ชื่อ-สกุล}}', person['name'])
        if el.tag == f'{W}tbl':
            fill_table(el, person['rows'])
        unhighlight(el)
    if i:
        page_break_before(next(el for el in letter if el.tag == f'{W}p'))
    for el in letter:
        sect.addprevious(el)

left = re.findall(r'\{\{[^}]*\}\}', text_of(root))
assert not left, f'unfilled placeholders: {set(left)}'
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/document.xml':
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)
print(f'wrote {out}: {len(people)} letters')
