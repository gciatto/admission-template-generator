"""Add native Excel protection and validation unsupported by the authoring API."""
import copy
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
PATH = ROOT / 'applicant-template.xlsx'
META = json.loads((ROOT / 'metadata.json').read_text())
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('', NS)
ET.register_namespace('r', REL)
Q = lambda x: f'{{{NS}}}{x}'


def colnum(s):
    n = 0
    for c in s:
        n = n * 26 + ord(c) - 64
    return n


def within(cell, region):
    a, _, b = region.partition(':')
    b = b or a
    parts = [re.fullmatch(r'([A-Z]+)(\d+)', x).groups() for x in (cell, a, b)]
    (c, r), (ca, ra), (cb, rb) = parts
    return colnum(ca) <= colnum(c) <= colnum(cb) and int(ra) <= int(r) <= int(rb)


def insert_before(root, el, successors):
    for i, old in enumerate(root):
        if old.tag.split('}')[-1] in successors:
            root.insert(i, el)
            return
    root.append(el)


with zipfile.ZipFile(PATH) as z:
    files = {n: z.read(n) for n in z.namelist()}
styles = ET.fromstring(files['xl/styles.xml'])
xfs = styles.find(Q('cellXfs'))
style_cache = {}


def unlocked_style(s):
    if s not in style_cache:
        xf = copy.deepcopy(xfs[int(s)])
        protection = xf.find(Q('protection'))
        if protection is not None:
            xf.remove(protection)
        ET.SubElement(xf, Q('protection'), locked='0', hidden='0')
        xf.set('applyProtection', '1')
        style_cache[s] = str(len(xfs))
        xfs.append(xf)
    return style_cache[s]


book = ET.fromstring(files['xl/workbook.xml'])
names = book.find(Q('definedNames'))
if names is None:
    names = ET.Element(Q('definedNames'))
    insert_before(book, names, {'calcPr','extLst'})
for item in META['definedNames']:
    for old in list(names):
        if old.get('name') == item['name']:
            names.remove(old)
    ET.SubElement(names, Q('definedName'), name=item['name']).text = item['ref']
calc = book.find(Q('calcPr'))
if calc is None:
    calc = ET.SubElement(book, Q('calcPr'))
calc.attrib.update(calcMode='auto', fullCalcOnLoad='1', forceFullCalc='1')
protection = ET.Element(Q('workbookProtection'), lockStructure='1')
insert_before(book, protection, {'bookViews','sheets','definedNames','calcPr','extLst'})
files['xl/workbook.xml'] = ET.tostring(book, encoding='utf-8', xml_declaration=True)

# Workbook order is explicit in the builder. All unspecified cells stay locked.
sheet_names = [s.get('name') for s in book.find(Q('sheets'))]
for sheet in book.find(Q('sheets')):
    if sheet.get('name') == 'Lists':
        sheet.set('state', 'hidden')
files['xl/workbook.xml'] = ET.tostring(book, encoding='utf-8', xml_declaration=True)
for index, name in enumerate(sheet_names, 1):
    key = f'xl/worksheets/sheet{index}.xml'
    root = ET.fromstring(files[key])
    editable = [x['range'] for x in META['unlocked'] if x['sheet'] == name]
    inherited = [x['range'] for x in META['inheritedDefaults'] if x['sheet'] == name]
    for cell in root.iter(Q('c')):
        if any(within(cell.get('r'), r) for r in editable):
            assert cell.find(Q('f')) is None or any(within(cell.get('r'), r) for r in inherited), (name, cell.get('r'), 'Only inherited input formulas may be editable')
            cell.set('s', unlocked_style(cell.get('s', '0')))
    # Add missing blank editable cells when the exporter omitted them.
    data = root.find(Q('sheetData'))
    rows = {int(r.get('r')): r for r in data}
    existing = {c.get('r') for c in root.iter(Q('c'))}
    for region in editable:
        a, _, b = region.partition(':')
        b = b or a
        ca, ra = re.fullmatch(r'([A-Z]+)(\d+)', a).groups()
        cb, rb = re.fullmatch(r'([A-Z]+)(\d+)', b).groups()
        assert ca == cb, region
        for r in range(int(ra), int(rb) + 1):
            address = ca + str(r)
            if address not in existing:
                if r not in rows:
                    rows[r] = ET.SubElement(data, Q('row'), r=str(r))
                ET.SubElement(rows[r], Q('c'), r=address, s=unlocked_style('0'))
    for row in data:
        row[:] = sorted(row, key=lambda c: colnum(re.match('[A-Z]+', c.get('r')).group()))
    data[:] = sorted(data, key=lambda r: int(r.get('r')))
    for old in root.findall(Q('sheetProtection')):
        root.remove(old)
    prot = ET.Element(Q('sheetProtection'), sheet='1', objects='1', scenarios='1',
                      formatCells='1', formatColumns='1', formatRows='0',
                      insertColumns='1', insertRows='1', deleteColumns='1',
                      deleteRows='1', sort='1', autoFilter='0',
                      selectLockedCells='0', selectUnlockedCells='0')
    insert_before(root, prot, {'protectedRanges','scenarios','autoFilter','sortState',
                              'dataConsolidate','customSheetViews','mergeCells',
                              'phoneticPr','conditionalFormatting','dataValidations',
                              'hyperlinks','printOptions','pageMargins','drawing','tableParts','extLst'})
    views = root.find(Q('sheetViews'))
    view = views.find(Q('sheetView'))
    view.set('showGridLines', '0')
    view.set('zoomScale', '85')
    view.set('zoomScaleNormal', '85')
    freeze = META['sheets'][name]['freeze']
    for old in list(view):
        if old.tag in (Q('pane'), Q('selection')):
            view.remove(old)
    if freeze:
        x = 0 if name in ('Personal & contacts','Degree class legend','Exam sector legend') else 2
        attrs = dict(ySplit=str(freeze), topLeftCell=f'{"C" if x else "A"}{freeze+1}',
                     activePane='bottomRight' if x else 'bottomLeft', state='frozen')
        if x:
            attrs['xSplit'] = str(x)
        ET.SubElement(view, Q('pane'), **attrs)
        ET.SubElement(view, Q('selection'), pane=attrs['activePane'],
                      activeCell=attrs['topLeftCell'], sqref=attrs['topLeftCell'])
    else:
        ET.SubElement(view, Q('selection'), activeCell='A1', sqref='A1')
    for old in root.findall(Q('dataValidations')):
        root.remove(old)
    entries = [v for v in META['validations'] if v['sheet'] == name]
    if entries:
        dvs = ET.Element(Q('dataValidations'), count=str(len(entries)))
        for entry in entries:
            rule = entry['rule']
            attrs = dict(type=rule['type'], sqref=entry['range'], allowBlank='1',
                         showInputMessage='1', showErrorMessage='1', errorStyle='stop',
                         errorTitle='Check your entry', error='Use the listed values or the expected data type.',
                         promptTitle='How to fill this cell', prompt=(entry.get('description') or 'Choose a listed value or use the specified format.')[:255])
            if 'operator' in rule:
                attrs['operator'] = rule['operator']
            dv = ET.SubElement(dvs, Q('dataValidation'), **attrs)
            for f in ('formula1','formula2'):
                if f in rule:
                    ET.SubElement(dv, Q(f)).text = str(rule[f])
        insert_before(root, dvs, {'hyperlinks','printOptions','pageMargins','pageSetup','headerFooter',
                                 'rowBreaks','colBreaks','drawing','legacyDrawing','tableParts','extLst'})
    files[key] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

for key, value in list(files.items()):
    if key.startswith('xl/tables/') and key.endswith('.xml'):
        root = ET.fromstring(value)
        style = root.find(Q('tableStyleInfo'))
        if style is not None:
            style.set('showRowStripes', '0')
            style.set('showColumnStripes', '0')
        files[key] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
xfs.set('count', str(len(xfs)))
files['xl/styles.xml'] = ET.tostring(styles, encoding='utf-8', xml_declaration=True)
temp = PATH.with_suffix('.tmp.xlsx')
with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as z:
    for key, value in files.items():
        z.writestr(key, value)
temp.replace(PATH)

# Structural assertions on the delivered archive, including edge input rows.
with zipfile.ZipFile(PATH) as z:
    final_styles = ET.fromstring(z.read('xl/styles.xml')).find(Q('cellXfs'))
    total_unlocked = total_validations = 0
    for index, name in enumerate(sheet_names, 1):
        root = ET.fromstring(z.read(f'xl/worksheets/sheet{index}.xml'))
        assert root.find(Q('sheetProtection')) is not None
        regions = [x['range'] for x in META['unlocked'] if x['sheet'] == name]
        inherited = [x['range'] for x in META['inheritedDefaults'] if x['sheet'] == name]
        for cell in root.iter(Q('c')):
            expected = any(within(cell.get('r'), r) for r in regions)
            p = final_styles[int(cell.get('s','0'))].find(Q('protection'))
            actual = p is not None and p.get('locked') == '0'
            assert actual == expected, (name, cell.get('r'), 'Wrong lock state')
            if cell.find(Q('f')) is not None:
                assert not actual or any(within(cell.get('r'), r) for r in inherited)
                assert not re.search(r'#REF!|#DIV/0!|#VALUE!|#NAME\?|#N/A|#NUM!', cell.findtext(Q('v')) or '')
            total_unlocked += actual
        total_validations += len(root.findall(f'{Q("dataValidations")}/{Q("dataValidation")}'))
    assert total_validations == len(META['validations'])
    assert len(sheet_names) == 12
print(f'Validated {len(sheet_names)} protected sheets, {total_unlocked} editable cells and {total_validations} validation ranges.')
print(f'Deliverable: {PATH} ({PATH.stat().st_size:,} bytes)')
