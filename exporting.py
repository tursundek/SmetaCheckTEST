"""Fill the styled XLSX template without changing the source quantities.

The template owns formatting; native XLSX parts are filled at export time so
the application needs neither a renderer nor Excel installed on the computer.
"""
import io
import copy
import math
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from core import reconcile, number
from rik_codes import add_rik_codes
from materials import material_price
from sections import section_blocks, visible_section

HEADER_ROW = 1
DATA_ROW = 2
TEMPLATE_HEADER_ROW = 6
TEMPLATE_DATA_ROW = 7
TEMPLATE = Path(__file__).parent / 'templates' / 'reconciliation.xlsx'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('x', NS)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')


def tag(name):
    return '{'+NS+'}'+name


def text_cell(cell, value):
    for child in list(cell): cell.remove(child)
    cell.set('t', 'inlineStr')
    text = ET.SubElement(ET.SubElement(cell, tag('is')), tag('t'))
    text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    # Literal strings, including those starting with = / + / - / @, are never
    # Excel formulas. Strip only XML 1.0 control characters, which XLSX rejects.
    text.text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(value))


def quantity_cell(cell, raw, tiny_style, number_style):
    if raw is None: return text_cell(cell, '—')
    value = number(raw)
    significant = len(''.join(map(str, value.as_tuple().digits)).rstrip('0'))
    for child in list(cell): cell.remove(child)
    # Excel supports 15 significant digits. Preserve longer quantities as exact
    # text; ordinary quantities remain numeric for further Excel calculations.
    if significant > 15 or (value and abs(value.adjusted()) > 307):
        return text_cell(cell, format(value, 'f'))
    cell.set('t', 'n')
    ET.SubElement(cell, tag('v')).text = str(value)
    if value and (abs(value) < number('0.000000000000001') or value.adjusted() > 12):
        cell.set('s', tiny_style)
    else:
        # Only display significant fractional digits: 19, 2.2, 0.305 rather
        # than 19. or 2.200. This changes display, never the stored quantity.
        fixed = format(value, 'f')
        decimals = len(fixed.split('.')[1].rstrip('0')) if '.' in fixed else 0
        cell.set('s', number_style(cell.get('s'), '#,##0'+('.'+'0'*decimals if decimals else '')))


def row_height(values, visible_columns):
    widths = [12, 67, 23, 16, 16, 16, 48, 18, 23, 27]
    lines = max(sum(max(1, math.ceil(len(line)/(widths[i]-3)))
                    for line in str(values[i] or '').split('\n')) for i in visible_columns)
    return min(409, max(32, lines*15+12))


def sheet_xml(xml, key, rows, project_name, issue_count, number_style, composition_targets=None):
    root = ET.fromstring(xml)
    data = root.find(tag('sheetData'))
    sample = data.find(f"{tag('row')}[@r='{TEMPLATE_DATA_ROW}']")
    tiny = data.find(f"{tag('row')}[@r='{TEMPLATE_DATA_ROW+1}']")
    styles = {c.get('r')[0]: c.get('s', '0') for c in sample}
    tiny_styles = {c.get('r')[0]: c.get('s', '0') for c in tiny}
    styles['I'] = styles['H']
    styles['J'] = styles['C']
    styles['H'] = styles['D']
    header = copy.deepcopy(data.find(f"{tag('row')}[@r='{TEMPLATE_HEADER_ROW}']"))
    header.set('r',str(HEADER_ROW))
    for cell in header: cell.set('r',cell.get('r').rstrip('0123456789')+str(HEADER_ROW))
    text_cell(header.find(f"{tag('c')}[@r='A1']"),'№ позиции')
    text_cell(header.find(f"{tag('c')}[@r='H1']"),'Цена')
    description_header = copy.deepcopy(header.find(f"{tag('c')}[@r='H1']"))
    description_header.set('r', 'I1')
    text_cell(description_header, 'Описание')
    header.append(description_header)
    rik_header = copy.deepcopy(header.find(f"{tag('c')}[@r='B1']"))
    rik_header.set('r', 'J1')
    text_cell(rik_header, 'Шифр для РИК')
    header.append(rik_header)
    for row in list(data): data.remove(row)
    data.append(header)
    last = HEADER_ROW
    visible = [0, 1, 2, 3, 4, 7, 8, 9] if key == 'unchanged' else [0, 1, 2, 3, 4, 5, 7, 8, 9]
    rendered=[]
    for block in section_blocks(rows):
        section=block['section']
        if visible_section(section):
            rendered.append(([section['label'],section['title']]+['']*8,True,None))
        for subgroup in block['subheadings']:
            if subgroup['subheading']['title']:
                rendered.append((['',subgroup['subheading']['title']]+['']*8,True,None))
            for t in subgroup['rows']:
                values = [(t.get('display_number',t.get('number')) if key=='added' else t.get('number')) or '—', t['name'], t['unit'], t.get('quantity'),
                          t['actual'], t['volume'],
                          '\n'.join(t.get('tender_sources', [])+t['provenance']),
                          material_price(t),
                          'Нет в загруженных актах' if t['description'] == 'Работа исключена' else t['description'],
                          '\n'.join(t.get('rik_codes', []))]
                rendered.append((values,False,t['id']))
    header_styles={c.get('r').rstrip('0123456789'):c.get('s','0') for c in header}
    links=[]
    for index, (values, is_section, identity) in enumerate(rendered, DATA_ROW):
        row = ET.SubElement(data, tag('row'), r=str(index), ht=str(row_height(values, visible)), customHeight='1')
        for col, value in zip('ABCDEFGHIJ', values):
            cell = ET.SubElement(row, tag('c'), r=col+str(index), s=(header_styles if is_section else styles)[col])
            if is_section and value=='':continue
            if col in 'DEF' and not is_section: quantity_cell(cell, value, tiny_styles[col], number_style)
            else: text_cell(cell, value)
        last=index
        if not is_section and composition_targets and (key,identity) in composition_targets:
            links.append((f'A{index}',composition_targets[key,identity]))
    if not rows:
        row = ET.SubElement(data, tag('row'), r=str(DATA_ROW), ht='32', customHeight='1')
        cell = ET.SubElement(row, tag('c'), r='B7', s=styles['H'])
        cell.set('r','B'+str(DATA_ROW))
        text_cell(cell, 'В этой группе нет позиций.')
        last = DATA_ROW
    for col in root.find(tag('cols')):
        col_index = int(col.get('min'))
        if col_index == 7 or (key == 'unchanged' and col_index==6):
            col.set('hidden', '1'); col.set('outlineLevel', '1')
        if col_index == 8:
            col.attrib.pop('hidden', None); col.attrib.pop('collapsed', None)
            col.set('width', '18')
        if col_index in [4, 5, 6] and rows:
            field = {4:'quantity', 5:'actual', 6:'volume'}[col_index]
            lengths = []
            for t in rows:
                if t.get(field) is None: continue
                value = number(t[field])
                if value and (abs(value) < number('0.000000000000001') or value.adjusted() > 12): continue
                formatted = format(value, ',f')
                if '.' in formatted: formatted = formatted.rstrip('0').rstrip('.')
                lengths.append(len(formatted))
            col.set('width', str(min(25, max([16]+[n+2 for n in lengths]))))
    root.find(tag('sheetFormatPr')).set('outlineLevelCol', '1')
    ET.SubElement(root.find(tag('cols')), tag('col'), min='9', max='9', width='23', customWidth='1')
    ET.SubElement(root.find(tag('cols')), tag('col'), min='10', max='10', width='27', customWidth='1')
    props = root.find(tag('sheetPr'))
    ET.SubElement(props, tag('outlinePr'), summaryRight='1')
    ET.SubElement(props, tag('pageSetUpPr'), fitToPage='1')
    view = root.find(f"{tag('sheetViews')}/{tag('sheetView')}")
    view.set('zoomScale', '85')
    pane = view.find(tag('pane'))
    pane.attrib.clear()
    pane.attrib.update(ySplit='1',topLeftCell='A2',activePane='bottomLeft',state='frozen')
    for selection in list(view.findall(tag('selection'))): view.remove(selection)
    ET.SubElement(view, tag('selection'), pane='bottomLeft', activeCell='A2', sqref='A2')
    root.insert(1, ET.Element(tag('dimension'), ref=f'A1:J{last}'))
    # Element order follows the native worksheet OOXML schema.
    root.insert(list(root).index(data)+1, ET.Element(tag('autoFilter'), ref=f'A{HEADER_ROW}:J{last if rows else HEADER_ROW}'))
    root.find(tag('pageMargins')).attrib.update(left='0.25', right='0.25', top='0.4', bottom='0.4')
    if links:
        hyperlinks=ET.Element(tag('hyperlinks'))
        for ref,target in links:
            ET.SubElement(hyperlinks,tag('hyperlink'),ref=ref,location=f"'Состав позиций'!A{target}",tooltip='Открыть состав этой строки результата')
        root.insert(list(root).index(root.find(tag('pageMargins'))),hyperlinks)
    ET.SubElement(root, tag('pageSetup'), orientation='landscape', paperSize='8', fitToWidth='1', fitToHeight='0')
    footer=ET.SubElement(root,tag('headerFooter'))
    ET.SubElement(footer,tag('oddHeader')).text='&L'+project_name.replace('&','&&')+'&R&A'
    ET.SubElement(footer,tag('oddFooter')).text=('&LПредварительный результат. Замечаний: '+str(issue_count) if issue_count else '')+'&RСтраница &P из &N'
    return ET.tostring(root, encoding='utf-8', xml_declaration=True), last


def composition_sheet_xml(xml,result,number_style):
    """Source audit tables; never extra calculation rows in the three results."""
    sample=ET.fromstring(xml).find(tag('sheetData'))
    data_styles={c.get('r')[0]:c.get('s','0') for c in sample.find(f"{tag('row')}[@r='{TEMPLATE_DATA_ROW}']")}
    header_styles={c.get('r')[0]:c.get('s','0') for c in sample.find(f"{tag('row')}[@r='{TEMPLATE_HEADER_ROW}']")}
    root=ET.Element(tag('worksheet'));views=ET.SubElement(root,tag('sheetViews'))
    view=ET.SubElement(views,tag('sheetView'),workbookViewId='0',zoomScale='85')
    ET.SubElement(view,tag('pane'),ySplit='1',topLeftCell='A2',activePane='bottomLeft',state='frozen')
    ET.SubElement(root,tag('sheetFormatPr'),defaultRowHeight='20')
    widths=[12,62,23,19,19,44,29,22];cols=ET.SubElement(root,tag('cols'))
    for i,width in enumerate(widths,1):ET.SubElement(cols,tag('col'),min=str(i),max=str(i),width=str(width),customWidth='1')
    data=ET.SubElement(root,tag('sheetData'));targets={}
    def append(values,heading=False,numeric=()):
        index=len(data)+1
        lines=max([1]+[sum(max(1,math.ceil(len(line)/(widths[i]-3))) for line in str(v or '').split('\n')) for i,v in enumerate(values)])
        row=ET.SubElement(data,tag('row'),r=str(index),ht=str(min(409,max(26,lines*15+10))),customHeight='1')
        for i,value in enumerate(values):
            col='ABCDEFGH'[i];base=(header_styles if heading else data_styles).get(col,'0')
            cell=ET.SubElement(row,tag('c'),r=col+str(index),s=base)
            if i in numeric and value not in (None,''):quantity_cell(cell,value,base,number_style)
            elif value is not None:text_cell(cell,value)
        return index
    append(['Состав позиций','Номер позиции на листах результата — ссылка на соответствующий состав.'],True)
    for category,label in [('added','Плюс / включённые'),('minus','Минус / исключённые'),('unchanged','Без изменений')]:
        for block in section_blocks(result['groups'][category]):
            for subgroup in block['subheadings']:
                for row in subgroup['rows']:
                    append([])
                    targets[category,row['id']]=append([label,row['name'],row['unit'],
                        'Тендер: '+str(row.get('quantity') if row.get('quantity') is not None else '—'),
                        'Выполнено: '+row['actual']],True)
                    append(['','Только источники этой строки. Таблицы тендера и актов не устанавливают пары отдельных строк.'])
                    for source in row.get('source_composition',[]):
                        role='Тендер' if source['role']=='tender' else 'Процентовка'
                        append([role,source['document'],'',
                            'Учтено: '+str(source['accounted_quantity'] if source['accounted_quantity'] is not None else '—')+' '+source['unit'],
                            'Разница накопленных объёмов' if source['mode']=='cumulative' else ''],True)
                        if source['mode']=='cumulative':
                            append(['','Объёмы в таблице — накопленные значения из файла, а не прибавка к предыдущему акту.'])
                        if not source['contribution_confirmed']:
                            append(['','Эти строки найдены, но их объём не включён в результат. Проверьте замечания к сверке.'])
                        append(['№ позиции','Исходное наименование','Единица','Объём в Excel','Часть объёма в этой строке','Раздел / примечание','Где в Excel','Пересчёт единицы'],True)
                        for origin in source['origins']:
                            allocated=origin.get('allocated_quantity')
                            # Full current allocations need no duplicate quantity.
                            partial=source['contribution_confirmed'] and source['mode']!='cumulative' and allocated is not None and number(allocated)!=number(origin['quantity'])
                            parts=origin.get('allocation_parts',[dict(quantity=origin.get('allocated_quantity',origin['quantity']),factor=origin.get('factor','1'))])
                            conversion=source['contribution_confirmed'] and source['mode']!='cumulative' and any(number(p['factor'])!=1 for p in parts)
                            append([origin.get('number') or '—',origin.get('name',''),origin.get('unit',''),origin.get('quantity'),
                                    allocated if partial else '',
                                    '\n'.join(v for v in [origin.get('section'),origin.get('subheading')] if v),origin.get('source',''),
                                    '\n'.join(p['quantity']+' × '+p['factor'] for p in parts)+'\n= '+origin['accounted_quantity']+' '+source['unit'] if conversion else ''],numeric=(3,4))
    ET.SubElement(root,tag('pageMargins'),left='0.25',right='0.25',top='0.4',bottom='0.4',header='0.2',footer='0.2')
    ET.SubElement(root,tag('pageSetup'),orientation='landscape',paperSize='8',fitToWidth='1',fitToHeight='0')
    root.insert(0,ET.Element(tag('dimension'),ref=f'A1:H{len(data)}'))
    return ET.tostring(root,encoding='utf-8',xml_declaration=True),targets


def export(p):
    result = add_rik_codes(reconcile(p))
    output = io.BytesIO()
    with ZipFile(TEMPLATE) as template, ZipFile(output, 'w', ZIP_DEFLATED) as out:
        replacements = {}; ends = []
        styles = ET.fromstring(template.read('xl/styles.xml'))
        formats = styles.find(tag('numFmts'))
        if formats is None:
            formats = ET.Element(tag('numFmts')); styles.insert(0, formats)
        xfs = styles.find(tag('cellXfs'))
        format_ids = {f.get('formatCode'): f.get('numFmtId') for f in formats}
        styled_numbers = {}
        def number_style(base, code):
            cache_key = (base, code)
            if cache_key not in styled_numbers:
                if code not in format_ids:
                    num_id = str(max([163]+[int(f.get('numFmtId')) for f in formats])+1)
                    ET.SubElement(formats, tag('numFmt'), numFmtId=num_id, formatCode=code)
                    format_ids[code] = num_id
                xf = copy.deepcopy(xfs[int(base)])
                xf.set('numFmtId', format_ids[code]); xf.set('applyNumberFormat', '1')
                styled_numbers[cache_key] = str(len(xfs)); xfs.append(xf)
            return styled_numbers[cache_key]
        composition,targets=composition_sheet_xml(template.read('xl/worksheets/sheet1.xml'),result,number_style)
        for index, key in enumerate(['added', 'minus', 'unchanged'], 1):
            path = f'xl/worksheets/sheet{index}.xml'
            replacements[path], last = sheet_xml(template.read(path), key, result['groups'][key], p.get('name', 'Сверка сметы'), len(result['issues']), number_style,targets)
            ends.append(last)
        formats.set('count', str(len(formats))); xfs.set('count', str(len(xfs)))
        replacements['xl/styles.xml'] = ET.tostring(styles, encoding='utf-8', xml_declaration=True)
        workbook = ET.fromstring(template.read('xl/workbook.xml'))
        views = ET.Element(tag('bookViews'))
        # Open on the minus sheet: the user's primary question.
        ET.SubElement(views, tag('workbookView'), activeTab='1')
        workbook.insert(0, views)
        names = ET.SubElement(workbook, tag('definedNames'))
        for i, (sheet, last) in enumerate(zip(workbook.find(tag('sheets')), ends)):
            quoted = "'"+sheet.get('name').replace("'", "''")+"'!"
            ET.SubElement(names, tag('definedName'), name='_xlnm.Print_Titles', localSheetId=str(i)).text = quoted+'$1:$1'
            ET.SubElement(names, tag('definedName'), name='_xlnm.Print_Area', localSheetId=str(i)).text = quoted+f'$A$1:$J${last}'
        replacements['xl/workbook.xml'] = ET.tostring(workbook, encoding='utf-8', xml_declaration=True)
        rel_ns='http://schemas.openxmlformats.org/package/2006/relationships'
        rels=ET.fromstring(template.read('xl/_rels/workbook.xml.rels'))
        rel_id='rIdComposition'
        while any(r.get('Id')==rel_id for r in rels):rel_id+='Source'
        ET.SubElement(workbook.find(tag('sheets')),tag('sheet'),name='Состав позиций',sheetId='4',attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id':rel_id})
        ET.SubElement(rels,'{'+rel_ns+'}Relationship',Id=rel_id,Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',Target='worksheets/sheet4.xml')
        replacements['xl/workbook.xml']=ET.tostring(workbook,encoding='utf-8',xml_declaration=True)
        replacements['xl/_rels/workbook.xml.rels']=ET.tostring(rels,encoding='utf-8',xml_declaration=True)
        types=ET.fromstring(template.read('[Content_Types].xml'))
        ET.SubElement(types,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/xl/worksheets/sheet4.xml',ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        replacements['[Content_Types].xml']=ET.tostring(types,encoding='utf-8',xml_declaration=True)
        for item in template.infolist(): out.writestr(item, replacements.get(item.filename, template.read(item.filename)))
        out.writestr('xl/worksheets/sheet4.xml',composition)
    return output.getvalue()
