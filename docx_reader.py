"""Extract Word table grids without counting merged cells twice."""
import io
import re
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def text_of(element):
    parts = []
    for paragraph in element.iter(W+'p'):
        text = ''.join(node.text or '' if node.tag == W+'t' else '\t' if node.tag == W+'tab' else '\n'
                       for node in paragraph.iter() if node.tag in [W+'t', W+'tab', W+'br'])
        if text.strip(): parts.append(text.strip())
    return '\n'.join(parts)


def read_tables(content):
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            info = archive.getinfo('word/document.xml')
            if info.file_size > 30_000_000: raise ValueError('Таблицы Word слишком большие')
            xml = archive.read(info)
        if b'<!DOCTYPE' in xml or b'<!ENTITY' in xml: raise ValueError('Неподдерживаемая структура Word')
        root = ET.fromstring(xml)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as error:
        raise ValueError('Не удалось прочитать Word. Сохраните файл в формате .docx') from error
    body = root.find(W+'body')
    if body is None: raise ValueError('В Word отсутствует основной текст')
    result = {}; section = ''; total_rows = 0
    # Descending into containers also supports tables in content controls.
    def blocks(parent):
        for child in parent:
            if child.tag in [W+'p', W+'tbl']: yield child
            else: yield from blocks(child)
    for block in blocks(body):
        if block.tag == W+'p':
            value = text_of(block)
            if re.match(r'^раздел\s+\d+', value, re.I): section = value
            continue
        # Wrapper tables are layout containers; their inner tables have the data.
        nested = list(block.iter(W+'tbl'))
        tables = [table for table in nested if len(list(table.iter(W+'tbl'))) == 1]
        for table in tables:
            rows = []; merged = {}
            for index, row in enumerate(table.findall(W+'tr'), 1):
                before = row.find(W+'trPr/'+W+'gridBefore')
                start = int(before.get(W+'val', '0')) if before is not None else 0
                if start > 150: raise ValueError('В таблице Word больше 150 колонок')
                cells = ['']*start; labels = ['']*start; continued = []
                for cell in row.findall(W+'tc'):
                    props = cell.find(W+'tcPr')
                    span_node = props.find(W+'gridSpan') if props is not None else None
                    span = int(span_node.get(W+'val', '1')) if span_node is not None else 1
                    if span < 1 or len(cells)+span > 150: raise ValueError('В таблице Word больше 150 колонок')
                    vertical = props.find(W+'vMerge') if props is not None else None
                    value = text_of(cell); col = len(cells)
                    if vertical is not None and vertical.get(W+'val') != 'restart':
                        value = merged.get(col, value)
                        continued.extend(range(col, col+span))
                    for offset in range(span):
                        if vertical is not None: merged[col+offset] = value
                        else: merged.pop(col+offset, None)
                    cells.extend([value]+['']*(span-1))
                    # A parent header applies to each of its spanned columns.
                    labels.extend([value]*span)
                rows.append(dict(index=index, cells=cells, header_cells=labels, continued_columns=continued,
                                 hidden=False, formulas=[], section_hint=section))
                total_rows += 1
                if total_rows > 50000: raise ValueError('В Word больше 50 000 строк')
            if rows: result[f'Таблица {len(result)+1}'] = rows
    if not result: raise ValueError('В Word нет таблиц. Нужны таблицы с работами, единицами и объёмами')
    return result, []
