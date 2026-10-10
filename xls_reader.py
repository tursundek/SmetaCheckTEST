"""Legacy Excel reader; return the same raw sheet records as the XLSX reader."""
from decimal import Decimal
import io
import re

XLS_VALUE_WARNING = 'XLS: программа читает готовые числа из файла и не пересчитывает формулы. Если меняли цифры или формулы, пересчитайте файл в Excel, сохраните его и загрузите снова.'

def read_xls(content):
    try:
        import xlrd
    except ImportError as exc:
        raise ValueError('Для XLS установите зависимости из requirements.txt и перезапустите приложение.') from exc
    kind = xlrd.inspect_format(content=content)
    if kind == 'xlsx':
        raise ValueError('Внутри файла XLS находится XLSX. Сохраните его с правильным расширением .xlsx.')
    if kind != 'xls':
        raise ValueError('Файл не является Excel XLS. Откройте его в Excel и сохраните как .xls или .xlsx.')
    try:
        book = xlrd.open_workbook(file_contents=content, formatting_info=True, logfile=io.StringIO())
    except (xlrd.XLRDError, ValueError, IndexError, OSError) as exc:
        raise ValueError('Не удалось прочитать XLS: файл повреждён, защищён паролем или имеет неподдерживаемую структуру. Сохраните его в Excel как .xlsx.') from exc
    try:
        result = {}
        for sheet in book.sheets():
            height, width = sheet.nrows, sheet.ncols
            if height > 50000 or width > 150:
                raise ValueError('Таблица превышает лимит 50 000 строк / 150 колонок')
            merged_labels, merged_spans, continued = {}, {}, {}
            for row1, row2, col1, col2 in sheet.merged_cells:
                if row2 == row1 + 1:
                    merged_spans.setdefault(row1 + 1, []).append((col1, col2 - 1))
                # Keep actual values at the anchor; never duplicate a volume.
                for row in range(row1 + 1, row2):
                    continued.setdefault(row + 1, []).extend(range(col1, col2))
                if row1 >= 100:
                    continue
                label = cell_text(book, sheet.cell(row1, col1), xlrd)
                for row in range(row1, min(row2, 100)):
                    for col in range(col1, col2):
                        merged_labels[(row, col)] = label
            rows = []
            for row in range(height):
                cells = sheet.row(row)
                values = [cell_text(book, cell, xlrd) for cell in cells]
                errors = {str(col + 1): 'Ошибка ячейки Excel: ' + values[col]
                          for col, cell in enumerate(cells) if cell.ctype == xlrd.XL_CELL_ERROR}
                info = sheet.rowinfo_map.get(row)
                rows.append(dict(index=row + 1, cells=values,
                                 header_cells=[merged_labels.get((row, col), value) for col, value in enumerate(values)],
                                 merged_spans=merged_spans.get(row + 1, []),
                                 continued_columns=continued.get(row + 1, []), hidden=bool(info and info.hidden),
                                 formulas=[], computed=[], formula_errors={}, formula_sources={}, cell_errors=errors))
            result[sheet.name] = rows
        return result, [XLS_VALUE_WARNING]
    finally:
        book.release_resources()


def cell_text(book, cell, xlrd):
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return ''
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(cell.value, '#ERROR!')
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            value = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
            return str(value)
        except (ValueError, OverflowError):
            return '#INVALID_DATE!'
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return str(bool(cell.value))
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        value = Decimal(str(cell.value))
        if not value.is_finite():
            return '#INVALID_NUMBER!'
        if value == value.to_integral_value():
            text = str(int(value))
            # Preserve a numeric position/code explicitly formatted as 000...
            xf = book.xf_list[cell.xf_index]
            fmt = book.format_map.get(xf.format_key)
            if fmt and re.fullmatch(r'0+', fmt.format_str) and value >= 0:
                return text.zfill(len(fmt.format_str))
            return text
        return format(value, 'f')
    return str(cell.value)
