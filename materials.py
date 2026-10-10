"""Material prices are source data, independent of display codes and names."""
import re
from decimal import Decimal


def price_key(value):
    from matching import number
    try:
        price = number(value)
        if price < 0:
            return None
        # Decimal equality ignores insignificant zeros without rounding.
        return price
    except ValueError:
        return None


def price_text(value):
    price = price_key(value)
    if price is None:
        return 'не определена'
    fixed = format(price, 'f')
    whole, _, fraction = fixed.partition('.')
    whole = re.sub(r'(?<=\d)(?=(\d{3})+$)', ' ', whole)
    fraction = fraction.rstrip('0')
    return whole + (',' + fraction if fraction else '') + ' сом'


def material_price(row):
    from matching import position_type
    if position_type(row) != 'material':
        return ''
    return price_text(row.get('unit_price')) if price_key(row.get('unit_price')) is not None else ''


def unit_price_column(rows, first_row=None):
    """Require an explicit per-unit label; never guess from totals/numbers."""
    for item in rows[:100]:
        if first_row is not None and item['index'] >= first_row:
            break
        cells = item.get('header_cells', item['cells'])
        for col, value in enumerate(cells):
            text = re.sub(r'\s+', ' ', str(value).lower()).strip()
            if re.search(r'(?:цена\s*(?:за|на)\s*(?:единиц|ед\b|ед\.)|стоимость\s*(?:за|на)\s*(?:единиц|ед\b|ед\.)|unit[ _-]*price|price\s*per\s*unit)', text):
                if re.search(r'итог|общая|с учетом расхода|с учётом расхода', text):
                    continue
                # Merged headers repeat the same label. The left anchor owns
                # the value, as with the existing quantity import.
                return col
    return None
