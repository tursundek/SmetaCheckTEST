"""Presentation-only RIK codes; never used as reconciliation identities."""
import re
from collections import defaultdict
from matching import is_price_code


def rik_code(value):
    original = '' if value is None else str(value)
    text = original.strip()
    if not text:return original
    if is_price_code(text):return 'С прайс-листа'
    # An explicit source material prefix is authoritative even without a star.
    # RIK uses Cyrillic С; retain the entire source suffix, including separators
    # and any existing star, rather than reconstructing the numeric part.
    if text[0] in 'СCсc':
        return 'С' + text[1:]
    match = re.fullmatch(r'(?:[^\W\d_]+(?:[\s.-]+[^\W\d_]+)*[\s.-]*)?([0-9]+(?:-[0-9]+)*)(\*)?', text)
    if not match:
        return original
    digits, star = match.groups()
    return ('С' if star else 'Е') + digits + (star or '')


def add_rik_codes(result):
    """Decorate computed output only, retaining every contributing source code."""
    source_rows = {m['row']['id']: m['row'] for m in result['matches']}
    matched_codes = defaultdict(list)
    for match in result['matches']:
        if match['match']['status'] in ['matched', 'manual_match']:
            for allocation in match['match']['allocations']:
                matched_codes[allocation['target']].append(match['row'].get('code'))
    for rows in result['groups'].values():
        for row in rows:
            codes = [m.get('code') for m in row.get('members', [])]
            codes.extend(row.get('codes', []))
            codes.append(row.get('code'))
            codes.extend(source_rows.get(o.get('id'), {}).get('code') for o in row.get('origins', []))
            if row.get('position_split')=='confirmed':
                codes.extend(source_rows.get(o.get('id'), {}).get('code') for o in row.get('matched_origins', []))
            else: codes.extend(matched_codes[row['id']])
            row['rik_codes'] = list(dict.fromkeys(rik_code(code) for code in codes
                                                  if code is not None and str(code).strip()))
    return result
