"""Name-first identity and exact source-volume validation."""
import re
import unicodedata
from decimal import Decimal, InvalidOperation, localcontext
from functools import lru_cache
from difflib import SequenceMatcher
from materials import price_key

FIELDS = ['number', 'code', 'name', 'unit', 'quantity', 'section', 'unit_price']
ALIASES = {'м2':'м²','кв.м':'м²','м.кв.':'м²','м3':'м³','куб.м':'м³','шт.':'шт','т.':'т'}

def norm(value):
    return re.sub(r'\s+', ' ', str(value or '').lower().replace('ё','е').replace('–','-').replace('—','-').replace('"','').replace('«','').replace('»','')).strip()

def is_price_code(value):
    return bool(re.search(r'п(?:а)?райс|price', str(value or ''), re.IGNORECASE))

def same_source_code(left, right):
    # Price labels identify no norm and cannot relate otherwise different rows.
    return bool(norm(left)) and not is_price_code(left) and not is_price_code(right) and norm(left)==norm(right)

def unit(value, aliases=False):
    value = norm(value)
    if not aliases: return value
    value=value.replace('²','2').replace('³','3').replace('^2','2').replace('^3','3')
    value=re.sub(r'(?<=\d)(?=м(?:[23]\b|\b))', ' ',value)
    value=re.sub(r'\bм\s*([23])\b',r'м\1',value)
    return ALIASES.get(value,value)

@lru_cache(maxsize=8192)
def name_norm(value):
    """Ignore presentation differences, never omit words or specifications."""
    value=norm(unicodedata.normalize('NFKC',str(value or '')))
    value=re.sub(r'(?<![\w.])\d+[.,]\d+(?![\w.])',
                 lambda m:m.group().replace(',','.').rstrip('0').rstrip('.'),value)
    value=re.sub(r'(?<!\d)[,.]|[,.](?!\d)', ' ', value)
    value=re.sub(r'[():]', ' ', value)
    value=re.sub(r'\s*([/\\*+\-])\s*',r'\1',value)
    return re.sub(r'\s+',' ',value).strip()

def number(value):
    if value is None or str(value).strip() == '': raise ValueError('Пустой объём')
    s = re.sub(r'[\s\u00a0\u202f]', '', str(value))
    if ',' in s and '.' in s:
        s = s.replace(',', '') if s.rfind('.') > s.rfind(',') else s.replace('.', '').replace(',', '.')
    else: s = s.replace(',', '.')
    try:
        d = Decimal(s)
        if not d.is_finite(): raise InvalidOperation()
        return d
    except InvalidOperation: raise ValueError('Нераспознанный объём: ' + str(value))

@lru_cache(maxsize=4096)
def _position_type(name,code,measurement):
    # This is a display label, never a substitute for checking the full name.
    actions=r'(?:установк|монтаж|демонтаж|устройств|разборк|сняти|очистк|пробивк|усилен|оштукатур|штукатур|шпаклев|шпатлев|выравнив|окраск|облицовк|прокладк|погрузк|перевозк|бетониров|заливк|армиров|укладк|кладк|бурен|сверлен|сварк|засыпк|разработк|уплотнен|ремонт|изготовлен|резк|обработк|гидроизоляц|стяжк)'
    modifiers=r'(?:(?:сплошн|ручн|механизированн|частичн|повторн|однослойн|двухслойн)\w*\s+){0,2}'
    if re.match('^'+modifiers+actions+r'\w*\b',name): return 'work'
    if re.match(r'^грунтовк\w*\b',name):
        return 'material' if re.match(r'^(?:\d+\s*)?(?:кг|т|л|литр)\b',measurement) else 'work'
    products=r'(?:блок|двер|окон|плитк|плит[аы]|лист|пенопласт|угольник|двутавр|уголок|уголк|краск|цемент|бетон|раствор|смес|шпатлевк|шпаклевк|шкаф|гофр|провод|кабел|автомат|реле|узо|хомут|кросс|скобян|креплен|труб|профил|арматур|песок|щебен|пластик|грунт)'
    if re.search(r'\b'+products+r'\w*\b',' '.join(name.split()[:3])): return 'material'
    if re.match(r'^(?:крссц|ссц|фссц|тссц)\b',code) or 'прайс' in code: return 'material'
    return 'work'

def position_type(row):
    # Only the original source code is authoritative. rik_codes is decoration.
    code=norm(row.get('code')).lstrip('. -')
    if code.startswith(('с','c')) or is_price_code(code) or re.match(r'^(?:крссц|фссц|тссц)\b',code):
        return 'material'
    if re.match(r'^(?:е|e|крер|фер|тер|гэсн|ц(?=\d))',code):
        return 'work'
    return _position_type(name_norm(row.get('name')),code,unit(row.get('unit'),True))

def material_compatible(left,right):
    a,b=position_type(left),position_type(right)
    if a!=b: return False
    if a!='material': return True
    price=price_key(left.get('unit_price'))
    return price is not None and price==price_key(right.get('unit_price'))

def identity_key(row,aliases=False):
    key=(name_norm(row.get('name')),unit(row.get('unit'),aliases))
    if position_type(row)=='material':
        price=price_key(row.get('unit_price'))
        return (*key,'material',price if price is not None else ('unknown',row['id']))
    return (*key,'work')

def work_signature(row,aliases=False):
    return (*identity_key(row,aliases),norm(row.get('section')),
            name_norm(row.get('context_name')) or norm(row.get('context')))

def tender_totals(rows,aliases=False):
    """Sum every source occurrence of the same work; never edit source rows."""
    if all(r.get('kind')!='work' or 'members' in r for r in rows): return rows
    result=[];groups={}
    for row in rows:
        if row.get('kind')!='work': result.append(dict(row));continue
        key=identity_key(row,aliases)
        # Missing identity fields must stay attached to their original row.
        if not key[0] or not key[1]: key=(*key,row['id'])
        if key not in groups:
            groups[key]=dict(row,members=[],errors=list(row.get('errors',[])))
            result.append(groups[key])
        group=groups[key]
        group['members'].append({k:row.get(k,'') for k in ['id','number','name','code','unit','quantity','section','source','context_name','unit_price']})
        for error in row.get('errors',[]):
            if error not in group['errors']: group['errors'].append(error)
    for group in groups.values():
        members=group['members'];group['member_count']=len(members)
        if len(members)==1: continue
        try:
            quantities=[number(m['quantity']) for m in members]
            width=max(1,max(q.adjusted()+1 for q in quantities))+max(0,max(-q.as_tuple().exponent for q in quantities))
            with localcontext() as context:
                context.prec=max(50,width+len(str(len(members)))+5)
                group['quantity']=str(sum(quantities,Decimal(0)))
        except ValueError:
            group['quantity']=None
            if 'Не удалось прочитать объём исходной позиции тендера' not in group['errors']: group['errors'].append('Не удалось прочитать объём исходной позиции тендера')
        numbers=[str(m['number']) for m in members if m['number']]
        group['number']=', '.join(numbers[:3])+(' (+'+str(len(numbers)-3)+')' if len(numbers)>3 else '')
        sections=list(dict.fromkeys(str(m['section']) for m in members if m['section']))
        group['sections']=sections
        group['section']=sections[0] if len(sections)==1 else ''
        group['context']=group['context_name']=''
        group['codes']=list(dict.fromkeys(m['code'] for m in members if m['code']))
    return result

def candidates(row, tender, aliases=False, numbering=False):
    a=name_norm(row.get('name'))
    if not a: return []
    result=[]
    for t in tender:
        if t.get('kind')!='work' or a!=name_norm(t.get('name')): continue
        if not material_compatible(row,t): continue
        same_unit=bool(unit(row.get('unit'),aliases)) and unit(row.get('unit'),aliases)==unit(t.get('unit'),aliases)
        code=norm(row.get('code'))
        same_code=same_source_code(row.get('code'),t.get('code'))
        pos=numbering and bool(row.get('number')) and norm(row['number'])==norm(t.get('number'))
        section=bool(row.get('section')) and norm(row['section'])==norm(t.get('section'))
        valid=True
        try: number(t.get('quantity'))
        except ValueError: valid=False
        safe=same_unit and valid and not t.get('errors')
        rule='Точное наименование и единица' if same_unit else 'Наименование совпало, единицы требуют проверки'
        if t.get('errors') or not valid: rule='Сначала уточните исходный объём в смете'
        result.append(dict(id=t['id'],score=1 if safe else .9,safe=safe,rule=rule,
                           evidence=dict(code=same_code,number=bool(pos),unit=same_unit,section=section,text=100,critical=True,action=True)))
    return sorted(result,key=lambda c:c['score'],reverse=True)

def different_unit_candidates(row,tender,aliases=False):
    """Related rows are review context only, never matching candidates."""
    result=[];a=name_norm(row.get('name'));code=norm(row.get('code'))
    for t in tender:
        if t.get('kind')!='work' or not unit(t.get('unit'),aliases) or unit(row.get('unit'),aliases)==unit(t.get('unit'),aliases): continue
        if not material_compatible(row,t): continue
        b=name_norm(t.get('name'));same_name=a==b
        same_code=same_source_code(row.get('code'),t.get('code'))
        similar=bool(a and b and len(a)>=8 and len(b)>=8 and len(set(a.split())&set(b.split()))>=2 and SequenceMatcher(None,a,b).ratio()>=.85)
        if same_name or same_code or similar:
            result.append(dict(id=t['id'],score=.9,safe=False,rule='Требует уточнения',evidence=dict(code=same_code,unit=False,text=100 if same_name else 0)))
    return result

def automatic(row,tender,settings):
    tender=tender_totals(tender,settings.get('aliases'))
    if row['kind']!='work': return dict(status='ignored',candidates=[],allocations=[])
    if row.get('errors') or not name_norm(row.get('name')) or not unit(row.get('unit')):
        return dict(status='invalid',candidates=[],allocations=[])
    if position_type(row)=='material' and price_key(row.get('unit_price')) is None:
        return dict(status='invalid',candidates=[],allocations=[],rule='Укажите цену за единицу материала в настройках импорта')
    try: number(row.get('quantity'))
    except ValueError: return dict(status='invalid',candidates=[],allocations=[])
    cs=candidates(row,tender,settings.get('aliases'),settings.get('numbering'))
    related=different_unit_candidates(row,tender,settings.get('aliases'))
    if related and not any(c['evidence']['unit'] for c in cs):
        return dict(status='new',candidates=related,allocations=[],unit_difference=True,rule='Требует уточнения')
    learned=settings.get('_learned',{}).get(work_signature(row,settings.get('aliases')))
    # A remembered decision cannot bypass the current name/unit/data checks.
    # Conversions, even previously confirmed by a person, are per-row decisions.
    try: reusable=bool(learned) and number(learned.get('factor','1'))==1
    except ValueError: reusable=False
    if reusable and any(c['id']==learned['target'] and c['safe'] for c in cs):
        return dict(status='matched',candidates=cs,rule='По ранее подтверждённому соответствию',learned=True,
                    allocations=[dict(target=learned['target'],factor='1',quantity=row['quantity'])])
    if not cs:
        return dict(status='new' if any(t.get('kind')=='work' for t in tender) else 'unmatched',candidates=[],allocations=[],
                    rule='В тендере нет позиции с таким наименованием')
    scoped=[c for c in cs if c['safe']]
    targets={t['id']:t for t in tender}
    for field in ['section','context','number','code']:
        if len(scoped)<=1: break
        if field=='context':
            named=[c for c in scoped if row.get('context_name') and targets[c['id']].get('context_name') and
                   name_norm(row['context_name'])==name_norm(targets[c['id']]['context_name'])]
            # Parent/code information only chooses between positions whose OWN
            # full names and units already match; it cannot bypass the name gate.
            narrowed=named or [c for c in scoped if row.get('context') and targets[c['id']].get('context')==row['context']]
        else: narrowed=[c for c in scoped if c['evidence'][field]]
        if narrowed: scoped=narrowed
    if len(scoped)==1:
        chosen=scoped[0]
        return dict(status='matched',candidates=cs,rule=chosen['rule'],
                    allocations=[dict(target=chosen['id'],quantity=row['quantity'],factor='1')])
    rule='Единицы измерения различаются. Автоматическое объединение запрещено' if all(not c['evidence']['unit'] for c in cs) else 'Нужно уточнить исходные данные'
    return dict(status='ambiguous',candidates=cs,allocations=[],rule=rule)

def validate_decision(row,decision,tender,settings):
    if decision.get('status')=='ignored': return
    values=[number(row.get('quantity'))]+[number(a['quantity']) for a in decision.get('allocations',[])]
    integer_digits=max(1,max(v.adjusted()+1 for v in values))
    fraction_digits=max(0,max(-v.as_tuple().exponent for v in values))
    with localcontext() as context:
        context.prec=max(50,integer_digits+fraction_digits+len(str(len(values)))+5)
        return _validate_decision(row,decision,tender,settings)

def _validate_decision(row,decision,tender,settings):
    status=decision['status']
    if status not in ['manual_match','new','ignored']: raise ValueError('Недопустимое решение')
    if status=='ignored': return
    if row.get('kind')!='work':
        raise ValueError('Начисление, раздел или итог нельзя подтвердить как работу')
    if not name_norm(row.get('name')) or not unit(row.get('unit')):
        raise ValueError('Укажите наименование и единицу перед подтверждением')
    if position_type(row)=='material' and price_key(row.get('unit_price')) is None:
        raise ValueError('Укажите цену за единицу материала в настройках импорта')
    if row.get('errors') and not (row.get('quantity') is not None and all('Отрицательный' in x for x in row['errors'])):
        raise ValueError('Исправьте ошибку импорта перед подтверждением')
    q=number(row.get('quantity'))
    if q<0 and not decision.get('negative_confirmed'): raise ValueError('Подтвердите отрицательный объём')
    if status=='new':
        if any(c['evidence']['unit'] for c in candidates(row,tender,settings.get('aliases'),settings.get('numbering'))):
            raise ValueError('Позиция с таким наименованием уже есть в тендере. Уточните соответствие или единицы')
        return
    allocations=decision.get('allocations',[])
    if not allocations: raise ValueError('Укажите распределение')
    if sum((number(a['quantity']) for a in allocations),Decimal(0))!=q:
        raise ValueError('Сумма распределения не равна исходному объёму')
    targets={t['id']:t for t in tender if t['kind']=='work' and not t.get('errors')}
    seen=set()
    for a in allocations:
        if a['target'] in seen: raise ValueError('Позиция повторена в распределении')
        seen.add(a['target'])
        if a['target'] not in targets: raise ValueError('Нет подтверждённой позиции тендера')
        if name_norm(row['name'])!=name_norm(targets[a['target']]['name']):
            raise ValueError('Наименования различаются: это отдельная позиция. Исправьте исходную строку, если в названии ошибка')
        if not material_compatible(row,targets[a['target']]):
            raise ValueError('Материалы с разными ценами или типами нельзя объединять')
        factor=number(a.get('factor','1'))
        if factor<=0: raise ValueError('Коэффициент должен быть положительным')
        if q>=0 and number(a['quantity'])<0 or q<0 and number(a['quantity'])>0:
            raise ValueError('Знак части распределения отличается от знака исходного объёма')
        same=unit(row['unit'],settings.get('aliases'))==unit(targets[a['target']]['unit'],settings.get('aliases'))
        if not same: raise ValueError('Разные единицы: это отдельная новая позиция. Исправьте единицу исходной строки, если она прочитана неверно')
        if (not same or factor!=1) and not decision.get('conversion_confirmed'):
            raise ValueError('Подтвердите преобразование единиц')
