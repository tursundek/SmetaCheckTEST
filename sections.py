"""Section provenance and table layout, independent of matching identities."""
import json
import re
from collections import OrderedDict, defaultdict


def parse_section(value):
    text=re.sub(r'\s+', ' ', str(value or '')).strip()
    match=re.fullmatch(r'(Раздел|Подраздел|Глава)\s*(?:№\s*)?(\d+(?:\.\d+)*|[IVXLCDM]+)(?=$|[\s.:)–—-])[\s.:)–—-]*(.*?)',text,re.I)
    if not match:return None
    label=match[1].capitalize()+' '+match[2]
    title=match[3].strip()
    return dict(key=(label+' '+title).casefold(),label=label,title=title,text=text)


def source_heading(item,mapping):
    """Only a standalone header in the selected name/section or merged cell."""
    cells=item['cells']
    def value(field):
        col=mapping.get(field)
        return cells[col] if isinstance(col,int) and 0<=col<len(cells) else ''
    if any(v is not None and str(v).strip() for v in [value('unit'),value('quantity')]):return None
    columns={c for f in ['name','section'] if isinstance((c:=mapping.get(f)),int)}
    columns.update(a for a,b in item.get('merged_spans',[]) if any(a<=c<=b for c in list(columns)))
    # Word grids retain a merged cell's text across header_cells; Excel uses
    # explicit merged_spans. Both locate the anchor covering the chosen name.
    labels=item.get('header_cells',[])
    for col in list(columns):
        if col<len(labels) and col<len(cells) and not cells[col] and labels[col]:
            columns.update(i for i,text in enumerate(cells) if text==labels[col])
    for col in sorted(columns):
        if col>=len(cells):continue
        descriptor=parse_section(cells[col])
        if not descriptor:continue
        if any(value(f) is not None and str(value(f)).strip() and value(f)!=cells[col] for f in ['number','code']):continue
        return descriptor
    return None


def _legacy(value):
    if not str(value or '').strip():return None
    return parse_section(value) or dict(key='legacy:'+str(value).casefold().strip(),label='Раздел',title=str(value).strip(),text=str(value).strip())


def document_sections(doc):
    """Project legacy source headers without reimporting or rewriting rows."""
    result={};config=doc.get('config',{});tables=config.get('tables',[config])
    rows=doc.get('rows',[]);carry=None;previous_hint=None
    for table in tables:
        if not table.get('enabled',True):continue
        sheet=table.get('sheet');raw=doc.get('sheets',{}).get(sheet,[])
        selected=[r for r in rows if r.get('sheet',config.get('sheet'))==sheet]
        if not raw or not table.get('mapping'):continue
        by_index=defaultdict(list)
        for r in selected:by_index[r.get('source_row',int(r.get('source','0').rsplit(':',1)[-1]) if r.get('source','').rsplit(':',1)[-1].isdigit() else 0)].append(r)
        hint=table.get('section_hint','')
        current=parse_section(hint)
        if doc.get('extension')=='.docx' and (not hint or hint==previous_hint):current=carry or current
        overrides={**table.get('overrides',{}),**config.get('overrides',{})}
        primary=not config.get('tables') or sheet==config.get('sheet')
        for item in raw:
            key=str(sheet)+':'+str(item['index'])
            override=overrides.get(key,overrides.get(str(item['index']),{}) if primary else {})
            descriptor=source_heading(item,table['mapping']) if not override else None
            if descriptor:current=descriptor
            for row in by_index.get(item['index'],[]):
                explicit=row.get('source_values',{}).get('section') if table['mapping'].get('section') is not None else None
                if override and row.get('kind')=='section':current=_legacy(row.get('name'))
                info=_legacy(override['section']) if 'section' in override else _legacy(explicit) if explicit else current
                result[row['id']]=info
        if selected:carry=result.get(selected[-1]['id'],current)
        previous_hint=hint
    for row in rows:
        if row['id'] not in result:result[row['id']]=row.get('section_info',_legacy(row.get('section')))
    return result


def preserve_import_sections(doc):
    infos=document_sections(doc)
    subtitles=document_subheadings(doc)
    for row in doc.get('rows',[]):
        row['section_info']=infos.get(row['id'])
        row['subheading_info']=subtitles.get(row['id'])


def source_subheading(item,mapping):
    """A standalone name-cell label, not a position, total or local comment.

    Import already distinguishes incomplete positions from text rows. This
    second, presentation-only classification never creates quantity data.
    """
    from importing import service_reason, header_mapping
    cells=item['cells']
    def value(field):
        col=mapping.get(field)
        return cells[col] if isinstance(col,int) and 0<=col<len(cells) else ''
    name=mapping.get('name')
    if not isinstance(name,int):return None
    anchors={name}
    anchors.update(a for a,b in item.get('merged_spans',[]) if a<=name<=b)
    nonempty=[(col,str(text).strip()) for col,text in enumerate(cells) if text is not None and str(text).strip()]
    if len(nonempty)!=1 or nonempty[0][0] not in anchors:return None
    col,text=nonempty[0]
    for field in ['number','code','unit','quantity','unit_price']:
        if value(field) is None or not str(value(field)).strip():continue
        # A merged label anchored in A is text, not a position number in A.
        covered=any(a==col and a<=name<=b for a,b in item.get('merged_spans',[]))
        if not covered or value(field)!=cells[col]:return None
    if col in item.get('continued_columns',[]):return None
    if not re.search(r'[а-яёa-z]',text,re.I) or parse_section(text):return None
    if service_reason({'name':text},cells):return None
    # Actual KS-2 files contain local correction records between positions,
    # e.g. "Поправки: ПЗ: *1.02/1.14; М: /1.34". Require both a service
    # prefix and numeric arithmetic; a room label mentioning repairs survives.
    if re.match(r'^(?:поправки|коэффициенты)\s*:',text,re.I) and re.search(r'[*×/=]\s*\d',text):return None
    if any(header_mapping(cells).get(field) is not None for field in ['name','unit','quantity']):return None
    # Explicit position comments/continuations do not open a new scope.
    if re.match(r'^(?:примечани[ея]|прим\.|пояснени[ея]|комментари[йи]|обоснование|продолжение наименования|составил|проверил|сдал:|принял:|м\.п\.|\(должность|составлен|наименование|шифр|код)(?=\s|[:.,(]|$)',text.lstrip(' .–—-'),re.I):return None
    return re.sub(r'\s+',' ',text).strip()


def document_subheadings(doc):
    """Read subtitle scope from retained source rows without reimport/save.

    Numbered sections reset the subtitle. Totals, financial records and local
    comments leave it unchanged. A header must precede an imported position;
    footer and preamble text cannot acquire unrelated positions.
    """
    result={};config=doc.get('config',{});rows=doc.get('rows',[])
    for table in config.get('tables',[config]):
        if not table.get('enabled',True) or not table.get('mapping'):continue
        sheet=table.get('sheet');raw=doc.get('sheets',{}).get(sheet,[])
        selected=[r for r in rows if r.get('sheet',config.get('sheet'))==sheet]
        by_index=defaultdict(list)
        for row in selected:
            index=row.get('source_row')
            if index is None:
                suffix=row.get('source','').rsplit(':',1)[-1]
                index=int(suffix) if suffix.isdigit() else 0
            by_index[index].append(row)
        work_indices=[index for index,items in by_index.items() if any(r.get('kind')=='work' for r in items)]
        if not raw or not work_indices:continue
        lower=int(table.get('header',0))+1;upper=max(work_indices)
        current=None;overrides={**table.get('overrides',{}),**config.get('overrides',{})}
        primary=not config.get('tables') or sheet==config.get('sheet')
        for item in raw:
            index=item['index']
            if index<lower or index>upper:continue
            items=by_index.get(index,[])
            key=str(sheet)+':'+str(index)
            override=overrides.get(key,overrides.get(str(index),{}) if primary else {})
            heading=source_heading(item,table['mapping']) if not override else parse_section(override.get('name')) if override.get('kind')=='section' else None
            if heading:current=None
            else:
                text=source_subheading(item,table['mapping'])
                # Do not reinterpret an incomplete work or a manual exclusion.
                accepted=not items or all(r.get('kind')=='section' for r in items)
                if text and accepted and not override:
                    current=dict(key=str(doc.get('id',''))+':'+key,text=text,source=key)
            for row in items:result[row['id']]=current
    for row in rows:
        if row['id'] not in result:result[row['id']]=row.get('subheading_info')
    return result


def _subheading_group(infos):
    distinct=OrderedDict()
    for info in infos:distinct.setdefault(info['key'] if info else '',info)
    values=list(distinct.values())
    if not values or values==[None]:return dict(key='none',title='',headings=[])
    if len(values)==1:return dict(key='single:'+values[0]['key'],title=values[0]['text'],headings=values)
    return dict(key='multiple:'+json.dumps(sorted(distinct),ensure_ascii=False),
                title='',headings=values)


def _group(infos):
    distinct=OrderedDict()
    for info in infos:distinct.setdefault(info['key'] if info else '',info)
    values=list(distinct.values())
    if not values or values==[None]:return dict(key='none',label='',title='Без раздела',sections=[])
    if len(values)==1:return {**values[0], 'key':'single:'+values[0]['key'],'sections':values}
    sections=[info or dict(key='',label='',title='Без раздела',text='Без раздела') for info in values]
    return dict(key='multiple:'+json.dumps(sorted(distinct),ensure_ascii=False),label='Несколько разделов',title='Объединённые позиции',sections=sections)


def decorate_sections(project,result):
    infos={};subtitles={};sources={}
    for doc in project['documents']:
        infos.update(document_sections(doc))
        subtitles.update(document_subheadings(doc))
        for row in doc.get('rows',[]):sources[row['id']]=dict(document=doc.get('name',''),document_id=doc.get('id'),role=doc.get('role'),**row)
    matched=defaultdict(list)
    for match in result['matches']:
        if match['match']['status'] not in ['matched','manual_match']:continue
        source=sources.get(match['row']['id'],{})
        info=infos.get(match['row']['id']);origin={k:source.get(k,'') for k in ['id','number','document','document_id','role','source','quantity','unit']}
        for allocation in match['match']['allocations']:
            target=allocation['target']
            matched[target].append(dict(origin,allocated_quantity=allocation['quantity'],factor=allocation.get('factor','1')))
    def enrich(origin):
        source=sources.get(origin.get('id'),{})
        info=infos.get(origin.get('id'),_legacy(origin.get('section')))
        return {**{k:source[k] for k in ['name','code','unit_price','sheet','source_row'] if k in source},**origin,
                'section':info['text'] if info else '',
                'subheading':(subtitles.get(origin.get('id')) or {}).get('text','')}
    for rows in result['groups'].values():
        for row in rows:
            origins=row.get('origins') or row.get('members') or [row]
            row['section_group']=_group([infos.get(o.get('id'),_legacy(o.get('section'))) for o in origins])
            row['subheading_group']=_subheading_group([subtitles.get(o.get('id')) for o in origins])
            row['origins']=[enrich(o) for o in origins]
            if not row.get('is_new'):
                linked=row.get('matched_origins',matched[row['id']])
                row['matched_origins']=[enrich(o) for o in linked]
            row['source_composition']=source_composition(row,project['documents'])
            if not row.get('is_new'):
                row['matched_origins']=[o for block in row['source_composition'] if block['role']=='act' for o in block['origins']]
    return result


def source_composition(row,documents):
    """Separate source tables, never an invented tender-to-act line pairing.

    Snapshot totals belong to an act, not to individual unconfirmed links.
    Keep original quantities alongside allocations, including split results.
    """
    from matching import number
    docs={d['id']:d for d in documents};blocks=OrderedDict()
    for origin in row.get('origins',[])+row.get('matched_origins',[]):
        doc=docs.get(origin.get('document_id'),{})
        role=origin.get('role') or ('act' if row.get('is_new') else 'tender')
        key=(role,origin.get('document_id') or origin.get('document',''))
        block=blocks.setdefault(key,dict(role=role,document_id=origin.get('document_id',''),
                                         document=origin.get('document',doc.get('name','')),
                                         mode=doc.get('mode','current'),unit=row.get('unit',''),origins=[]))
        block['origins'].append(dict(origin))
    for block in blocks.values():
        block['accounted_quantity']=row.get('quantity') if block['role']=='tender' else row.get('acts',{}).get(block['document_id'],'0')
        block['contribution_confirmed']=block['role']=='tender' or block['document_id'] in row.get('acts',{})
        # Coalesce multiple allocations of the same source within THIS result.
        combined=OrderedDict()
        for origin in block['origins']:
            key=origin.get('id')
            if 'allocated_quantity' in origin:
                origin['allocation_parts']=[dict(quantity=origin['allocated_quantity'],factor=origin.get('factor','1'))]
            if key in combined and 'allocated_quantity' in origin and 'allocated_quantity' in combined[key]:
                combined[key]['allocated_quantity']=str(number(combined[key]['allocated_quantity'])+number(origin['allocated_quantity']))
                combined[key]['allocation_parts'].extend(origin['allocation_parts'])
            else:combined[key]=origin
        block['origins']=list(combined.values())
        for origin in block['origins']:
            if origin.get('allocation_parts'):
                parts=OrderedDict()
                for part in origin['allocation_parts']:
                    factor=str(number(part['factor']))
                    parts[factor]=parts.get(factor,number('0'))+number(part['quantity'])
                origin['allocation_parts']=[dict(quantity=str(q),factor=factor) for factor,q in parts.items()]
                if len(parts)>1:origin.pop('factor',None)
            if block['role']=='act' and block['mode']!='cumulative' and block['contribution_confirmed']:
                origin['accounted_quantity']=str(sum((number(p['quantity'])*number(p['factor']) for p in origin.get('allocation_parts',[dict(quantity=origin['quantity'],factor='1')])),number('0')))
    return list(blocks.values())


def visible_section(section):
    return bool(section.get('sections') and len(section['sections'])==1 and section['sections'][0])


def section_blocks(rows):
    groups=OrderedDict()
    for row in rows:
        section=row.get('section_group') or _group([_legacy(row.get('section'))])
        groups.setdefault(section['key'],dict(section=section,rows=[]))['rows'].append(row)
    for block in groups.values():
        subtitles=OrderedDict()
        for row in block['rows']:
            subtitle=row.get('subheading_group') or _subheading_group([])
            subtitles.setdefault(subtitle['key'],dict(subheading=subtitle,rows=[]))['rows'].append(row)
        # Unscoped / mixed origins precede real subtitles, so their sum never
        # appears under the last unrelated source label. Same-text occurrences
        # retain their distinct source keys, including repeated room labels.
        block['subheadings']=sorted(subtitles.values(),key=lambda g:bool(g['subheading']['title']))
    # Rows spanning sections are shown before named sections, without claiming
    # the combined amount belongs to any one of them.
    return sorted(groups.values(),key=lambda g:visible_section(g['section']))
