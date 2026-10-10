"""Partition calculated totals by confirmed source occurrences, never by volume fit.

Matching still addresses whole identity groups. This projection keeps its totals
and exposes individual increases/decreases only when every incoming allocation
has a supported source position. Otherwise the original aggregate remains.
"""
import re
from collections import defaultdict
from decimal import Decimal

from matching import identity_key, name_norm, norm, number, position_type, tender_totals
from sections import document_sections


def _sheet(row):
    return str(row.get('source', '')).rpartition(':')[0]


def _layout_key(row, aliases):
    position = re.sub(r'\.$', '', norm(row.get('number')))
    if not position: return None
    # Price labels/codes cannot prove material identity. Its name, unit and
    # actual unit price are included in identity_key instead.
    code = norm(row.get('code')) if position_type(row)=='work' else ''
    return (position, identity_key(row, aliases), code)


def _layout_links(tender_doc, acts, aliases, infos):
    """Corroborate unique numbers with the composition/order of a source table.

    All selected act works must have unique full keys in one tender table and
    preserve that table's sequence. At least two different semantic identities
    must corroborate it. Excel row indices and quantities are never evidence.
    Tables with only repeated identical work, duplicate numbers/keys, changed
    identities, or multiple possible source tables do not provide a link.
    """
    tender_tables=defaultdict(list)
    for row in tender_doc.get('rows', []):
        if row.get('kind')=='work': tender_tables[_sheet(row)].append(row)
    links={}
    for doc in acts:
        tables=defaultdict(list)
        for row in doc['rows']:
            if row.get('kind')=='work': tables[_sheet(row)].append(row)
        for rows in tables.values():
            keys=[_layout_key(row, aliases) for row in rows]
            if None in keys or len(set(keys))!=len(keys): continue
            if len({key[1] for key in keys})<2: continue
            possible=[]
            for tender_rows in tender_tables.values():
                index=defaultdict(list)
                for i,row in enumerate(tender_rows): index[_layout_key(row,aliases)].append((i,row))
                if not all(len(index[key])==1 for key in keys): continue
                selected=[index[key][0] for key in keys]
                if any(a[0]>=b[0] for a,b in zip(selected,selected[1:])): continue
                if any(row.get('errors') or original.get('errors') for row,(_,original) in zip(rows,selected)): continue
                # Real section/context evidence cannot conflict with numbering.
                conflict=False
                for row,(_,original) in zip(rows,selected):
                    a,b=infos.get(row['id']),infos.get(original['id'])
                    if a and b and a['key']!=b['key']: conflict=True
                    a,b=name_norm(row.get('context_name')),name_norm(original.get('context_name'))
                    if a and b and a!=b: conflict=True
                if not conflict: possible.append([original['id'] for _,original in selected])
            if len(possible)==1: links.update(zip((row['id'] for row in rows),possible[0]))
    return links


def _member_scope(members,aliases,infos):
    index={}
    for member in members:
        scope=index.setdefault(identity_key(member,aliases),{field:defaultdict(set) for field in ['section','context','code']})
        section=infos.get(member['id'])
        if section:scope['section'][section['key']].add(member['id'])
        context=name_norm(member.get('context_name'))
        if context:scope['context'][context].add(member['id'])
        code=norm(member.get('code'))
        if code and position_type(member)=='work':scope['code'][code].add(member['id'])
    return index


def _scoped_member(row, index, aliases, infos):
    scope=index.get(identity_key(row,aliases),{})
    evidence=[]
    section=infos.get(row['id'])
    if section:
        same=scope.get('section',{}).get(section['key'])
        if same: evidence.append(same)
    context=name_norm(row.get('context_name'))
    if context:
        same=scope.get('context',{}).get(context)
        if same: evidence.append(same)
    code=norm(row.get('code'))
    if code and position_type(row)=='work':
        same=scope.get('code',{}).get(code)
        if same: evidence.append(same)
    if not evidence: return None, False
    narrowed=set.intersection(*evidence)
    return (next(iter(narrowed)) if len(narrowed)==1 else None), not bool(narrowed)


def _origin(row,doc,allocation):
    return {**{k:row.get(k,'') for k in ['id','number','section','source','quantity','unit']},
            'document_id':doc['id'],'document':doc['name'],'role':'act',
            'allocated_quantity':str(allocation['quantity']),
            'factor':str(allocation.get('factor','1')),'mode':doc.get('mode','current')}


def _description(category,row,actual):
    if category=='unchanged': return 'Без изменений'
    if category=='added': return 'Объём увеличен'
    if actual==0: return 'Материал исключён' if position_type(row)=='material' else 'Работа исключена'
    return 'Объём уменьшен'


def partition_groups(result,docs,settings):
    parents={r['id']:r for r in result['summary'] if r['kind']=='work' and not r['needs_review']}
    repeated={key:row for key,row in parents.items() if row['member_count']>1}
    if not repeated or not result['ready']:
        _compact_numbers(result)
        return
    aliases=settings.get('aliases')
    tender_doc=next(d for d in docs if d['role']=='tender')
    originals={r['id']:r for r in tender_doc['rows'] if r['kind']=='work'}
    member_group={m['id']:key for key,row in repeated.items() for m in row['members']}
    acts={d['id']:d for d in result['acts']}
    infos={}
    for doc in [tender_doc,*acts.values()]: infos.update(document_sections(doc))
    layout=_layout_links(tender_doc,list(acts.values()),aliases,infos)
    members={key:{m['id'] for m in parent['members']} for key,parent in repeated.items()}
    scopes={key:_member_scope([originals[member] for member in ids],aliases,infos) for key,ids in members.items()}
    entries=defaultdict(list); uncertain=set()
    for match in result['matches']:
        row,m=match['row'],match['match']
        if m['status'] not in ['matched','manual_match']: continue
        doc=acts[match['document_id']]
        if m['status']=='manual_match':
            # These allocations were validated against original tender IDs by
            # core. Their remapping to group IDs must not erase manual evidence.
            for allocation in row['decision']['allocations']:
                member=allocation['target']; key=member_group.get(member)
                if key: entries[key].append((member,doc,allocation,_origin(row,doc,allocation)))
            continue
        for allocation in m['allocations']:
            key=allocation['target']
            if key not in repeated: continue
            scoped,conflict=_scoped_member(row,scopes[key],aliases,infos)
            structural=layout.get(row['id'])
            if structural not in members[key]: structural=None
            if conflict or (scoped and structural and scoped!=structural):
                uncertain.add(key); continue
            member=scoped or structural
            if member: entries[key].append((member,doc,allocation,_origin(row,doc,allocation)))
            else: uncertain.add(key)
    original_categories={row['id']:(category,row) for category,rows in result['groups'].items()
                         for row in rows if not row['is_new']}
    groups={key:[row for row in rows if row['is_new']] for key,rows in result['groups'].items()}
    for key,parent in parents.items():
        category,old=original_categories[key]
        partitions=None
        if key in repeated and key not in uncertain and entries[key]:
            partitions=_partition(parent,entries[key],list(acts.values()),originals,aliases)
        if partitions is None:
            groups[category].append(old)
        else:
            for category,row in partitions: groups[category].append(row)
    result['groups']=groups
    _compact_numbers(result)


def _partition(parent,entries,acts,originals,aliases):
    ids=[m['id'] for m in parent['members']]
    snapshots={member:defaultdict(Decimal) for member in ids}
    presence=defaultdict(set)
    for member,doc,allocation,origin in entries:
        snapshots[member][doc['id']]+=number(allocation['quantity'])*number(allocation.get('factor','1'))
        presence[doc['id']].add(member)
    for doc in acts:
        if doc.get('mode')=='cumulative' and presence.get(doc['id']) and presence[doc['id']]!=set(ids): return None
    contributions={member:{} for member in ids}
    for member in ids:
        previous=Decimal(0)
        for doc in acts:
            if doc['id'] not in snapshots[member]: continue
            value=snapshots[member][doc['id']]
            if doc.get('mode')=='cumulative':
                delta=value-previous
                if delta<0: return None
                previous=value;value=delta
            contributions[member][doc['id']]=value
    # A display partition may never change even a single per-act group total.
    for doc in acts:
        if sum((contributions[m].get(doc['id'],Decimal(0)) for m in ids),Decimal(0))!=number(parent['acts'].get(doc['id'],'0')): return None
    buckets=defaultdict(list)
    actuals={member:sum(contributions[member].values(),Decimal(0)) for member in ids}
    for member in ids:
        planned=number(originals[member]['quantity']);actual=actuals[member]
        category='unchanged' if actual==planned else 'added' if actual>planned else 'minus'
        buckets[category].append(member)
    if len(buckets)==1: return None
    projected=[]
    for category,members in buckets.items():
        subset=tender_totals([originals[member] for member in members],aliases)[0]
        planned=sum((number(originals[member]['quantity']) for member in members),Decimal(0))
        actual=sum((actuals[member] for member in members),Decimal(0))
        selected_ids=set(members)
        selected=[entry for entry in entries if entry[0] in selected_ids]
        origin_by_id={o['id']:o for o in parent['origins']}
        description=_description(category,parent,actual)
        provenance=[f"{doc['name']} / № {origin['number'] or 'не указан'} / {origin['source']} / {allocation['quantity']} × {allocation.get('factor','1')}"
                    for _,doc,allocation,origin in selected]
        row=dict(parent,**{k:subset[k] for k in ['number','members','member_count','codes','sections','section','context','context_name'] if k in subset})
        row.update(id=key_for_partition(parent['id'],category),group_id=parent['id'],
                   quantity=str(planned),actual=str(actual),added=str(max(actual-planned,0)),minus=str(max(planned-actual,0)),
                   volume=str(actual if category=='unchanged' else actual-planned),description=description,status=description,is_new=False,
                   acts={doc['id']:str(sum((contributions[m].get(doc['id'],Decimal(0)) for m in members),Decimal(0))) for doc in acts if any(doc['id'] in contributions[m] for m in members)},
                   origins=[origin_by_id[m] for m in members],matched_origins=[origin for _,_,_,origin in selected],
                   tender_sources=[f"Тендер / {originals[m]['source']} / № {originals[m]['number']} / {originals[m]['quantity']} {originals[m]['unit']}" for m in members],
                   provenance=provenance,position_split='confirmed')
        # The representative row must also belong to this partition.
        row['code']=originals[members[0]].get('code','')
        row['codes']=list(dict.fromkeys(originals[m].get('code','') for m in members))
        row['source']=originals[members[0]].get('source','')
        projected.append((category,row))
    return projected


def key_for_partition(parent,category):
    return parent+':positions:'+category


def _compact_numbers(result):
    for row in result['groups']['added']:
        sources=row.get('origins') or row.get('members') or [row]
        row['display_number']=sources[0].get('number','') if len(sources)==1 else ''
