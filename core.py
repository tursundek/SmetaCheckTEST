"""Strict name-based reconciliation; source rows are immutable."""
from decimal import Decimal, localcontext
from collections import defaultdict
from matching import FIELDS, norm, unit, name_norm, number, position_type, work_signature, candidates, automatic, validate_decision, tender_totals, material_compatible
from materials import price_key, material_price
from sections import decorate_sections
from result_positions import partition_groups

def reconcile(project):
    from importing import material_project
    project=material_project(project)
    precision=50
    for doc in project['documents']:
        for row in doc.get('rows',[]):
            inputs=[row.get('quantity')]
            for a in row.get('decision',{}).get('allocations',[]): inputs.extend([a.get('quantity'),a.get('factor','1')])
            for value in inputs:
                try:
                    n=number(value)
                    width=max(0,n.adjusted()+1)+max(0,-n.as_tuple().exponent)
                    precision=max(precision,2*width+10)
                except ValueError: pass
    with localcontext() as ctx:
        ctx.prec=precision+len(str(sum(len(d.get('rows',[])) for d in project['documents'])))
        return _reconcile(project)

def _reconcile(project):
    docs = [d for d in project['documents'] if d.get('imported') and d.get('active',True)]
    original_tender = next((d['rows'] for d in docs if d['role']=='tender'),[])
    acts = [d for d in docs if d['role']=='act']
    settings = dict(project.get('settings',{}))
    tender=tender_totals(original_tender,settings.get('aliases'))
    tender_doc=next((d for d in docs if d['role']=='tender'),{})
    def origin(row,doc):
        return {**{k:row.get(k,'') for k in ['id','number','section','source','quantity','unit']},
                'document_id':doc.get('id',''),'document':doc.get('name',''),'role':doc.get('role','')}
    for t in tender:
        if t['kind']=='work':
            t['origins']=[origin(m,tender_doc) for m in t['members']]
    tender_targets={member['id']:t['id'] for t in tender if t['kind']=='work' for member in t['members']}
    # Only identical units with factor 1 may be remembered for automatic use.
    # A manually confirmed conversion belongs exclusively to its own row.
    remembered=defaultdict(set)
    for doc in acts:
        for row in doc['rows']:
            decision=row.get('decision',{})
            if decision.get('status')!='manual_match' or len(decision.get('allocations',[]))!=1: continue
            try: validate_decision(row,decision,original_tender,settings)
            except ValueError: continue
            allocation=decision['allocations'][0]
            target=next(t for t in original_tender if t['id']==allocation['target'])
            if unit(row['unit'],settings.get('aliases'))!=unit(target['unit'],settings.get('aliases')) or number(allocation.get('factor','1'))!=1:
                continue
            remembered[work_signature(row,settings.get('aliases'))].add((tender_targets[allocation['target']],format(number(allocation.get('factor','1')).normalize(),'f')))
    settings['_learned']={signature:dict(target=next(iter(values))[0],factor=next(iter(values))[1]) for signature,values in remembered.items() if len(values)==1}
    contributions=defaultdict(lambda:defaultdict(Decimal)); sources=defaultdict(list)
    matches=[]; issues=[]; news={}; pending=set(); confidences=defaultdict(set)
    volume_checks=[]; volume_issues=[]
    for doc in docs:
        for warning in doc.get('warnings',[]): issues.append({'source':doc['name'],'message':warning})
        for r in doc['rows']:
            for error in r.get('errors',[]):
                if r.get('decision',{}).get('negative_confirmed') and 'Отрицательный' in error: continue
                if r.get('decision',{}).get('status')=='ignored': continue
                issues.append({'source':f"{doc['name']} / {r['source']}",'message':error})
    for doc in acts:
        values=defaultdict(Decimal)
        for row in doc['rows']:
            m = row.get('decision') or automatic(row,tender,settings)
            if row.get('decision'):
                targets={t['id']:t for t in original_tender}
                outdated=m.get('status')=='manual_match' and any(
                    a.get('target') in targets and (name_norm(row['name'])!=name_norm(targets[a['target']]['name']) or unit(row['unit'],settings.get('aliases'))!=unit(targets[a['target']]['unit'],settings.get('aliases')) or not material_compatible(row,targets[a['target']]))
                    for a in m.get('allocations',[]))
                if outdated:
                    # Preserve the saved decision/history, but the newly agreed
                    # name rule takes precedence over incompatible old links.
                    m={**automatic(row,tender,settings),'previous_decision_ignored':True}
                else:
                    try: validate_decision(row,m,original_tender,settings)
                    except ValueError as e: m={'status':'invalid','candidates':[],'allocations':[]}; issues.append({'source':row['id'],'message':str(e)})
            if m['status'] in ['matched','manual_match']:
                m={**m,'allocations':[{**a,'target':tender_targets.get(a['target'],a['target'])} for a in m['allocations']]}
            matches.append({'document':doc['name'],'document_id':doc['id'],'row':dict(row,position_type=position_type(row)),'match':m})
            if row['kind']=='work':
                try: source_q=number(row.get('quantity'))
                except ValueError: source_q=None
                checked_q=None
                if m['status'] in ['matched','manual_match']:
                    checked_q=sum((number(a['quantity']) for a in m['allocations']),Decimal(0))
                elif m['status'] in ['new','ignored']: checked_q=source_q
                balanced=source_q is not None and source_q==checked_q
                if m['status'] in ['matched','manual_match','new'] and not balanced:
                    raise ValueError('Не сходится исходный объём строки '+row['source'])
                volume_checks.append(dict(document=doc['id'],source=row['id'],unit=row['unit'],
                                          quantity=str(source_q) if source_q is not None else None,
                                          accounted=str(checked_q) if checked_q is not None else None,
                                          balanced=balanced,status=m['status']))
            if m['status'] in ['ambiguous','unmatched','invalid']:
                if m['status']=='invalid' and not name_norm(row.get('name')):
                    pending.update(t['id'] for t in tender if t['kind']=='work' and (not unit(row.get('unit')) or unit(row['unit'],settings.get('aliases'))==unit(t['unit'],settings.get('aliases'))))
                cs=m.get('candidates',[]) or candidates(row,tender,settings.get('aliases'),settings.get('numbering'))
                if cs:
                    for c in cs:
                        if c['score']>=cs[0]['score']-.08 and (not unit(row.get('unit')) or c['evidence']['unit']): pending.add(c['id'])
                reason=m.get('rule') or {'ambiguous':'Нужно выбрать позицию тендера','unmatched':'Работа не найдена в тендере','invalid':'Проверьте объём или единицу измерения'}[m['status']]
                if not row.get('errors'): issues.append({'source':f"{doc['name']} / {row['source']}",'message':reason})
            if m['status'] in ['matched','manual_match']:
                for a in m['allocations']:
                    values[a['target']] += number(a['quantity'])*number(a.get('factor','1'))
                    confidences[a['target']].add('Ручная' if m['status']=='manual_match' else 'Высокая')
                    sources[a['target']].append(f"{doc['name']} / № {row.get('number') if row.get('number') is not None and str(row.get('number')).strip() else 'не указан'} / {row.get('section','')} / {row['source']} / {a['quantity']} × {a.get('factor','1')} ({m['status']}) {m.get('comment','')}")
            elif m['status']=='new':
                key='new:'+name_norm(row['name'])+'|'+unit(row['unit'],settings.get('aliases'))
                if position_type(row)=='material':
                    price=format(price_key(row.get('unit_price')),'f')
                    if '.' in price: price=price.rstrip('0').rstrip('.')
                    key+='|material|'+price
                news.setdefault(key,{'id':key,'name':row['name'],'unit':row['unit'],'code':row.get('code',''),
                                     'unit_price':row.get('unit_price'),'position_type':position_type(row),'sources':[],'origins':[],'review_rows':[]})['sources'].append(f"{doc['name']} / № {row.get('number') or 'не указан'} / {row.get('section','')} / {row['source']}")
                news[key]['origins'].append(origin(row,doc))
                if m.get('unit_difference') and number(row['quantity'])!=0: news[key]['review_rows'].append(row['id'])
                values[key]+=number(row['quantity'])
        for target,q in values.items():
            contributions[target][doc['id']]=q
    # Cumulative values represent successive snapshots. Missing rows do not reset a snapshot.
    for target,bydoc in contributions.items():
        previous=Decimal(0)
        modes=set()
        for doc in acts:
            if doc['id'] not in bydoc: continue
            modes.add(doc.get('mode','current'))
            if doc.get('mode','current') == 'cumulative':
                snapshot=bydoc[doc['id']]
                delta=snapshot-previous
                if delta < 0:
                    issues.append({'source':doc['name'],'message':'Накопительный объём уменьшился для '+target})
                    pending.add(target)
                    volume_issues.append(dict(target=target,document=doc['id'],reason='Накопительный объём уменьшился. Проверьте режим и порядок процентовок.'))
                bydoc[doc['id']]=delta
                previous=snapshot
        if len(modes)>1:
            issues.append({'source':target,'message':'Смешаны текущие и накопительные акты: позиция исключена из расчёта'})
            pending.add(target); bydoc.clear()
            volume_issues.append(dict(target=target,document=None,reason='Для одной позиции смешаны объёмы за период и накопительные объёмы. Проверьте режимы файлов.'))
    summary=[]
    for t in tender:
        if t['kind']!='work': summary.append(dict(t,actual=None,added=None,minus=None,status='Раздел / итог',acts={},provenance=[])); continue
        actual=sum(contributions[t['id']].values(),Decimal(0)); planned=number(t['quantity']) if not t.get('errors') else None
        excluded='Материал исключён' if position_type(t)=='material' else 'Работа исключена'
        status='Требует проверки' if planned is None or t['id'] in pending else 'Без изменений' if actual==planned else excluded if actual==0 else 'Объём увеличен' if actual>planned else 'Объём уменьшен'
        tender_sources=[f"Тендер / {m['source']} / № {m['number']} / {m['quantity']} {m['unit']}" for m in t['members']]
        summary.append(dict(t,position_type=position_type(t),actual=str(actual),added=str(max(actual-planned,0)) if planned is not None else None,minus=str(max(planned-actual,0)) if planned is not None else None,status=status,needs_review=planned is None or t['id'] in pending,confidence=', '.join(sorted(confidences[t['id']])),acts={k:str(v) for k,v in contributions[t['id']].items()},provenance=sources[t['id']],tender_sources=tender_sources,warning='В актах не найден подтверждённый объём. Проверьте, работа исключена или ещё не отражена в загруженных актах' if actual==0 else ''))
    for key,n in news.items(): n['actual']=str(sum(contributions[key].values(),Decimal(0))); n['acts']={k:str(v) for k,v in contributions[key].items()};n['needs_review']=key in pending
    unimported=[d for d in project['documents'] if d.get('active',True) and not d.get('imported')]
    seen_files=set()
    for doc in docs:
        if not doc.get('hash'): continue
        key=(doc['role'],doc['hash'])
        if key in seen_files:
            unimported.append(doc)
            issues.append(dict(source=doc['name'],message='Одинаковый файл включён дважды. Выключите повторный файл.'))
        seen_files.add(key)
    work_matches=[m for m in matches if m['row']['kind']=='work']
    review_count=sum(m['match']['status'] in ['ambiguous','unmatched','invalid'] or (bool(m['match'].get('unit_difference')) and number(m['row']['quantity'])!=0) for m in work_matches)+sum(bool(t.get('errors')) for t in tender if t['kind']=='work')+len({i['target'] for i in volume_issues})
    result={'summary':summary,'matches':matches,'issues':issues,'new':list(news.values()),'acts':acts,
            'ready':bool(tender and acts) and not unimported,
            'pending_documents':[dict(id=d['id'],name=d['name']) for d in unimported],
            'volume_checks':volume_checks,'volume_issues':volume_issues,
            'metrics':dict(total=len(work_matches),automatic=sum(m['match']['status'] in ['matched','new'] and (not m['row'].get('decision') or m['match'].get('previous_decision_ignored',False)) for m in work_matches),review=review_count)}
    result['groups']=classify(result)
    partition_groups(result,docs,settings)
    for rows in result['groups'].values():
        for row in rows:
            row['price_display']=material_price(row)
    return decorate_sections(project,result)

def classify(result):
    """Only confirmed rows enter the three final categories."""
    groups={'added':[],'minus':[],'unchanged':[]}
    if result.get('ready') is False: return groups
    for row in result['summary']:
        if row['kind']!='work' or row['needs_review']: continue
        actual=number(row['actual']); planned=number(row['quantity'])
        if actual==planned: group,description,volume='unchanged','Без изменений',actual
        elif actual>planned: group,description,volume='added','Объём увеличен',actual-planned
        else: group,description,volume='minus',('Материал исключён' if position_type(row)=='material' else 'Работа исключена') if actual==0 else 'Объём уменьшен',actual-planned
        groups[group].append(dict(row,description=description,volume=str(volume),is_new=False))
    for row in result['new']:
        if row.get('needs_review'): continue
        if number(row['actual'])==0: continue
        material=row.get('position_type')=='material'
        negative=number(row['actual'])<0
        description=('Отрицательный объём материала' if material else 'Отрицательный объём работы') if negative else ('Новый материал' if material else 'Новая работа')
        review_rows=row.get('review_rows',[])
        groups['minus' if negative else 'added'].append(dict(row,number='',quantity=None,type=('Новый материал' if material else 'Новая работа') if review_rows else description,description='Требует уточнения' if review_rows else description,volume=row['actual'],is_new=True,provenance=row['sources']))
    return groups
