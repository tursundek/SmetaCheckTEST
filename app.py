"""Local-only HTTP API. Run: python app.py"""
import base64, json, argparse, threading, webbrowser, socket
import os, hashlib
import importing as importer_module
import core as core_module
import matching as matching_module
import materials as materials_module
import exporting as exporting_module
import sections as sections_module
import rik_codes as rik_module
import result_positions as positions_module
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from storage import ROOT, DATA, create, get, save, listing, audit, delete_project, restore_project, deleted_listing
from importing import upload, import_rows, suggest_config, preview, IMPORT_VERSION, sheet_info, sheet_page, extraction_preview, material_project
from core import reconcile, validate_decision, FIELDS
from exporting import export
from rik_codes import add_rik_codes

LOCK=threading.RLock()
LOADED_MODULES={str(Path(path).resolve()):hashlib.sha256(Path(path).read_bytes()).hexdigest()
                for path in [__file__,importer_module.__file__,core_module.__file__,matching_module.__file__,materials_module.__file__,exporting_module.__file__,sections_module.__file__,rik_module.__file__,positions_module.__file__]}

def public_project(p):
    p=material_project(p)
    # Sheet content is requested in pages instead of sending every sheet at once.
    return {**p,'documents':[{**d,'sheet_info':sheet_info(d),'sheets':{name:[] for name in d.get('sheets',{})} if d.get('extension') in ['.xlsx','.xls'] else d.get('sheets',{}),
                             'import_report':{'counts':d.get('import_report',{}).get('counts',{})}}
                            if d.get('extension') in ['.xlsx','.xls'] and 'sheets' in d else d for d in p['documents']]}

def structure(config):
    return [(t['sheet'],{f:t['mapping'].get(f) for f in FIELDS},int(t.get('first_row',int(t.get('header',1))+1)),t.get('last_row') or None,
             t.get('mode',config.get('mode','current')),t.get('method',config.get('method','differences')))
            for t in config.get('tables',[config]) if t.get('enabled',True)]

class LocalServer(ThreadingHTTPServer):
    def server_bind(self):
        # Windows otherwise permits two HTTP servers to share the same port,
        # allowing an old process to keep serving stale calculations.
        if hasattr(socket,'SO_EXCLUSIVEADDRUSE'):
            self.allow_reuse_address=False
            self.socket.setsockopt(socket.SOL_SOCKET,socket.SO_EXCLUSIVEADDRUSE,1)
        super().server_bind()

def command(body):
    action=body['action']
    if action=='create': return create(body.get('name',''))
    if action=='delete_project':
        if body.get('confirmed') is not True: raise ValueError('Подтвердите удаление проекта')
        return delete_project(body['project'])
    if action=='restore_project': return restore_project(body['project'])
    p=material_project(get(body['project']))
    if action=='upload':
        doc=upload(p,body['name'],base64.b64decode(body['content'],validate=True),body['role'])
        audit(p,'Загрузка файла',{'file':doc['name'],'hash':doc['hash'],'warnings':doc['warnings']})
        if doc['suggestion']:
            config=doc['suggestion']
            for table in config.get('tables',[config]):
                template=p.get('import_templates',{}).get(table.get('layout_key',''))
                if template and (table.get('needs_choice') or table.get('unmapped')):
                    table['mapping']={**table['mapping'],**template['mapping']};table['mode']=template['mode'];table['needs_choice']=False
                    # An explicit header found in this file keeps its own row,
                    # even when the template resolves an uncertain volume column.
                    if 'header' in template and (table.get('unmapped') or not table.get('header')):
                        table['header']=template['header']
                    table['unmapped']=False
            tables=config.get('tables',[config])
            config['needs_choice']=any(t.get('needs_choice') for t in tables) or len({t['mode'] for t in tables})>1 or config.get('manual_required',False)
            if config.get('tables'):
                config['mapping']=dict(config['tables'][0]['mapping']);config['mode']=config['tables'][0]['mode']
        if doc['extension'] in ['.xlsx','.xls','.docx'] and doc['suggestion'] and not doc['suggestion'].get('needs_choice'):
            import_rows(doc,doc['suggestion'])
            if doc['role']=='tender' and any(d['id']!=doc['id'] and d['role']=='tender' and d.get('imported') and d.get('active') for d in p['documents']):
                doc['active']=False
                doc['warnings'].append('Уже есть исходная смета. Включите этот файл, если хотите заменить её.')
            audit(p,'Автоматически определены столбцы и импортированы строки',{'file':doc['name'],'config':doc['config']})
    elif action=='settings':
        p['settings']={'aliases':bool(body.get('aliases')),'numbering':bool(body.get('numbering'))}
        audit(p,'Правила единиц и нумерации',p['settings'])
    elif action in ['import','active','replace_tender','decision','undo','auto','edit','repair_import','confirm_negative']:
        doc=next(d for d in p['documents'] if d['id']==body['document'])
        if doc['role']=='tender' and action in ['decision','undo','auto','edit','confirm_negative']:
            raise ValueError('Тендерная смета — исходная база. Изменение её строк запрещено')
        if doc['role']=='tender' and action=='import' and any(t.get('overrides') for t in [body['config']]+body['config'].get('tables',[])):
            raise ValueError('Можно выбрать столбцы тендера, но нельзя менять исходные значения его строк')
        if action=='replace_tender':
            if doc['role']!='tender' or not doc.get('imported'): raise ValueError('Сначала выберите столбцы исходной сметы')
            for other in p['documents']:
                if other['role']=='tender': other['active']=other['id']==doc['id']
            audit(p,'Выбрана исходная смета',doc['name'])
        elif action=='repair_import':
            config=suggest_config(doc)
            if not config: raise ValueError('Не удалось определить колонки: выберите их в настройках файла')
            old_config=doc.get('config',{})
            config.update({k:v for k,v in old_config.items() if k not in ['sheet','header','mapping','tables','import_version']})
            previous={r['id']:r for r in doc['rows']}
            import_rows(doc,config)
            doc['suggestion']=suggest_config(doc)
            for row in doc['rows']:
                old=previous.get(row['id'],{})
                for key in ['decision','history']:
                    if key in old: row[key]=old[key]
            audit(p,'Исправлена шапка и колонки импорта',{'file':doc['name'],'before':old_config,'after':config})
        elif action=='confirm_negative':
            row=next(r for r in doc['rows'] if r['id']==body['row'])
            from core import number
            if row['quantity'] is None or number(row['quantity'])>=0: raise ValueError('Это не отрицательный объём')
            config=doc['config'];key=row.get('import_key',row['id'].split(':')[-1])
            config.setdefault('negative_confirmed_rows',[])
            if key not in config['negative_confirmed_rows']: config['negative_confirmed_rows'].append(key)
            row['errors']=[e for e in row['errors'] if 'Отрицательный' not in e]
            audit(p,'Подтверждён отрицательный объём',{'file':doc['name'],'source':row['source'],'quantity':row['quantity']})
        elif action=='import':
            if doc['active'] and doc['role']=='tender' and any(d['role']=='tender' and d['imported'] and d['active'] and d['id']!=doc['id'] for d in p['documents']): raise ValueError('Сначала исключите предыдущую тендерную смету')
            unchanged=doc.get('imported') and structure(doc['config'])==structure(body['config'])
            previous={r['id']:r for r in doc['rows']} if unchanged else {}
            has_decisions=any(r.get('decision') or r.get('history') for r in doc['rows'])
            if doc.get('imported') and not unchanged and has_decisions and not body.get('reset_confirmed'):
                raise ValueError('Новые настройки требуют перечитывания и сброса ручных решений этого файла. Подтвердите сброс перед применением.')
            import_rows(doc,body['config'])
            for row in doc['rows']:
                for key in ['decision','history']:
                    if key in previous.get(row['id'],{}): row[key]=previous[row['id']][key]
            for table in doc['config'].get('tables',[doc['config']]):
                if not table.get('enabled',True): continue
                if table.get('layout_key'):
                    p.setdefault('import_templates',{})[table['layout_key']]=dict(mapping=table['mapping'],mode=table.get('mode','current'),header=table.get('header',1))
            for other in p['documents']:
                if other['id']!=doc['id'] and doc.get('act_number') and doc.get('date') and other.get('act_number')==doc['act_number'] and other.get('date')==doc['date']:
                    warning='Конфликт номера и даты акта: подтвердите включение'; doc['warnings'].append(warning); doc['active']=False
            audit(p,'Импорт / изменение настроек',{'file':doc['name'],'config':body['config']})
        elif action=='active':
            if body['active'] and any(d['id']!=doc['id'] and d['role']==doc['role'] and d.get('active') and d.get('hash')==doc.get('hash') for d in p['documents']):
                raise ValueError('Этот файл уже учитывается. Выключите прежнюю копию перед включением новой')
            if body['active'] and doc['role']=='tender' and any(d['id']!=doc['id'] and d['role']=='tender' and d['active'] and d['imported'] for d in p['documents']): raise ValueError('Допустима одна активная тендерная смета')
            doc['active']=bool(body['active']); audit(p,'Включение / исключение документа',{'file':doc['name'],'active':doc['active']})
        else:
            row=next(r for r in doc['rows'] if r['id']==body['row'])
            history=row.setdefault('history',[])
            if action=='decision':
                tender=next((d['rows'] for d in p['documents'] if d['role']=='tender' and d['active'] and d['imported']),[])
                validate_decision(row,body['decision'],tender,p['settings'])
                history.append(row.get('decision')); row['decision']=body['decision']
            elif action=='undo':
                if not history: raise ValueError('Нет действия для отмены')
                previous=history.pop()
                if previous is None: row.pop('decision',None)
                else: row['decision']=previous
            elif action=='auto': history.append(row.get('decision')); row.pop('decision',None)
            elif action=='edit':
                config=doc['config']; config.setdefault('overrides',{})[row.get('import_key',row['id'].split(':')[-1])]=body['values']
                old={r['id']:r for r in doc['rows']}; import_rows(doc,config)
                for new in doc['rows']:
                    if new['id']!=row['id'] and new['id'] in old:
                        for key in ['decision','history']:
                            if key in old[new['id']]: new[key]=old[new['id']][key]
            audit(p,action,{'source':row['source'],'before':row.get('raw'),'details':body})
    else: raise ValueError('Неизвестное действие')
    result=reconcile(p)
    analysis=[{'row':m['row']['id'],'status':m['match']['status'],'candidates':m['match'].get('candidates',[]),'allocations':m['match'].get('allocations',[])} for m in result['matches']]
    if p.get('last_analysis')!=analysis:
        audit(p,'Результат сопоставления / сработавшие правила',analysis)
        p['last_analysis']=analysis
    save(p); return p

class Handler(BaseHTTPRequestHandler):
    def send(self,status,data,ctype='application/json; charset=utf-8'):
        if not isinstance(data,bytes): data=json.dumps(data,ensure_ascii=False).encode()
        self.send_response(status); self.send_header('Content-Type',ctype); self.send_header('Content-Length',str(len(data))); self.send_header('X-Content-Type-Options','nosniff'); self.send_header('Cache-Control','no-store'); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        try:
            u=urlparse(self.path); q=parse_qs(u.query)
            with LOCK:
                if u.path=='/api/runtime':
                    return self.send(200,dict(pid=os.getpid(),root=str(ROOT),data=str(DATA),
                                              import_version=IMPORT_VERSION,loaded_modules=LOADED_MODULES))
                if u.path=='/api/projects': return self.send(200,listing())
                if u.path=='/api/deleted_projects': return self.send(200,deleted_listing())
                if u.path=='/api/project':
                    p=get(q['id'][0])
                    # Reading a saved project never reimports or saves its documents.
                    return self.send(200,{'project':public_project(p),'result':add_rik_codes(reconcile(p))})
                if u.path=='/api/sheet':
                    p=get(q['project'][0]);doc=next(d for d in p['documents'] if d['id']==q['document'][0])
                    return self.send(200,sheet_page(doc,q['sheet'][0],q.get('start',['1'])[0],q.get('limit',['100'])[0]))
                if u.path=='/api/export':
                    p=get(q['id'][0]); r=reconcile(p)
                    if not r['ready']: raise ValueError('Загрузите смету и процентовки и выберите объёмы нераспознанных файлов')
                    if r['issues'] and q.get('ack',['0'])[0]!='1': raise ValueError('Подтвердите экспорт с оставшимися предупреждениями')
                    audit(p,'Экспорт Excel',{'acts':[a['name'] for a in r['acts']],'warnings':r['issues']}); save(p)
                    return self.send(200,export(p),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                if u.path in ['/','/index.html']: return self.send(200,(ROOT/'web'/'index.html').read_bytes(),'text/html; charset=utf-8')
                if u.path=='/simple.js': return self.send(200,(ROOT/'web'/'simple.js').read_bytes(),'text/javascript; charset=utf-8')
                if u.path=='/review.js': return self.send(200,(ROOT/'web'/'review.js').read_bytes(),'text/javascript; charset=utf-8')
                if u.path=='/excel-import.js': return self.send(200,(ROOT/'web'/'excel-import.js').read_bytes(),'text/javascript; charset=utf-8')
                if u.path=='/waiting.js': return self.send(200,(ROOT/'web'/'waiting.js').read_bytes(),'text/javascript; charset=utf-8')
                return self.send(404,{'error':'Не найдено'})
        except Exception as e: self.send(400,{'error':str(e)})
    def do_POST(self):
        try:
            if self.path not in ['/api/command','/api/import-preview']: return self.send(404,{'error':'Не найдено'})
            if self.headers.get('Origin') and self.headers['Origin']!=f"http://{self.headers.get('Host')}": raise ValueError('Недопустимый источник запроса')
            length=int(self.headers.get('Content-Length',0))
            if length>40000000: raise ValueError('Максимальный размер запроса 40 МБ')
            body=json.loads(self.rfile.read(length))
            if self.path=='/api/import-preview':
                with LOCK:
                    p=get(body['project']);doc=next(d for d in p['documents'] if d['id']==body['document'])
                    report=extraction_preview(doc,body['config'])
                # Paginate both diagnostics and positions; preserve all counts.
                offset=max(0,int(body.get('offset',0)));limit=max(1,min(200,int(body.get('limit',100))))
                return self.send(200,{**{k:v[offset:offset+limit] for k,v in report.items() if k!='counts'},'counts':report['counts']})
            with LOCK: p=command(body)
            if 'deleted_project' in p: return self.send(200,p)
            self.send(200,{'project':public_project(p),'result':add_rik_codes(reconcile(p))})
        except Exception as e: self.send(400,{'error':str(e)})

if __name__=='__main__':
    parser=argparse.ArgumentParser(); parser.add_argument('--port',type=int,default=8765); parser.add_argument('--no-browser',action='store_true'); args=parser.parse_args()
    url=f'http://127.0.0.1:{args.port}'
    server=LocalServer(('127.0.0.1',args.port),Handler)
    print('Сверка объёмов: '+url,flush=True)
    if not args.no_browser: webbrowser.open(url)
    try: server.serve_forever()
    except KeyboardInterrupt: server.server_close()
