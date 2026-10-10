import json, sqlite3, uuid
from contextlib import contextmanager
from pathlib import Path
from datetime import datetime, timezone

ROOT=Path(__file__).resolve().parent
DATA=ROOT/'data'

@contextmanager
def connection():
    DATA.mkdir(exist_ok=True)
    db=sqlite3.connect(DATA/'projects.sqlite3')
    db.execute('CREATE TABLE IF NOT EXISTS projects (id TEXT PRIMARY KEY, body TEXT NOT NULL)')
    db.execute('CREATE TABLE IF NOT EXISTS deleted_projects (id TEXT PRIMARY KEY, body TEXT NOT NULL)')
    try:
        with db: yield db
    finally: db.close()

def uid(): return uuid.uuid4().hex

def audit(p,action,details):
    p['log'].append({'time':datetime.now(timezone.utc).isoformat(),'user':'Локальный пользователь','action':action,'details':details})

def save(p):
    with connection() as db: db.execute('INSERT OR REPLACE INTO projects VALUES (?,?)',(p['id'],json.dumps(p,ensure_ascii=False)))

def get(pid):
    with connection() as db: row=db.execute('SELECT body FROM projects WHERE id=?',(pid,)).fetchone()
    if not row: raise ValueError('Проект не найден')
    return json.loads(row[0])

def listing():
    with connection() as db: return [{'id':p['id'],'name':p['name']} for (body,) in db.execute('SELECT body FROM projects') for p in [json.loads(body)]]

def create(name):
    p={'id':uid(),'name':name or 'Новая сверка','settings':{'aliases':True,'numbering':False},'documents':[],'log':[]}
    audit(p,'Создание проекта',p['name']); save(p); return p

def delete_project(pid):
    with connection() as db:
        row=db.execute('SELECT body FROM projects WHERE id=?',(pid,)).fetchone()
        if not row: raise ValueError('Проект не найден')
        p=json.loads(row[0]);audit(p,'Удаление проекта',{'can_restore':True})
        db.execute('INSERT OR REPLACE INTO deleted_projects VALUES (?,?)',(pid,json.dumps(p,ensure_ascii=False)))
        db.execute('DELETE FROM projects WHERE id=?',(pid,))
    return {'deleted_project':pid}

def deleted_listing():
    with connection() as db:
        return [{'id':p['id'],'name':p['name']} for (body,) in db.execute('SELECT body FROM deleted_projects') for p in [json.loads(body)]]

def restore_project(pid):
    with connection() as db:
        row=db.execute('SELECT body FROM deleted_projects WHERE id=?',(pid,)).fetchone()
        if not row: raise ValueError('Удалённый проект не найден')
        if db.execute('SELECT 1 FROM projects WHERE id=?',(pid,)).fetchone(): raise ValueError('Проект уже существует')
        p=json.loads(row[0]);audit(p,'Восстановление проекта',{})
        db.execute('INSERT INTO projects VALUES (?,?)',(pid,json.dumps(p,ensure_ascii=False)))
        db.execute('DELETE FROM deleted_projects WHERE id=?',(pid,))
    return p
