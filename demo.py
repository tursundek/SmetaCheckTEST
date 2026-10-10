"""Import demo workbooks through the actual application pipeline and export."""
from pathlib import Path
from storage import create, audit, save
from importing import upload, import_rows
from core import reconcile
from exporting import export

def build():
    folder=Path(__file__).parent/'demo'
    p=create('Демонстрация — тендер и три акта')
    for i,name in enumerate(['Тендер','Акт 1','Акт 2','Акт 3']):
        d=upload(p,name+'.xlsx',(folder/(name+'.xlsx')).read_bytes(),'tender' if i==0 else 'act')
        import_rows(d,dict(sheet='Работы',header=3,mapping=dict(number=0,code=1,name=2,unit=3,quantity=4,section=5),mode='current',act_number=str(i) if i else '',date=f'2026-09-{i*5:02}' if i else ''))
        audit(p,'Импорт демонстрационного файла',name)
        for row in d['rows']:
            if row['name']=='Гидроизоляция обмазочная': row['decision']={'status':'new','comment':'Подтверждённая новая работа'};audit(p,'Ручное подтверждение новой работы',row['id'])
    r=reconcile(p)
    assert r['summary'][0]['actual']=='85'
    assert r['summary'][0]['minus']=='15'
    assert r['summary'][1]['actual']=='12'
    assert r['summary'][1]['added']=='2'
    assert r['new'][0]['actual']=='250'
    assert r['matches'][-1]['match']['status']=='ambiguous'
    audit(p,'Экспорт предварительного результата',{'warnings':r['issues']})
    save(p);(folder/'Пример сверки.xlsx').write_bytes(export(p))
    print('Demo verified; project:',p['id'])
    return p

if __name__=='__main__': build()
