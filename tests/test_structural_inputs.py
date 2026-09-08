"""Run accepted and rejected input shapes through the real CLI and compare values."""
from pathlib import Path
import argparse,copy,hashlib,io,json,os,re,subprocess,tempfile,zipfile

parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1])
root=parser.parse_args().root.resolve()
readme=(root/'README.md').read_text(encoding='utf-8-sig')
base=json.loads(re.findall(r'```json\n(.*?)\n```',readme,re.S)[0])
evidence=root/'tests/results'
evidence.mkdir(exist_ok=True)
results=[]

def check(ok,detail):
    if not ok: raise AssertionError(detail)

def passed(name,detail):
    results.append(dict(name=name,status='PASS',detail=detail))
    print('PASS',name,detail,flush=True)

with tempfile.TemporaryDirectory(prefix='rdv-structural-') as scratch:
    folder=Path(scratch)
    def run(name,raw,cfg=None,expected=0,validate=False,extras=None):
        case=folder/name
        case.mkdir()
        config=copy.deepcopy(base if cfg is None else cfg)
        config['paths']={'dataDir':str(case),'ledger':str(case/'ledger.xlsx'),'log':str(case/'operation.log')}
        path=case/'settings.json'
        path.write_text(json.dumps(config,ensure_ascii=False),encoding='utf-8')
        inputs={config['data']['tables']['A']['file']:raw}
        inputs.update(extras or {})
        for file,value in inputs.items():
            (case/file).write_bytes(value.encode('utf-8') if isinstance(value,str) else value)
        protected={p:hashlib.sha256(p.read_bytes()).hexdigest() for p in [path]+[case/f for f in inputs]}
        output=case/'result.json'
        command=['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(root/'src/ReaderDataViewer.ps1'),
                 '-ValidateOnly' if validate else '-RunUpdate','-Config',str(path),'-DataDir',str(case)]
        if not validate: command+=['-Output',str(output)]
        execution=subprocess.run(command,capture_output=True,timeout=40)
        text=(execution.stdout+execution.stderr).decode('utf-8')
        (evidence/('structural-'+name+'.log')).write_text(text,encoding='utf-8')
        check(execution.returncode==expected,name+': unexpected exit '+str(execution.returncode)+'\n'+text)
        check(all(hashlib.sha256(p.read_bytes()).hexdigest()==digest for p,digest in protected.items()),name+': source changed')
        check(not (case/'ledger.xlsx').exists(),name+': configured ledger was written')
        log=(Path(os.environ['LOCALAPPDATA'])/'ReaderDataViewer/logs/feedback.log').read_text(encoding='utf-8')
        check(str(path) in log,name+': invocation missing in feedback log')
        if expected or validate:
            check(not output.exists(),name+': unexpected output')
            return text
        report=json.loads(output.read_text(encoding='utf-8'))
        (evidence/('structural-'+name+'.json')).write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        for warning in report['warnings']:
            check(warning in text and warning in log,name+': warning lost in console/file log')
        return report

    raw='id,name\n001,one\n002\n\n003,three\n,discard\n001,one\n'
    report=run('short-fast',raw)
    check(report['rows']==[['001','one'],['003','three']],'short records shifted values')
    for key in ('skippedShort','skippedBlank','skippedEmpty','skippedDuplicate'):
        check(report['summary'][key]==1 and report['inputs'][0][key]==1,'wrong '+key)
    passed('short-fast','2 kept; short/blank/empty-key/duplicate each 1; warnings in file log')
    text=run('short-validate',raw,validate=True)
    check('skippedShort=1' in text and 'skippedBlank=1' in text and 'WARNING' in text,'validation lost counts')
    passed('short-validate','exit 0; matching counters and warning')
    variants=[('quoted','id,name\n001,"one\nline"\n002\n\n003,three\n','utf-8','one\nline'),
              ('utf16','id,name\n001,one\n002\n\n003,three\n','utf-16-le','one'),
              ('cr-only','id,name\r001,one\r002\r\r003,three\r','utf-8','one')]
    for name,text,codec,value in variants:
        cfg=copy.deepcopy(base)
        if codec=='utf-16-le': cfg['data']['encoding']='utf-16'
        report=run('short-'+name,text.encode(codec),cfg)
        check(report['rows']==[['001',value],['003','three']],'decoded CSV values')
        check(report['summary']['skippedShort']==1 and report['summary']['skippedBlank']==1,'decoded CSV counts')
        passed('short-'+name,'2 kept; 1 short; 1 blank; source values retained')
    report=run('all-short','id,name\n001\n002\n')
    check(report['rows']==[] and report['summary']['skippedShort']==2,'all short')
    check(any('除外後' in w for w in report['warnings']),'all-short reported header-only')
    passed('all-short','0 kept; 2 short; explicit empty-result warning')
    report=run('empty-final-cell','id,name\n001,\n')
    check(report['rows']==[['001','']] and report['summary']['skippedShort']==0,'valid empty cell skipped')
    passed('empty-final-cell','trailing delimiter remains an empty cell')

    duplicate='unused,name,id,unused\nleft,one,001,right\nshort,two,002\nleft,three,003,right\n'
    for name,raw in [('plain',duplicate),('quoted',duplicate.replace('unused,name','"unused",name'))]:
        report=run('unused-'+name,raw)
        check(report['rows']==[['001','one'],['003','three']],'projection shifted columns/key')
        check(report['values']['A']['columns']==['A.name','A.id'],'unused columns remain')
        check(report['summary']['skippedColumns']==2 and report['summary']['skippedShort']==1,'pre-projection shape check')
        passed('unused-'+name,'both duplicate columns removed; original width checked; id remains aligned')
    text=run('unused-validate',duplicate,validate=True)
    check('skippedColumns=2' in text and 'WARNING' in text,'validation did not exclude duplicate group')
    passed('unused-validate','exit 0; 2 excluded columns reported')
    report=run('trimmed-header','unused,name,id, unused ,unused\nL,one,001,R,Z\n')
    check(report['summary']['skippedColumns']==3 and report['rows']==[['001','one']],'trimmed group')
    passed('trimmed-header','all 3 same-name columns counted')

    def xlsx(head,values=None):
        data=io.BytesIO()
        with zipfile.ZipFile(data,'w') as book:
            book.writestr('xl/workbook.xml','<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>')
            book.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
            rows=''
            for number,values in [(5,head)]+([] if values is None else [(9,values)]):
                cells=''.join('<c r="'+chr(65+i)+str(number)+'" t="inlineStr"><is><t>'+v+'</t></is></c>' for i,v in enumerate(values))
                rows+='<row r="'+str(number)+'">'+cells+'</row>'
            book.writestr('xl/worksheets/sheet1.xml','<worksheet><sheetData>'+rows+'</sheetData></worksheet>')
        return data.getvalue()
    cfg=copy.deepcopy(base)
    cfg['data']['tables']['A']['file']='rows.xlsx'
    report=run('unused-xlsx',xlsx(['unused','name','id','unused'],['left','one','001']),cfg)
    check(report['rows']==[['001','one']] and report['summary']['skippedColumns']==2,'XLSX projection')
    check(report['summary']['skippedShort']==0,'XLSX omitted cell treated as short CSV')
    passed('unused-xlsx','2 ignored columns; XLSX omitted tail cell retained as empty')
    text=run('used-xlsx',xlsx(['id','name','name']),cfg,expected=3)
    check('5 行目、列 3' in text and 'Duplicate headers used' in text,'XLSX ambiguity diagnostic')
    passed('used-xlsx','exit 3; actual header row 5 and column 3')

    protected_raw='id,name,unused,unused\n001,one,9,\n'
    for kind in ('key','type','ledger','expression','where','select','groupBy','aggregate','order','set','later-job'):
        cfg=copy.deepcopy(base)
        cfg['data']['labels'].update({'A.unused':'Unused','subset':'Subset','sums':'Sums','sums.total':'Total','A.derived':'Derived'})
        job=cfg['data']['jobs'][0]
        step=None
        if kind=='key': cfg['data']['tables']['A']['key']='unused'
        elif kind=='type': cfg['data']['types']={'A.unused':{'type':'text'}}
        elif kind=='ledger': cfg['data']['ledger']['columns']['source'].append('A.unused')
        elif kind in ('expression','later-job'):
            step={'operation':'calculate','target1':'A','column':'derived','expression':'A.unused + 1','output':'A'}
        elif kind in ('where','set'):
            step={'operation':'extract','target1':'A','where':{'column':'A.unused' if kind=='where' else 'A.id','operator':'notEmpty'},'output':'subset'}
        elif kind=='select': step={'operation':'select','target1':'A','columns':['A.id','A.name','A.unused'],'output':'A'}
        elif kind in ('groupBy','aggregate'):
            step={'operation':'aggregate','target1':'A','groupBy':['A.unused' if kind=='groupBy' else 'A.id'],
                  'aggregates':[{'function':'sum','column':'A.unused','as':'total'}],'output':'sums'}
        elif kind=='order': step={'operation':'sort','target1':'A','orderBy':[{'column':'A.unused','direction':'ascending'}],'output':'A'}
        if kind=='later-job':
            later=copy.deepcopy(job)
            later['id']='later'
            later['steps'].insert(0,step)
            cfg['data']['jobs'].append(later)
        elif step: job['steps'].insert(0,step)
        if kind=='set':
            job['steps'].insert(1,{'operation':'update','target1':'A','target2':'subset','set':[{'column':'A.unused','expression':'1'}],'output':'A'})
        text=run('reference-'+kind,protected_raw,cfg,expected=3)
        check('1 行目、列 4' in text and 'Duplicate headers used' in text,kind+': did not reject at ambiguous source header')
        passed('reference-'+kind,'exit 3; referenced duplicate rejected even with empty second column')

    cfg=copy.deepcopy(base)
    cfg['data']['labels'].update({'A.unused':'Unused'})
    cfg['data']['jobs'][0]['steps'].insert(0,{'operation':'calculate','target1':'A','column':'name','expression':"'A.unused'",'output':'A'})
    report=run('label-and-literal',protected_raw,cfg)
    check(report['rows']==[['001','A.unused']] and report['summary']['skippedColumns']==2,'label/literal mistaken for source use')
    passed('label-and-literal','unused label and quoted expression text do not reference a column')

    for name,raw in [('extra','id,name\n001,one,extra\n'),('quote','id,name\n001,"open'),('empty-head','id,\n001,one\n'),('encoding',b'id,name\n001,\xff\n')]:
        run('reject-'+name,raw,expected=3)
        passed('reject-'+name,'exit 3; malformed input remains refused')
    text=run('physical-row','id,name\n001,1\nshort\n\n003,bad\n',
             dict(base,data=dict(base['data'],types={'A.name':{'type':'number'}})),expected=3)
    check('5 行目' in text and 'bad' in text,'source row numbers shifted after skips')
    passed('physical-row','type error retains source row 5 after short/blank skips')

    cfg=copy.deepcopy(base)
    cfg['data']['labels'].update({f'A.{c}':c for c in ('quantity','completed','unitValue','note','extra')})
    cfg['data']['ledger']['columns']['source']+=['A.note','A.extra']
    cfg['data']['types']={f'A.{c}':{'type':'number'} for c in ('quantity','completed','unitValue')}
    section=readme.split('<a id="calculation-storage"></a>')[1].split('<a id="join-grain"></a>')[0]
    steps=json.loads(re.findall(r'```json\n(.*?)\n```',section,re.S)[0])
    cfg['data']['jobs'][0]['steps']=steps+cfg['data']['jobs'][0]['steps']
    report=run('readme-calculation','id,name,quantity,completed,unitValue,note,extra\n001,one,10,4,25,old note,old extra\n',cfg)
    check(report['rows']==[['001','one','6','250']],'README calculation did not reach saved rows')
    passed('readme-calculation','published steps produce saved values 6 and 250; source remains intact')

(evidence/'structural-input-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print(f'TOTAL {len(results)}/{len(results)}')
