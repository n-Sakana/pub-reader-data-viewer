"""Reproduce unsupported duplicate headers and supported explicit zero-normalization rules."""
from pathlib import Path
import copy,json,os,re,subprocess,tempfile,zipfile

root=Path(__file__).resolve().parents[1]
readme=(root/'README.md').read_text(encoding='utf-8-sig')
minimal=json.loads(re.findall(r'```json\n(.*?)\n```',readme,re.S)[0])
results=[]
evidence=root/'tests/results'
def check(ok,message):
    if not ok: raise AssertionError(message)
def passed(name,detail):
    results.append(dict(name=name,status='PASS',detail=detail))
    print('PASS',name,detail,flush=True)

with tempfile.TemporaryDirectory(prefix='rdv-input-boundaries-') as temporary:
    folder=Path(temporary)
    def execute(name,cfg,files,expected,mode='-RunUpdate'):
        case=folder/name
        data=case/'data'
        data.mkdir(parents=True)
        cfg=copy.deepcopy(cfg)
        cfg['paths']={'dataDir':str(data),'ledger':str(case/'ledger.xlsx'),'log':str(case/'operation.log')}
        config=case/'settings.json'
        config.write_text(json.dumps(cfg,ensure_ascii=False,indent=2),encoding='utf-8')
        for name,text in files.items():
            if isinstance(text,bytes): (data/name).write_bytes(text)
            else: (data/name).write_text(text,encoding='utf-8',newline='\r\n')
        output=case/'report.json'
        command=['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(root/'src/ReaderDataViewer.ps1'),
                 mode,'-Config',str(config),'-DataDir',str(data)]
        if mode=='-RunUpdate': command+=['-Output',str(output)]
        result=subprocess.run(command,capture_output=True,timeout=35)
        (evidence/('boundary-'+case.name+'.log')).write_bytes(result.stdout+result.stderr)
        check(result.returncode==expected,case.name+': unexpected exit '+str(result.returncode))
        if expected:
            log=Path(os.environ['LOCALAPPDATA'])/'ReaderDataViewer/logs/feedback.log'
            text=log.read_text(encoding='utf-8')
            check(str(config) in text, 'error not in file log')
            return result.stderr.decode('utf-8')
        report=json.loads(output.read_text(encoding='utf-8'))
        (evidence/('boundary-'+case.name+'.json')).write_text(json.dumps(report,ensure_ascii=False),encoding='utf-8')
        return report

    for name,head,mode in [('csv-fast','id,name,name','-RunUpdate'),('csv-quoted','id,name,"name"','-RunUpdate'),('csv-validate','id,name,name','-ValidateOnly')]:
        msg=execute(name,minimal,{'rows.csv':head+'\n001,left,right\n'},3,mode)
        check('1 行目、列 3' in msg and 'Remove duplicate columns only if they are unused' in msg and 'JSON labels or select cannot' in msg,'missing header location/remedy')
        passed(name,'exit 3; row 1 column 3; correction and unsupported JSON workaround stated')

    workbook=folder/'duplicate.xlsx'
    with zipfile.ZipFile(workbook,'w') as book:
        book.writestr('xl/workbook.xml','<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
        book.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
        cells=''.join('<c r="'+letter+'5" t="inlineStr"><is><t>'+name+'</t></is></c>' for letter,name in zip('ABC',['id','name','name']))
        book.writestr('xl/worksheets/sheet1.xml','<worksheet><sheetData><row r="5">'+cells+'</row></sheetData></worksheet>')
    cfg=copy.deepcopy(minimal)
    cfg['data']['tables']['A']['file']='duplicate.xlsx'
    msg=execute('xlsx-header-row',cfg,{'duplicate.xlsx':workbook.read_bytes()},3)
    check('5 行目、列 3' in msg and 'Remove duplicate columns' in msg,'XLSX actual header position/remedy')
    passed('xlsx-header-row','exit 3; actual header row 5 column 3; same correction')

    cfg=copy.deepcopy(minimal)
    cfg['data']['tables']['B']={'file':'B.csv','key':'number','keyValidation':{'length':'variable'}}
    cfg['data']['labels'].update({'B.number':'番号B','B.value':'値B','A.matchId':'照合A','B.matchId':'照合B','joined':'照合結果'})
    cfg['data']['ledger']['columns']['source'].append('B.value')
    job=cfg['data']['jobs'][0]
    job['inputs'].append({'table':'B'})
    merge=job['steps'][-1]
    merge['target1']='joined'
    job['steps']=[{'operation':'calculate','target1':t,'column':'matchId','expression':"regexExtract("+col+", '^[0-9]{1,8}$') + 0",'output':t} for t,col in [('A','A.id'),('B','B.number')]]+[
        {'operation':'join','target1':'A','target2':'B','keys':['A.matchId','B.matchId'],'condition':'left','output':'joined'},merge]
    files={'rows.csv':'id,name\n00000001,one\n00000023,two\n00000099,missing\n','B.csv':'number,value\n1,hit1\n23,hit23\n'}
    report=execute('normalized-key',cfg,files,0)
    check(report['rows']==[['00000001','one','hit1'],['00000023','two','hit23'],['00000099','missing','']],'normalization changed identity or join')
    check(report['joins'][0]['unmatchedLeft']==1,'unmatched count')
    passed('normalized-key','2 matches/1 missing; original padded identities preserved')

    cfg=copy.deepcopy(minimal)
    snippet=re.findall(r'```json\n(.*?)\n```',readme.split('**元の番号は必ず8桁')[1],re.S)[0]
    cfg['data']['jobs'][0]['steps'].insert(0,json.loads(snippet))
    cfg['data']['labels']['A.paddedId']='8桁'
    report=execute('padding-expression',cfg,{'rows.csv':'id,name\n1,one\n23,two\n00000000,zero\n99999999,max\n'},0)
    check([r[-1] for r in report['values']['A']['rows']]==['00000001','00000023','00000000','99999999'],'padding results')
    passed('padding-expression','README expression: 1/23/zero/max padded to 8 digits')
    msg=execute('padding-overflow',cfg,{'rows.csv':'id,name\n123456789,too-long\n'},3)
    check('regexExtract found no match' in msg,'9 digits were silently truncated')
    passed('padding-overflow','exit 3; 9 digits rejected before truncation')

(evidence/'input-boundaries-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print(f'TOTAL {len(results)}/{len(results)}')
