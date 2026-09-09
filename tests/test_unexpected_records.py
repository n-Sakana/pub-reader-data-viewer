"""Exercise row exclusions and structural stops through both app preparation and headless execution."""
from pathlib import Path
from xml.sax.saxutils import escape
import copy, hashlib, json, subprocess, tempfile, zipfile

ROOT = Path(__file__).resolve().parents[1]
EVIDENCE = ROOT/'tests/results'
CASES = []

def merge(target, identity='A.id'):
    return dict(operation='merge',target1=target,target2='ledger',keys=[identity,identity],sourceOnly='add',both='update',output='ledger')

def workbook(path, rows):
    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    body=''
    for r, values in enumerate(rows,1):
        cells=''
        for c,value in enumerate(values):
            address=chr(65+c)+str(r)
            if isinstance(value,tuple):
                kind,text=value
                cell='<f>'+escape(text)+'</f>' if kind=='formula' else '<v>'+escape(text)+'</v>'
                cells+=f'<c r="{address}" t="{kind if kind != "formula" else "n"}">{cell}</c>'
            else: cells+=f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
        body+=f'<row r="{r}">{cells}</row>'
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>')
        z.writestr('_rels/.rels',f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{ns}" xmlns:r="{rel}"><sheets><sheet name="Input" sheetId="1" r:id="r1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml',f'<worksheet xmlns="{ns}"><sheetData>{body}</sheetData></worksheet>')

def case(root,name,values,steps=None,source=None,identity='A.id',types=None,rule=None,rows=None,expected=None,invalid=0,duplicate=0,empty=0,words=(),stop=False,baseline=None):
    directory=root/name; directory.mkdir()
    source=source or ['A.id','A.name','A.v']
    steps=steps or [merge('A',identity)]
    table={'file':'rows.csv','key':'id','keyValidation':rule if rule is not None else {'characters':'unicode','length':'variable'}}
    labels={'A.id':'ID','A.name':'Name','A.v':'Value','A.other':'Other','ledger':'Ledger'}
    for column in source: labels[column]=column
    for step in steps:
        if step['output']!='ledger': labels[step['output']]=step['output']
    config={'schema':3,'paths':{'dataDir':str(directory),'ledger':str(directory/'shared.xlsx'),'log':str(directory/'events.log')},'watch':{'targets':[]},
        'data':{'encoding':'utf-8','tables':{'A':table},'labels':labels,'types':types or {},
                'jobs':[{'id':'update','kind':'update','inputs':[{'table':'A'}],'steps':steps}],
                'ledger':{'identity':identity,'search':{'columns':[identity]},'columns':{'source':source,'application':[{'name':'workState','onSourceChange':'reset'}]}}},
        'screen':{'workState':{'trigger':'manual','store':{'column':'state'},'states':[{'id':'todo','stored':'FALSE'},{'id':'done','stored':'TRUE'}],'initial':'todo'},
                  'export':{'defaultFields':[identity]},'candidates':{'columns':[{'value':{'field':identity}}]},'sections':[{'type':'titleBar'}]}}
    if rows is not None:
        table['file']='rows.xlsx'; workbook(directory/'rows.xlsx',rows)
    else: (directory/'rows.csv').write_text('id,name,v\n'+'\n'.join(values)+'\n',encoding='utf-8')
    if baseline:
        workbook(directory/'baseline.xlsx',baseline)
    configpath=directory/'settings.json'; configpath.write_text(json.dumps(config,ensure_ascii=False),encoding='utf-8')
    entry=dict(name=name,directory=str(directory),config=str(configpath),report=str(directory/'result.json'),baseline=str(directory/'baseline.xlsx') if baseline else '')
    CASES.append((entry,dict(expected=expected,invalid=invalid,duplicate=duplicate,empty=empty,words=words,stop=stop)))
    return directory,config

with tempfile.TemporaryDirectory(prefix='rdv-unexpected-') as tmp:
    root=Path(tmp)
    for name,expression,values,answers in [
        ('split-missing',"splitPart(A.v, '*', 1)",['p*q','missing','r*s'],['q','s']),
        ('split-empty-result',"splitPart(A.v, '*', 1)",['p*q','p*','r*s'],['q','s']),
        ('substring-range','substring(A.v, 1, 1)',['ab','x','cd'],['b','d']),
        ('substring-empty','substring(A.v, 1, 1)',['ab','','cd'],['b','d']),
        ('regex-missing',"regexExtract(A.v, 'P-[0-9]+')",['P-100','none','P-300'],['P-100','P-300']),
        ('arithmetic-number','A.v * 2',['10','bad','30'],['20','60']),
        ('arithmetic-zero','10 / A.v',['2','0','5'],['5','2']),
        ('arithmetic-overflow','A.v * 2',['10','79228162514264337593543950335','30'],['20','60']),
    ]:
        lines=[f'{i+1},{chr(97+i)},{value}' for i,value in enumerate(values)]
        calc=dict(operation='calculate',target1='A',column='x',expression=expression,output='D')
        case(root,name,lines,[calc,merge('D')],['A.id','A.name','A.v','D.x'],
             expected=[['1','a',values[0],answers[0]],['3','c',values[2],answers[1]]],invalid=1,words=('3 行目','2','A.v',expression.split('(')[0] if '(' in expression else '数値' if name=='arithmetic-number' else '計算'))
    aggregate=dict(operation='aggregate',target1='A',groupBy=['A.name'],aggregates=[dict(function='sum',column='A.v',**{'as':'total'}),dict(function='sum',column='A.other',**{'as':'other'}),dict(function='count',**{'as':'count'})],output='G')
    directory,_=case(root,'aggregate-atomic',[],[aggregate,merge('G','A.name')],['A.name','G.total','G.other','G.count'],identity='A.name',expected=[['G','40','60','2']],invalid=1,words=('3 行目','bad','A.other'))
    (directory/'rows.csv').write_text('id,name,v,other\n1,G,10,20\n2,G,99,bad\n3,G,30,40\n',encoding='utf-8')
    ag=copy.deepcopy(aggregate); ag['groupBy']=[]; ag['aggregates']=[dict(function='sum',column='A.v',**{'as':'total'}),dict(function='count',**{'as':'count'})]
    case(root,'aggregate-all-invalid',['1,a,bad','2,b,bad','3,c,bad'],[ag,merge('G','G.count')],['G.count','G.total'],identity='G.count',expected=[['0','0']],invalid=3,words=('2 行目','3 行目','4 行目'))
    sort=dict(operation='sort',target1='A',orders=[dict(column='A.v',direction='ascending',type='number')],output='S')
    case(root,'numeric-sort',['1,a,10','2,b,bad','3,c,5'],[sort,merge('S')],expected=[['3','c','5'],['1','a','10']],invalid=1,words=('A.v','bad','3 行目'))
    extract=dict(operation='extract',target1='A',where=dict(column='A.v',operator='greater',value='5'),output='S')
    case(root,'numeric-condition',['1,a,10','2,b,bad','3,c,30'],[extract,merge('A')],expected=[['1','a','10'],['2','b','bad'],['3','c','30']],invalid=1,words=('A.v','bad','3 行目'))
    case(root,'typed-cells-one-row',['1,20260909,10','2,wrong,bad','3,20260910,30'],types={'A.name':{'type':'date','format':'yyyyMMdd'},'A.v':{'type':'number'}},expected=[['1','20260909','10'],['3','20260910','30']],invalid=1,words=('wrong','bad','A.name','A.v','3 行目'))
    for name,middle in [('ascii','00あ2,b,20'),('width','02,b,20'),('control','00\x0102,b,20')]:
        case(root,'key-'+name,['0001,a,10',middle,'0003,c,30'],rule={},expected=[['0001','a','10'],['0003','c','30']],invalid=1,words=('id','3 行目'))
    case(root,'empty-key-error-compat',['0001,a,10',',b,20','0003,c,30'],rule={'empty':'error'},expected=[['0001','a','10'],['0003','c','30']],empty=1,words=('id','3 行目'))
    case(root,'duplicate-conflict',['0001,a,10','0001,b,20','0003,c,30'],rule={},expected=[['0003','c','30']],duplicate=2,words=('0001','2, 3','2 行'))
    case(root,'duplicate-conflict-after-identical',['0001,a,10','0001,a,10','0001,b,20','0004,d,40'],rule={},expected=[['0004','d','40']],duplicate=3,words=('0001','2, 3, 4','3 行'))
    case(root,'identical-retransmission',['1,a,10','1,a,10','3,c,30'],expected=[['1','a','10'],['3','c','30']],duplicate=1,words=('3','1 行'))
    case(root,'quoted-comma',['1,a,10','2,"b, comma",20','3,c,30'],expected=[['1','a','10'],['2','b, comma','20'],['3','c','30']])
    for name,bad,words in [('error',['2',('e','#VALUE!'),('e','#N/A')],('B3','C3','#VALUE!','#N/A')),('formula',['2','b',('formula','1/0')],('C3','1/0')),('extra',['2','b','20','extra'],('D3','extra'))]:
        case(root,'xlsx-'+name,[],rows=[['id','name','v'],['1','a','10'],bad,['3','c','30']],expected=[['1','a','10'],['3','c','30']],invalid=1,words=words)
    for name,rows in [('shared-index',[['id','name','v'],['1','a',('s','999')]]),('header-error',[['id',('e','#N/A'),'v'],['1','a','10']])]:
        case(root,'structural-'+name,[],rows=rows,stop=True)
    case(root,'structural-csv-width',['1,a,10,extra'],stop=True)
    case(root,'structural-unclosed-quote',['1,"unfinished,10'],stop=True)
    for name,baseline in [('state',[['state','id','name','v'],['UNKNOWN','1','a','10']]),('duplicate',[['state','id','name','v'],['FALSE','1','a','10'],['TRUE','1','b','20']])]:
        case(root,'ledger-'+name,['1,a,10'],baseline=baseline,stop=True)

    inputs={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in root.rglob('*') if p.is_file()}
    manifest=root/'manifest.json'; manifest.write_text(json.dumps([e for e,_ in CASES]),encoding='utf-8')
    observation=root/'observed.json'
    run=subprocess.run(['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(ROOT/'tests/Test-UnexpectedRecords.ps1'),'-Manifest',str(manifest),'-Output',str(observation)],capture_output=True,timeout=180)
    EVIDENCE.mkdir(exist_ok=True)
    (EVIDENCE/'unexpected-records-output.log').write_bytes(run.stdout+run.stderr)
    if run.returncode: raise RuntimeError((run.stdout+run.stderr).decode('utf-8',errors='replace'))
    observations={item['name']:item for item in json.loads(observation.read_text(encoding='utf-8'))}
    results=[]
    for entry,expected in CASES:
        name=entry['name']; actual=observations[name]; result=dict(name=name,status='PASS')
        try:
            code=3 if expected['stop'] else 0
            assert actual['runExit']==code and actual['validateExit']==code, actual
            reportpath=Path(entry['report'])
            if expected['stop']:
                assert not reportpath.exists(), 'a stopped run wrote a result'
            else:
                report=json.loads(reportpath.read_text(encoding='utf-8-sig'))
                assert report['rows']==expected['expected'], report['rows']
                for key in ('invalid','duplicate','empty'):
                    assert report['summary']['skipped'+key.capitalize()]==expected[key], report['summary']
                warnings='\n'.join(report['warnings'])
                for word in expected['words']: assert word in warnings, (word,warnings)
                assert actual['windowLines']==['\t'.join(row) for row in report['rows']], actual['windowLines']
                assert actual['previewValid'], actual['preview']
                log=(Path(entry['directory'])/'events.log').read_text(encoding='utf-8-sig') if report['warnings'] else ''
                for warning in report['warnings']: assert warning in log, ('log',warning)
                if name=='numeric-condition': assert report['values']['S']['count']==2, report['values']['S']
                result.update(summary=report['summary'],rows=report['rows'],warnings=report['warnings'])
            assert not (Path(entry['directory'])/'shared.xlsx').exists(), 'headless or preview wrote the ledger'
        except AssertionError as error: result.update(status='FAIL',detail=str(error))
        results.append(result); print(result['status'],name,result.get('detail',''),flush=True)
    assert all(hashlib.sha256(Path(p).read_bytes()).hexdigest()==digest for p,digest in inputs.items()), 'an input or baseline changed'
    report=dict(tests=results,passed=sum(r['status']=='PASS' for r in results),failed=sum(r['status']=='FAIL' for r in results),inputFilesUnchanged=len(inputs))
    (EVIDENCE/'unexpected-records-results.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('TOTAL',report['passed'],'passed',report['failed'],'failed')
    raise SystemExit(bool(report['failed']))
