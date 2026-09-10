"""Check that data errors lead to a concrete, non-destructive correction."""
from pathlib import Path
import copy,json,os,re,subprocess,tempfile

root=Path(__file__).resolve().parents[1]
cfg=json.loads(re.findall(r'```json\n(.*?)\n```',(root/'manual/SETTINGS.md').read_text(encoding='utf-8-sig'),re.S)[0])
evidence=root/'tests/results'
results=[]
def check(ok,message):
    if not ok: raise AssertionError(message)
def passed(name,detail):
    results.append(dict(name=name,status='PASS',detail=detail))
    print('PASS',name,detail,flush=True)

with tempfile.TemporaryDirectory(prefix='rdv-feedback-hints-') as scratch:
    folder=Path(scratch)
    def run(name,raw,encoding,expected):
        case=folder/name
        data=case/'data'
        data.mkdir(parents=True)
        config=copy.deepcopy(cfg)
        config['paths']={'dataDir':str(data),'ledger':str(case/'ledger.xlsx'),'log':str(case/'operation.log')}
        config['data']['encoding']=encoding
        path=case/'settings.json'
        path.write_text(json.dumps(config,ensure_ascii=False),encoding='utf-8')
        csv=data/'rows.csv'
        csv.write_bytes(raw)
        command=['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(root/'src/ReaderDataViewer.ps1'),
                 '-ValidateOnly','-Config',str(path),'-DataDir',str(data)]
        result=subprocess.run(command,capture_output=True,timeout=30)
        text=(result.stdout+result.stderr).decode('utf-8')
        (evidence/('hints-'+name+'.log')).write_text(text,encoding='utf-8')
        check(result.returncode==expected,name+': exit '+str(result.returncode))
        check(csv.read_bytes()==raw,name+': source changed')
        if expected:
            log=(Path(os.environ['LOCALAPPDATA'])/'ReaderDataViewer/logs/feedback.log').read_text(encoding='utf-8')
            diagnostic=result.stderr.decode('utf-8').replace('\r\n','\n').split('FAIL',1)[1].strip()
            check(str(path) in log and diagnostic in log,name+': full diagnostics missing in file')
        return text

    sample='id,name\r\n001,ユニット\r\n'
    for mode,codec,suggestion in [('le-utf8','utf-16-le','utf-16'),('be-utf8','utf-16-be','utf-16BE'),('le-sjis','utf-16-le','utf-16')]:
        wrong='shift_jis' if mode.endswith('sjis') else 'utf-8'
        text=run(mode,sample.encode(codec),wrong,3)
        check('Possible UTF-16' in text and 'set data.encoding to "'+suggestion+'"' in text and 'not changed automatically' in text,'UTF-16 hint absent')
        passed(mode,'exit 3; concrete suggested encoding; original bytes unchanged')
    for mode,codec,encoding in [('correct-le','utf-16-le','utf-16'),('correct-be','utf-16-be','utf-16BE')]:
        run(mode,sample.encode(codec),encoding,0)
        passed(mode,'exit 0 with explicitly corrected encoding')
    text=run('sjis-is-not-guessed-as-utf16',sample.encode('cp932'),'utf-8',3)
    check('Possible UTF-16' not in text,'false UTF-16 hint for Shift-JIS')
    passed('sjis-is-not-guessed-as-utf16','exit 3; no unsupported UTF-16 guess')
    text=run('binary-is-not-guessed-as-text',b'\xff\x00\x01\x00\x02\x00\x03\x00'*3,'utf-8',3)
    check('Possible UTF-16' not in text,'false CSV guess for control bytes')
    passed('binary-is-not-guessed-as-text','exit 3; alternating zeros alone are insufficient')
    text=run('duplicate-key-guidance',b'id,name\r\n001,left\r\n001,right\r\n','utf-8',0)
    for word in ('aggregate','groupBy','入力キーの検査はaggregateより先','distinct','合計しません','001','2, 3','どの行も採用しません'):
        check(word in text,'missing guidance: '+word)
    passed('duplicate-key-guidance','exit 0; both conflicting rows excluded and named; detail key -> aggregate/groupBy; distinct is not summing')

(evidence/'feedback-hints-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print(f'TOTAL {len(results)}/{len(results)}')
