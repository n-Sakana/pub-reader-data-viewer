"""Run the complete four-table README example through the real application entry point."""
from pathlib import Path
import hashlib
import json
import re
import subprocess
import tempfile
import json5

root = Path(__file__).resolve().parents[1]
section = (root/'README.md').read_text(encoding='utf-8-sig').split('<a id="four-tables"></a>')[1].split('<a id="ledger"></a>')[0]
config_text = re.findall(r'```jsonc\n(.*?)\n```', section, re.S)[0]
csvs = re.findall(r'```csv\n(.*?)\n```', section, re.S)
assert len(csvs) == 5, 'Four initial CSVs and the replacement D.csv are required'
results = []
outdir = root/'tests/results'
outdir.mkdir(exist_ok=True)

def check(ok, detail):
    if not ok:
        raise AssertionError(detail)

def passed(name, detail):
    results.append(dict(name=name, status='PASS', detail=detail))
    print('PASS', name, detail, flush=True)

with tempfile.TemporaryDirectory(prefix='rdv-readme-') as scratch:
    folder = Path(scratch)
    data = folder/'data'
    data.mkdir()
    config = folder/'settings.json'
    config.write_text(config_text, encoding='utf-8')
    for name, csv in zip('ABCD', csvs):
        (data/(name+'.csv')).write_text(csv+'\n', encoding='utf-8')

    def invoke(name, mode, *args, expected=0):
        command = ['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(root/'src/ReaderDataViewer.ps1'),
                   mode,'-Config',str(config),'-DataDir',str(data),*map(str,args)]
        result = subprocess.run(command, capture_output=True, timeout=35)
        (outdir/('readme-'+name+'.log')).write_bytes(result.stdout+result.stderr)
        check(result.returncode == expected, f'{name}: exit {result.returncode}; see readme-{name}.log')
        return result

    def update(name, baseline=None):
        output=folder/(name+'.json')
        options=['-Output',output]
        if baseline:
            options += ['-BaselineLedger',baseline]
        invoke(name,'-RunUpdate',*options)
        text=output.read_text(encoding='utf-8')
        (outdir/('readme-'+name+'.json')).write_text(text, encoding='utf-8')
        return output, json.loads(text)

    def snapshot(report, name, reviewed=False):
        path=folder/(name+'.xlsx')
        command=['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(root/'tests/Test-ReadmeBaseline.ps1'),
                 '-Config',str(config),'-DataDir',str(data),'-Report',str(report),'-Output',str(path)]
        if reviewed:
            command+=['-MarkReviewed']
        result=subprocess.run(command, capture_output=True, timeout=35)
        (outdir/('readme-'+name+'.log')).write_bytes(result.stdout+result.stderr)
        check(result.returncode == 0 and path.exists(), f'{name}: fixture ledger creation failed')
        return path

    invoke('validate','-ValidateOnly')
    passed('complete-json-and-japanese-dates', 'exit 0; four dates include single- and double-digit month/day')
    first_path, first=update('first')
    rows={r[0]:dict(zip(first['columns'],r)) for r in first['rows']}
    check(first['summary']==dict(rows=4,skippedEmpty=0,skippedDuplicate=0,skippedShort=0,skippedBlank=0,skippedColumns=0,skippedInvalid=0,baselineRows=0,resetRows=0), 'first summary')
    check(rows['R01']['B.amount']=='80' and rows['R03']['B.amount']=='1200', 'aggregate amounts')
    check([r for r in rows if rows[r]['B.id']=='']==['R02','R04'], 'left join missing rows')
    check([j['unmatchedLeft'] for j in first['joins']]==[2,0,4], 'join counts')
    check(first['states']==['FALSE']*4, 'initial states')
    passed('aggregate-left-joins', '4 rows; sums 80/1200; unmatched 2/0/4; four FALSE states')

    baseline=snapshot(first_path,'before',True)
    checksum=hashlib.sha256(baseline.read_bytes()).hexdigest()
    (data/'D.csv').write_text(csvs[4]+'\n', encoding='utf-8')
    second_path, second=update('second',baseline)
    check(second['summary']['resetRows']==1 and [r[0] for r in second['resetRows']]==['R02'], 'reset one row')
    check(second['states']==['TRUE','FALSE','FALSE','FALSE'], 'unmodified reviewed row preserved')
    check(second['rows'][1][-1]=='1' and second['joins'][2]['unmatchedLeft']==3, 'cancel value and joins')
    check(hashlib.sha256(baseline.read_bytes()).hexdigest()==checksum, 'baseline must be read-only')
    passed('cancellation-resets-only-changed-row', 'R02 reset; R01 TRUE preserved; baseline bytes unchanged')

    after=snapshot(second_path,'after')
    _, repeated=update('repeat',after)
    check(repeated['summary']['resetRows']==0 and repeated['rows']==second['rows'] and repeated['states']==second['states'], 'repeat changed content/state')
    passed('repeat-is-stable', 'resetRows=0; all rows and states unchanged')

    bad=json5.loads(config_text)
    bad['data']['labels']['B']='duplicate'
    del bad['data']['tables']['C']['label']
    bad['data']['labels']['C']='also duplicate with omitted table label'
    config.write_text(json.dumps(bad,ensure_ascii=False), encoding='utf-8')
    invoke('labels','-ValidateOnly',expected=3)
    diagnostics=(outdir/'readme-labels.log').read_bytes()
    check(b'FAIL 2 errors' in diagnostics and b'B already has a table label' in diagnostics and b'C already has a table label' in diagnostics, 'both table labels must be rejected together')
    passed('documented-table-label-rule', 'explicit and omitted table label both reported; FAIL 2 errors; exit 3')

(outdir/'readme-example-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2), encoding='utf-8')
print(f'TOTAL {len(results)}/{len(results)}')
