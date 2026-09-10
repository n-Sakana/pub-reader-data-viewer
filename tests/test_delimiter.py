"""Tab / semicolon separated sources (data.tables.<ID>.delimiter) and the tab hint when the setting is missing."""
from pathlib import Path
import copy, json, re, subprocess, tempfile

root = Path(__file__).resolve().parents[1]
readme = (root/'manual/SETTINGS.md').read_text(encoding='utf-8-sig')
minimal = json.loads(re.findall(r'```json\n(.*?)\n```', readme, re.S)[0])
evidence = root/'tests/results'
evidence.mkdir(exist_ok=True)
results = []


def check(ok, message):
    if not ok:
        raise AssertionError(message)


def passed(name, detail):
    results.append(dict(name=name, status='PASS', detail=detail))
    print('PASS', name, detail, flush=True)


with tempfile.TemporaryDirectory(prefix='rdv-delimiter-') as temporary:
    folder = Path(temporary)

    def execute(name, cfg, files, expected, mode='-RunUpdate'):
        case = folder/name
        data = case/'data'
        data.mkdir(parents=True)
        cfg = copy.deepcopy(cfg)
        cfg['paths'] = {'dataDir': str(data), 'ledger': str(case/'ledger.xlsx'), 'log': str(case/'operation.log')}
        config = case/'settings.json'
        config.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
        for fname, content in files.items():
            if isinstance(content, bytes):
                (data/fname).write_bytes(content)
            else:
                (data/fname).write_text(content, encoding='utf-8', newline='\r\n')
        output = case/'report.json'
        command = ['powershell.exe', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', str(root/'src/ReaderDataViewer.ps1'),
                   mode, '-Config', str(config), '-DataDir', str(data)]
        if mode == '-RunUpdate':
            command += ['-Output', str(output)]
        result = subprocess.run(command, capture_output=True, timeout=60)
        (evidence/('delimiter-'+case.name+'.log')).write_bytes(result.stdout+result.stderr)
        check(result.returncode == expected, case.name+': unexpected exit '+str(result.returncode)+'\n'+result.stderr.decode('cp932', 'replace'))
        if expected:
            return result.stderr.decode('cp932', 'replace')
        report = json.loads(output.read_text(encoding='utf-8'))
        (evidence/('delimiter-'+case.name+'.json')).write_text(json.dumps(report, ensure_ascii=False), encoding='utf-8')
        return report

    unicode_text = ('id\tname\r\n001\tone\r\n002\ttwo, with comma\r\n').encode('utf-16')   # BOM + UTF-16LE, Excel "Unicode text"
    cfg = copy.deepcopy(minimal)
    cfg['data']['tables']['A'].update({'file': 'rows.txt', 'encoding': 'utf-16', 'delimiter': 'tab'})
    report = execute('utf16-tab', cfg, {'rows.txt': unicode_text}, 0)
    check(report['rows'] == [['001', 'one'], ['002', 'two, with comma']], 'tab rows %r' % report['rows'])
    passed('utf16-tab', 'Excel Unicode text (UTF-16, tab) read with delimiter tab; a comma inside a cell is data')

    cfg = copy.deepcopy(minimal)
    cfg['data']['tables']['A'].update({'file': 'rows.txt', 'encoding': 'utf-16'})
    message = execute('utf16-tab-missing-setting', cfg, {'rows.txt': unicode_text}, 3, '-ValidateOnly')
    check('delimiter' in message and 'tab' in message, 'tab hint missing: ' + message)
    passed('utf16-tab-missing-setting', 'without delimiter the error names delimiter and tab instead of duplicate headers')

    cfg = copy.deepcopy(minimal)
    message = execute('utf8-tab-missing-setting', cfg, {'rows.csv': 'id\tname\n001\tone\n'}, 3, '-ValidateOnly')
    check('delimiter' in message and 'tab' in message, 'tab hint missing on the fast path: ' + message)
    passed('utf8-tab-missing-setting', 'the byte-level CSV path gives the same tab hint')

    cfg = copy.deepcopy(minimal)
    cfg['data']['tables']['A']['delimiter'] = ';'
    report = execute('semicolon', cfg, {'rows.csv': 'id;name\n001;one\n002;"two;quoted"\n'}, 0)
    check(report['rows'] == [['001', 'one'], ['002', 'two;quoted']], 'semicolon rows %r' % report['rows'])
    passed('semicolon', 'one-character delimiter; a quoted cell keeps the delimiter')

    cfg = copy.deepcopy(minimal)
    cfg['data']['tables']['A']['delimiter'] = 'newline'
    message = execute('bad-delimiter', cfg, {'rows.csv': 'id,name\n001,one\n'}, 3, '-ValidateOnly')
    check('delimiter' in message, 'bad delimiter not reported: ' + message)
    passed('bad-delimiter', 'an unknown delimiter word is refused at settings time')

(evidence/'delimiter-results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
print('TOTAL %d/%d' % (len(results), len(results)))
