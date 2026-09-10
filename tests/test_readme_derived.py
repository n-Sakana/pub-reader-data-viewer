"""Run the two README examples added after the blind acceptance round through the real entry point:
'明細だけから伝票単位' (derived identity and aggregate columns) and '複数ファイルを縦に足す' (append/distinct/delete)."""
from pathlib import Path
import copy, json, re, subprocess, tempfile
import json5

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


def section(start, end):
    text = readme.split('<a id="%s"></a>' % start)[1].split('<a id="%s"></a>' % end)[0]
    return json5.loads(re.findall(r'```jsonc\n(.*?)\n```', text, re.S)[0])


def settings(data, identity, fields, state_names=('未確認', '確認済')):
    cfg = copy.deepcopy(minimal)
    cfg['data'] = data
    cfg['search'] = {'pattern': '.+', 'candidateRowsShown': 100}
    screen = cfg['screen']
    screen['export']['defaultFields'] = fields + ['$work']
    screen['candidates']['columns'] = [{'header': f, 'value': {'field': f}} for f in fields[:2]]
    screen['sections'][0]['figure'] = {'label': identity, 'value': {'field': identity}}
    screen['sections'][1]['rows'] = [{'label': f, 'value': {'field': f, 'empty': ''}} for f in fields]
    screen['workState']['states'][0]['text'] = state_names[0]
    screen['workState']['states'][1]['text'] = state_names[1]
    return cfg


with tempfile.TemporaryDirectory(prefix='rdv-readme-derived-') as temporary:
    folder = Path(temporary)

    def run(name, cfg, files):
        case = folder/name
        data = case/'data'
        data.mkdir(parents=True)
        cfg = copy.deepcopy(cfg)
        cfg['paths'] = {'dataDir': str(data), 'ledger': str(case/'ledger.xlsx'), 'log': str(case/'operation.log')}
        config = case/'settings.json'
        config.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
        for fname, text in files.items():
            (data/fname).write_text(text, encoding='utf-8', newline='\r\n')
        output = case/'report.json'
        for mode in ('-ValidateOnly', '-RunUpdate'):
            command = ['powershell.exe', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', str(root/'src/ReaderDataViewer.ps1'),
                       mode, '-Config', str(config), '-DataDir', str(data)]
            if mode == '-RunUpdate':
                command += ['-Output', str(output)]
            result = subprocess.run(command, capture_output=True, timeout=60)
            (evidence/('readme-derived-%s-%s.log' % (name, mode.strip('-')))).write_bytes(result.stdout+result.stderr)
            check(result.returncode == 0, '%s %s: exit %d\n%s' % (name, mode, result.returncode, result.stderr.decode('cp932', 'replace')))
        report = json.loads(output.read_text(encoding='utf-8'))
        (evidence/('readme-derived-%s.json' % name)).write_text(json.dumps(report, ensure_ascii=False), encoding='utf-8')
        return report

    data = section('detail-only', 'leading-zero')
    cfg = settings(data, 'L.伝票番号', ['L.伝票番号', 'L.出荷日', 'L.得意先', 'T.合計金額', 'T.明細数'])
    lines = ('伝票番号,行番号,出荷日,得意先,数量,単価\n'
             'D001,1,2026/08/02,青葉商事,2,"1,250"\nD001,2,2026/08/02,青葉商事,3,100\n'
             'D002,1,2026/08/03,北山電機,1,480\nD003,1,2026/08/04,駿河物産,10,12000\nD003,2,2026/08/04,駿河物産,1,250\n')
    report = run('detail-only', cfg, {'出荷実績.csv': lines})
    check(report['columns'] == ['L.伝票番号', 'L.出荷日', 'L.得意先', 'T.合計金額', 'T.明細数'], 'columns %r' % report['columns'])
    check(report['rows'] == [['D001', '2026/08/02', '青葉商事', '2800', '2'], ['D002', '2026/08/03', '北山電機', '480', '1'],
                             ['D003', '2026/08/04', '駿河物産', '120250', '2']], 'rows %r' % report['rows'])
    passed('readme-detail-only', 'identity from the group key; sums 2800/480/120250 and counts saved as T.* columns')

    data = section('append-files', 'four-tables')
    cfg = settings(data, 'A.受注番号', ['A.受注番号', 'A.受注日', 'A.得意先', 'A.金額', 'A.状態'], ('未処理', '処理済'))
    april = ('受注番号,受注日,得意先,金額,状態\n26040001,2026/04/01,青葉商事,1000,受注\n26040002,2026/04/02,北山電機,2000,取消\n'
             '26040003,2026/04/03,駿河物産,3000,出荷済\n')
    may = ('受注番号,受注日,得意先,金額,状態\n26050001,2026/05/01,高橋建材,4000,受注\n26040003,2026/04/03,駿河物産,3000,取消\n'
           '26040001,2026/04/01,青葉商事,1000,受注\n26050002,2026/05/02,中央印刷,5000,取消\n')
    report = run('append-files', cfg, {'受注_4月.csv': april, '受注_5月.csv': may})
    rows = {r[0]: r for r in report['rows']}
    check(sorted(rows) == ['26040001', '26050001'], 'rows %r' % report['rows'])
    check(report['summary']['rows'] == 2 and report['values']['cancelled']['count'] == 3, 'summary %r' % report['summary'])
    passed('readme-append-files', 'union of both months, May wins, 3 cancelled rows removed, 2 rows kept')

    data = section('conditional-replace', 'four-tables')
    cfg = settings(data, 'A.受注番号', ['A.受注番号', 'A.受注日', 'A.得意先', 'A.金額', 'A.状態コード', 'A.状態名'])
    orders = ('受注番号,受注日,得意先,金額,状態コード\nQ0001,2026/09/01,青葉商事,1000,1\nQ0002,2026/09/02,北山電機,2000,2\n'
              'Q0003,2026/09/03,駿河物産,3000,9\nQ0004,2026/09/04,高橋建材,4000,7\n')
    report = run('conditional-replace', cfg, {'受注.csv': orders})
    names = {r[0]: r[5] for r in report['rows']}
    check(names == {'Q0001': '受注', 'Q0002': '出荷済', 'Q0003': '取消', 'Q0004': '7'}, 'state names %r' % names)
    passed('readme-conditional-replace', 'codes 1/2/9 replaced by names, an unlisted code stays visible as the code')

(evidence/'readme-derived-results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
print('TOTAL %d/%d' % (len(results), len(results)))
