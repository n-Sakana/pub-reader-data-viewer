"""Run the four relaxations added after the 2026-09-09 blind acceptance round through the real entry point:
derived ledger columns, a derived identity, headerRow above the header, workbook date serials, and the
launcher's refusal under PowerShell 7."""
from pathlib import Path
import copy, json, os, re, shutil, subprocess, tempfile, zipfile

root = Path(__file__).resolve().parents[1]
readme = (root/'README.md').read_text(encoding='utf-8-sig')
minimal = json.loads(re.findall(r'```json\n(.*?)\n```', readme, re.S)[0])
results = []
evidence = root/'tests/results'
evidence.mkdir(exist_ok=True)


def check(ok, message):
    if not ok:
        raise AssertionError(message)


def passed(name, detail):
    results.append(dict(name=name, status='PASS', detail=detail))
    print('PASS', name, detail, flush=True)


def workbook(path, rows):
    """rows: list of lists; a str becomes an inline string cell, an int/float a numeric cell."""
    def cell(ref, value):
        if isinstance(value, str):
            return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, value)
        return '<c r="%s"><v>%s</v></c>' % (ref, value)
    body = ''
    for r, row in enumerate(rows, 1):
        body += '<row r="%d">' % r + ''.join(cell(chr(65 + c) + str(r), v) for c, v in enumerate(row)) + '</row>'
    with zipfile.ZipFile(path, 'w') as book:
        book.writestr('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
        book.writestr('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        book.writestr('xl/worksheets/sheet1.xml', '<worksheet><sheetData>' + body + '</sheetData></worksheet>')


with tempfile.TemporaryDirectory(prefix='rdv-relaxations-') as temporary:
    folder = Path(temporary)

    def execute(name, cfg, files, expected, mode='-RunUpdate', shell='powershell.exe'):
        case = folder/name
        data = case/'data'
        data.mkdir(parents=True)
        cfg = copy.deepcopy(cfg)
        cfg['paths'] = {'dataDir': str(data), 'ledger': str(case/'ledger.xlsx'), 'log': str(case/'operation.log')}
        config = case/'settings.json'
        config.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
        for fname, text in files.items():
            if isinstance(text, bytes):
                (data/fname).write_bytes(text)
            else:
                (data/fname).write_text(text, encoding='utf-8', newline='\r\n')
        output = case/'report.json'
        command = [shell, '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', str(root/'src/ReaderDataViewer.ps1'),
                   mode, '-Config', str(config), '-DataDir', str(data)]
        if mode == '-RunUpdate':
            command += ['-Output', str(output)]
        result = subprocess.run(command, capture_output=True, timeout=60)
        (evidence/('relax-'+case.name+'.log')).write_bytes(result.stdout+result.stderr)
        check(result.returncode == expected, case.name+': unexpected exit '+str(result.returncode)+'\n'+result.stderr.decode('cp932', 'replace'))
        if expected:
            return result.stderr.decode('cp932', 'replace')
        if mode != '-RunUpdate':
            return None
        report = json.loads(output.read_text(encoding='utf-8'))
        (evidence/('relax-'+case.name+'.json')).write_text(json.dumps(report, ensure_ascii=False), encoding='utf-8')
        return report

    def table(cfg, tid, file, key, label=None, **extra):
        entry = {'file': file, 'key': key}
        if label:
            entry['label'] = label
        entry.update(extra)
        cfg['data']['tables'][tid] = entry

    # 1. One detail file, ledger per group: identity and saved columns come from the aggregate output.
    cfg = copy.deepcopy(minimal)
    cfg['data']['tables'] = {}
    table(cfg, 'L', 'lines.csv', ['slip', 'line'], '明細', keyValidation={'length': 'variable'})
    cfg['data']['types'] = {'L.qty': {'type': 'number'}, 'L.price': {'type': 'number'}, 'T.total': {'type': 'number'}}
    cfg['data']['labels'] = {'L.slip': '伝票', 'L.line': '行', 'L.qty': '数量', 'L.price': '単価', 'L.amount': '明細金額',
                             'T': '伝票集計', 'T.slip': '伝票', 'T.total': '合計金額', 'T.count': '行数', 'ledger': '台帳'}
    cfg['data']['jobs'] = [{'id': 'update', 'kind': 'update', 'inputs': [{'table': 'L'}], 'steps': [
        {'operation': 'calculate', 'target1': 'L', 'column': 'amount', 'expression': 'L.qty * L.price', 'output': 'L'},
        {'operation': 'aggregate', 'target1': 'L', 'groupBy': ['L.slip'],
         'aggregates': [{'function': 'sum', 'column': 'L.amount', 'as': 'total'}, {'function': 'count', 'as': 'count'}], 'output': 'T'},
        {'operation': 'merge', 'target1': 'T', 'target2': 'ledger', 'keys': ['L.slip', 'L.slip'],
         'sourceOnly': 'add', 'both': 'update', 'targetOnly': 'keep', 'output': 'ledger'}]}]
    cfg['data']['ledger'] = {'identity': 'L.slip', 'search': {'columns': ['L.slip'], 'match': 'exact'},
                             'columns': {'source': ['L.slip', 'T.total', 'T.count'],
                                         'application': [{'name': 'workState', 'onSourceChange': 'reset'}]}}
    cfg['screen']['export']['defaultFields'] = ['L.slip', 'T.total', 'T.count', '$work']
    cfg['screen']['candidates']['columns'] = [{'header': '伝票', 'value': {'field': 'L.slip'}}, {'header': '合計', 'value': {'field': 'T.total'}}]
    cfg['screen']['sections'][0]['figure'] = {'label': '伝票', 'value': {'field': 'L.slip'}}
    cfg['screen']['sections'][1]['rows'] = [{'label': '合計金額', 'value': {'field': 'T.total', 'empty': ''}},
                                            {'label': '行数', 'value': {'field': 'T.count', 'empty': ''}}]
    lines = 'slip,line,qty,price\nD001,1,2,"1,250"\nD001,2,3,100\nD002,1,1,480\n'
    report = execute('derived-identity-and-columns', cfg, {'lines.csv': lines}, 0)
    check(report['columns'] == ['L.slip', 'T.total', 'T.count'], 'ledger columns %r' % report['columns'])
    check(report['rows'] == [['D001', '2800', '2'], ['D002', '480', '1']], 'aggregated ledger %r' % report['rows'])
    passed('derived-identity-and-columns', 'aggregate output saved as T.total/T.count; identity L.slip from the group key; 2 rows')
    execute('derived-validate-only', cfg, {'lines.csv': lines}, 0, '-ValidateOnly')
    passed('derived-validate-only', 'ValidateOnly accepts derived ledger columns and identity')

    # 2. A calculate column saved under its own name next to a joined table.
    cfg = copy.deepcopy(minimal)
    table(cfg, 'B', 'B.csv', 'id', '相手')
    cfg['data']['types'] = {'A.qty': {'type': 'number'}, 'B.done': {'type': 'number'}}
    cfg['data']['labels'].update({'A.qty': '数量', 'B.id': '相手ID', 'B.done': '完了数', 'joined': '結合', 'joined.remain': '残数'})
    job = cfg['data']['jobs'][0]
    job['inputs'].append({'table': 'B'})
    job['steps'] = [{'operation': 'join', 'target1': 'A', 'target2': 'B', 'keys': ['A.id', 'B.id'], 'condition': 'match', 'output': 'joined'},
                    {'operation': 'calculate', 'target1': 'joined', 'column': 'remain', 'expression': 'A.qty - B.done', 'output': 'joined'},
                    dict(job['steps'][-1], target1='joined')]
    cfg['data']['ledger']['columns']['source'] = ['A.id', 'A.name', 'joined.remain']
    cfg['screen']['export']['defaultFields'] = ['A.id', 'A.name', 'joined.remain', '$work']
    report = execute('calculated-column', cfg, {'rows.csv': 'id,name,qty\n001,one,10\n002,two,7\n', 'B.csv': 'id,done\n001,4\n002,7\n'}, 0)
    check(report['rows'] == [['001', 'one', '6'], ['002', 'two', '0']], 'calculated ledger %r' % report['rows'])
    passed('calculated-column', 'joined.remain saved with values 6 and 0')

    # 3. A saved column whose table was never joined is refused before any row is written.
    cfg = copy.deepcopy(minimal)
    table(cfg, 'B', 'B.csv', 'id', '相手')
    cfg['data']['labels'].update({'B.id': '相手ID', 'B.note': '備考'})
    cfg['data']['jobs'][0]['inputs'].append({'table': 'B'})
    cfg['data']['ledger']['columns']['source'] = ['A.id', 'A.name', 'B.note']
    message = execute('unjoined-source-column', cfg, {'rows.csv': 'id,name\n001,one\n', 'B.csv': 'id,note\n001,x\n'}, 3, '-ValidateOnly')
    check('B.note' in message and 'merge A' in message, 'missing remedy in: ' + message)
    passed('unjoined-source-column', 'exit 3 names B.note and the merge step')

    # 4. headerRow: report title lines above the CSV header, and above an XLSX header.
    cfg = copy.deepcopy(minimal)
    cfg['data']['tables']['A']['headerRow'] = 3
    report = execute('csv-header-row', cfg, {'rows.csv': '売上集計表（2026年8月度）\n出力日,2026/09/01\nid,name\n001,one\n002,two\n'}, 0)
    check(report['rows'] == [['001', 'one'], ['002', 'two']], 'title lines not skipped: %r' % report['rows'])
    check(any('読み飛ばし' in w and '2' in w for w in report['warnings']), 'skip notice missing: %r' % report['warnings'])
    check(report['summary']['skippedShort'] == 0 and report['summary']['skippedBlank'] == 0, 'title lines counted as short/blank')
    passed('csv-header-row', 'headerRow 3 reads the header on line 3; 2 lines skipped with a notice; nothing counted as short')
    book = folder/'title.xlsx'
    workbook(book, [['売上集計表'], ['出力日', '2026/09/01'], ['id', 'name'], ['001', 'one'], ['002', 'two']])
    cfg['data']['tables']['A']['file'] = 'title.xlsx'
    report = execute('xlsx-header-row', cfg, {'title.xlsx': book.read_bytes()}, 0)
    check(report['rows'] == [['001', 'one'], ['002', 'two']], 'xlsx title rows not skipped: %r' % report['rows'])
    passed('xlsx-header-row', 'headerRow 3 applies to a workbook as well')
    cfg['data']['tables']['A']['headerRow'] = 1
    cfg['data']['tables']['A']['file'] = 'rows.csv'
    message = execute('csv-header-row-omitted', cfg, {'rows.csv': '売上集計表（2026年8月度）\n出力日,2026/09/01\nid,name\n001,one\n'}, 3, '-ValidateOnly')
    check('id' in message, 'the old failure (key column missing) must still name the column: ' + message)
    passed('csv-header-row-omitted', 'without headerRow the title line is still taken as the header and refused')

    # 5. Workbook date serials become the declared date text; text dates and undeclared columns are untouched.
    cfg = copy.deepcopy(minimal)
    cfg['data']['tables']['A']['file'] = 'orders.xlsx'
    cfg['data']['types'] = {'A.day': {'type': 'date', 'format': 'yyyy/MM/dd'}, 'A.amount': {'type': 'number'}}
    cfg['data']['labels'].update({'A.day': '日付', 'A.amount': '金額', 'A.raw': '生の値'})
    cfg['data']['ledger']['columns']['source'] = ['A.id', 'A.name', 'A.day', 'A.amount', 'A.raw']
    cfg['screen']['export']['defaultFields'] = ['A.id', 'A.name', 'A.day', 'A.amount', 'A.raw', '$work']
    book = folder/'orders.xlsx'
    workbook(book, [['id', 'name', 'day', 'amount', 'raw'], ['001', 'one', 46246, 12300, 46246], ['002', 'two', '2026/01/02', 5.5, 46246], ['003', 'three', 60, 1, 1]])
    report = execute('xlsx-date-serial', cfg, {'orders.xlsx': book.read_bytes()}, 0)
    rows = {r[0]: r for r in report['rows']}
    # serial 60 is Excel's imaginary 1900-02-29; .NET has no such day, so it lands on 1900-03-01 like serial 61
    check(rows['001'][2] == '2026/08/12' and rows['002'][2] == '2026/01/02' and rows['003'][2] == '1900/03/01',
          'serial conversion %r' % report['rows'])
    check(rows['001'][4] == '46246' and rows['001'][3] == '12300' and rows['002'][3] == '5.5', 'undeclared or numeric cells changed %r' % report['rows'])
    passed('xlsx-date-serial', '46246 -> 2026/08/12 in the declared date column; text date kept; undeclared column keeps 46246')

    # 6. pwsh refuses to start with a clear reason instead of a compiler error.
    pwsh = shutil.which('pwsh')
    if pwsh:
        message = execute('pwsh-refused', minimal, {'rows.csv': 'id,name\n001,one\n'}, 3, '-ValidateOnly', shell=pwsh)
        check('Windows PowerShell 5.1' in message and 'powershell.exe' in message, 'pwsh guidance missing: ' + message)
        passed('pwsh-refused', 'exit 3 with the powershell.exe instruction under ' + pwsh)
    else:
        print('SKIP pwsh-refused (pwsh not installed)')

(evidence/'relaxations-results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
print('TOTAL %d/%d' % (len(results), len(results)))
