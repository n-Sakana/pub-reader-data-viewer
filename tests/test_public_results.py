import json
import pathlib
import subprocess
import sys
import tempfile
from public_results import redact, sanitize_tree


with tempfile.TemporaryDirectory(prefix='rdv-public-evidence-') as temporary:
    root = pathlib.Path(temporary)
    try:
        sanitize_tree(root / 'missing', check=True)
        raise AssertionError('Missing evidence passed the publication check')
    except ValueError:
        pass
    before = dict(passed=3, failed=0, rows=['TRUE', 'FALSE'],
                  file=r'C:\Users\test user\AppData\Local\sample.xlsx',
                  forward='D:/Users/利用者 名/Temp/trace.txt')
    raw = json.dumps(before, ensure_ascii=False)
    (root / 'result.json').write_text(raw, encoding='utf-8')
    original_log = 'FAIL path c:\\Users\\example\\Temp\\case.log\r\nrows=3 archive=1\r\n'
    (root / 'result.txt').write_bytes(original_log.encode('utf-16'))
    try:
        sanitize_tree(root, check=True)
        raise AssertionError('Unredacted evidence passed the publication check')
    except ValueError:
        pass
    assert sanitize_tree(root) == 2
    after = json.loads((root / 'result.json').read_text(encoding='utf-8'))
    assert after['passed'] == before['passed'] and after['failed'] == before['failed'] and after['rows'] == before['rows']
    assert after['file'] == r'<USERPROFILE>\AppData\Local\sample.xlsx'
    assert after['forward'] == '<USERPROFILE>/Temp/trace.txt'
    assert (root / 'result.txt').read_bytes().decode('utf-16') == '<USERPROFILE>'.join(original_log.split(r'c:\Users\example'))
    assert sanitize_tree(root) == 0 and sanitize_tree(root, check=True) == 0
    assert redact('PASS rows=3; C:/app/data/ledger.xlsx') == 'PASS rows=3; C:/app/data/ledger.xlsx'
    (root / 'new.txt').write_text('C:/Users/example/Temp/new.log', encoding='utf-8')
    subprocess.run([sys.executable, str(pathlib.Path(__file__).with_name('public_results.py')), '--path', str(root)], check=True)
    assert (root / 'new.txt').read_text() == '<USERPROFILE>/Temp/new.log'
print('PASS public evidence: reject leaks, JSON values, UTF-16 logs, Unicode/space profiles, idempotence, regenerated output')
