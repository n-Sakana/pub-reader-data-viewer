"""Windows process-level feedback checks. Faults affect only temporary inputs and owned children."""
import argparse
import copy
import ctypes
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time

import json5

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
args = parser.parse_args()
root = args.root.resolve()
feedback = Path(os.environ['LOCALAPPDATA'])/'ReaderDataViewer/logs/feedback.log'
base = json5.loads((root/'settings.json').read_text(encoding='utf-8-sig'))
results = []

def read_log():
    return feedback.read_text(encoding='utf-8') if feedback.exists() else ''

def check(condition, detail):
    if not condition:
        raise AssertionError(detail)

def test(name, run):
    try:
        detail = run()
        results.append(dict(name=name, status='PASS', detail=detail))
    except Exception as error:
        results.append(dict(name=name, status='FAIL', detail=str(error)))
    print(results[-1]['status'], name, results[-1]['detail'], flush=True)

with tempfile.TemporaryDirectory(prefix='rdv-log-tests-') as temporary:
    folder = Path(temporary)
    def setup(name):
        case = folder/name
        case.mkdir()
        shutil.copytree(root/'tests/fixtures/data', case/'data')
        cfg = copy.deepcopy(base)
        cfg['paths'] = dict(dataDir=str(case/'data'), ledger=str(case/'ledger.xlsx'), log=str(case/'operation.log'))
        cfg['watch']['targets'] = []
        config = case/'settings.json'
        config.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
        return case, config

    def run(case, config, expected, extra=None, during=None, needle=''):
        command = ['powershell.exe', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File',
                   str(root/'src/ReaderDataViewer.ps1'), '-Config', str(config)] + (extra or ['-ValidateOnly'])
        before = len(read_log())
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        try:
            if during:
                during(process)
            out, err = process.communicate(timeout=35)
        finally:
            if process.poll() is None:
                process.kill()
                process.communicate()
        log = read_log()[before:]
        (case/'feedback.log').write_text(log, encoding='utf-8')
        evidence = root/'tests/results'/('log-'+case.name+'.txt')
        evidence.write_text(log, encoding='utf-8')
        check(f'pid={process.pid}\tBEGIN BOOTSTRAP' in log, 'missing bootstrap record')
        if expected is not None:
            check(process.returncode == expected, f'exit {process.returncode}: {err.decode(errors="replace")}')
            check(f'pid={process.pid}\tEND\texit={expected}' in log, 'missing END/exit code')
        else:
            check(f'pid={process.pid}\tEND' not in log, 'killed run was reported as completed')
        check(needle in log, 'missing feedback: '+needle)
        return dict(exit=process.returncode, pid=process.pid, logBytes=len(log.encode('utf-8')))

    def ordinary():
        case, config = setup('validate-ok')
        return run(case, config, 0, needle='PASS\tValidateOnly')
    test('validation-success-is-recorded', ordinary)

    def execute():
        case, config = setup('update-ok')
        result = run(case, config, 0, ['-RunUpdate', '-Output', str(case/'rows.json')], needle='PASS\tRunUpdate')
        report = json.loads((case/'rows.json').read_text(encoding='utf-8'))
        check(report['summary']['rows'] == 990, 'sample ledger differs')
        check(not (case/'ledger.xlsx').exists(), 'headless mode wrote the configured ledger')
        return dict(result, rows=990)
    test('update-success-and-output-unchanged', execute)

    def syntax():
        case, config = setup('syntax')
        config.write_text('{', encoding='utf-8')
        return run(case, config, 3, needle='FAIL 1 error')
    test('broken-json-file-log', syntax)

    def labels():
        case, config = setup('labels')
        cfg = json.loads(config.read_text(encoding='utf-8'))
        cfg['data']['labels'] = {}
        config.write_text(json.dumps(cfg), encoding='utf-8')
        result = run(case, config, 3, needle='NOT CHECKED')
        check('has no screen label' in (case/'feedback.log').read_text(encoding='utf-8'), 'errors missing')
        return result
    test('all-validation-errors-in-file', labels)

    def damaged(name, content, expected_text):
        case, config = setup(name)
        (case/'data/tableA.csv').write_bytes(content)
        return run(case, config, 3, needle=expected_text)
    test('broken-csv-file-log', lambda: damaged('broken-csv', b'id,name\n"unfinished', 'tableA.csv'))
    test('invalid-encoding-file-log', lambda: damaged('encoding', b'id,name\n12345678,\x81\n', 'tableA.csv'))

    def missing():
        case, config = setup('missing')
        (case/'data/tableB.csv').unlink()
        return run(case, config, 3, needle='tableB.csv')
    test('missing-input-file-log', missing)

    def locked():
        case, config = setup('locked')
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.CreateFileW.restype = ctypes.c_void_p
        handle = kernel.CreateFileW(str(case/'data/tableA.csv'), 0x80000000, 0, None, 3, 0, None)
        check(handle != ctypes.c_void_p(-1).value, 'could not lock own fixture')
        try:
            return run(case, config, 3, needle='tableA.csv')
        finally:
            kernel.CloseHandle(ctypes.c_void_p(handle))
    test('unreadable-input-file-log', locked)

    def delayed(name, kill=False, update=False):
        case, config = setup(name)
        source = case/'data/tableA.csv'
        lines = source.read_bytes().splitlines(keepends=True)
        source.write_bytes(lines[0]+lines[1]*100000)
        def act(process):
            deadline = time.monotonic()+20
            while time.monotonic() < deadline:
                log = read_log()
                if f'pid={process.pid}\tPHASE\treading input A ' in log:
                    if kill:
                        process.kill()
                    else:
                        (case/'data/tableB.csv').unlink()
                    return
                check(process.poll() is None, 'process ended before fault injection')
                time.sleep(.005)
            raise AssertionError('input-read phase not reached')
        extra = ['-RunUpdate', '-Output', str(case/'rows.json')] if update else ['-ValidateOnly']
        return run(case, config, None if kill else 3, extra, act, 'reading input A' if kill else 'tableB.csv')
    test('input-disappears-during-validation', lambda: delayed('disappeared'))
    test('killed-validation-keeps-last-phase', lambda: delayed('killed-validate', True))
    test('killed-update-keeps-last-phase', lambda: delayed('killed-update', True, True))

report = dict(passed=sum(x['status']=='PASS' for x in results), failed=sum(x['status']=='FAIL' for x in results), tests=results)
(root/'tests/results/logging-results.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(f"TOTAL {report['passed']} passed; {report['failed']} failed")
raise SystemExit(bool(report['failed']))
