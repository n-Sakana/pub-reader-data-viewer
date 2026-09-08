"""Fault the compiled logger and bootstrap in owned processes; verify retained evidence."""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import uuid

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--binary', type=Path, required=True)
args = parser.parse_args()
root = Path(__file__).resolve().parents[1]
binary = args.binary.resolve()
folder = binary.parent
log = Path(os.environ['LOCALAPPDATA'])/'ReaderDataViewer/logs/feedback.log'
fallback = Path(os.environ['TEMP'])/'ReaderDataViewer/logs/feedback.log'
results = []

def check(ok, detail):
    if not ok:
        raise AssertionError(detail)

def test(name, action):
    try:
        detail = action()
        results.append(dict(name=name, status='PASS', detail=detail))
    except Exception as error:
        results.append(dict(name=name, status='FAIL', detail=str(error)))
    print(results[-1]['status'], name, results[-1]['detail'], flush=True)

def compiled(mode):
    token = str(uuid.uuid4())
    result = subprocess.run([str(binary), mode, token], capture_output=True, timeout=20)
    text = log.read_text(encoding='utf-8')
    check(token in text, 'context token missing from log')
    if mode == 'throw':
        check(result.returncode != 0 and 'FATAL' in text and token+' worker crash' in text, 'unhandled exception was not recorded')
    elif mode == 'silent':
        check(result.returncode == 0 and 'UNEXPECTED EXIT\tNo END was reached; last phase='+token in text, 'silent exit lacked incomplete-run record')
    else:
        check(result.returncode == 0 and 'HEARTBEAT\tlast phase='+token in text, 'long-running phase lacked heartbeat')
    (folder/(mode+'.log')).write_text('\n'.join(x for x in text.splitlines() if token in x), encoding='utf-8')
    return dict(exit=result.returncode, token=token)

test('unhandled-worker-exception', lambda: compiled('throw'))
test('silent-exit-zero', lambda: compiled('silent'))
test('unfinished-phase-heartbeat', lambda: compiled('wait'))

def rotation():
    destination = folder/('rotation-'+uuid.uuid4().hex+'.log')
    subprocess.run([str(binary), 'rotate', 'record', str(destination)], check=True, timeout=15)
    files = sorted(folder.glob(destination.name+'*'))
    check(len(files) == 4, f'expected current + 3 previous, found {len(files)}')
    sizes = [f.stat().st_size for f in files]
    check(max(sizes) <= 4194304, f'file limit exceeded: {sizes}')
    check(destination.read_text().startswith('record:9:'), 'newest entry missing')
    check(Path(str(destination)+'.3').read_text().startswith('record:6:'), 'three previous generations not retained')
    return dict(files=4, bytes=sizes, total=sum(sizes))
test('bounded-rotation-keeps-three-previous-files', rotation)

def simultaneous():
    destination = folder/('parallel-'+uuid.uuid4().hex+'.log')
    processes = [subprocess.Popen([str(binary), 'append', str(i), str(destination)]) for i in range(4)]
    try:
        check(all(p.wait(timeout=20) == 0 for p in processes), 'append process failed')
    finally:
        for p in processes:
            if p.poll() is None:
                p.kill(); p.wait()
    rows = destination.read_text().splitlines()
    expected = {f'{p}:{i}' for p in range(4) for i in range(250)}
    check(len(rows) == 1000 and set(rows) == expected, 'concurrent records lost or interleaved')
    return dict(processes=4, records=len(rows))
test('cross-process-log-writes', simultaneous)

def backup():
    token = str(uuid.uuid4())
    subprocess.run([str(binary), 'fallback', token], check=True, timeout=10)
    text = fallback.read_text(encoding='utf-8')
    check('LOG FALLBACK' in text and token in text, 'fallback file lacks error/context')
    return dict(token=token, file=str(fallback))
test('unwritable-primary-uses-temp-file', backup)

def bad_destination():
    token = str(uuid.uuid4())
    subprocess.run([str(binary), 'bad-destination', token, str(folder)], check=True, timeout=10)
    text = log.read_text(encoding='utf-8')
    check(token in text and 'configured log '+str(folder) in text, 'configured-log failure lost')
    return token
test('unwritable-configured-log-keeps-local-feedback', bad_destination)

def bootstrap(broken=False):
    launcher = root/'src/ReaderDataViewer.ps1'
    if broken:
        app = folder/('broken-compile-'+uuid.uuid4().hex)
        shutil.copytree(root/'src', app/'src')
        shutil.copytree(root/'lib', app/'lib')
        with (app/'src/Rdv3Values.cs').open('a', encoding='utf-8') as stream:
            stream.write('\nTHIS_IS_A_COMPILER_FAULT\n')
        launcher = app/'src/ReaderDataViewer.ps1'
    before = len(log.read_text(encoding='utf-8'))
    result = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(launcher),
                             '-CompileOnly' if broken else '-Impossible'], capture_output=True, timeout=15)
    text = log.read_text(encoding='utf-8')[before:]
    expected = 3 if broken else 2
    check(result.returncode == expected, f'wrong exit: {result.returncode}')
    check(('launcher failed' if broken else 'Unknown arguments: -Impossible') in text, 'bootstrap failure not logged')
    check('END\texit='+str(expected) in text, 'bootstrap END missing')
    (folder/('compiler.log' if broken else 'arguments.log')).write_text(text, encoding='utf-8')
    return dict(exit=result.returncode, bytes=len(text.encode('utf-8')))
test('unknown-arguments-before-csharp', bootstrap)
test('compiler-failure-before-csharp', lambda: bootstrap(True))

report = dict(passed=sum(x['status']=='PASS' for x in results), failed=sum(x['status']=='FAIL' for x in results), tests=results)
(root/'tests/results/logging-faults-results.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(f"TOTAL {report['passed']} passed; {report['failed']} failed")
raise SystemExit(bool(report['failed']))
