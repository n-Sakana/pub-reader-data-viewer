"""Exact regeneration check using only Python 3 and Node.js from PATH."""
import hashlib
import json
import pathlib
import subprocess
import tempfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
config = ROOT / 'configs/sample/settings.json'
names = [table['file'] for table in json.loads(config.read_text(encoding='utf-8-sig'))['data']['tables'].values()]
inputs = [ROOT / 'samples/current' / folder / name for folder in ('data', '追加CSVデータ/5月分（3件追加）') for name in names]
inputs += [config, ROOT / 'settings.json']
before = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in inputs}
with tempfile.TemporaryDirectory(prefix='rdv-monthly-') as directory:
    for run in ('first', 'repeat'):
        output = pathlib.Path(directory) / run
        subprocess.run(['node', str(ROOT / 'build/gen_monthly_samples.mjs'), '--output', str(output)], check=True)
        for source in inputs[:-2]:
            relative = source.relative_to(ROOT / 'samples/current')
            assert source.read_bytes() == (output / relative).read_bytes(), str(relative)
        assert len(list(output.rglob('*.csv'))) == 8
assert before == {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in inputs}
print('PASS two regenerations: all eight CSVs byte-identical; source data/settings unchanged')
