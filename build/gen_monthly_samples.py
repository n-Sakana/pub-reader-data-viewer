"""Regenerate the fixed April/May sample CSVs into a separate output folder."""
import argparse
import copy
import csv
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--output', type=pathlib.Path, default=ROOT / 'build/out/monthly-samples')
args = parser.parse_args()
output = args.output.resolve()
if output == ROOT or any(output == ROOT / name or ROOT / name in output.parents for name in ('configs', 'samples', 'data', 'src', 'tests')):
    parser.error('Use a separate output folder; existing samples and settings are inputs.')
config = json.loads((ROOT / 'configs/sample/settings.json').read_text(encoding='utf-8-sig'))
names = {key: value['file'] for key, value in config['data']['tables'].items()}
rows = {}
for key, name in names.items():
    if pathlib.Path(name).name != name or not name.endswith('.csv'):
        raise ValueError('Expected a fixed CSV filename: ' + name)
    with (ROOT / 'samples/current/data' / name).open(encoding='cp932', newline='') as stream:
        rows[key] = list(csv.reader(stream))
    if len(rows[key]) != 6:
        raise ValueError('Expected header and five input rows: ' + name)

def card(n):
    return 'AB' + str(10000000 + n) + 'CD'

def application(n):
    return '受付' + ('AA' if n % 2 else 'BB') + '26-' + str(n).zfill(6)

def new_row(row, n):
    replacements = [(card(1), card(n)), (application(1), application(n))]
    replacements += [(p + 'A', p + chr(64 + n)) for p in ('案件', '利用者', '申込者')]
    replacements += [(p + '0001', p + str(n).zfill(4)) for p in ('TEST', 'USER', 'TESTPAY', 'TESTBILL', 'TESTMASTER')]
    replacements += [('sample1@', 'sample' + str(n) + '@')]
    result = []
    for value in row:
        for old, new in replacements:
            value = value.replace(old, new)
        result.append(value)
    return result

april = copy.deepcopy(rows)
may = {}
for row in april['APP'][1:]:
    row[17] = '2026/04/10'
for key in ('TXN', 'PAY'):
    # The distributed May set retains April transaction/payment dates.
    may[key] = copy.deepcopy([rows[key][0], rows[key][2], rows[key][3]]) + [new_row(rows[key][1], n) for n in (9, 10, 11)]
may['PAY'][1][4] = '決済済'
may['APP'] = [rows['APP'][0], copy.copy(rows['APP'][2]), new_row(rows['APP'][1], 3)] + [new_row(rows['APP'][1], n) for n in (9, 10, 11)]
for i, row in enumerate(may['APP'][1:], 1):
    row[17] = '2026/04/10' if i < 3 else '2026/05/10'
may['DEL'] = [rows['DEL'][0]] + [[card(n), application(n)] for n in (2, 3, 9, 10, 11)]
remarks = rows['APP'][0].index('意見欄')
for n, dataset, directory in ((4, april, output / 'data'), (5, may, output / '追加CSVデータ/5月分（3件追加）')):
    for row in dataset['APP'][1:]:
        row[remarks] = (f'{n}月分の申込内容と提出資料を確認しました。\n連絡は平日の午後を希望しています。\n'
                        '日程が変わる場合は、候補日を二つ案内してください。\n次回は変更後の計画を確認します。以上は架空の内容です。')
    directory.mkdir(parents=True, exist_ok=True)
    for key, name in names.items():
        with (directory / name).open('w', encoding='cp932', newline='') as stream:
            csv.writer(stream, lineterminator='\r\n').writerows(dataset[key])
print('Created two sets of four CSV files, five records each: ' + str(output))
