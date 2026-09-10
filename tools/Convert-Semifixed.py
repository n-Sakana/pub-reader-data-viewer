"""Convert the confirmed screen shape without changing data/jobs or bindings.
Python is needed for conversion only; normal startup uses Windows PowerShell.
"""
import argparse
from copy import deepcopy
import json
from pathlib import Path

FIELDS = ['userId','userName','userCategory','applicationDate','applicationNumber',
          'cardNumber','applicantName','birthDate','qualification','office','remarks','plan']
CANDIDATES = ['applicationNumber','cardNumber','applicant','applicationDate','payment']

def convert(original):
    result = deepcopy(original)
    screen = result['screen']
    sections = screen['sections']
    columns = [s for s in sections if s['type'] == 'columns']
    texts = [s for s in sections if s['type'] == 'textBox']
    bands = [s for s in sections if s['type'] == 'statusBand']
    if len(columns) != 1 or len(texts) != 2 or len(bands) != 1 or bands[0]['judgment'] != 'paymentStatus':
        raise ValueError('Expected the confirmed layout: one two-column section, two text boxes and paymentStatus')
    groups = columns[0]['items']
    if len(groups) != 2 or [len(s['rows']) for s in groups] != [3,7]:
        raise ValueError('Expected three user fields and seven application fields')
    values = [row['value'] for group in groups for row in group['rows']] + [s['value'] for s in texts]
    candidate_columns = screen['candidates']['columns']
    if len(candidate_columns) != 7:
        raise ValueError('Expected seven candidate columns')
    if candidate_columns[0]['value'] != {'state':'rowNumber'} or candidate_columns[-1]['value'] != {'state':'workStateShort'}:
        raise ValueError('Unexpected candidate row-number / work-state binding')
    candidates = {}
    for key,column in zip(CANDIDATES,candidate_columns[1:-1]):
        candidates[key] = {'value': column['value']}
        if 'looks' in column:
            candidates[key]['looks'] = column['looks']
    actions = {}
    for section in sections:
        for button in section.get('buttons',[]):
            if button['action'] in ('updateRecords','deleteRecords'):
                if button['action'] in actions:
                    raise ValueError('The fixed screen has one update and one delete action')
                actions[button['action']] = button['job']
    if set(actions) != {'updateRecords','deleteRecords'}:
        raise ValueError('Both update and delete actions are required')
    result['screen'] = {
        'bindings': dict(zip(FIELDS, values)), 'judgments': screen['judgments'],
        'workState': screen['workState'], 'export': screen['export'],
        'candidates': candidates, 'actions': actions,
    }
    return result

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    converted = convert(json.loads(args.source.read_text(encoding='utf-8-sig')))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open('x', encoding='utf-8') as stream:
        json.dump(converted, stream, ensure_ascii=False, indent=2)
        stream.write('\n')
    print(args.output)
