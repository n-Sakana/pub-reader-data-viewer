"""Synthetic presentation fixtures. Does not read or write a business ledger."""
import copy
import re
import json5

BRIDGE = """window.rdvTestMessages=[];window.rdvHandlers=[];window.chrome=window.chrome||{};
window.chrome.webview={postMessage:m=>window.rdvTestMessages.push(m),
addEventListener:(name,h)=>window.rdvHandlers.push(h)};
window.rdvDeliver=m=>window.rdvHandlers.forEach(h=>h({data:m}));"""


def load_page(page, root, dialog=False, screen=None, state=None):
    """Load exact bundled CSS/JS without network or live WebView permissions."""
    html = (root / 'web/index.html').read_bytes().decode('utf-8-sig')
    html = re.sub(r'<script\b[^>]*\bsrc=[^>]*>\s*</script>', '', html, flags=re.I)
    html = re.sub(r'<link\b[^>]*\brel=["\']stylesheet["\'][^>]*>', '', html, flags=re.I)
    page.set_content(html)
    page.evaluate(BRIDGE)
    page.evaluate("d=>location.hash=d?'dialog':''", dialog)
    page.add_style_tag(content=(root / 'web/app.css').read_text(encoding='utf-8-sig'))
    page.add_script_tag(content=(root / 'web/app.js').read_text(encoding='utf-8-sig'))
    page.wait_for_function('!!window.rdvBridge')
    if screen is not None:
        page.evaluate('s=>window.rdvBridge.render(s)', screen)
    if state is not None:
        page.evaluate('s=>window.rdvBridge.state(s)', state)
    page.wait_for_timeout(35)
    page.evaluate('window.rdvTestMessages=[]')


def full_fixture(root):
    cfg = json5.loads((root / 'settings.json').read_text(encoding='utf-8-sig'))['screen']
    values = {}
    example = {
        'B.key1': '00016168', 'B.key2': '00023415', 'B.b_line': '001',
        'A.a_name': 'SAMPLE-A-0016168', 'A.a_code': 'A52903', 'A.a_grade': 'B2',
        'A.a_dept': 'D568', 'A.a_date': '2024/08/14', 'A.a_amount': '3,969,262',
        'A.a_rate': '0.1596', 'A.a_flag': 'N',
        'B.b_memo': '納品内容を確認済み。担当部門への照会は不要です。\n次回更新時に最新の作業状態を確認してください。',
        'C.c_remark': '定例処理対象 / 照合番号 RMK-449812', 'B.b_status': 'DONE',
    }
    states = {'searchKey': '00016168', 'workState': '未処理', 'pendingCount': '未送信 2 件',
              'appState': '監視中', 'watchLabel': 'メモ帳 接続中', 'ledgerFile': 'ReaderDataViewer-Ledger.xlsx', 'clock': '12:17:52'}
    def bind(source, name):
        if 'field' in source:
            text = example.get(source['field'], '---')
        elif 'fields' in source:
            text = source.get('joiner', ' ').join(example.get(f, '---') for f in source['fields'])
        else:
            text = states.get(source.get('state'), '')
        values[name] = {'text': text, 'tone': 0}
        return name
    def section(src, path):
        item = copy.deepcopy(src)
        item['id'] = path
        t = src['type']
        if t == 'keyPanel':
            item.update(label=src['figure']['label'], value=bind(src['figure']['value'], path+'.value'),
                        inputLabel=src['input']['label'], inputWidth=src['input']['width'],
                        maxLength=src['input']['maxLength'], placeholder=src['input']['placeholder'])
        elif t == 'columns':
            item['items'] = [section(s, path+'.item'+str(i)) for i, s in enumerate(src['items'])]
            item['stackBelow'] = src.get('stackBelow', 620)
        elif t == 'fieldList':
            item['rows'] = [{'label': r['label'], 'value': bind(r['value'], path+'.row'+str(i))} for i, r in enumerate(src['rows'])]
        elif t in ('textBox', 'sendBar'):
            item['value'] = bind(src['value'], path+'.value')
        elif t == 'statusBand':
            item['judgment'] = path
        elif t == 'statusBar':
            item['segments'] = [dict(r, value=bind(r['value'],path+'.segment'+str(i)), clock=False) for i,r in enumerate(src['segments'])]
        for btn in item.get('buttons', []):
            if btn['action']=='workState':
                btn['text']='未処理(&W)'
        return item
    card = copy.deepcopy(cfg['card'])
    padding = card['padding']
    if len(padding)==1:
        card['padding']=padding*4
    screen = {'card':card, 'sections':[section(s,'s'+str(i)) for i,s in enumerate(cfg['sections'])]}
    state = {'opsEnabled':True,'workEnabled':True,'key':'00016168','values':values,'workText':'未処理(&W)',
             'workDown':False,'pending':2,'judgments':{'s4':{'look':'ok','text':'OK','sub':'DONE ・ 未処理'}}}
    return screen, state


CANDIDATES = {
    'title':'候補一覧', 'hint':'行をクリックするとレコード詳細を表示します', 'width':744,
    'total':3, 'selected':0, 'selectable':True, 'maxHeight':168,'rowHeight':20,'headerHeight':20,
    'columns':[{'header':'番号2','width':100},{'header':'名称','width':220}, {'header':'状態','width':100,'render':'tag'}, {'header':'処理済','render':'tag'}],
    'rows':[['00023415','SAMPLE-C-0233',{'text':'DONE','look':'accent'},{'text':'未','look':'neutral'}],
            ['00023416','SAMPLE-C-0234',{'text':'HOLD','look':'outline'},{'text':'済','look':'accent'}],
            ['00023417','SAMPLE-C-0235',{'text':'VOID','look':'faded'},{'text':'未','look':'neutral'}]],
}
EXPORT = {'title':'テーブル出力', 'hint':'統合台帳から、選んだ項目だけを CSV に書き出します。',
          'destination':'C:/export/sample.csv', 'defaults':['T.id'],
          'fields':[{'ref':'T.id','label':'番号','type':'text'}, {'ref':'T.name','label':'名称','type':'text'},
                    {'ref':'T.date','label':'日付','type':'date','format':'yyyy-MM-dd'},
                    {'ref':'T.amount','label':'金額','type':'number'}]}
SETTINGS = {'title':'設定', 'hint':'場所・検索・監視対象を設定します。', 'dataDir':'data',
            'ledger':'data/ReaderDataViewer-Ledger.xlsx','log':'data/ReaderDataViewer.log',
            'pattern':'^[0-9]+$', 'candidateRows':8, 'target':{'summary':'メモ帳 / Edit','read':'UI Automation'}}
PROCESS = {'title':'レコード更新', 'hint':'設定に従ってデータを統合します。', 'inputTitle':'入力データ',
           'inputs':[{'id':'A','file':'tableA.csv','key':'key1','rows':'100','validation':'列一致','valid':True}],
           'steps':[['1','結合','表B','表A','番号1','左外部','中間1']],
           'output':['C:/ReaderDataViewer/data','ReaderDataViewer-Ledger.xlsx','2026/09/08 09:00'],
           'executeText':'実行', 'canRun':True}
MODALS = {'confirm':{'title':'確認','body':'表示中のレコードを処理済にします。よろしいですか？','ask':True},
          'candidates':CANDIDATES,'shared':dict(CANDIDATES,title='共有台帳の変更'),
          'sharedTell':dict(CANDIDATES,title='共有台帳のお知らせ'),
          'unmatched':{'title':'未送信の競合','body':'競合した未送信変更を保持します。','rows':[['1','00023415','作業状態の競合']]},
          'update':PROCESS,'delete':dict(PROCESS,title='レコード削除',executeText='削除する'),
          'settings':SETTINGS,'export':EXPORT}
