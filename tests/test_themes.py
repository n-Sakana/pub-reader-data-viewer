#!/usr/bin/env python3
"""Win98 browser regression matrix. Synthetic bridge; NOT native Windows UI."""
import argparse
import copy
import datetime
import io
import json
import pathlib
import sys
from PIL import Image, ImageChops
from playwright.sync_api import sync_playwright
import test_web as old
from theme_support import load_page, full_fixture, MODALS, CANDIDATES, EXPORT

REGRESSIONS = [
 ('input-normalization',old.input_normalization),('IME-main',old.ime_main),('enter-once',old.enter_once),
 ('max-length',old.max_length),('duplicate-job-routing',old.duplicate_actions),('disabled-actions',old.disabled_actions),
 ('IME-confirm',old.ime_modal),('conflicts-keep-escape-html',lambda p:old.unmatched(p,False)),
 ('conflicts-discard',lambda p:old.unmatched(p,True)),('esc-keeps-pending',old.escape_keeps),
 ('export-safe-default',lambda p:old.export(p,True)),('export-raw',lambda p:old.export(p,False)),
 ('export-reject-non-csv',lambda p:old.export(p,True,'xlsx')),('rerender-ids',old.rerender_ids)]


def check(condition,message):
    if not condition: raise AssertionError(message)


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--root',type=pathlib.Path,default=pathlib.Path(__file__).resolve().parents[1])
    ap.add_argument('--original',type=pathlib.Path)
    ap.add_argument('--chromium',default='/usr/bin/chromium')
    ap.add_argument('--screenshots',action='store_true')
    ap.add_argument('--theme',choices=['win98'])
    args=ap.parse_args(); root=args.root.resolve()
    screen,state=full_fixture(root)
    results=[]
    def test(name,callback):
        try:
            callback();results.append({'name':name,'status':'PASS'})
        except Exception as e:
            results.append({'name':name,'status':'FAIL','detail':str(e)})
        print(results[-1]['status'],name,results[-1].get('detail',''),flush=True)
    with sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=args.chromium,headless=True,args=['--no-sandbox'])
        def on_page(theme,callback,*,width=1100,height=900,dialog=False,full=False,motion='auto',reduced=False,forced=False):
            page=browser.new_page(viewport={'width':width,'height':height},reduced_motion='reduce' if reduced else 'no-preference',forced_colors='active' if forced else 'none')
            page.set_default_timeout(2500)
            errors=[]; page.on('pageerror',lambda e:errors.append(str(e)))
            try:
                load_page(page,root,dialog,screen if full else old.SCREEN,state if full else old.STATE)
                callback(page)
                check(not errors,'Browser errors: '+str(errors))
            finally:
                page.close()
        def layout(page):
            data=page.evaluate('''()=>{const c=document.querySelector('.client');return {client:c.clientWidth,scroll:c.scrollWidth,clipped:[...document.querySelectorAll('.win .btn,.win .tog')].filter(e=>e.getBoundingClientRect().width>0 && (e.scrollWidth>e.clientWidth+2 || e.scrollHeight>e.clientHeight+2)).map(e=>e.textContent)};}''')
            check(data['scroll']<=data['client']+1,'Horizontal content overflow: '+str(data))
            check(not data['clipped'],'Clipped button text: '+str(data))
            page.locator('[data-action="sendChanges"]').scroll_into_view_if_needed()
            page.locator('[data-action="sendChanges"]').click()
            check(old.messages(page,'action')[-1]['name']=='sendChanges','send is inaccessible')
        def dialog_check(page,kind,content):
            old.modal(page,kind,content)
            old.check(page.locator('.veil.show .dlg').get_attribute('role')=='dialog','dialog semantic lost')
            info=page.locator('.veil.show .dlg').evaluate('''e=>({width:e.getBoundingClientRect().width,height:e.getBoundingClientRect().height,scroll:e.scrollWidth,client:e.clientWidth})''')
            check(info['scroll']<=info['client']+2,'Dialog horizontal overflow: '+str(info))
            old.clear(page)
            page.keyboard.press('Escape')
            check(len(old.messages(page,'modalResult'))==1,'Esc did not resolve exactly once')
        def motion(page,expected):
            duration=page.locator('#b-search').evaluate('e=>getComputedStyle(e).transitionDuration')
            nums=[float(p.strip().rstrip('s')) for p in duration.split(',')]
            check((max(nums)>0)==expected,'Motion preference incorrect: '+duration)
            check(max(nums)<=.16,'Long animation')
            old.modal(page,'confirm',MODALS['confirm'])
            old.clear(page)
            # Immediately after reveal: action is never animation-gated.
            page.locator('.veil.show [data-modal-default=true]').evaluate('e=>e.click()')
            check(len(old.messages(page,'modalResult'))==1,'Animation blocked confirmation')
        def candidates_keyboard(page):
            old.modal(page,'candidates',CANDIDATES)
            page.locator('.veil.show tbody tr').nth(0).focus()
            page.keyboard.press('ArrowDown')
            check(page.locator('.veil.show tbody tr').nth(1).get_attribute('aria-selected')=='true','Candidate selection changed')
            old.clear(page);page.keyboard.press('Enter')
            check(old.messages(page,'modalResult')[-1]['result']['index']==1,'Candidate index changed')
        for theme in ['win98']:
            for label,callback in REGRESSIONS:
                test(theme+':'+label,lambda t=theme,c=callback:on_page(t,c))
            for width,height in [(818,636),(480,640),(1100,900)]:
                test(theme+f':layout-{width}',lambda t=theme,w=width,h=height:on_page(t,layout,width=w,height=h,full=True))
            for kind,content in MODALS.items():
                test(theme+':native-dialog-'+kind,lambda t=theme,k=kind,c=content:on_page(t,lambda p:dialog_check(p,k,c),width=1000,height=1000,dialog=True,full=True))
            test(theme+':candidate-keyboard',lambda t=theme:on_page(t,candidates_keyboard))
            test(theme+':no-motion',lambda t=theme:on_page(t,lambda p:motion(p,False)))
            if args.screenshots:
                preview=root/'docs/previews';preview.mkdir(parents=True,exist_ok=True)
                def capture(page,t=theme):
                    page.locator('#input').evaluate('e=>e.blur()')
                    page.screenshot(path=str(preview/(t+'.png')))
                    old.modal(page,'candidates',CANDIDATES)
                    page.wait_for_timeout(180)
                    page.screenshot(path=str(preview/(t+'-dialog.png')))
                on_page(theme,capture,width=920,height=780,full=True)
        if args.original:
            def classic_pixels():
                shots=[]
                for folder,original in [(args.original,True),(root,False)]:
                    page=browser.new_page(viewport={'width':920,'height':780})
                    try:
                        load_page(page,folder,False,screen,state)
                        page.locator('#input').evaluate('e=>e.blur()')
                        page.wait_for_timeout(100)
                        shots.append(Image.open(io.BytesIO(page.screenshot())).convert('RGB'))
                    finally: page.close()
                diff=ImageChops.difference(shots[0],shots[1]);box=diff.getbbox()
                if box:
                    diff.save(str(root/'tests/results/classic-pixel-diff.png'))
                check(box is None,'Classic differs from original pixels: '+str(box))
            test('win98:pixel-identical-to-original',classic_pixels)
        browser_version=browser.version;browser.close()
    output={'scope':'Chromium with synthetic WebView bridge, not native C#/WPF/SMB','browser':browser_version,
            'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'tests':results,
            'passed':sum(r['status']=='PASS' for r in results),'failed':sum(r['status']=='FAIL' for r in results)}
    (root/'tests/results'/('theme-'+args.theme+'-results.json' if args.theme else 'theme-results.json')).write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n')
    print(output['passed'],'passed;',output['failed'],'failed',flush=True)
    return bool(output['failed'])
if __name__=='__main__':sys.exit(main())
