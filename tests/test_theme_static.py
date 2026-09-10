#!/usr/bin/env python3
"""Source integrity and presentation checks; NOT a PowerShell/C# compiler."""
import argparse, datetime, hashlib, json, pathlib, re, subprocess, sys

def check(value, message):
    if not value: raise AssertionError(message)

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--root',type=pathlib.Path,default=pathlib.Path(__file__).resolve().parents[1])
    root=ap.parse_args().root.resolve(); results=[]; details={}
    def test(name,callback):
        try: callback(); results.append(dict(name=name,status='PASS'))
        except Exception as error: results.append(dict(name=name,status='FAIL',detail=str(error)))
        print(results[-1]['status'],name,results[-1].get('detail',''))
    manifest=json.loads((root/'design/preserved-files.json').read_text())
    tracked=set(subprocess.check_output(['git','-C',str(root),'ls-files','-z']).decode('utf-8').split('\0'))
    # Runtime data changes independently of the delivered source revision.
    details['excludedFingerprints']=[entry['path'] for entry in manifest['files']
                                     if entry['path'].startswith('data/') or entry['path'] not in tracked]
    def fingerprint(entry):
        data = (root/entry['path']).read_bytes()
        if entry.get('normalization') == 'crlf-to-lf':
            data = data.replace(b'\r\n', b'\n')
        return hashlib.sha256(data).hexdigest()
    for entry in manifest['files']:
        if entry['path'] in details['excludedFingerprints']: continue
        test('preserved:'+entry['path'],lambda e=entry:check(fingerprint(e)==e['sha256'],'Uploaded baseline differs'))
    def win98_only():
        for name in ('design/themes.json','web/theme.json','web/themes.css','web/theme-motion.js','src/03_Theme.cs'):
            check(not (root/name).exists(),'Retired presentation file remains: '+name)
        html=(root/'web/index.html').read_text(encoding='utf-8-sig')
        check(not any(name in html for name in ('data-rdv-', 'themes.css', 'theme-motion.js')),'Retired presentation asset referenced')
        native=(root/'src/03_Win98.cs').read_text()
        check('Color.FromArgb(212, 208, 200)' in native,'Win98 background changed')
        for colour in ('Caption = 0x501b08','CaptionText = 0xffffff','Border = 0x808080'):
            check(colour in native,'Win98 caption colour changed: '+colour)
        host=(root/'src/02_MainWindow.cs').read_text()
        check('RdvTheme' not in host and 'surfaceShown' not in host,'Theme branch remains in native host')
    test('win98-only-presentation',win98_only)
    def no_motion():
        html=(root/'web/index.html').read_text(encoding='utf-8-sig')
        css=(root/'web/app.css').read_text(encoding='utf-8-sig')
        styles='\n'.join(re.findall(r'<style[^>]*>(.*?)</style>',html,re.S))+css
        check(not re.search(r'@keyframes\b',styles,re.I),'Unexpected Win98 keyframes')
        for value in re.findall(r'\b(?:animation|transition)\s*:\s*([^;}]+)',styles,re.I):
            check(re.sub(r'\s*!important\s*$', '', value, flags=re.I).strip().lower()=='none',
                  'Unexpected Win98 motion: '+value)
    test('classic-no-presentation-motion',no_motion)
    def local_assets():
        html=(root/'web/index.html').read_text(encoding='utf-8-sig')
        css='\n'.join(re.findall(r'<style[^>]*>(.*?)</style>',html,re.S))+(root/'web/app.css').read_text(encoding='utf-8-sig')
        check('@import' not in css,'external style import')
        for url in re.findall(r'url\((.*?)\)',css): check(url.strip("\"'").startswith('data:'),'non-local style asset')
        for url in re.findall(r"""<(?:script|link)\b[^>]*(?:src|href)=["']([^"']+)""",html,re.I):
            check(not re.match(r'(?:[a-z]+:)?//',url,re.I),'non-local script or stylesheet')
    test('no-new-network-assets',local_assets)
    def scripts():
        for name in ('tools/Build.ps1','tests/Test-Build.ps1'):
            raw=(root/name).read_bytes()
            check(raw.startswith(b'\xef\xbb\xbf'),'UTF-8 BOM required for Windows PowerShell')
            check(b'\n' not in raw.replace(b'\r\n',b''),'Windows script line endings')
        batch=(root/'build.bat').read_bytes();batch.decode('ascii')
        check(b'\n' not in batch.replace(b'\r\n',b''),'batch line endings')
    test('windows-script-encodings',scripts)
    report={'scope':'Source fingerprints, fixed Win98 presentation and local assets. Not native execution or a full accessibility audit.','utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'details':details,'tests':results,'passed':sum(x['status']=='PASS' for x in results),'failed':sum(x['status']=='FAIL' for x in results)}
    (root/'tests/results/theme-static-results.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(report['passed'],'passed;',report['failed'],'failed')
    return bool(report['failed'])
if __name__=='__main__': sys.exit(main())
