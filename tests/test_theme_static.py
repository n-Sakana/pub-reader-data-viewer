#!/usr/bin/env python3
"""Source integrity and presentation checks; NOT a PowerShell/C# compiler."""
import argparse, datetime, hashlib, json, pathlib, re, subprocess, sys

def check(value, message):
    if not value: raise AssertionError(message)

def luminance(colour):
    text=colour.lstrip('#')
    if len(text)==3: text=''.join(x*2 for x in text)
    channels=[int(text[i:i+2],16)/255 for i in (0,2,4)]
    return sum((x/12.92 if x<=.04045 else ((x+.055)/1.055)**2.4)*w for x,w in zip(channels,(.2126,.7152,.0722)))

def contrast(a,b):
    low,high=sorted((luminance(a),luminance(b)))
    return (high+.05)/(low+.05)

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--root',type=pathlib.Path,default=pathlib.Path(__file__).resolve().parents[1])
    root=ap.parse_args().root.resolve(); results=[]; details={}
    def test(name,callback):
        try: callback(); results.append(dict(name=name,status='PASS'))
        except Exception as error: results.append(dict(name=name,status='FAIL',detail=str(error)))
        print(results[-1]['status'],name,results[-1].get('detail',''))
    manifest=json.loads((root/'design/preserved-files.json').read_text())
    for entry in manifest['files']:
        test('preserved:'+entry['path'],lambda e=entry:check(hashlib.sha256((root/e['path']).read_bytes()).hexdigest()==e['sha256'],'Uploaded baseline differs'))
    catalog=json.loads((root/'design/themes.json').read_text())
    ids=['win98','apple','material','fluent','carbon','spectrum']
    def catalog_check():
        check(catalog['schema']==1,'catalog schema')
        check([x['id'] for x in catalog['themes']]==ids,'exact six design IDs/order')
        for item in catalog['themes']:
            check(item['modern']==(item['id']!='win98'),'mode mismatch')
            for key in ('background','caption','captionText','border'): check(re.fullmatch('#[a-fA-F0-9]{6}',item[key]),'invalid colour')
    test('six-theme-catalog',catalog_check)
    def defaults():
        html=(root/'web/index.html').read_text(); cfg=json.loads((root/'web/theme.json').read_text())
        check(cfg['id']=='win98' and cfg['motion']=='off','classic default config')
        for key,value in [('theme','win98'),('modern','false'),('motion','off')]:
            check(re.findall('data-rdv-'+key+'="([^"]+)"',html)==[value],'theme marker mismatch')
        check('<link rel="stylesheet" href="themes.css">' in html,'CSS not loaded')
        check('src="theme-motion.js"' in html,'motion not loaded')
    test('classic-default-theme-markers',defaults)
    css=(root/'web/themes.css').read_text()
    base=dict(re.findall(r'(--ui-[\w-]+):([^;]+)',css.split('}')[0]))
    def colours(theme):
        m=re.search(r'html\[data-rdv-theme="'+theme+r'"\] \{([^}]+)',css)
        values=base|dict(re.findall(r'(--ui-[\w-]+):([^;]+)',m[1] if m else ''))
        def resolve(key):
            value=values[key]
            return resolve(value[4:-1]) if value.startswith('var(') else value
        pairs=[('body','--ui-ink','--ui-surface'),('muted','--ui-muted','--ui-subtle'),('primary','--ui-on-accent','--ui-accent'),('selected-label','--ui-accent-text','--ui-tint')]
        ratios={label:round(contrast(resolve(a),resolve(b)),3) for label,a,b in pairs}
        check(all(n>=4.5 for n in ratios.values()),str(ratios))
        details[theme+'-sampled-text-contrast']=ratios
    for theme in ids[1:]: test(theme+':sampled-text-contrast',lambda t=theme:colours(t))
    def motion_source():
        source=(root/'web/theme-motion.js').read_text()
        run=subprocess.run(['node','--check',str(root/'web/theme-motion.js')],capture_output=True,text=True,timeout=15)
        check(run.returncode==0,run.stderr)
        check(not re.search(r'\b(setTimeout|setInterval|fetch|XMLHttpRequest|postMessage)\s*\(',source),'motion contains a business/network/timer side effect')
        check('prefers-reduced-motion' in source and 'forced-colors' in source,'preference hooks missing')
        check('Math.min(160' in source,'duration bound missing')
    test('presentation-script-syntax-and-boundaries',motion_source)
    def local_assets():
        check('@import' not in css,'external style import')
        for url in re.findall(r'url\((.*?)\)',css): check(url.strip('"\'').startswith('data:'),'non-local theme asset')
    test('no-new-network-assets',local_assets)
    def scripts():
        for name in ('tools/Build.ps1','tests/Test-Build.ps1'):
            raw=(root/name).read_bytes()
            check(raw.startswith(b'\xef\xbb\xbf'),'UTF-8 BOM required for Windows PowerShell')
            check(b'\n' not in raw.replace(b'\r\n',b''),'Windows script line endings')
        batch=(root/'build.bat').read_bytes();batch.decode('ascii')
        check(b'\n' not in batch.replace(b'\r\n',b''),'batch line endings')
    test('windows-script-encodings',scripts)
    report={'scope':'Source fingerprints, configuration, sampled contrast and JS syntax. Not native execution or a full accessibility audit.','utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'details':details,'tests':results,'passed':sum(x['status']=='PASS' for x in results),'failed':sum(x['status']=='FAIL' for x in results)}
    (root/'tests/results/theme-static-results.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(report['passed'],'passed;',report['failed'],'failed')
    return bool(report['failed'])
if __name__=='__main__': sys.exit(main())
