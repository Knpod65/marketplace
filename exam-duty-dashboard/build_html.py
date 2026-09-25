"""Inject data.json into template.html -> index.html.

Usage: python3 build_html.py data.json [index.html]
"""
import json, sys
data = json.load(open(sys.argv[1], encoding='utf-8'))
out = sys.argv[2] if len(sys.argv) > 2 else 'index.html'
tpl = open('template.html', encoding='utf-8').read()
payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
open(out, 'w', encoding='utf-8').write(tpl.replace('/*__DATA__*/null', payload))
print('wrote', out)
