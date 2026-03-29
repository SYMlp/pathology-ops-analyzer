import json
import re

html_path = 'templates/病理科 运营分析月报 — 2026年2月.html'
with open(html_path, 'r', encoding='utf-8') as f:
    text = f.read()

# get style
style_match = re.search(r'<style>(.*?)</style>', text, re.DOTALL)
css = style_match.group(1).strip() if style_match else ''
css = css.replace('`', '\\`')

# Save just the style
with open('css_saved.txt', 'w', encoding='utf-8') as f:
    f.write(css)

# get script
script_match = re.search(r'<script>(.*?)</script>', text, re.DOTALL)
script_code = script_match.group(1).strip() if script_match else ''
script_code = script_code.replace('`', '\\`')
with open('script_saved.txt', 'w', encoding='utf-8') as f:
    f.write(script_code)

# get body HTML 
body_match = re.search(r'<body>\s*<div.*?>(.*?)</div>\s*<script>', text, re.DOTALL)
html_body = body_match.group(1).strip() if body_match else ''
with open('body_saved.html', 'w', encoding='utf-8') as f:
    f.write(html_body)

print("Saved css, script, body. Lengths:", len(css), len(script_code), len(html_body))
