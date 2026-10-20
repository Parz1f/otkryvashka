"""Turns the web page (web/otkryvashka.html) into the app's start page.

The page is written for the browser, where libraries and fonts come from CDNs.
In the app they ship inside the APK: the page loads libs/ itself when it sees
it runs in the app, and this script swaps the Google Fonts link for local files.
Run it after editing web/otkryvashka.html:  python3 tools/build_www.py
"""
import pathlib
import re

root = pathlib.Path(__file__).resolve().parent.parent
src = (root / 'web' / 'otkryvashka.html').read_text(encoding='utf-8')
out = root / 'app' / 'src' / 'main' / 'assets' / 'www' / 'index.html'

FONTS = [('Golos Text', 'golos-text', w) for w in (400, 500, 600)] + \
        [('Unbounded', 'unbounded', 600)] + \
        [('JetBrains Mono', 'jetbrains-mono', w) for w in (400, 500)]
CYRILLIC = 'U+0301,U+0400-045F,U+0490-0491,U+04B0-04B1,U+2116'
faces = []
for family, file, weight in FONTS:
    for subset, rng in (('cyrillic', CYRILLIC), ('latin', 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD')):
        faces.append(f"@font-face{{font-family:'{family}';font-style:normal;font-weight:{weight};font-display:swap;"
                     f"src:url(fonts/{file}-{subset}-{weight}-normal.woff2) format('woff2');unicode-range:{rng}}}")

page = re.sub(r'<link rel="preconnect"[^>]*>\n', '', src)
page = re.sub(r'<link rel="stylesheet" href="https://fonts\.googleapis\.com[^>]*>', '<style>' + ''.join(faces) + '</style>', page)
page = page.replace('<meta name="claude:viewport-fit-sides" content="cover">\n', '')

head = ('<!doctype html>\n<html lang="ru">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '<style>:root{color-scheme:light}body{margin:0}img{max-width:100%}</style>\n')
title_end = page.index('</title>') + len('</title>')
body_start = page.index('<header')
html = head + page[:body_start] + '</head>\n<body>\n' + page[body_start:] + '\n</body>\n</html>\n'
out.write_text(html, encoding='utf-8')
# Своя отрисовка слайдов лежит рядом со страницей
(out.parent / 'slides.js').write_text((root / 'web' / 'slides.js').read_text(encoding='utf-8'), encoding='utf-8')
print('wrote', out.relative_to(root))
