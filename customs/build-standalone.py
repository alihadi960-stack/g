#!/usr/bin/env python3
"""يبني customs/standalone.html: ملف واحد يتفتح بدبل كليك من غير سيرفر.
شغّله بعد أي تعديل في index.html أو app.js أو sheet.js أو styles.css:
    python3 customs/build-standalone.py
"""
import pathlib
import re

HERE = pathlib.Path(__file__).parent
html = (HERE / "index.html").read_text(encoding="utf-8")
css = (HERE.parent / "styles.css").read_text(encoding="utf-8")
sheet = (HERE / "sheet.js").read_text(encoding="utf-8")
app = (HERE / "app.js").read_text(encoding="utf-8")

sheet = re.sub(r"^export ", "", sheet, flags=re.M)
app, n = re.subn(r'^import \{[^}]*\} from "\./sheet\.js";\n', "", app, flags=re.M)
assert n == 1, "sheet.js import not found in app.js"

html = html.replace('<link rel="stylesheet" href="/styles.css" />', f"<style>\n{css}\n</style>")
html, n = re.subn(
    r'<script type="module" src="/customs/app\.js"></script>',
    lambda _: f'<script type="module">\n{sheet}\n{app}\n</script>',
    html,
)
assert n == 1, "app.js script tag not found in index.html"

(HERE / "standalone.html").write_text(html, encoding="utf-8")
print("wrote", HERE / "standalone.html")
