"""
convert.py - Markdown to HTML converter
Usage: python convert.py <input.md>
Requires: pandoc (in PATH)
"""

import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path


def text_to_id(text: str) -> str:
    text = re.sub(r'^\d[\d.]*\s+', '', text.strip())
    text = ''.join(c for c in text if not unicodedata.category(c).startswith(('P', 'S')))
    return re.sub(r'\s+', '', text)


def fix_heading_ids(html: str) -> str:
    heading_pat = re.compile(
        r'<(h[1-6])\s+id="([^"]*)">(.*?)</h[1-6]>',
        re.DOTALL
    )
    id_map = {}
    for m in heading_pat.finditer(html):
        old_id = m.group(2)
        raw = re.sub(r'<[^>]+>', '', m.group(3))
        new_id = text_to_id(raw)
        if new_id and new_id != old_id:
            id_map[old_id] = new_id

    if not id_map:
        return html

    def replace_heading(m):
        tag, old_id, inner = m.group(1), m.group(2), m.group(3)
        new_id = id_map.get(old_id, old_id)
        return f'<{tag} id="{new_id}">{inner}</{tag}>'

    html = heading_pat.sub(replace_heading, html)

    def replace_href(m):
        old_id = m.group(1)
        return f'href="#{id_map.get(old_id, old_id)}"'

    return re.sub(r'href="#([^"]+)"', replace_href, html)


def fix_title(html: str) -> str:
    if re.search(r'<title>[^<]+</title>', html):
        return html
    m = re.search(r'<h1[^>]*>(.*?)</h1>', html, re.DOTALL)
    if m:
        title = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        html = re.sub(r'<title></title>', f'<title>{title}</title>', html)
    return html


def main():
    if len(sys.argv) < 2:
        print('Usage: python convert.py <input.md>', file=sys.stderr)
        sys.exit(1)

    input_path = Path(sys.argv[1]).resolve()
    if not input_path.exists():
        print(f'File not found: {input_path}', file=sys.stderr)
        sys.exit(1)

    if not shutil.which('pandoc'):
        print('pandoc not found. Install from https://pandoc.org', file=sys.stderr)
        sys.exit(1)

    output_path = input_path.with_suffix('.html')
    script_dir = Path(__file__).parent
    template = script_dir / 'template.html'

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_input = Path(tmpdir) / 'input.md'
        tmp_output = Path(tmpdir) / 'output.html'

        shutil.copy2(input_path, tmp_input)

        result = subprocess.run(
            ['pandoc', str(tmp_input), '-o', str(tmp_output),
             '--template', str(template), '--toc', '--toc-depth=3'],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print(f'pandoc error:\n{result.stderr}', file=sys.stderr)
            sys.exit(result.returncode)

        html = tmp_output.read_text(encoding='utf-8')

    html = fix_heading_ids(html)
    html = fix_title(html)

    output_path.write_text(html, encoding='utf-8')
    print(output_path)


if __name__ == '__main__':
    main()
