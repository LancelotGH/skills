# coding=utf-8
"""
Markdown → HTML → PDF via Chrome headless.
Usage: python convert.py <file.md>
Requires: pandoc (in PATH), Google Chrome
"""
import os
import sys
import re
import subprocess
import tempfile

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }

@page {
    size: A4;
    margin: 20mm 18mm 20mm 18mm;
}

body {
    font-family: 'Microsoft YaHei', 'PingFang SC', 'Noto Sans SC', 'Hiragino Sans GB', sans-serif;
    font-size: 10.5pt;
    color: #1a1a1a;
    line-height: 1.7;
}

h1 {
    font-size: 20pt;
    font-weight: 700;
    color: #1a1a1a;
    border-bottom: 2px solid #2c3e50;
    padding-bottom: 6px;
    margin-top: 24px;
    margin-bottom: 14px;
}

h2 {
    font-size: 14pt;
    font-weight: 700;
    color: #2c3e50;
    border-left: 4px solid #3498db;
    padding-left: 10px;
    margin-top: 20px;
    margin-bottom: 10px;
}

h3 {
    font-size: 11.5pt;
    font-weight: 700;
    color: #2c3e50;
    margin-top: 16px;
    margin-bottom: 6px;
}

h4 {
    font-size: 10.5pt;
    font-weight: 700;
    color: #555;
    margin-top: 12px;
    margin-bottom: 4px;
}

p {
    margin-bottom: 6px;
}

ul, ol {
    padding-left: 1.4em;
    margin-bottom: 6px;
}

li {
    margin-bottom: 3px;
    line-height: 1.6;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin: 12px 0;
    font-size: 9.5pt;
}

thead tr {
    background-color: #dce8f5;
}

th {
    padding: 6px 10px;
    text-align: center;
    font-weight: 700;
    border: 1px solid #b0c8e0;
    color: #1a1a1a;
}

td {
    padding: 5px 10px;
    border: 1px solid #d0d0d0;
    text-align: left;
    vertical-align: top;
}

tr:nth-child(even) { background-color: #f7f9fc; }

blockquote {
    border-left: 3px solid #ccc;
    padding: 4px 12px;
    color: #555;
    margin: 8px 0;
    font-size: 9.5pt;
}

code {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 9pt;
    color: #c7254e;
    background: #fdf2f4;
    padding: 1px 4px;
    border-radius: 3px;
}

pre {
    background: #f6f8fa;
    border: 1px solid #e1e4e8;
    border-radius: 4px;
    padding: 10px 12px;
    margin: 8px 0;
    font-size: 9pt;
    line-height: 1.5;
    overflow-x: auto;
}

pre code {
    background: none;
    color: #333;
    padding: 0;
}

hr {
    border: none;
    border-top: 1px solid #e0e0e0;
    margin: 16px 0;
}

strong { font-weight: 700; }
a { color: #2980b9; text-decoration: none; }
"""

CHROME_CANDIDATES = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
]


def find_browser():
    for path in CHROME_CANDIDATES:
        if os.path.exists(path):
            return path
    raise RuntimeError(
        "未找到 Chrome 或 Edge。请确认已安装 Google Chrome 或 Microsoft Edge。"
    )


def md_to_html_body(md_path):
    result = subprocess.run(
        ['pandoc', md_path, '-f', 'markdown', '-t', 'html', '--no-highlight'],
        capture_output=True, text=True, encoding='utf-8'
    )
    if result.returncode != 0:
        print(f"pandoc 错误: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    return result.stdout


def colorize_markers(html):
    html = html.replace('✓', '✅')
    html = html.replace('✗', '❌')
    html = re.sub(r'<a\b[^>]*>(.*?)</a>', r'\1', html, flags=re.DOTALL)
    return html


def wrap_html(body, title=''):
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
{CSS}
</style>
</head>
<body>
{body}
</body>
</html>"""


def convert(md_path):
    md_path = os.path.abspath(md_path)
    if not md_path.endswith('.md'):
        print("请输入 .md 文件", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(md_path):
        print(f"文件不存在: {md_path}", file=sys.stderr)
        sys.exit(1)

    pdf_path = md_path[:-3] + '.pdf'
    title = os.path.basename(md_path[:-3])

    print(f"[1/3] Markdown → HTML: {os.path.basename(md_path)}")
    html_body = md_to_html_body(md_path)
    html_body = colorize_markers(html_body)
    full_html = wrap_html(html_body, title)

    with tempfile.NamedTemporaryFile(mode='w', suffix='.html',
                                     encoding='utf-8', delete=False) as f:
        f.write(full_html)
        tmp_html = f.name

    browser = find_browser()
    abs_pdf = os.path.abspath(pdf_path)
    abs_html = os.path.abspath(tmp_html)

    print(f"[2/3] HTML → PDF via {os.path.basename(os.path.dirname(browser))}")
    result = subprocess.run([
        browser,
        '--headless=new',
        '--disable-gpu',
        '--no-sandbox',
        '--run-all-compositor-stages-before-draw',
        f'--print-to-pdf={abs_pdf}',
        '--print-to-pdf-no-header',
        '--no-pdf-header-footer',
        f'file:///{abs_html.replace(os.sep, "/")}'
    ], capture_output=True)
    os.unlink(tmp_html)

    if result.returncode != 0:
        print(f"浏览器错误 (code {result.returncode})", file=sys.stderr)
        sys.exit(1)

    print(f"[3/3] 完成: {pdf_path}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python convert.py <文档.md>")
        sys.exit(1)
    for path in sys.argv[1:]:
        convert(path)
