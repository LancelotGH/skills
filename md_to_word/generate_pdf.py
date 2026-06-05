# coding=utf-8
"""
Markdown → HTML → PDF via weasyprint.
Usage: python generate_pdf.py <file.md>
"""
import os
import sys
import re
import subprocess
import tempfile

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');

* { box-sizing: border-box; margin: 0; padding: 0; }

@page {
    size: A4;
    margin: 20mm 18mm 20mm 18mm;
}

body {
    font-family: 'Noto Sans SC', 'Microsoft YaHei', 'PingFang SC', sans-serif;
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

/* Tables */
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

/* Blockquotes */
blockquote {
    border-left: 3px solid #ccc;
    padding: 4px 12px;
    color: #555;
    margin: 8px 0;
    font-size: 9.5pt;
}

/* Inline code */
code {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 9pt;
    color: #c7254e;
    background: #fdf2f4;
    padding: 1px 4px;
    border-radius: 3px;
}

/* Code blocks */
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

/* Horizontal rule */
hr {
    border: none;
    border-top: 1px solid #e0e0e0;
    margin: 16px 0;
}

/* Bold / Strong */
strong { font-weight: 700; }

a { color: #2980b9; text-decoration: none; }
"""

def md_to_html(md_path):
    """Convert markdown to HTML body via pandoc."""
    result = subprocess.run(
        ['pandoc', md_path, '-f', 'markdown', '-t', 'html', '--no-highlight'],
        capture_output=True, text=True, encoding='utf-8'
    )
    if result.returncode != 0:
        print(f"Pandoc error: {result.stderr}")
        sys.exit(1)
    return result.stdout

def colorize_markers(html):
    """Replace ✓/✗ with emoji, strip hyperlinks to plain text."""
    html = html.replace('✓', '✅')
    html = html.replace('✗', '❌')
    # Remove <a href="...">text</a> → text
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

def convert_to_pdf(md_path):
    if not md_path.endswith('.md'):
        print("请输入 .md 文件")
        sys.exit(1)

    base = md_path[:-3]
    pdf_path = base + '.pdf'
    title = os.path.basename(base)

    print(f"[1/3] 转换 Markdown → HTML: {md_path}")
    html_body = md_to_html(md_path)

    print(f"[2/3] 处理样式与标记着色...")
    html_body = colorize_markers(html_body)
    full_html = wrap_html(html_body, title)

    with tempfile.NamedTemporaryFile(mode='w', suffix='.html',
                                     encoding='utf-8', delete=False) as f:
        f.write(full_html)
        tmp_html = f.name

    CHROME = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    abs_pdf = os.path.abspath(pdf_path)
    abs_html = os.path.abspath(tmp_html)

    print(f"[3/3] 生成 PDF: {abs_pdf}")
    result = subprocess.run([
        CHROME,
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
        print(f"Chrome error (code {result.returncode})")
        sys.exit(1)

    print(f"\n完成！已生成 {pdf_path}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python generate_pdf.py <文档.md>")
        sys.exit(1)
    for path in sys.argv[1:]:
        convert_to_pdf(path)
