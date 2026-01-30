
import os
import re
import argparse
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

class MarkdownToDocx:
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file
        self.doc = Document()
        self._setup_styles()
        
    def _setup_styles(self):
        """配置文档基本样式，支持中文"""
        style = self.doc.styles['Normal']
        style.font.name = '宋体'
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        style.font.size = Pt(12)
        
        # 设置标题样式
        for i in range(1, 10):
            if f'Heading {i}' in self.doc.styles:
                h_style = self.doc.styles[f'Heading {i}']
                h_style.font.name = '黑体'
                h_style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                h_style.font.color.rgb = RGBColor(0, 0, 0) # 黑色

    def parse_inline_styles(self, paragraph, text):
        """解析行内样式：**Bold**"""
        # 简单的加粗解析
        parts = re.split(r'(\*\*.*?\*\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.font.bold = True
                run.font.name = '宋体'
                run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            else:
                if part:
                    run = paragraph.add_run(part)
                    run.font.name = '宋体'
                    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    def convert(self):
        with open(self.input_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        table_mode = False
        table_rows = []
        code_mode = False

        for line in lines:
            line = line.strip()
            
            # 处理代码块
            if line.startswith('```'):
                code_mode = not code_mode
                continue
            
            if code_mode:
                p = self.doc.add_paragraph()
                p.style = 'No Spacing'
                run = p.add_run(line)
                run.font.name = 'Courier New'
                continue

            # 处理表格
            if line.startswith('|') and line.endswith('|'):
                if not table_mode:
                    table_mode = True
                    table_rows = []
                
                # 跳过分隔符行 |---|---|
                if '---' in line:
                    continue
                    
                row_data = [cell.strip() for cell in line.strip('|').split('|')]
                table_rows.append(row_data)
                continue
            else:
                if table_mode:
                    # 表格结束，写入表格
                    self._create_table(table_rows)
                    table_mode = False
                    table_rows = []

            if not line:
                continue

            # 处理标题
            header_match = re.match(r'^(#{1,6})\s+(.*)', line)
            if header_match:
                level = len(header_match.group(1))
                content = header_match.group(2)
                self.doc.add_heading(content, level=level)
                continue

            # 处理列表
            if line.startswith('- ') or line.startswith('* '):
                p = self.doc.add_paragraph(style='List Bullet')
                self.parse_inline_styles(p, line[2:])
                continue
            
            # 处理有序列表 (简单匹配 1. )
            if re.match(r'^\d+\.\s', line):
                content = re.sub(r'^\d+\.\s', '', line)
                p = self.doc.add_paragraph(style='List Number')
                self.parse_inline_styles(p, content)
                continue

            # 普通段落
            p = self.doc.add_paragraph()
            self.parse_inline_styles(p, line)

        # 如果文件以表格结尾
        if table_mode:
             self._create_table(table_rows)

        self.doc.save(self.output_file)
        print(f"✅ Converted: {self.output_file}")

    def _create_table(self, rows):
        if not rows:
            return
        
        num_cols = max(len(r) for r in rows)
        table = self.doc.add_table(rows=len(rows), cols=num_cols)
        table.style = 'Table Grid'
        
        for i, row_data in enumerate(rows):
            row_cells = table.rows[i].cells
            for j, cell_text in enumerate(row_data):
                if j < len(row_cells):
                    cell_cells = row_cells[j]
                    # cell_cells.text = cell_text # 这样会失去字体设置
                    # 使用 paragraph 添加以保持字体
                    p = cell_cells.paragraphs[0]
                    self.parse_inline_styles(p, cell_text)
                    
                    # 第一行加粗 (假设是表头)
                    if i == 0:
                        for run in p.runs:
                            run.font.bold = True

def main():
    parser = argparse.ArgumentParser(description='Convert Markdown to Docx (Custom)')
    parser.add_argument('input', help='Input Markdown file')
    parser.add_argument('output', help='Output Docx file')
    args = parser.parse_args()
    
    converter = MarkdownToDocx(args.input, args.output)
    converter.convert()

if __name__ == "__main__":
    main()
