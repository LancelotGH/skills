#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读取和解析Word文档内容，包括图片提取"""

import sys
import json
import os
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml import parse_xml
from docx.oxml.ns import qn

def extract_images(doc, output_dir=None):
    """提取文档中的所有图片"""
    images = []
    
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 遍历文档中的所有关系（包括图片）
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_data = {
                "id": rel.rId,
                "filename": os.path.basename(rel.target_ref),
                "type": rel.target_ref.split('.')[-1]
            }
            
            # 如果指定了输出目录，保存图片
            if output_dir:
                image_path = os.path.join(output_dir, image_data["filename"])
                with open(image_path, 'wb') as f:
                    f.write(rel.target_part.blob)
                image_data["saved_path"] = image_path
            
            images.append(image_data)
    
    return images

def find_images_in_paragraph(para):
    """查找段落中的图片"""
    images_in_para = []
    
    # 查找drawing元素（图片通常在这里）
    for run in para.runs:
        for drawing in run.element.findall('.//wp:inline', 
                                          {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}):
            # 查找图片的blip元素
            blip = drawing.find('.//a:blip', 
                               {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if blip is not None:
                embed_id = blip.get(qn('r:embed'))
                if embed_id:
                    images_in_para.append(embed_id)
    
    return images_in_para

def read_docx(file_path, max_paragraphs=300, max_tables=50, extract_images_flag=True, image_output_dir=None):
    """
    读取Word文档并输出为结构化格式
    
    参数:
        file_path: Word文件路径
        max_paragraphs: 最大读取段落数（默认300）
        max_tables: 最大读取表格数（默认50）
        extract_images_flag: 是否提取图片（默认True）
        image_output_dir: 图片保存目录（默认None，不保存）
    """
    try:
        doc = Document(file_path)
        
        # 提取所有图片
        all_images = []
        if extract_images_flag:
            all_images = extract_images(doc, image_output_dir)
        
        result = {
            "file": file_path,
            "total_images": len(all_images),
            "images": all_images,
            "content": []
        }
        
        para_count = 0
        table_count = 0
        
        for element in doc.element.body:
            # 读取段落
            if isinstance(element, CT_P):
                if para_count >= max_paragraphs:
                    result["content"].append({
                        "type": "note",
                        "text": f"... 已省略剩余段落（超过{max_paragraphs}个）"
                    })
                    break
                
                para = Paragraph(element, doc)
                text = para.text.strip()
                
                # 检查段落中是否有图片
                images_in_para = find_images_in_paragraph(para)
                
                # 判断是否为标题
                if para.style.name.startswith('Heading'):
                    content_item = {
                        "type": "heading",
                        "level": para.style.name,
                        "text": text if text else "[空标题]"
                    }
                    if images_in_para:
                        content_item["has_images"] = True
                        content_item["image_ids"] = images_in_para
                    result["content"].append(content_item)
                elif text or images_in_para:  # 只添加有文本或有图片的段落
                    content_item = {
                        "type": "paragraph",
                        "text": text if text else "[段落仅含图片]"
                    }
                    if images_in_para:
                        content_item["has_images"] = True
                        content_item["image_ids"] = images_in_para
                        content_item["text"] = text if text else f"[图片段落，包含{len(images_in_para)}张图片]"
                    result["content"].append(content_item)
                
                para_count += 1
            
            # 读取表格
            elif isinstance(element, CT_Tbl):
                if table_count >= max_tables:
                    result["content"].append({
                        "type": "note",
                        "text": f"... 已省略剩余表格（超过{max_tables}个）"
                    })
                    break
                
                table = Table(element, doc)
                table_data = {
                    "type": "table",
                    "rows": len(table.rows),
                    "cols": len(table.columns),
                    "data": []
                }
                
                # 读取表格内容（最多30行）
                for row_idx, row in enumerate(table.rows[:30]):
                    row_data = [cell.text.strip() for cell in row.cells]
                    table_data["data"].append(row_data)
                
                if len(table.rows) > 30:
                    table_data["note"] = f"表格共{len(table.rows)}行，仅显示前30行"
                
                result["content"].append(table_data)
                table_count += 1
        
        return result
        
    except Exception as e:
        return {
            "error": str(e),
            "file": file_path
        }

def format_output(data, format_type="markdown"):
    """
    格式化输出
    
    参数:
        data: 读取的数据
        format_type: 输出格式 (json/markdown)
    """
    if "error" in data:
        return f"错误: {data['error']}"
    
    if format_type == "json":
        return json.dumps(data, ensure_ascii=False, indent=2)
    
    elif format_type == "markdown":
        output = [f"# Word文档分析: {os.path.basename(data['file'])}\n"]
        
        # 图片信息摘要
        if data["total_images"] > 0:
            output.append(f"## 📷 文档包含 {data['total_images']} 张图片\n")
            for idx, img in enumerate(data["images"][:10], 1):  # 只显示前10张
                output.append(f"{idx}. {img['filename']} (ID: {img['id']}, 类型: {img['type']})")
                if "saved_path" in img:
                    output.append(f"   - 已保存到: {img['saved_path']}")
            if len(data["images"]) > 10:
                output.append(f"\n... 还有 {len(data['images']) - 10} 张图片")
            output.append("\n---\n")
        
        # 文档内容
        output.append("## 📄 文档内容\n")
        
        for item in data["content"]:
            if item["type"] == "heading":
                # 根据标题级别添加#
                level = int(item["level"][-1]) if item["level"][-1].isdigit() else 2
                output.append(f"\n{'#' * (level + 1)} {item['text']}")
                if item.get("has_images"):
                    output.append(f" 📷[含{len(item.get('image_ids', []))}张图片]")
                output.append("\n")
            
            elif item["type"] == "paragraph":
                if item.get("has_images"):
                    output.append(f"\n📷 **[图片段落]** {item['text']}\n")
                elif item["text"]:
                    output.append(f"{item['text']}\n")
            
            elif item["type"] == "table":
                output.append(f"\n**表格** ({item['rows']}行 × {item['cols']}列):\n")
                
                if len(item["data"]) > 0:
                    # 表头
                    header = item["data"][0]
                    output.append("| " + " | ".join(header) + " |")
                    output.append("| " + " | ".join(["---"] * len(header)) + " |")
                    
                    # 数据行
                    for row in item["data"][1:]:
                        output.append("| " + " | ".join(row) + " |")
                
                if "note" in item:
                    output.append(f"\n*{item['note']}*\n")
            
            elif item["type"] == "note":
                output.append(f"\n*{item['text']}*\n")
        
        return "\n".join(output)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python read_docx.py <word文件路径> [输出格式:json|markdown] [图片保存目录]")
        sys.exit(1)
    
    file_path = sys.argv[1]
    format_type = sys.argv[2] if len(sys.argv) > 2 else "markdown"
    image_dir = sys.argv[3] if len(sys.argv) > 3 else None
    
    data = read_docx(file_path, extract_images_flag=True, image_output_dir=image_dir)
    output = format_output(data, format_type)
    print(output)
