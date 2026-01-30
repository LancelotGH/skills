#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读取和解析Word文档内容，包括图片提取和上下文关联"""

import sys
import json
import os
import io
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

# 设置标准输出为UTF-8编码
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def extract_images(doc, output_dir=None):
    """提取文档中的所有图片"""
    images = []
    
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_data = {
                "id": rel.rId,
                "filename": os.path.basename(rel.target_ref),
                "type": rel.target_ref.split('.')[-1]
            }
            
            if output_dir:
                image_path = os.path.join(output_dir, image_data["filename"])
                with open(image_path, 'wb') as f:
                    f.write(rel.target_part.blob)
                image_data["saved_path"] = image_path
            
            images.append(image_data)
    
    return images

def find_images_in_paragraph(para):
    """查找段落中的图片ID"""
    images_in_para = []
    
    for run in para.runs:
        for drawing in run.element.findall('.//wp:inline', 
                                          {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}):
            blip = drawing.find('.//a:blip', 
                               {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if blip is not None:
                embed_id = blip.get(qn('r:embed'))
                if embed_id:
                    images_in_para.append(embed_id)
    
    return images_in_para

def read_docx_enhanced(file_path, max_paragraphs=500, max_tables=50, 
                      context_before=2, context_after=2, 
                      extract_images_flag=True, image_output_dir=None):
    """
    增强版Word文档读取，提取图片及其上下文
    
    参数:
        file_path: Word文件路径
        max_paragraphs: 最大读取段落数
        max_tables: 最大读取表格数
        context_before: 图片前的上下文段落数
        context_after: 图片后的上下文段落数
        extract_images_flag: 是否提取图片
        image_output_dir: 图片保存目录
    """
    try:
        doc = Document(file_path)
        
        # 提取所有图片
        all_images = []
        if extract_images_flag:
            all_images = extract_images(doc, image_output_dir)
        
        # 第一遍：收集所有内容
        all_content = []
        para_count = 0
        table_count = 0
        
        for element in doc.element.body:
            if isinstance(element, CT_P):
                if para_count >= max_paragraphs:
                    break
                    
                para = Paragraph(element, doc)
                text = para.text.strip()
                images_in_para = find_images_in_paragraph(para)
                
                if para.style.name.startswith('Heading'):
                    all_content.append({
                        "type": "heading",
                        "level": para.style.name,
                        "text": text if text else "[空标题]",
                        "has_images": len(images_in_para) > 0,
                        "image_ids": images_in_para,
                        "index": len(all_content)
                    })
                elif text or images_in_para:
                    all_content.append({
                        "type": "paragraph",
                        "text": text if text else "[仅含图片的段落]",
                        "has_images": len(images_in_para) > 0,
                        "image_ids": images_in_para,
                        "index": len(all_content)
                    })
                
                para_count += 1
            
            elif isinstance(element, CT_Tbl):
                if table_count >= max_tables:
                    break
                
                table = Table(element, doc)
                table_data = {
                    "type": "table",
                    "rows": len(table.rows),
                    "cols": len(table.columns),
                    "data": [],
                    "index": len(all_content)
                }
                
                for row_idx, row in enumerate(table.rows[:30]):
                    row_data = [cell.text.strip() for cell in row.cells]
                    table_data["data"].append(row_data)
                
                if len(table.rows) > 30:
                    table_data["note"] = f"表格共{len(table.rows)}行，仅显示前30行"
                
                all_content.append(table_data)
                table_count += 1
        
        # 第二遍：为包含图片的段落添加上下文
        for item in all_content:
            if item.get("has_images"):
                idx = item["index"]
                
                # 提取前面的上下文
                context_before_items = []
                start_idx = max(0, idx - context_before)
                for i in range(start_idx, idx):
                    ctx_item = all_content[i]
                    if ctx_item["type"] in ["heading", "paragraph"]:
                        context_before_items.append({
                            "type": ctx_item["type"],
                            "text": ctx_item["text"],
                            "level": ctx_item.get("level")
                        })
                
                # 提取后面的上下文
                context_after_items = []
                end_idx = min(len(all_content), idx + context_after + 1)
                for i in range(idx + 1, end_idx):
                    ctx_item = all_content[i]
                    if ctx_item["type"] in ["heading", "paragraph"]:
                        context_after_items.append({
                            "type": ctx_item["type"],
                            "text": ctx_item["text"],
                            "level": ctx_item.get("level")
                        })
                
                item["context_before"] = context_before_items
                item["context_after"] = context_after_items
        
        result = {
            "file": file_path,
            "total_images": len(all_images),
            "images": all_images,
            "content": all_content
        }
        
        return result
        
    except Exception as e:
        return {
            "error": str(e),
            "file": file_path
        }

def format_output(data, format_type="markdown", show_context=True):
    """格式化输出"""
    if "error" in data:
        return f"错误: {data['error']}"
    
    if format_type == "json":
        return json.dumps(data, ensure_ascii=False, indent=2)
    
    elif format_type == "markdown":
        output = [f"# Word文档完整分析: {os.path.basename(data['file'])}\n"]
        
        # 图片信息摘要
        if data["total_images"] > 0:
            output.append(f"## 📷 文档包含 {data['total_images']} 张图片\n")
            for idx, img in enumerate(data["images"], 1):
                output.append(f"{idx}. {img['filename']} (ID: {img['id']}, 类型: {img['type']})")
                if "saved_path" in img:
                    output.append(f"   - 保存位置: {img['saved_path']}")
            output.append("\n---\n")
        
        # 文档内容
        output.append("## 📄 文档内容\n")
        
        for item in data["content"]:
            if item["type"] == "heading":
                level = int(item["level"][-1]) if item["level"][-1].isdigit() else 2
                output.append(f"\n{'#' * (level + 1)} {item['text']}")
                if item.get("has_images"):
                    output.append(f" 📷[含{len(item.get('image_ids', []))}张图片]")
                output.append("\n")
                
                # 如果标题包含图片，显示上下文
                if show_context and item.get("has_images"):
                    output.append("**图片上下文：**\n")
                    if item.get("context_after"):
                        for ctx in item["context_after"]:
                            output.append(f"- {ctx['text']}\n")
            
            elif item["type"] == "paragraph":
                if item.get("has_images"):
                    output.append(f"\n📷 **[图片段落]**\n")
                    
                    # 显示图片的上下文
                    if show_context:
                        if item.get("context_before"):
                            output.append("**图片前的说明：**\n")
                            for ctx in item["context_before"]:
                                if ctx["type"] == "heading":
                                    output.append(f"### {ctx['text']}\n")
                                else:
                                    output.append(f"{ctx['text']}\n")
                        
                        output.append(f"\n**图片段落内容：** {item['text']}\n")
                        
                        if item.get("context_after"):
                            output.append("\n**图片后的说明：**\n")
                            for ctx in item["context_after"]:
                                if ctx["type"] == "heading":
                                    output.append(f"### {ctx['text']}\n")
                                else:
                                    output.append(f"{ctx['text']}\n")
                    
                    output.append("\n" + "-" * 60 + "\n")
                elif item["text"]:
                    output.append(f"{item['text']}\n")
            
            elif item["type"] == "table":
                output.append(f"\n**表格** ({item['rows']}行 × {item['cols']}列):\n")
                
                if len(item["data"]) > 0:
                    header = item["data"][0]
                    output.append("| " + " | ".join(header) + " |")
                    output.append("| " + " | ".join(["---"] * len(header)) + " |")
                    
                    for row in item["data"][1:]:
                        output.append("| " + " | ".join(row) + " |")
                
                if "note" in item:
                    output.append(f"\n*{item['note']}*\n")
        
        return "\n".join(output)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python read_docx_enhanced.py <word文件路径> [输出格式:json|markdown] [图片保存目录]")
        sys.exit(1)
    
    file_path = sys.argv[1]
    format_type = sys.argv[2] if len(sys.argv) > 2 else "markdown"
    image_dir = sys.argv[3] if len(sys.argv) > 3 else None
    
    data = read_docx_enhanced(file_path, extract_images_flag=True, image_output_dir=image_dir)
    output = format_output(data, format_type, show_context=True)
    print(output)
