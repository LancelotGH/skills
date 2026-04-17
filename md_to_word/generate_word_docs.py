# coding=utf-8
import os
import sys
import re
import subprocess

def fix_markdown_formatting(file_path):
    """
    修改原 Markdown 内容：在列表项和标题前补充缺失的空行，
    以修正在 Pandoc 转换到 Word 时因结构粘连导致的排版层级丢失。
    """
    if not os.path.exists(file_path):
        print(f"Error: {file_path} 不存在。")
        return False
        
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    new_lines = []
    
    for i, line in enumerate(lines):
        # 0. 预处理：修复因手误引发的不合规前置空格缩进
        if line.startswith('  #'):
            line = line.lstrip()
        elif line.startswith(' #') or line.startswith(' -') or line.startswith(' >') or line.startswith(' |'):
            line = line[1:]
            
        # 匹配无序列表和标题
        is_list = re.match(r'^[ \t]*[-*]\s', line)
        is_heading = re.match(r'^#{1,6}\s', line)
        
        if is_list or is_heading:
            if len(new_lines) > 0:
                prev = new_lines[-1].strip()
                # 如果前一行不是空行，不是标题，且不已经是同类列表/分割线
                if prev != '' and not prev.startswith('#') and not re.match(r'^[-*_]{3,}\s*$', prev):
                    if is_list and re.match(r'^[ \t]*[-*]\s', prev):
                        pass # 上一行已经是列表项了，不需要加空行
                    else:
                        new_lines.append('')
        
        new_lines.append(line)

    new_content = '\n'.join(new_lines) + '\n'
    
    # 将修复后的内容写回原文件
    if new_content != content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"已自动补充缺失空行修复格式：{file_path}")
    
    return True

def convert_to_word(md_file):
    """
    自动化流程：
    1. 修正 Markdown 排版语法
    2. 使用 Pandoc 生成基础 docx
    3. 调用 format_md_style.py 对生成的 docx 进一步提取和排版调优
    """
    if not md_file.endswith('.md'):
        print("错误：请输入 .md 后缀的 Markdown 文件。")
        sys.exit(1)
        
    base_name = md_file[:-3]
    base_docx = f"{base_name}_base.docx"
    final_docx = f"{base_name}.docx"
    
    # 获取脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    style_script = os.path.join(current_dir, "format_md_style.py")
    
    # 1. 语法检查与补充空行修复
    print(f"\n[1/3] 正在检查并修复 Markdown 语法结构: {md_file} ...")
    if not fix_markdown_formatting(md_file):
        sys.exit(1)
        
    # 2. 调用 Pandoc 转换
    print(f"[2/3] 正在调用 Pandoc 生成基础文档: {base_docx} ...")
    cmd_pandoc = f'pandoc "{md_file}" -o "{base_docx}"'
    result = subprocess.run(cmd_pandoc, shell=True)
    if result.returncode != 0:
        print("Pandoc 转换失败，请检查是否已正确安装并配置环境变量。")
        sys.exit(1)
        
    # 3. 执行 Python 样式优化（清理斜体、渲染代码块、重塑列表）
    print(f"[3/3] 正在应用文档样式优化: {final_docx} ...")
    cmd_style = f'python "{style_script}" "{base_docx}" "{final_docx}"'
    result = subprocess.run(cmd_style, shell=True)
    if result.returncode != 0:
        print("应用样式脚本失败，请检查运行环境或 format_md_style.py 是否存在。")
        sys.exit(1)
        
    # 4. 清除中间临时文件
    if os.path.exists(base_docx):
        os.remove(base_docx)
        
    print(f"\n全部完成！已成功将 {md_file} 转化为高质量排版的 {final_docx} .\n")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python generate_word_docs.py <你的文档.md> [其他的文档.md...]")
        sys.exit(1)
        
    for file_path in sys.argv[1:]:
        convert_to_word(file_path)
