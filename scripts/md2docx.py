#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown转Word脚本 - 修复版
"""

import os
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 文件路径 - 自动获取最新生成的报告
import glob
CLIENT_DATA_DIR = r"C:\Users\mingh\client-data"

# 查找明阳电路运维报告MD文件
import glob
CLIENT_DATA_DIR = r"C:\Users\mingh\client-data"
md_files = glob.glob(os.path.join(CLIENT_DATA_DIR, "明阳电路_2025_运维报告_*.md"))
import re
def extract_version(f):
    match = re.search(r'V(\d+)', f)
    return int(match.group(1)) if match else 0
md_files.sort(key=extract_version, reverse=True)
if md_files:
    latest_md = md_files[0]
    MD_FILE = latest_md
    DOC_FILE = latest_md.replace('.md', '.docx')
else:
    MD_FILE = os.path.join(CLIENT_DATA_DIR, "明阳电路_2025_运维报告.md")
    DOC_FILE = os.path.join(CLIENT_DATA_DIR, "明阳电路_2025_运维报告.docx")


def read_md(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()


def add_heading(doc, text, level=1):
    heading = doc.add_heading(text, level)
    return heading


def add_paragraph(doc, text, bold=False, italic=False, font_size=None):
    p = doc.add_paragraph(text)
    if bold:
        p.runs[0].bold = True
    if italic:
        p.runs[0].italic = True
    if font_size:
        p.runs[0].font.size = Pt(font_size)
    return p


def add_table_from_markdown(doc, md_content):
    """从markdown表格转换"""
    lines = md_content.strip().split('\n')
    table = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 跳过分隔行
        if re.match(r'\|[\s\-:|]+\|', line):
            continue
        
        if line.startswith('|') and line.endswith('|'):
            cells = [c.strip() for c in line.split('|')[1:-1]]
            
            if table is None:
                table = doc.add_table(rows=1, cols=len(cells))
                table.style = 'Table Grid'
                header_row = table.rows[0]
                for i, cell in enumerate(cells):
                    header_row.cells[i].text = cell
            else:
                row = table.add_row()
                for i, cell in enumerate(cells):
                    if i < len(row.cells):
                        row.cells[i].text = cell
    
    return table is not None


def parse_and_convert(md_content, doc):
    """解析markdown并添加到doc"""
    lines = md_content.split('\n')
    i = 0
    table_buffer = []
    
    while i < len(lines):
        line = lines[i].strip()
        
        # 跳过空行
        if not line:
            i += 1
            continue
        
        # 表格处理
        if line.startswith('|'):
            table_buffer.append(line)
            
            # 检查是否还在表格中
            if i + 1 >= len(lines) or not lines[i + 1].strip().startswith('|'):
                # 表格结束
                add_table_from_markdown(doc, '\n'.join(table_buffer))
                table_buffer = []
            i += 1
            continue
        
        # 标题处理
        if line.startswith('# '):
            doc.add_heading(line[2:], 0)
        elif line.startswith('## '):
            doc.add_heading(line[3:], 1)
        elif line.startswith('### '):
            doc.add_heading(line[4:], 2)
        elif line.startswith('#### '):
            doc.add_heading(line[5:], 3)
        # 列表处理 - 改进版
        elif line.startswith('- ') or line.startswith('* '):
            # 去掉开头的- 或 * 
            text = line[2:]
            # 处理加粗
            p = doc.add_paragraph()
            while '**' in text:
                start = text.find('**')
                end = text.find('**', start + 2)
                if end == -1:
                    break
                before = text[:start]
                bold_text = text[start+2:end]
                after = text[end+2:]
                
                if before:
                    p.add_run(before)
                p.add_run(bold_text).bold = True
                text = after
            
            if text:
                p.add_run(text)
        # 普通段落 - 处理加粗
        else:
            text = line
            p = doc.add_paragraph()
            while '**' in text:
                start = text.find('**')
                end = text.find('**', start + 2)
                if end == -1:
                    break
                before = text[:start]
                bold_text = text[start+2:end]
                after = text[end+2:]
                
                if before:
                    p.add_run(before)
                p.add_run(bold_text).bold = True
                text = after
            
            if text:
                p.add_run(text)
        
        i += 1


def main():
    print("Markdown转Word")
    print("=" * 40)
    
    print(f"读取: {MD_FILE}")
    md_content = read_md(MD_FILE)
    print(f"  - 字符数: {len(md_content)}")
    
    doc = Document()
    
    # 设置文档默认样式
    style = doc.styles['Normal']
    font = style.font
    font.name = '微软雅黑'
    font.size = Pt(12)
    
    print("转换中...")
    parse_and_convert(md_content, doc)
    
    print(f"保存: {DOC_FILE}")
    doc.save(DOC_FILE)
    
    print("完成!")


if __name__ == "__main__":
    main()
