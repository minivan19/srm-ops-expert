#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提取各模块完整工单数据
用于LLM批量分析
"""

import pandas as pd
import glob
import os

EXCEL_DIR = r"C:\Users\mingh\client-data\raw\客户档案\诺斯贝尔\运维工单"
OUTPUT_DIR = r"C:\Users\mingh\client-data"


def load_data():
    """读取所有Excel"""
    files = glob.glob(os.path.join(EXCEL_DIR, "*.xlsx"))
    all_dfs = []
    for f in files:
        df = pd.read_excel(f)
        all_dfs.append(df)
    return pd.concat(all_dfs, ignore_index=True)


def extract_module_data(df, module_name, top_n=None):
    """提取指定模块的完整工单数据"""
    module_data = df[df['模块'] == module_name]
    
    if top_n:
        module_data = module_data.head(top_n)
    
    results = []
    for _, row in module_data.iterrows():
        record = {
            '工单号': row.get('编号', ''),
            '分类': row.get('分类', ''),
            '模块': module_name,
            '标题': row.get('标题', ''),
            '描述': str(row.get('描述', ''))[:300] if pd.notna(row.get('描述')) else '',
            '解决方案': str(row.get('解决方案', ''))[:200] if pd.notna(row.get('解决方案')) else '',
            '根本原因': str(row.get('根本原因', ''))[:100] if pd.notna(row.get('根本原因')) else '',
        }
        results.append(record)
    return results


def format_for_llm(module_name, records):
    """格式化给LLM的数据"""
    lines = [f"### {module_name} ({len(records)}单)"]
    lines.append("")
    for r in records:
        lines.append(f"**工单号**: {r['工单号']}")
        lines.append(f"**分类**: {r.get('分类', '')}")
        lines.append(f"**模块**: {r.get('模块', '')}")
        lines.append(f"**标题**: {r['标题']}")
        if r['描述']:
            lines.append(f"**描述**: {r['描述']}")
        if r['解决方案']:
            lines.append(f"**解决方案**: {r['解决方案']}")
        if r['根本原因']:
            lines.append(f"**根本原因**: {r['根本原因']}")
        lines.append("")
    return "\n".join(lines)


def main():
    df = load_data()
    
    # 统计各模块
    print("各模块工单数:")
    module_counts = df['模块'].value_counts()
    for mod, cnt in module_counts.items():
        print(f"  {mod}: {cnt}")
    
    # 需要分析的模块（>10单）
    analysis_modules = []
    for mod, cnt in module_counts.items():
        if cnt > 10:
            analysis_modules.append((mod, cnt))
    
    print(f"\n需要分析的模块: {len(analysis_modules)}个")
    for mod, cnt in analysis_modules:
        print(f"  - {mod}: {cnt}单")
    
    # 提取各模块数据
    all_data = {}
    for mod, cnt in analysis_modules:
        records = extract_module_data(df, mod)
        all_data[mod] = records
        print(f"\n{mod} 提取完成: {len(records)}条")
    
    # 保存到文件
    output_file = os.path.join(OUTPUT_DIR, "诺斯贝尔_2025_模块工单数据.txt")
    with open(output_file, 'w', encoding='utf-8') as f:
        for mod, records in all_data.items():
            f.write(format_for_llm(mod, records))
            f.write("\n\n" + "="*50 + "\n\n")
    
    print(f"\n数据已保存到: {output_file}")


if __name__ == "__main__":
    main()
