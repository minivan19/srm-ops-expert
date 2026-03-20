#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""虎牙2025运维报告 - 精简版（直接统计+按需LLM）"""
import os, sys, glob, re, json
import pandas as pd
import requests
import time
from datetime import datetime

# ============== 配置 ==============
CLIENT_NAME = "虎牙"
RAW_DATA_DIR = r"/Users/limingheng/AI\client-data\raw\客户档案\虎牙\运维工单"
OUTPUT_FILE = rf"/Users/limingheng/AI\client-data\{CLIENT_NAME}_2025_运维报告.md"
API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-340ed7819c2346508c0a46a80df85999")
MODEL = "deepseek-chat"

# ============== DeepSeek LLM ==============
def call_llm(prompt, max_tokens=800, temperature=0.3):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens
    }
    for attempt in range(1, 6):
        try:
            resp = requests.post(
                "https://api.deepseek.com/v1/chat/completions",
                headers=headers, json=data,
                timeout=(30, 120), stream=False
            )
            if resp.status_code == 200:
                result = resp.json()
                if result.get("choices"):
                    return result["choices"][0]["message"]["content"]
            return f"HTTP {resp.status_code}: {resp.text[:100]}"
        except Exception as e:
            last_err = str(e)
            if attempt < 5:
                time.sleep(5)
    return f"FAILED: {last_err}"

# ============== 数据读取 ==============
def load_2025_data():
    files = sorted(glob.glob(os.path.join(RAW_DATA_DIR, "*.xlsx")))
    files = [f for f in files if "2025" in os.path.basename(f)]
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f)
            dfs.append(df)
        except Exception as e:
            print(f"  [跳过] {os.path.basename(f)}: {e}", flush=True)
    combined = pd.concat(dfs, ignore_index=True)
    print(f"  总工单: {len(combined)} 条", flush=True)
    return combined

# 按位置映射列名（UTF-8字节顺序对应中文）
_COL_MAP = {
    6: '分类',
    8: '模块',
    4: '状态',
    9: '描述',
    12: '提单时间',
}

def clean_df(df):
    # 只保留已知列
    new_df = pd.DataFrame()
    for idx, name in _COL_MAP.items():
        if idx < len(df.columns):
            new_df[name] = df.iloc[:, idx].astype(str).str.strip()
    return new_df

# ============== 统计分析（无需LLM） ==============
def compute_stats(df):
    stats = {}
    stats['total'] = len(df)
    
    # 分类分布
    if '分类' in df.columns:
        stats['by_category'] = df['分类'].value_counts().to_dict()
    
    # 模块分布
    if '模块' in df.columns:
        stats['by_module'] = df['模块'].value_counts().to_dict()
    
    # 月度分布
    if '提单时间' in df.columns:
        try:
            df['提单时间'] = pd.to_datetime(df['提单时间'], errors='coerce')
            df['月份'] = df['提单时间'].dt.month
            stats['by_month'] = df['月份'].value_counts().sort_index().to_dict()
        except:
            pass
    
    # SLA（状态分布）
    if '状态' in df.columns:
        stats['by_status'] = df['状态'].value_counts().to_dict()
    
    # 满意度
    if '满意度' in df.columns:
        stats['satisfaction'] = df['满意度'].value_counts().to_dict()
    
    return stats

def render_stats_markdown(stats):
    lines = []
    lines.append(f"## 📊 统计摘要\n")
    lines.append(f"- **总工单数**: {stats['total']} 条")
    
    # SLA
    if 'by_status' in stats:
        total = sum(stats['by_status'].values())
        # 假设已解决/已完成 = SLA达标
        sla_keys = [k for k in stats['by_status'].keys()]
        done = sum(stats['by_status'].get(k, 0) for k in sla_keys if any(
            w in str(k) for w in ['完成', '解决', '已处理', 'closed', 'resolved', 'done']
        ))
        rate = done / total * 100 if total > 0 else 0
        lines.append(f"- **SLA达标率**: {rate:.1f}%（{done}/{total}条已处理）")
    
    # 月度趋势
    if 'by_month' in stats:
        lines.append(f"\n### 📅 月度趋势\n")
        lines.append(f"| 月份 | 工单数 |")
        lines.append(f"|------|--------|")
        for m in range(1, 13):
            cnt = stats['by_month'].get(m, 0)
            if cnt > 0:
                lines.append(f"| {m}月 | {cnt} |")
    
    # 分类分布
    if 'by_category' in stats:
        lines.append(f"\n### 📋 分类分布（Top 10）\n")
        lines.append(f"| 分类 | 工单数 |")
        lines.append(f"|------|--------|")
        for cat, cnt in list(stats['by_category'].items())[:10]:
            lines.append(f"| {cat} | {cnt} |")
    
    # 模块分布
    if 'by_module' in stats:
        lines.append(f"\n### 🏗️ 模块分布\n")
        lines.append(f"| 模块 | 工单数 |")
        lines.append(f"|------|--------|")
        for mod, cnt in list(stats['by_module'].items()):
            lines.append(f"| {mod} | {cnt} |")
    
    return "\n".join(lines)

# ============== FAQ 生成（需要LLM） ==============
def generate_faq(df, top_n=5):
    """从业务咨询类工单中提取FAQ（仅工单描述>10字的记录）"""
    if '描述' not in df.columns and '分类' not in df.columns:
        return "（无法提取FAQ：数据中无描述或分类字段）"
    
    # 筛选业务咨询类
    biz_df = df[df.get('分类', pd.Series(['未分类']*len(df))).astype(str).str.contains(
        '业务|咨询|操作|使用|如何|怎么|怎样|哪里|哪个', na=False
    )] if '分类' in df.columns else df
    
    samples = []
    for _, row in biz_df.head(30).iterrows():
        desc = str(row.get('描述', '')).strip()
        if len(desc) > 10:
            cat = str(row.get('分类', '其他'))
            mod = str(row.get('模块', '其他'))
            samples.append(f"[{mod}/{cat}] {desc[:200]}")
    
    if not samples:
        return "（该模块暂无业务咨询类工单，无法生成FAQ）"
    
    prompt = f"""你是SRM系统运维支持工程师。请根据以下用户提问，生成{top_n}个最常见的问题和答案。

要求：
- 每个FAQ包含"问"和"答"两部分
- 答要具体、可操作，基于工单内容总结，不要臆造
- 按频率从高到低排列
- 只输出{top_n}个FAQ

工单内容：
{chr(10).join(samples)}

输出格式：
1. 问：...答：...
2. 问：...答：...
（直接列出，不要加标题）"""
    
    print("  [LLM] 生成FAQ...", flush=True)
    result = call_llm(prompt, max_tokens=1000)
    return result if result else "（FAQ生成失败）"

# ============== 问题分析（轻量LLM） ==============
def analyze_problems(df, module_name, count):
    if count < 3:
        return None
    
    samples = []
    for _, row in df.head(min(count, 15)).iterrows():
        desc = str(row.get('描述', '')).strip()
        if len(desc) > 5:
            samples.append(desc[:200])
    
    if len(samples) < 2:
        return None
    
    prompt = f"""你是SRM客户成功经理。请分析以下工单，总结该模块最常见的{min(3, count)}个问题及解决方案。

要求：
- 每个问题说明：问题描述、原因分析、解决建议
- 只基于工单内容，不要臆造
- 禁止空洞词汇（"优化流程"、"加强培训"等）

工单描述：
{chr(10).join(samples)}"""
    
    print(f"  [LLM] 分析 {module_name} ({count}单)...", flush=True)
    result = call_llm(prompt, max_tokens=600)
    return result

# ============== 主流程 ==============
def main():
    print("=" * 50, flush=True)
    print(f"{CLIENT_NAME} 2025 运维报告生成中...", flush=True)
    print("=" * 50, flush=True)
    
    # 1. 读取数据
    print("\n[Step 1] 读取工单...", flush=True)
    df = load_2025_data()
    df = clean_df(df)
    
    # 2. 统计
    print("\n[Step 2] 统计分析...", flush=True)
    stats = compute_stats(df)
    
    # 3. 生成报告
    print("\n[Step 3] 撰写报告...", flush=True)
    report_lines = []
    report_lines.append(f"# {CLIENT_NAME} 2025年运维报告\n")
    report_lines.append(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
    
    # 统计摘要
    report_lines.append(render_stats_markdown(stats))
    
    # 按模块分析（仅高频模块用LLM）
    if 'by_module' in stats:
        report_lines.append("\n---\n")
        report_lines.append("## 🔍 各模块详细分析\n")
        
        for mod, cnt in stats['by_module'].items():
            mod_df = df[df.get('模块', pd.Series(['其他']*len(df))).astype(str) == mod] if '模块' in df.columns else pd.DataFrame()
            report_lines.append(f"\n### {mod}（{cnt}单）\n")
            
            # LLM分析（5单以上才调用）
            if cnt >= 5:
                llm_result = analyze_problems(mod_df, mod, cnt)
                if llm_result:
                    report_lines.append(f"{llm_result}\n")
            else:
                report_lines.append(f"（工单数较少，暂不进行深度分析）\n")
            
            # FAQ
            if len(mod_df) > 3:
                faq = generate_faq(mod_df, top_n=3)
                if faq and '（' not in faq[:4]:
                    report_lines.append(f"\n**常见FAQ**\n{faq}\n")
    
    # 4. 总结
    print("\n[Step 4] 生成总结...", flush=True)
    summary_prompt = f"""你是SRM客户成功经理。请根据以下{stats['total']}条运维工单统计，撰写一段简洁的总结。

数据：
- 总工单：{stats['total']}条
- SLA达标：{stats.get('by_status', {})}
- 月度分布：{stats.get('by_month', {})}
- 模块分布：{stats.get('by_module', {})}

要求：
- 总结整体运维表现
- 识别1-2个主要风险点
- 提出具体改进建议（禁止空洞词汇）
- 50字以内"""
    
    summary = call_llm(summary_prompt, max_tokens=200)
    report_lines.append(f"\n---\n## 📝 总结\n{summary}\n")
    
    # 5. 保存
    content = "\n".join(report_lines)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"\n✅ 报告已保存: {OUTPUT_FILE}", flush=True)
    print(f"   字符数: {len(content)}", flush=True)

if __name__ == "__main__":
    main()
