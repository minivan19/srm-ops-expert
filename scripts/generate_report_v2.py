#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SRM运维报告生成脚本
自动调用LLM生成完整报告
"""

import os
import sys
import glob
import requests
import pandas as pd
from datetime import datetime

# ============== 配置 ==============
CLIENT_DATA_DIR = r"C:\Users\mingh\client-data"
RAW_DATA_DIR = r"C:\Users\mingh\client-data\raw\客户档案\明阳电路\运维工单"
MODULE_DATA_FILE = os.path.join(CLIENT_DATA_DIR, "明阳电路_2025_模块工单数据.txt")
OUTPUT_DIR = CLIENT_DATA_DIR

# API配置
API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-340ed7819c2346508c0a46a80df85999")

# LLM配置
MODEL = "deepseek-chat"
TEMPERATURE = 0.3


def call_llm(prompt, temperature=TEMPERATURE):
    """调用DeepSeek LLM"""
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature
    }
    response = requests.post(
        "https://api.deepseek.com/v1/chat/completions",
        headers=headers,
        json=data,
        timeout=180
    )
    return response.json()["choices"][0]["message"]["content"]


def load_module_data():
    """读取模块工单数据"""
    if not os.path.exists(MODULE_DATA_FILE):
        print(f"错误: 数据文件不存在 {MODULE_DATA_FILE}")
        print("请先运行 extract_module_data.py")
        sys.exit(1)
    
    with open(MODULE_DATA_FILE, 'r', encoding='utf-8') as f:
        return f.read()


def load_raw_data():
    """读取原始Excel数据用于统计"""
    files = glob.glob(os.path.join(RAW_DATA_DIR, "*.xlsx"))
    if not files:
        print(f"错误: 未找到Excel文件 {RAW_DATA_DIR}")
        sys.exit(1)
    
    df = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)
    return df


def get_statistics(df):
    """获取统计数据"""
    # 月度分布 - 表格格式
    df['创建时间'] = pd.to_datetime(df['创建时间'], errors='coerce')
    df['月份'] = df['创建时间'].dt.month
    monthly = df['月份'].value_counts().sort_index()
    monthly_rows = []
    for m in range(1, 13):
        cnt = int(monthly.get(m, 0))
        monthly_rows.append(f"| {m}月 | {cnt} |")
    monthly_table = "\n".join(["| 月份 | 工单数 |"] + ["|------|--------|"] + monthly_rows)
    
    # 分类统计 - 表格格式
    category_counts = df['分类'].value_counts()
    total = len(df)
    category_rows = []
    for i, (cat, cnt) in enumerate(category_counts.items(), 1):
        pct = cnt / total * 100
        category_rows.append(f"| {i} | {cat} | {cnt} | {pct:.1f}% |")
    category_table = "\n".join(["| 序号 | 分类名称 | 工单数 | 占比 |"] + ["|------|---------|--------|-------|"] + category_rows)
    
    # 模块统计 - 表格格式
    module_counts = df['模块'].value_counts()
    module_rows = []
    for i, (mod, cnt) in enumerate(module_counts.items(), 1):
        pct = cnt / total * 100
        module_rows.append(f"| {i} | {mod} | {cnt} | {pct:.1f}% |")
    module_table = "\n".join(["| 序号 | 系统模块 | 工单数 | 占比 |"] + ["|------|---------|--------|-------|"] + module_rows)
    
    # SLA和满意度
    sla_rate = df['SLA是否达标'].value_counts(normalize=True).get('达标', 0) * 100
    product_sat = df['产品满意度'].mean() if '产品满意度' in df.columns else 5.0
    service_sat = df['服务满意度'].mean() if '服务满意度' in df.columns else 5.0
    
    return {
        'monthly': monthly_table,
        'category': category_table,
        'module': module_table,
        'total': len(df),
        'sla': f"{sla_rate:.1f}%",
        'product_sat': f"{product_sat:.1f}",
        'service_sat': f"{service_sat:.1f}"
    }


def batch_0_trend_sla(stats):
    """第0批：趋势+SLA分析"""
    prompt = f"""## 任务：基于统计数据生成趋势分析和SLA分析

### 数据
- 月度分布：{stats['monthly']}
- SLA达标率：{stats['sla']}
- 产品满意度：{stats['product_sat']}分
- 服务满意度：{stats['service_sat']}分

### 输出格式
### 1.4 趋势分析
**整体趋势判断** [简要描述]
**趋势预测与建议** 1. 2. 3.

### 1.5 SLA与满意度分析
**核心指标表现**
- SLA达标率：{stats['sla']} → [评价]
- 产品满意度：{stats['product_sat']}分 → [评价]
- 服务满意度：{stats['service_sat']}分 → [评价]"""

    return call_llm(prompt)


def batch_1_classification_module(module_data):
    """第1批：分类+模块分析"""
    prompt = f"""## 任务：基于完整工单数据进行问题分析

### 重要约束
1. 只基于提供的工单数据进行分析，不要推测
2. 模块范围基于Excel"模块"字段的实际值
3. 解决方案必须基于工单中已有的"解决方案"字段，如果为空则写"待完善"
4. 禁止使用"优化"、"完善"、"加强"等空洞词汇，必须给出具体可执行的建议

### 数据格式（完整工单信息）
- 工单号、分类、模块、标题、描述、**解决方案**、根本原因

### 关键要求
- 如果工单有解决方案 → 提取并作为主要方案
- 如果工单解决方案为空 → 写"待完善"，不要自行编造
- 解决方案要具体，如：进入【XX模块】→点击【XX按钮】→执行XX操作
- **参考工单必须与问题描述直接相关，禁止幻觉，只列出数据中真实存在的工单号**

### 需要分析的模块（全部列出）
- 订单/物流: 56单 → 建议3个问题
- 寻源（询价/招标）: 19单 → 建议1个问题
- 系统基础/报表/应用商店: 18单 → 建议1个问题
- 合作伙伴: 17单 → 建议1个问题

### 数据
{module_data}

### 输出格式
## 第二部分：分类维度深度分析
### 2.1 业务咨询 - 97单，建议3个问题
#### 问题1：{{类别}}
**问题描述** [基于描述字段提炼共性]
**原因分析** 1. 2.
**解决方案**
1. [基于现有解决方案，如果为空写"待完善"]
2. [如果有多个工单有解决方案，提取共性]
**参考工单** - I-xxx

### 2.2 外围系统 - 10单，建议1个问题

## 第三部分：模块维度深度分析
### 3.1 订单/物流 - 56单，建议3个问题
### 3.2 寻源（询价/招标） - 19单，建议1个问题
### 3.3 系统基础/报表/应用商店 - 18单，建议1个问题
### 3.4 合作伙伴 - 17单，建议1个问题"""

    return call_llm(prompt)


def batch_2_faq(module_data):
    """第2批：FAQ"""
    prompt = f"""## 任务：从业务咨询类工单生成高频FAQ（面向最终用户）

### 重要约束
1. 基于工单的"标题"、"描述"、"解决方案"、"根本原因"四个字段综合分析
2. 解决方案定位：面向最终用户（业务人员），不是技术人员
3. 解决方案类型限制：
   - ✅ 可以是：操作指导（如：点击XX按钮→填写XX→提交）
   - ✅ 可以是：答疑解释（如：这是因为...）
   - ✅ 可以是：联系系统管理员（如：请联系管理员处理"XX问题"）
   - ❌ 不能是：配置建议（如：进入系统设置→配置XX）
   - ❌ 不能是：实施优化建议（如：建议优化XX流程）
4. 禁止使用"优化"、"完善"、"配置"、"调整"等词汇
5. 问题场景需要聚类提炼，不强求使用原文
6. 参考工单必须与问题描述直接相关，禁止幻觉

### 数据字段
- 标题：工单标题
- 描述：详细描述
- 解决方案：处理方案
- 根本原因：根因分析

### 数据
{module_data}

### 输出格式
### Q1：{{基于多个工单标题聚类提炼}}
**问题场景** [综合标题+描述提炼，使用用户能理解的语言]
**解决方案** 
1. [操作指导：如果问题可以通过操作解决，给出具体步骤]
2. [或答疑解释：如果无法操作解决，解释原因]
3. [或联系管理员：如果需要系统配置，写"请联系系统管理员处理"]
**参考工单** - I-xxxxx"""

    return call_llm(prompt)


def batch_3_summary(batch1_result, batch2_result, stats):
    """第3批：总结"""
    prompt = f"""## 任务：基于前期分析结果生成总结与改进建议

### 第1批分析结果
{batch1_result[:2000]}

### 第2批FAQ结果
{batch2_result[:1000]}

### 基础数据
- 总工单：{stats['total']}
- SLA达标率：{stats['sla']}
- 满意度：{stats['product_sat']}分

### 输出格式
### 5.1 运维整体表现
### 5.2 主要问题与风险
### 5.3 改进建议"""

    return call_llm(prompt)


def generate_report():
    """生成完整报告"""
    print("=" * 50)
    print("SRM Report Generator")
    print("=" * 50)
    
    # 1. 加载数据
    print("\n[1/5] Loading data...")
    module_data = load_module_data()
    df = load_raw_data()
    stats = get_statistics(df)
    print(f"  Total tickets: {stats['total']}")
    print(f"  SLA rate: {stats['sla']}")
    
    # 2. 第0批：趋势+SLA
    print("\n[2/5] Batch 0: Trend+SLA...")
    result_0 = batch_0_trend_sla(stats)
    print("  [OK]")
    
    # 3. 第1批：分类+模块
    print("\n[3/5] Batch 1: Classification+Module...")
    result_1 = batch_1_classification_module(module_data)
    print("  [OK]")
    
    # 4. 第2批：FAQ
    print("\n[4/5] Batch 2: FAQ...")
    result_2 = batch_2_faq(module_data)
    print("  [OK]")
    
    # 5. 第3批：总结
    print("\n[5/5] Batch 3: Summary...")
    result_3 = batch_3_summary(result_1, result_2, stats)
    print("  [OK]")
    
    # 整合报告
    report = f"""# 明阳电路2025年度SRM运维分析报告

## 报告信息
| 项目 | 内容 |
|------|------|
| 报告周期 | 2025年度 |
| 客户名称 | 明阳电路 |
| 报告生成时间 | {datetime.now().strftime('%Y-%m-%d')} |
| 服务团队 | 甄云科技 |

---

## 统计摘要
| 指标 | 数值 |
|------|------|
| 总工单数 | {stats['total']} |
| SLA达标率 | {stats['sla']} |
| 产品满意度 | {stats['product_sat']} |
| 服务满意度 | {stats['service_sat']} |

---

## 第一部分：工单概览

### 1.1 按工单分类分布
{stats['category']}

### 1.2 按系统模块分布
{stats['module']}

### 1.3 月度分布
{stats['monthly']}

---

{result_0}

---

{result_1}

---

## 第四部分：高频Q&A知识库

{result_2}

---

## 第五部分：总结与改进建议

{result_3}

---

**报告结束**
"""
    
    # 保存报告
    version = 1
    output_file = os.path.join(OUTPUT_DIR, f"明阳电路_2025_运维报告_V{version}.md")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n[OK] Report saved: {output_file}")
    return output_file


if __name__ == "__main__":
    generate_report()
