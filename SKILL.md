---
name: srm-operation-report-generator
description: 基于客户运维工单数据自动生成 SRM 运维分析报告。采用"脚本数据处理 + LLM 分析"架构，支持 Excel 工单文件直接导入，自动完成数据提取、统计分析和报告生成

触发场景：
- 用户上传运维工单 Excel 文件，要求生成分析报告
- 用户指定客户名称，生成该客户的运维年度报告
- 用户需要分析运维工单的分类、模块、趋势、FAQ 等内容
---

# SRM运维报告生成技能

## 核心架构

```
Excel工单文件
    ↓
extract_module_data.py（提取完整工单数据）
    ↓
generate_report_v2.py（4批次LLM调用）
    ↓
Markdown报告
    ↓
md2docx.py（可选，转Word）
```

**脚本负责**：Excel读取、数据提取、统计计算
**LLM负责**：趋势分析、问题分析、FAQ生成、总结建议

## 使用流程

### Step 1: 修改配置文件

在脚本开头修改以下配置：

```python
# extract_module_data.py
EXCEL_DIR = r"/Users/limingheng/AI\client-data\raw\客户档案\{客户名}\运维工单"
OUTPUT_DIR = r"/Users/limingheng/AI\client-data"

# generate_report_v2.py
RAW_DATA_DIR = r"/Users/limingheng/AI\client-data\raw\客户档案\{客户名}\运维工单"
MODULE_DATA_FILE = os.path.join(CLIENT_DATA_DIR, "{客户名}_2025_模块工单数据.txt")
API_KEY = "your-api-key"  # 或设置环境变量 DEEPSEEK_API_KEY
```

### Step 2: 提取工单数据

```bash
cd skills/srm-ops-expert/scripts
python extract_module_data.py 客户名
python extract_module_data.py 客户名 --year 2024  # 指定年份
```

> **年份规则**：不指定则默认上一自然年（当前2026年，默认2025年）

输出：`/Users/limingheng/AI/client-data/{客户名}/{客户名}_{年份}_模块工单数据.txt`

### Step 3: 生成报告

```bash
python generate_report_v2.py 客户名
python generate_report_v2.py 客户名 --year 2024  # 指定年份
```

输出：`/Users/limingheng/AI/client-data/{客户名}/{客户名}_{年份}_运维报告_V{版本}.md`

版本号自动递增（V1, V2, ...）

### Step 4: 转换为Word（可选）

```bash
python md2docx.py 客户名
python md2docx.py 客户名 --year 2025  # 指定年份
```

输出：`/Users/limingheng/AI/client-data/{客户名}/{客户名}_{年份}_运维报告_V{版本}.docx`

## 数据目录结构

```
/Users/limingheng/AI/client-data/
├── raw/
│   └── 客户档案/
│       └── {客户名}/
│           └── 运维工单/
│               └── *.xlsx     # 原始工单文件
└── {客户名}/                   # Step 2/3/4 生成
    ├── {客户名}_{年份}_模块工单数据.txt
    ├── {客户名}_{年份}_运维报告_Vxx.md
    └── {客户名}_{年份}_运维报告_Vxx.docx
```

## 报告结构

| 章节 | 内容 |
|------|------|
| 统计摘要 | 总工单数、SLA达标率、满意度 |
| 第一部分：工单概览 | 分类分布、模块分布、月度趋势 |
| 第二部分：分类维度分析 | 按工单分类深度分析核心问题 |
| 第三部分：模块维度分析 | 按系统模块深度分析核心问题 |
| 第四部分：高频Q&A | 面向最终用户的FAQ知识库 |
| 第五部分：总结与建议 | 整体表现、风险识别、改进建议 |

## LLM约束规则

详见 `references/deterministic_analysis_manual.md`，核心规则：

### 问题分析
- 只基于工单数据，禁止推测
- 解决方案来源：有则提取共性，无则写"待完善"
- 禁止空洞词汇（提供/优化/加强/完善）
- 参考工单必须真实

### 问题数量
- ≥30单 → 3个问题
- 20-29单 → 2个问题
- 10-19单 → 1个问题
- <10单 → 不输出

### FAQ
- 只从业务咨询类筛选
- 面向最终用户的操作指导
- 固定5个FAQ

## 依赖

```bash
pip install pandas numpy openpyxl requests python-docx xlrd
```

## 注意事项

- **无需修改脚本配置**，通过命令行参数指定客户名即可
- extract_module_data.py 读取所有 .xlsx 文件并合并
- generate_report_v2.py 需要先运行 extract_module_data.py
- API密钥支持环境变量 `DEEPSEEK_API_KEY`
