<div align="center">

# Asyre Tables

> *"Agent 动嘴，表格动手。"*

![License](https://img.shields.io/badge/License-MIT-blue)
![Python](https://img.shields.io/badge/Python-3.8%2B-green)
![Presets](https://img.shields.io/badge/格式预设-4种-brightgreen)
![Formats](https://img.shields.io/badge/格式支持-CSV%20%7C%20XLSX%20%7C%20JSON-orange)
![AgentSkills](https://img.shields.io/badge/AgentSkills-Standard-yellow)

<br>

**客户在 Discord 发了一条"帮我记一笔账"，你的 Agent 需要写 30 行 Python 来操作 Excel？**

**客户说"看看这个月的账对不对"，你让 AI 把 10 万行数据塞进上下文？**

**客户每个月发同样格式的表，你的 Agent 每次都要重新理解表格结构？**

<br>

### 一行命令操作表格。零代码。最少 token。格式不破坏。

<br>

[**快速开始**](#快速开始) · [**功能**](#能力矩阵) · [**使用场景**](#使用场景) · [**安装**](#安装)

</div>

<br>

---

## 能力矩阵

| 能力 | 命令 | Token 消耗 | 说明 |
|------|------|-----------|------|
| **增删改查** | `tables sheet add/update/delete/query` | ~15 token/次 | 一行命令，不写代码 |
| **文件摘要** | `tables sheet info` | ~10 token | 10 万行文件也只返回结构摘要 |
| **统计分析** | `tables sheet stats` | ~10 token | sum/avg/min/max + 分组，不拉数据 |
| **模板系统** | `tables template register/fill` | 注册1次，以后 ~15 token/次 | 客户表格结构记住一次，以后只传字段值 |
| **审计对账** | `tables sheet audit --mark` | 1次 | 自动验算余额、检查合计，标出问题 |
| **格式化** | `tables sheet format --preset` | 1次 | 4 种专业预设一键美化 |
| **格式转换** | `tables convert` | 1次 | CSV ↔ XLSX ↔ JSON 互转 |
| **合并文件** | `tables merge` | 1次 | 多文件合并 |

**核心承诺：修改 Excel 只动值，不动样式。** 合并单元格、行高、列宽、字体颜色、公式、边框 — 全部原样保留。

## 使用场景

### 场景 1：老板在 Discord 说"帮我记一笔账"

> **角色**：摩托车店老板，每天在群里报流水
> **困境**：会计不在，老板只会发语音/文字，不会开 Excel
> **解法**：Agent 识别出关键信息，一行命令写入

```bash
tables template fill 20260406.xlsx --template 寻北本田日记账 摘要=轮胎更换 收入=800 支出=600 -q
# → filled: 3
```

> *"老板说一句话，账就记好了。"*

### 场景 2：月底对账，老板问"这个月的数对不对"

> **角色**：小企业老板，看不懂公式，只想知道有没有问题
> **困境**：让会计查要等三天，自己看表格头疼
> **解法**：一行命令审计，问题用人话标出来

```bash
tables sheet audit ledger.xlsx --mark
# → "机油更换"这笔账没有算余额，算下来应该剩 849.35 元
# → 合计金额少算了："收入"的合计漏算了 1000 元
```

> *"哪笔账有问题、差多少钱，红色标出来，旁边写清楚。"*

### 场景 3：新客户发来一张从没见过的表格

> **角色**：代账公司，每个客户表格格式都不一样
> **困境**：每次要重新教 Agent 理解表格结构，浪费 token
> **解法**：注册一次模板，以后填数据只传字段名

```bash
# 注册（一次性）
tables template register 客户原件.xlsx --name 新客户月报 --client 新客户

# 以后每次（15 token）
tables template fill output.xlsx --template 新客户月报 日期=4/6 项目=咨询费 金额=5000
```

> *"第一次见面花 5 分钟认识，以后每次 3 秒搞定。"*

### 场景 4：10 万行大文件，上下文塞不下

> **角色**：数据分析 Agent
> **困境**：客户传了个大 Excel，全量读进去直接爆 context
> **解法**：先看结构，再查局部

```bash
tables sheet info huge.csv           # → 100,000 行，6 列
tables sheet query huge.csv --where "status=待处理" --limit 10 --json
tables sheet stats huge.csv --column amount  # → sum=5,230,000 avg=523
```

> *"不读全量，只拿需要的。"*

### 场景 5：Agent 写完报告，需要配一份格式化的 Excel 附件

> **角色**：自动化报告流水线
> **困境**：生成了数据，但客户要带格式的 Excel，不是 CSV
> **解法**：写入数据 + 一键专业格式

```bash
tables sheet add report.xlsx 月份=3月 营收=125000 成本=80000 利润=45000
tables sheet format report.xlsx --preset professional
```

> *"蓝底白字表头、隔行变色、边框，一条命令。"*

## 快速开始

```bash
pip install asyre-tables
# 或
git clone https://github.com/Qihe-agent/asyre-tables.git
cd asyre-tables && pip install -e .
```

```bash
# 创建一个表，加两行数据
tables sheet add clients.csv name=张三 phone=138 status=新客户
tables sheet add clients.csv name=李四 phone=139 status=已签约

# 查看
tables sheet list clients.csv

# 查询
tables sheet query clients.csv --where "status=新客户"

# 改
tables sheet update clients.csv --where "name=张三" --set "status=已签约"

# 统计
tables sheet stats clients.csv --column status
```

## 命令速查

### 表格操作

| 命令 | 说明 |
|------|------|
| `tables sheet info <file>` | 查看文件结构（行数、列名、大小） |
| `tables sheet list <file>` | 列出内容（默认前 20 行） |
| `tables sheet query <file> --where "..." ` | 条件查询 |
| `tables sheet add <file> key=val ...` | 添加一行 |
| `tables sheet update <file> --where "..." --set "..."` | 更新行 |
| `tables sheet delete <file> --where "..."` | 删除行 |
| `tables sheet stats <file> --column <col>` | 列统计 |
| `tables sheet sort <file> --column <col>` | 排序 |
| `tables sheet audit <file> --mark` | 审计对账 |

### Excel 格式化

| 命令 | 说明 |
|------|------|
| `tables sheet formula <file> --range "..." --value "..."` | 写公式 |
| `tables sheet style <file> --range "..." --bold --bg-color ...` | 样式 |
| `tables sheet merge-cells <file> --range "..." --value "..."` | 合并单元格 |
| `tables sheet width <file> --auto` | 自动列宽 |
| `tables sheet format <file> --preset professional` | 一键格式预设 |

### 模板系统

| 命令 | 说明 |
|------|------|
| `tables template register <file> --name <名> --client <客户>` | 注册模板 |
| `tables template list` | 列出所有模板 |
| `tables template info <名>` | 查看模板结构 |
| `tables template new <名> -o <输出文件>` | 从模板创建新文件 |
| `tables template fill <file> --template <名> key=val ...` | 按模板填数据 |

### 格式转换

| 命令 | 说明 |
|------|------|
| `tables convert <源> <目标>` | 格式互转（CSV ↔ XLSX ↔ JSON） |
| `tables merge <文件...> -o <输出>` | 合并多个文件 |

## 输出格式

| 标志 | 给谁用 | 示例 |
|------|--------|------|
| 默认 | 人（终端表格） | `tables sheet list data.csv` |
| `--json` | 程序/Agent | `tables sheet list data.csv --json` |
| `-q` | 最省 token | `tables sheet add data.csv name=X -q` → `added: 1` |

## 格式支持

| 格式 | 读 | 写 |
|------|----|----|
| `.csv` / `.tsv` | Yes | Yes |
| `.xlsx` | Yes | Yes |
| `.xls` | Yes | No（旧格式，只读） |
| `.json` | Yes | Yes |

## 安装

### 依赖

```bash
pip install openpyxl    # 必须
pip install xlrd        # 可选，读 .xls 旧格式
```

### 从 GitHub

```bash
git clone https://github.com/Qihe-agent/asyre-tables.git
cd asyre-tables
pip install -e .
```

### 从 PyPI（即将上线）

```bash
pip install asyre-tables
```

## 生态

Asyre Tables 是 Agent 办公流水线的数据层，配合其他 Asyre 工具：

| 工具 | 职责 | 配合方式 |
|------|------|---------|
| **[Asyre DocForge](https://github.com/Qihe-agent/asyre-docforge)** | Markdown → PDF/Word | Office 处理数据，DocForge 出报告 |
| **[Asyre Presentation](https://github.com/Qihe-agent/asyre-html-presentation)** | HTML 演示文稿 | Office 提供数据，Presentation 做展示 |
| **writing-engine** | 长文写作 | Office 的数据喂给写作引擎 |
| **asyre-search** | 社媒数据分析 | 分析结果存入 Office 表格 |

---

<div align="center">

**Agent 动嘴，表格动手。**

![Asyre Tables](https://img.shields.io/badge/Asyre-Tables-black?style=for-the-badge)

Powered by [**Qihe Agent**](https://github.com/Qihe-agent)

</div>
