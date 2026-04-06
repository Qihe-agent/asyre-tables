---
name: asyre-tables
description: |
  统一文件操作 CLI，专为 AI Agent 设计。支持 CSV/Excel 的 CRUD、模板系统、审计对账。
  最小 token 消耗：Agent 不需要写代码，只调命令。
  触发词: office, 表格, excel, csv, 记账, 对账, 审计, 填表, 模板, spreadsheet, sheet
metadata:
  openclaw:
    emoji: "📊"
    requires:
      bins: ["python3"]
      pip: ["openpyxl"]
---

# Asyre Tables

统一文件操作 CLI。Agent 通过一行命令操作 CSV/Excel 文件，不需要写代码。

## 安装

```bash
pip install asyre-tables
# 或
git clone https://github.com/yzha0302/asyre-tables.git
cd asyre-tables && pip install -e .
```

## 核心原则

1. **最小 token**：每个操作都是一行命令，不需要生成代码
2. **不破坏格式**：修改 Excel 只动值，不动样式/合并/公式/行高/列宽
3. **模板系统**：客户的表格结构注册一次，以后填数据不需要理解结构
4. **安全**：update/delete 必须带 --where（防误删）；模板文件只读
5. **说人话**：审计结果用老板能看懂的语言，不用技术术语

---

## 一、表格 CRUD（tables sheet）

### 查看文件

```bash
# 查看结构（列名、行数、大小）— 大文件先用这个，不拉全量数据
tables sheet info data.xlsx

# 列出内容（默认前20行）
tables sheet list data.csv
tables sheet list data.xlsx --limit 10
tables sheet list data.csv --columns "name,status"

# JSON 输出（给程序解析，省 token）
tables sheet list data.csv --json
tables sheet info data.xlsx --json
```

### 条件查询

```bash
# 等于
tables sheet query data.csv --where "status=新客户"

# 比较 + 排序 + 限制
tables sheet query data.xlsx --where "amount>1000" --sort "-amount" --limit 5

# 多条件（逗号分隔）
tables sheet query data.csv --where "status=已签约,amount>5000"

# 只看特定列
tables sheet query data.csv --where "status=新客户" --columns "name,phone"
```

### 增删改

```bash
# 添加行（文件不存在会自动创建）
tables sheet add data.csv name=张三 phone=138 status=新客户

# 更新（必须带 --where，防止误改全表）
tables sheet update data.csv --where "name=张三" --set "status=已签约,note=大客户"

# 删除（必须带 --where）
tables sheet delete data.csv --where "status=已流失"

# 静默模式（只返回影响行数，最省 token）
tables sheet add data.csv name=李四 -q        # → added: 1
tables sheet update data.csv --where "name=李四" --set "status=跟进中" -q  # → updated: 1
```

### 统计

```bash
# 单列统计
tables sheet stats data.xlsx --column amount
# → count: 50, sum: 125000, avg: 2500, min: 100, max: 9999

# 分组统计
tables sheet stats data.xlsx --column revenue --group-by region
# → 华东: sum=500000  华南: sum=400000

# 排序
tables sheet sort data.csv --column amount --desc
```

---

## 二、Excel 格式化（tables sheet）

### 公式

```bash
# 写单个公式
tables sheet formula data.xlsx --range "D2" --value "=B2*C2"

# 批量写公式（{row} 会被替换为行号）
tables sheet formula data.xlsx --range "D2:D100" --value "=B{row}*C{row}"
```

### 样式

```bash
# 加粗 + 蓝底白字
tables sheet style data.xlsx --range "A1:F1" --bold --bg-color 4472C4 --font-color FFFFFF

# 数字格式
tables sheet style data.xlsx --range "D2:D100" --number-format "#,##0.00"

# 居中对齐
tables sheet style data.xlsx --range "A1:A100" --align center
```

### 合并单元格

```bash
tables sheet merge-cells data.xlsx --range "A1:C1" --value "标题文字"
```

### 列宽

```bash
# 自动适配所有列
tables sheet width data.xlsx --auto

# 手动设置
tables sheet width data.xlsx --column A --size 20
```

### 一键格式预设

```bash
tables sheet format data.xlsx --preset professional  # 蓝底白字表头 + 隔行变色 + 边框
tables sheet format data.xlsx --preset minimal        # 极简（只有表头加粗 + 底线）
tables sheet format data.xlsx --preset colorful       # 绿色主题
tables sheet format data.xlsx --preset financial      # 深色金融风
```

---

## 三、审计对账（tables sheet audit）

自动检查表格中的计算是否正确，找出问题并标注。

### 检查内容

1. **余额验算**：自动识别"余额 = 上笔余额 + 收入 - 支出"的模式，逐行验算
   - 识别的列名：余额/结余/balance、收入/入账/income、支出/出账/expense
2. **合计验算**：检查 SUM 公式是否覆盖了所有数据行（有没有漏算）
3. **数据完整性**：有收支但没摘要等情况

### 用法

```bash
# 只查看（不改文件）
tables sheet audit ledger.xlsx

# 查看 + 标记到文件
tables sheet audit ledger.xlsx --mark

# JSON 输出
tables sheet audit ledger.xlsx --json
```

### --mark 做了什么

**在原表上（不改结构）：**
- 有问题的单元格标色：黄色 = 缺失，红色 = 算错
- 旁边的空单元格写**简短批注**，如 `← 余额应为 849.35`、`← 合计漏算 1000元`
- 不改变原表的行数、列数、行高、列宽、合并单元格

**新建"审计结果"sheet（独立 tab）：**
- 表头：序号、行号、类型、问题描述、计算详情
- 问题描述用**人话**写，老板能看懂：
  - `"机油更换"这笔账没有算余额。上笔余额799.35元，收入200元 支出150元，算下来应该剩 849.35 元`
  - `合计金额少算了："收入"的合计没有包含新增的数据，漏算了 1000 元`
- 自动换行，不会截断
- 底部有审计摘要（检查行数、问题数、审计时间）

### 输出示例

```
[检查1] 余额验算: 余额 = prev + 收入 - 支出
  → 2 个问题

[检查2] 合计公式验证
  检查了2个公式 → 2 个问题

[检查3] 数据完整性
  → 通过

========================================
审计完成: 18 行已检查，发现 4 个问题：

  [漏算] 第18行: "机油更换"这笔账没有算余额。上笔余额799.35元，收入200元 支出150元，算下来应该剩 849.35 元
  [漏算] 第19行: "轮胎更换"这笔账没有算余额。上笔余额849.35元，收入800元 支出600元，算下来应该剩 1049.35 元
  [少算] 第20行: 合计金额少算了："收入"的合计没有包含新增的数据，漏算了 1000 元。合计数字比实际偏小
  [少算] 第20行: 合计金额少算了："支出"的合计没有包含新增的数据，漏算了 750 元。合计数字比实际偏小
```

---

## 四、模板系统（tables template）

为每个客户注册表格模板。注册一次后，以后填数据只需传字段名和值，不需要理解表格结构。

### 注册模板

```bash
# 从客户给的 Excel 注册（自动分析结构：表头位置、列类型、公式、保护行等）
tables template register 客户原件.xlsx --name 寻北本田日记账 --client 寻北五羊本田
```

注册后保存在 `~/.asyre-tables/templates/`：
- `寻北本田日记账.xlsx` — 原始模板文件（**只读，不可修改**）
- `寻北本田日记账.json` — 结构定义
- `registry.json` — 客户 → 模板的映射

### 查看模板

```bash
# 列出所有模板
tables template list

# 按客户筛选
tables template list --client 寻北五羊本田

# 查看模板结构（列定义、公式、保护行）
tables template info 寻北本田日记账

# JSON 输出
tables template list --json
tables template info 寻北本田日记账 --json
```

### 从模板创建新文件

```bash
# 复制模板，生成新文件（模板本身不动）
tables template new 寻北本田日记账 -o 20260406.xlsx

# 自动替换标题中的日期
tables template new 寻北本田日记账 -o 20260406.xlsx --date 2026年04月06日
```

### 填数据

```bash
# 按字段名填值（自动插入到数据区，在合计行前面）
tables template fill 20260406.xlsx --template 寻北本田日记账 摘要=轮胎更换 收入=800 支出=600

# 静默模式
tables template fill 20260406.xlsx --template 寻北本田日记账 摘要=机油更换 收入=200 支出=150 -q
# → filled: 3
```

**安全机制**：`fill` 命令拒绝直接写入模板文件，只能写副本。

### 删除模板

```bash
tables template delete 寻北本田日记账
```

---

## 五、格式转换和合并

### 格式转换

```bash
tables convert data.csv data.xlsx     # CSV → Excel
tables convert data.xlsx data.csv     # Excel → CSV
tables convert data.csv data.json     # CSV → JSON
tables convert data.json data.xlsx    # JSON → Excel
```

### 合并多个文件

```bash
tables merge file1.csv file2.csv file3.csv -o merged.csv
tables merge *.xlsx -o all.xlsx
```

---

## 六、Agent 调用完整示例

### 场景 1：客户在 Discord 说"帮我记一笔账"

```bash
# Agent 识别出：摘要=轮胎更换，收入=800，支出=600
tables template fill /data/客户A/20260406.xlsx --template 寻北本田日记账 \
  摘要=轮胎更换 收入=800 支出=600 -q
# → filled: 3
```

Token 消耗：约 15 个 token（一行命令），不需要写任何 Python 代码。

### 场景 2：客户说"看看今天的账对不对"

```bash
tables sheet audit /data/客户A/20260406.xlsx --mark
```

Agent 把审计结果转述给客户：
> "查了18行数据，发现4个问题：机油更换和轮胎更换两笔没算余额，应该分别是849.35和1049.35；合计的收入少算了1000元，支出少算了750元。已经在表里标出来了。"

### 场景 3：客户发了一张新格式的表格截图

```bash
# 1. Agent 让客户发 Excel 原件
# 2. 注册为模板
tables template register uploaded.xlsx --name 新客户月报 --client 新客户公司

# 3. 查看 Agent 能用的字段
tables template info 新客户月报
# → columns: 日期[text], 项目[text], 金额[number], 备注[text]

# 4. 以后填数据
tables template fill output.xlsx --template 新客户月报 日期=4/6 项目=咨询费 金额=5000
```

### 场景 4：大文件，上下文装不下

```bash
# 先看结构（不拉数据）
tables sheet info huge.csv
# → rows: 100,000 / columns: id, name, amount, status

# 只查需要的部分
tables sheet query huge.csv --where "status=待处理" --limit 10 --json

# 统计（不需要看每一行）
tables sheet stats huge.csv --column amount
# → sum: 5,230,000  avg: 523  min: 1  max: 9,999
```

### 场景 5：生成报告（配合 DocForge）

```bash
# 1. 从 Excel 提取数据
tables sheet query sales.xlsx --where "month=3" --json > /tmp/data.json

# 2. Agent 用数据写 Markdown 报告
# 3. DocForge 转成 PDF
docforge -p proposal --eisvogel --title "3月销售报告" --titlepage /tmp/report.md
```

---

## 七、输出格式

| 标志 | 输出格式 | 用途 |
|------|---------|------|
| （默认） | 对齐表格 | 给人看 |
| `--json` | JSON | 给程序解析 |
| `-q, --quiet` | 只返回影响行数 | 写操作最省 token |

---

## 八、支持的文件格式

| 格式 | 读 | 写 | 备注 |
|------|----|----|------|
| `.csv` | Yes | Yes | 自动检测编码（UTF-8/GBK）和分隔符 |
| `.tsv` | Yes | Yes | Tab 分隔 |
| `.xlsx` | Yes | Yes | 完整格式保留 |
| `.xls` | Yes | No | 旧格式，只读 |
| `.json` | Yes | Yes | JSON 数组格式 |

---

## 九、注意事项

- `update` 和 `delete` 必须带 `--where`，防止误操作全表
- 模板文件（`~/.asyre-tables/templates/*.xlsx`）是只读的，`fill` 命令拒绝直接写入
- Excel 修改采用**原地编辑**：只改对应的单元格值，其他一切不动
- 审计 `--mark` 在原表只加颜色和简短批注，详细报告在独立的"审计结果"sheet
- 审计批注写在**问题单元格旁边的空单元格**里，不用 Excel 的悬停批注
- 大文件建议先 `info` 再 `query`，不要 `list` 全量数据
