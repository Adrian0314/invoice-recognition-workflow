# 发票识别工作流机器人 · 纯 Python 单文件版（v2）

**在 Python IDLE 里按 F5 就能跑**，不装任何第三方包也能工作。

```text
IDLE:  File → Open → idle_invoice_bot.py → Run Module (F5)
命令行: python idle_invoice_bot.py -i 发票目录 -o 输出目录
```

---

## 一、它做什么

```
发票文件(PDF/图片)
      │
      ├─ 1. 扫描目录 / 文件指纹去重（幂等）
      ├─ 2. 文本提取：PyMuPDF → 内置标准库 PDF 解析器 → OCR（可选）
      │      每档产物先过「文本质量闸门」，不合格才升档
      ├─ 3. 字段解析：版式模板优先 → 内置启发式兜底，逐字段记置信度
      ├─ 4. 校验：16 条规则（含 4 条交叉校验），结果反哺置信度
      ├─ 5. 去重：文件指纹 + 发票代码/号码 双键
      └─ 6. 输出：台账 / HTML 报告 / CSV / JSON / 待复核清单 / 审计轨迹
```

**一张发票一行**，多张发票挤在一个 PDF 里会自动拆分。

---

## 二、产出物（默认写到 `output_v2/`，与 v1 的 `output/` 刻意分开）

| 文件 | 说明 |
|---|---|
| `发票台账.xlsx` | 三区：发票明细 / 异常记录 / 汇总统计。**需装 `openpyxl`**；没装则只出下面几项 |
| `处理报告.html` | 自包含报告，KPI 卡片 + 异常清单 + 全量明细，可直接打印归档 |
| `待复核.csv` | **人工只需要看这一个文件**：异常 + 待复核 + 待人工录入三类行，8 列 |
| `发票明细.csv` / `.json` | 结构化导出。CSV 里长数字串写成 `="..."` 保住前导 0 |
| `异常记录.csv` | 异常清单单独导出 |
| `审计轨迹.jsonl` | **本次批次**每张票一行 JSON：提取器、质量分、逐字段置信度、命中的原文片段、耗时 |
| `台账索引.json` | **内部状态文件，请勿手工编辑**。跨批次恢复台账用（详见下方 FAQ） |

> ⚠ 别删 `台账索引.json`。没装 `openpyxl` 时读不了 xlsx，全靠它把上一次的台账恢复回来；
> 删了它下次运行会从空台账开始，**已有结果会被覆盖**。

---

## 三、命令行

| 命令 | 作用 |
|---|---|
| `python idle_invoice_bot.py` | 处理脚本同目录下的 `invoices_input/` |
| `-i D:\发票 -o D:\台账输出` | 指定输入目录与输出目录 |
| `--dry-run` | 只解析校验、打印结果，**不写任何文件** |
| `--force` | 忽略文件指纹，强制重新解析并替换台账中对应行（装好 OCR / 改完模板后用） |
| `--selftest` | 内置自检，**不读任何文件、不需任何依赖** |
| `--explain` | 打印生效配置、16 条规则清单、已加载模板 |
| `--quiet` | 精简输出 |

在 IDLE 里想换参数：`main()` 底部的 `parse_known_args` 已容错 IDLE 的 `sys.argv`，
也可以直接改文件顶部 `DEFAULT_CONFIG` 里的路径。

---

## 四、依赖是"可选"的，不是"必需"的

| 缺什么 | 会怎样 |
|---|---|
| 什么都不装 | **PDF 文本层照常识别**（内置标准库 PDF 解析器，覆盖 ToUnicode / UniGB-UTF16 / GBK / ASCII85+Flate 四种编码形态）；台账出 CSV+JSON 而非 xlsx |
| 没装 `openpyxl` | 不写 `发票台账.xlsx`，其余产出齐全（启动时会明确提示） |
| 没装 `Pillow` | 图片（png/jpg）登记为「待人工录入」并提示 `pip install pillow`；PDF 不受影响 |
| 没装 OCR 引擎 | 扫描件 PDF 登记为「待人工录入」；电子版 PDF 不受影响 |

想补齐能力时（按需装，不必全装）：

```bat
pip install openpyxl                      :: 出 xlsx 台账
pip install pillow                        :: 识别图片
pip install pymupdf                       :: 更稳的 PDF 文本层（装上会自动优先使用）
pip install rapidocr "onnxruntime==1.20.1" :: 扫描件 OCR
```

> ⚠ `onnxruntime` 别用 1.30：在 Windows 10.0.17763 上 `import` 即原生崩溃（0xC0000005），
> 实测 1.20.1 可用。OCR 已放在**独立子进程**里跑，引擎崩溃只影响当前文件，整批不受影响。

---

## 五、配置

优先级：`DEFAULT_CONFIG`（文件内） < `config.json`（同目录，可选） < 环境变量 `INVOICE_*`。

常用几项：

| 配置项 | 默认 | 说明 |
|---|---|---|
| `min_confidence` | `0.75` | 记录级加权置信度阈值，低于此值进「待复核」 |
| `min_critical_confidence` | `0.50` | 号码/日期/价税合计任一项低于此值 → 待复核 |
| `text_quality_min` | `0.55` | 文本质量分低于此值 → 升级提取方式 |
| `balance_tolerance` | `0.02` | 金额勾稽容差 |
| `rate_rel_tolerance` | `0.02` | 税额 vs 金额×税率 的相对容差 |
| `on_duplicate` | `skip` | `skip`=不入账 / `mark`=入账并标注重复 |
| `rules_disabled` | `[]` | 要禁用的规则 id，如 `["R16"]` |
| `ocr_enabled` | `auto` | `auto` / `off` |
| `llm_fallback.enabled` | `false` | 开启后，低置信度记录会调 LLM 兜底（需 `INVOICE_LLM_API_KEY`） |

示例 `config.json`：

```json
{
  "min_confidence": 0.8,
  "rules_disabled": ["R16"],
  "llm_fallback": { "enabled": true, "model": "gpt-4o-mini" }
}
```

---

## 六、加一种新票据版式（改这个文件，不碰代码）

编辑 `invoice_templates.json`，加一条模板即可：

```json
{
  "name": "某供应商专用版式",
  "keywords": ["某某科技有限公司", "发票号码"],
  "exclude_keywords": ["机动车销售"],
  "type": "电子普通发票",
  "fields": {
    "invoice_no": ["发票号码[:：]?([0-9]{20})"],
    "total": ["价税合计[^0-9¥￥]{0,40}?[¥￥]?(-?[0-9][0-9,]*\\.[0-9]{2})"]
  },
  "static": { "invoice_type": "电子普通发票" }
}
```

规则：
- `keywords` 命中越多越优先；`exclude_keywords` 命中则整个模板跳过；
- `fields` 每个键可给**多条正则**，按顺序取第一个命中；
- 没写的字段自动回落到内置启发式解析器（**所以模板只会让结果更好，不会让原来能识别的变差**）；
- 匹配前文本已被压成一行（去掉了所有空白），正则里不用写 `\s+`；
- 忘了格式就看 `invoice_templates.example.json`（含供应商专用版式、英文标签版式两个完整示例）。

---

## 七、16 条校验规则

| 编号 | 级别 | 内容 |
|---|---|---|
| R01 | 严重 | 必填字段缺失（号码/日期/价税合计/购销双方名称） |
| R02 | 严重 | 发票号码含非数字 |
| R03 | 警告 | 发票号码位数与版式不符 |
| R04 | 严重 | 开票日期非法 / 晚于当前日期 / 过早 |
| R05 | 严重 | 金额勾稽不平（金额+税额 ≠ 价税合计） |
| R06 | 提示 | 红字（负数）发票 / 零金额 |
| R07 | 严重 | 税号格式非法或长度罕见 |
| R08 | 警告 | 统一社会信用代码校验位不通过（GB 32100-2015） |
| R09 | 警告 | 购销双方名称完全相同 |
| R10 | 提示 | 税率不在常用集合内 |
| R11 | 警告 | 非数电发票缺少发票代码 |
| R12 | 警告 | 销售方税号缺失 |
| **R13** | **严重** | **税额与「金额×税率」不勾稽** ← 新增 |
| **R14** | **警告** | **数电号码结构异常（省级码 / 年份前缀）** ← 新增 |
| **R15** | **警告** | **发票代码结构异常（首位 0 / 票种码）** ← 新增 |
| **R16** | **警告** | **购销方名称高度相似但税号不同** ← 新增 |

R13–R16 是**组合型**校验，专抓"每个字段单独看都合法、组合起来不自洽"的识别错误——
这类错误单字段校验一定漏掉。

---

## 八、目录

```
optimized_python/
├── idle_invoice_bot.py            ★ 全部实现（单文件，可直接 F5）
├── invoice_templates.json         生效中的版式模板
├── invoice_templates.example.json  模板写法示例（改名为 invoice_templates.json 即生效）
├── config.json                    可选：外部配置覆盖（当前未提供，需要时自建）
├── output_v2/                     运行产出
└── README.md                      本文件
```

---

## 九、常见问题

**Q：为什么 `output_v2` 而不是复用 v1 的 `output`？**
A：两版台账列名基本一致但 v2 多了「识别置信度 / 提取方式」两列，直接写同一个文件会让 v1 读回时列错位。
需要合并时把 v2 的 CSV 导入即可，或统一改用 v2。

**Q：重跑会不会产生重复行？**
A：不会。先按文件内容指纹跳过已处理的文件（日志里显示 `[已处理] … 跳过`），再按「发票代码+号码」判重。

**Q：台账状态存在哪儿？为什么有时候明明有 CSV 却说"从空台账开始"？**
A：状态恢复有两条路：
1. 装了 `openpyxl` 且 `发票台账.xlsx` 存在 → 读 xlsx（人工在 Excel 里的修改会被保留）；
2. 否则 → 读 `台账索引.json`（内部状态，含文件指纹）。

两条路都走不通时才会从空台账开始，**这时已入账的结果会被覆盖**。所以：
要么装 `openpyxl`，要么不要删 `台账索引.json`。二者至少留一个。

**Q：`审计轨迹.jsonl` 只有几张票的记录？**
A：它是**批次级**的（每次运行覆盖，只记本次处理过的文件）。要长期留存审计轨迹，
把 `output_v2/审计轨迹.jsonl` 按日期改名归档，或在 CI 里用 artifact 收集。

**Q：装好 OCR / 改完版式模板之后，怎么让之前那些行重算？**
A：用 `--force`。它会先把台账里对应文件的那一行摘掉再重新解析，所以不会产生重复行：

```bat
python idle_invoice_bot.py --force           :: 重算全部
python idle_invoice_bot.py --force -i 只放那几张发票的目录
```

（UiPath 版没有这个开关，等价做法是删掉 `发票台账.xlsx` 里对应行后重跑。）

**Q：台账写不进去，报 PermissionError？**
A：`发票台账.xlsx` 被 Excel/WPS 占用了。关掉再跑；已入账的文件会自动跳过。

**Q：怎么做到"落盘即自动入账"？**
A：用 v1 的 `watch.py`（守护模式），或在任务计划程序里定时跑 `idle_invoice_bot.py`。
CI 版见仓库根 `.github/workflows/invoice-ci.yml` 的 `scheduled-run` job。

---

## 十、与 UiPath 版的关系

业务语义同源（字段口径、置信度权值、R01–R16 规则编号完全对齐），台账列名与 Config 键名一致，
**两版可以交替跑、台账能互相读回**。

- 逻辑改动要**两边同步**：Python 侧改本文件；UiPath 侧改 `../optimized_uipath/ParseInvoice.vb`，
  改完执行 `python ../optimized_uipath/build_xaml.py` 重新生成并自校验 XAML。

**改动前先跑 `--selftest`，改动后跑 `assert_ledger.py`**：

```bat
python idle_invoice_bot.py --selftest
python ..\.github\scripts\assert_ledger.py output_v2\发票明细.json
```
