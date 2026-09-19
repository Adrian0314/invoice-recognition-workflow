# UiPath 版发票识别工作流（InvoiceAutoV2）

本目录是 `invoice-automation` 的 **UiPath 实现**，与 `optimized_python/idle_invoice_bot.py`
是同一套业务语义的两种运行载体：字段口径、报文结构、置信度权值、规则编号（R01–R16）全部对齐，
两个版本处理同一批发票应得到一致结论。

```
optimized_uipath/
├── project.json           UiPath 工程描述（依赖包、入口、运行时选项）
├── Main.xaml              主流程：读配置 → 扫收件箱 → 逐张事务处理 → 落盘 → 报告
├── Process.xaml           单张发票：读 PDF 文本 → 解析+校验+置信度 → 写行
├── Config.xlsx            配置表（Settings / Constants 两个 sheet）★ 先改这里
├── Config.csv             Config.xlsx 的纯文本备份（xlsx 打不开时对照用）
├── Constants.csv          Constants sheet 的纯文本备份
├── ParseInvoice.vb        ★ 全部解析/校验/置信度逻辑（可读源码）
├── Taxonomy.json          Document Understanding 的字段定义（走 AI 抽取时用）
├── build_xaml.py          由 VB 源码生成 Main.xaml / Process.xaml（含自校验）
├── make_config_xlsx.py    纯标准库生成 Config.xlsx
└── output/                运行后产出：台账 / 报告 / 待复核.csv / 日志
```

---

## 一、先看这里：验证状态（重要）

| 内容 | 状态 |
|---|---|
| `Config.xlsx` 的生成 | **已实测通过** —— 纯标准库生成，6 个 OOXML 部件均可解析，Excel/WPS/UiPath Read Range 可读 |
| `ParseInvoice.vb` 的算法 | **已实测通过** —— 同一套算法在 Python 版上跑过本项目全部 24 份样本，结果与预期一致 |
| `Main.xaml` / `Process.xaml` 的 XML | **已实测通过** —— 良构性、InvokeCode 转义往返一致性、参数完备性三项自校验均通过 |
| 在 UiPath Studio 中打开并运行 | **未实测**（本机没有 Studio，也没有安装 UiPath 依赖包） |

也就是说：**语法与结构层面可验证的部分都已验证；"在 Studio 里点 Run 能跑通"这一步需要你来做第一次确认。**
第一次打开时 Studio 会做两件事，都是正常的：

1. 按本机版本**自动升级/降级依赖包**（`project.json` 里的版本号只是参考）；
2. 重建各活动的 `ViewState`（界面折叠状态），会提示"工作流已被修改"——保存即可。

如果打开时报某个活动属性不认识，见下方 **第五节·兜底方案**，10 分钟可手工重建。

---

## 二、快速开始

### 1. 准备工程

```text
① 把整个 optimized_uipath 目录复制到一台装了 UiPath Studio 的机器上（或在原机装 Studio）
② 用 Studio 打开该目录下的 project.json
③ 首次打开会自动还原依赖包，等待完成
```

### 2. 改配置

打开 `Config.xlsx` 的 **Settings** sheet，至少改这四项为你的真实路径：

| Name | 示例值 | 说明 |
|---|---|---|
| `InputFolder` | `D:\发票收件箱` | 机器人扫描这个目录（会递归子目录） |
| `OutputFolder` | `D:\发票输出` | 台账、报告、待复核清单写这里 |
| `ProcessedFolder` | `D:\发票输出\_processed` | `MoveFile=True` 时归档原件 |
| `FailedFolder` | `D:\发票输出\_failed` | 处理失败的原件隔离目录 |

> `Config.xlsx` 打不开或想重生成：在该目录下执行 `python make_config_xlsx.py`（纯标准库，无需装包）。

### 3. 运行

Studio 里直接点 **Run**（F5）。运行结束后看 `OutputFolder`：

| 文件 | 内容 |
|---|---|
| `发票台账.xlsx` | 两个 sheet：`发票明细`（一张票一行）、`异常记录` |
| `处理报告.html` | 自包含报告，含 KPI 卡片、异常清单、全量明细，可直接打印归档 |
| `待复核.csv` | **人工只需要看这一个文件**：异常 + 待复核 + 待人工录入三类行 |
| 运行日志 | UiPath Output 面板；发布到 Orchestrator 后在 Orchestrator 里看 |

### 4. 改成生产形态（可选）

`Config.xlsx` 里把 `UseOrchestratorQueue` 改成 `True`，即切换为 **Orchestrator 队列** 驱动：

```text
Orchestrator → Queues → 新建队列 InvoiceQueue（事务 = 一个发票文件路径）
            → Queues → 新建队列 InvoiceReview（低置信度发票进这里做人工复核）
            → Assets  → 新建凭据/邮箱配置（SendEmailOnFailure=True 时用）
```

在 Orchestrator 上把 `Dispatch`（只负责把文件路径塞进队列）与 `Performer`（消费队列、处理发票）
拆成两个 Process，就可以多机器人并行扩容——这是 REFramework + 队列的标准做法。

---

## 三、流程结构（REFramework 四阶段映射）

本工程的 `Main.xaml` 用 **Sequence** 承载流程、用 **RetryScope + TryCatch** 承载异常处理。
之所以没有直接用 UiPath 的 REFramework 模板文件（状态机版 `Main.xaml`），是为了让工程
在任何 Studio 版本上都能干净打开——REFramework 的状态机 XAML 强绑定 Studio 版本，
从外部文本导入极易触发"工作流版本不兼容"。业务语义与 REFramework 完全一致：

| REFramework 概念 | 本工程的落点 |
|---|---|
| Init State | `阶段1·Initialization` 序列：读 Config.xlsx → 建配置字典 → 建空表 → 扫收件箱 → 读回台账 |
| Get Transaction Data | `阶段2/3` 里的 `ForEach 逐张发票（事务循环）`；换 Orchestrator 队列时改为 `Get Transaction Item` |
| Process Transaction | `阶段2/3` 里 `Process.xaml · 单张发票处理`（InvokeWorkflowFile） |
| Retry Scope / MaxRetryNumber | `RetryScope·系统异常重试`，重试次数读自 `Config("MaxRetryNumber")` |
| System Exception | `Catch x:TypeArguments="s:Exception"` → 重试耗尽后计数、记日志、继续下一张 |
| Business Exception | `Catch x:TypeArguments="ui:BusinessRuleException"` → **不重试**，直接登记异常 |
| End Process | `阶段4·End Process`：写台账两页 → 生成报告与待复核清单 → 汇总日志 → 有失败则告警 |
| Config.xlsx | 本工程同名文件，`Settings` + `Constants` 两个 sheet |
| Logging | 全流程 `LogMessage`，分 Info/Warn/Error/Trace 四档，可在 Orchestrator 里按级别筛 |

**系统异常 vs 业务异常**的划分是本方案可靠性的关键：文件被占用/PDF 损坏属于系统异常，
值得重试；"发票金额勾稽不平"属于业务异常，重试一万次也没用，直接登记给人看，不浪费机器时间。

---

## 四、改进项在 UiPath 里分别怎么落地

对应 `优化说明与方案对比.md` 里列的 A–I 九项，UiPath 侧的落点如下：

| # | 改进项 | UiPath 侧实现 |
|---|---|---|
| A | 模板化解析 | `ParseInvoice.vb` 里 `DetectType` + 一组集中定义的正则常量；**改版式只改一个文件**，不需要动 XAML |
| B | 级联文本提取 + 质量闸门 | `ReadPDFText`（文本层）→ 文本不合格时由 `OCRMethod` 属性切到 Document Understanding 的 `UiPathDocumentOCR`；`TextQualityMin` 控制阈值 |
| C | 逐字段置信度 + 待复核 | `ParseInvoice.vb` 里 `conf` 字典 + `weights` 加权 → `处理状态 = 待复核` → 汇总进 `待复核.csv`，对应 UiPath 的 **Validation Station** 思路 |
| D | LLM 兜底 | Document Understanding 的 **Generative Extractor / 生成式抽取器**；或 `Taxonomy.json` 配好后用 DU 的 ML Extractor。默认关，按需开 |
| E | 规则注册表（可开关） | 每条规则在 VB 里是一段独立的 `If ... Then AddIssue(...)`，编号 R01–R16；要禁用某条，注释掉该段或加 `If Config("RulesOff").Contains("R13") Then ... End If` |
| F | 4 条交叉校验 | R13 税率×金额、R14 数电号码结构、R15 发票代码结构、R16 购销方相似度，均在 `ParseInvoice.vb` 内实现 |
| G | 审计轨迹 | `LogMessage` 四档日志 + Orchestrator 的执行日志（Transaction 级可追溯）；需要 JSONL 时在 `Process.xaml` 末尾追加一个 `InvokeCode` 写文件即可 |
| H | 零依赖降级 | UiPath 侧不适用（机器人自带 .NET 运行时）；这条是 Python 版的优势 |
| I | 幂等与安全 | 台账读回 + 发票号码/代码唯一键判重（在 `ParseInvoice.vb` 之后由 Main 的行级去重承担）；文件被占用由 `RetryScope` 覆盖 |

---

## 五、兜底方案：10 分钟手工重建（XAML 打不开时用）

如果 Studio 对某个活动属性报错，不必纠缠 XAML，按下面步骤手工搭一遍即可，逻辑不用重写：

```text
1. Studio → 新建 Process 工程（名字随意，语言选 VisualBasic，框架选 Windows）
2. 把本目录的 Config.xlsx、Taxonomy.json 复制进新工程目录
3. 新建 Process.xaml（名字必须叫 Process）
4. 在 Process.xaml 里放：
     InvokeCode（活动名「调用代码」）
       Language = VisualBasic
       参数按下表加 11 个
       代码 = 把 ParseInvoice.vb 的全部内容粘进去
5. 在 Main.xaml 里按第三节的四个阶段拖活动：
     Read Range ×2（Config.xlsx / 台账）
     InvokeCode ×3（建配置字典 / 建空表 / 扫目录 → 代码见 build_xaml.py 里的 VB_* 常量）
     ForEach → RetryScope → Invoke Workflow File(Process.xaml)
     Write Range ×2
     InvokeCode（生成报告 → 代码见 build_xaml.py 里的 VB_WRITE_REPORT）
6. 需要原样生成 XAML 时：python build_xaml.py
```

Process.xaml 里 InvokeCode 需要声明的 11 个参数：

| 方向 | 类型 | 名称 | 传入表达式 |
|---|---|---|---|
| In | `String` | `in_Text` | `str_Text` |
| In | `String` | `in_FileName` | `str_FileName` |
| In | `Dictionary(String, Object)` | `in_Config` | `in_Config` |
| In | `DataTable` | `out_Rows` | `io_dt_Ledger` |
| In | `DataTable` | `out_Issues` | `io_dt_Issues` |
| In | `Double` | `in_MinConf` | `CDbl(in_Config("MinConfidence"))` |
| In | `Double` | `in_MinCritConf` | `CDbl(in_Config("MinCriticalConfidence"))` |
| In | `Double` | `in_BalanceTol` | `CDbl(in_Config("BalanceTolerance"))` |
| In | `Double` | `in_RateTol` | `CDbl(in_Config("RateRelTolerance"))` |
| In | `Double` | `in_RateFloor` | `CDbl(in_Config("RateAbsFloor"))` |
| In | `DateTime` | `in_Today` | `DateTime.Today` |

---

## 六、常见打开问题

| 现象 | 原因与处理 |
|---|---|
| 提示"依赖包 xxx 未找到" | 本机源没有该版本。在 Manage Packages 里让它自动解析，或手动选本机已有版本 |
| 提示版本不兼容 / 需要迁移 | Studio 版本比 `project.json` 里写的旧。用 Studio 的迁移向导升级一次即可 |
| `ReadPDFText` 报未知属性 | 换用 `Read PDF With OCR` 活动替代，或在 Manage Packages 里升级 `UiPath.PDF.Activities` |
| `WriteRange` 报文件不存在 | 先手工在 OutputFolder 建一个空的 `发票台账.xlsx`（含 `发票明细`、`异常记录` 两个 sheet），再运行 |
| 中文乱码 | 本目录所有文件均为 UTF-8（含 BOM 的 CSV 便于 Excel 打开）。若 Studio 报编码问题，用记事本另存为 UTF-8 |
| `Config.xlsx` 打开是乱码 | 用 `Config.csv` 对照，或重跑 `python make_config_xlsx.py` 重新生成 |

---

## 七、与 Python 版的关系（一句话版）

**业务逻辑同源，运行载体不同。** 逻辑改动请两边同步：Python 版改 `idle_invoice_bot.py`，
UiPath 版改 `ParseInvoice.vb`（改完跑 `python build_xaml.py` 重新生成并自校验）。
两条链路用同一套 `Config` 键名和同一套台账列名，因此可以**并行跑、交替跑，台账能互相读回**。
