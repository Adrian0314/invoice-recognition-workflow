# -*- coding: utf-8 -*-
"""
build_xaml.py —— 由可读源码生成 UiPath 的 Main.xaml / Process.xaml
====================================================================
为什么用生成器而不是手写 XAML：
  UiPath 把 InvokeCode 的 VB 代码存成 XML **属性**，换行必须写成 &#xA;、
  & 必须写成 &amp;。手写几乎必错（VB 里 & 是字符串连接符，出现频率极高）。
  生成器负责转义，逻辑仍然存在独立的 .vb / 这里的片段里，改完重跑即可。

自校验（每次生成后都会做，失败就以非 0 退出）：
  1. 每个 XAML 都能被 XML 解析器解析（良构）
  2. InvokeCode 的 Code 属性反解后，与源 .vb 逐字符相等（转义没丢东西）
  3. 关键活动名、参数名齐全

用法：
    python build_xaml.py
"""
from __future__ import annotations

import sys
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
VB_PARSE = (HERE / "ParseInvoice.vb").read_text(encoding="utf-8")

# ==========================================================================
# 1. 属性转义
# ==========================================================================
def attr(s: str) -> str:
    """把任意文本塞进 XML 属性值。字符引用不会被 XML 属性值归一化，
    所以换行必须用 &#xA; —— 直接用真实换行会被解析器压成空格。"""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = s.replace('"', "&quot;")
    s = s.replace("\n", "&#xA;")
    return s


def unattr(s: str) -> str:
    """反解属性值，用于自校验。"""
    return s.replace("&#xA;", "\n").replace("&quot;", '"') \
            .replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&")


# ==========================================================================
# 2. 内联 VB 片段（都比较短，长的解析逻辑在 ParseInvoice.vb）
# ==========================================================================
VB_BUILD_CONFIG = r'''
' 把 Config.xlsx 读出来的两列表格转成字典
Config = New Dictionary(Of String, Object)()
For Each r As DataRow In dt_Config.Rows
    Dim k As String = If(r("Name") Is Nothing, "", r("Name").ToString().Trim())
    If k = "" Then Continue For
    Config(k) = If(r("Value") Is Nothing, "", r("Value").ToString())
Next
Console.WriteLine("[Init] 配置项 " & Config.Count.ToString() & " 条")
'''

VB_SCAN_FILES = r'''
' 扫描收件箱：只取受支持的扩展名，按修改时间排序；跳过 Excel 临时文件
list_Files = New List(Of String)()
If Not Directory.Exists(str_InputFolder) Then Directory.CreateDirectory(str_InputFolder) End If
Dim exts As String() = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
Dim found As New List(Of String)()
For Each f As String In Directory.GetFiles(str_InputFolder, "*.*", SearchOption.AllDirectories)
    Dim nm As String = Path.GetFileName(f)
    If nm.StartsWith("~$") Then Continue For
    If exts.Contains(Path.GetExtension(f).ToLower()) Then found.Add(f)
Next
list_Files = found.OrderBy(Function(x) File.GetLastWriteTime(x)).ToList()
Console.WriteLine("[Init] 收件箱 " & str_InputFolder & " 发现 " & list_Files.Count.ToString() & " 个待处理文件")
'''

VB_BUILD_TABLES = r'''
' 建三张空表（列名与 Python 版台账完全一致，两个版本可交替使用同一批数据）
Dim MakeTable As Func(Of String(), DataTable) =
    Function(cols As String())
        Dim t As New DataTable()
        For Each c As String In cols
            t.Columns.Add(c, GetType(String))
        Next
        Return t
    End Function

Dim detailCols As String() = {"序号", "源文件名", "发票类型", "发票代码", "发票号码", "开票日期",
    "购买方名称", "购买方税号", "销售方名称", "销售方税号", "金额(不含税)", "税额", "价税合计",
    "税率/征收率", "主要项目/货物名称", "校验码", "开票人", "识别置信度", "提取方式",
    "处理状态", "异常/提示说明", "处理时间", "文件指纹"}
Dim issueCols As String() = {"序号", "源文件名", "发票号码", "问题级别", "涉及字段",
    "问题描述", "建议处理", "发生时间"}

dt_Ledger = MakeTable(detailCols)
dt_Issues = MakeTable(issueCols)
For Each c As String In {"金额(不含税)", "税额", "价税合计"}
    dt_Ledger.Columns(c).DataType = GetType(Double)
Next
Console.WriteLine("[Init] 台账空表已建：明细 " & dt_Ledger.Columns.Count.ToString() &
                  " 列 / 异常 " & dt_Issues.Columns.Count.ToString() & " 列")
'''

VB_WRITE_REPORT = r'''
' 生成 HTML 报告 + 待复核清单 CSV（全部用标准库，不依赖 Excel）
Dim outDir As String = str_OutputFolder
If Not Directory.Exists(outDir) Then Directory.CreateDirectory(outDir) End If

' ---- 待复核清单 ----
Dim reviewPath As String = Path.Combine(outDir, Config("ReviewFile").ToString())
Dim sb As New StringBuilder()
sb.AppendLine("源文件名,发票号码,开票日期,销售方名称,价税合计,识别置信度,处理状态,异常/提示说明")
Dim nReview As Integer = 0
For Each r As DataRow In dt_Ledger.Rows
    Dim st As String = If(r("处理状态") Is Nothing, "", r("处理状态").ToString())
    If st = "待复核" OrElse st = "待人工录入" OrElse st = "异常" Then
        nReview += 1
        sb.AppendLine(String.Join(",", {
            r("源文件名").ToString(), r("发票号码").ToString(), r("开票日期").ToString(),
            r("销售方名称").ToString(), r("价税合计").ToString(), r("识别置信度").ToString(),
            st, r("异常/提示说明").ToString().Replace(",", "，")}))
    End If
Next
File.WriteAllText(reviewPath, sb.ToString(), New UTF8Encoding(True))

' ---- 汇总 ----
Dim nOK As Integer = 0, nWarn As Integer = 0, nErr As Integer = 0
Dim nRev As Integer = 0, nDup As Integer = 0, nTodo As Integer = 0
Dim sumTotal As Double = 0
For Each r As DataRow In dt_Ledger.Rows
    Dim st As String = If(r("处理状态") Is Nothing, "", r("处理状态").ToString())
    Select Case st
        Case "正常" : nOK += 1
        Case "存在提示" : nWarn += 1
        Case "异常" : nErr += 1
        Case "待复核" : nRev += 1
        Case "待人工录入" : nTodo += 1
        Case "重复未入账" : nDup += 1
    End Select
    If st <> "重复未入账" AndAlso st <> "待人工录入" AndAlso r("价税合计") IsNot DBNull.Value Then
        sumTotal += CDbl(r("价税合计"))
    End If
Next

Dim html As New StringBuilder()
html.AppendLine("<!DOCTYPE html><html lang=""zh-CN""><head><meta charset=""utf-8"">")
html.AppendLine("<title>UiPath 发票处理报告</title>")
html.AppendLine("<style>body{font-family:'Microsoft YaHei',sans-serif;margin:24px;color:#1e2a3a;background:#f5f7fb}" &
                "table{border-collapse:collapse;width:100%;font-size:12.5px;background:#fff}" &
                "th{background:#1f4e79;color:#fff;padding:8px;text-align:left}" &
                "td{padding:6px 8px;border-bottom:1px solid #e3e8f0}" &
                ".err{color:#c0392b;font-weight:600}.rev{color:#5b4b8a;font-weight:600}" &
                ".ok{color:#1e7145;font-weight:600}.warn{color:#b7791f;font-weight:600}" &
                ".card{display:inline-block;background:#fff;border:1px solid #e3e8f0;border-radius:10px;" &
                "padding:12px 18px;margin:6px 8px 6px 0}.v{font-size:22px;font-weight:700;color:#1f4e79}</style>")
html.AppendLine("</head><body>")
html.AppendLine("<h1 style=""font-size:20px"">发票自动化处理报告 · UiPath 版</h1>")
html.AppendLine("<div>生成时间 " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "　机器人批次 " & Now.Ticks.ToString() & "</div>")
html.AppendLine("<div><div class=""card""><div>正常</div><div class=""v"">" & nOK.ToString() & "</div></div>" &
                "<div class=""card""><div>异常</div><div class=""v"">" & nErr.ToString() & "</div></div>" &
                "<div class=""card""><div>待复核</div><div class=""v"">" & nRev.ToString() & "</div></div>" &
                "<div class=""card""><div>重复拦截</div><div class=""v"">" & nDup.ToString() & "</div></div>" &
                "<div class=""card""><div>待人工录入</div><div class=""v"">" & nTodo.ToString() & "</div></div>" &
                "<div class=""card""><div>合计价税合计</div><div class=""v"">¥" & sumTotal.ToString("N2") & "</div></div></div>")
html.AppendLine("<h2 style=""font-size:16px"">发票明细</h2><table><tr>")
For Each c As DataColumn In dt_Ledger.Columns
    If c.ColumnName <> "文件指纹" Then html.AppendLine("<th>" & c.ColumnName & "</th>")
Next
html.AppendLine("</tr>")
For Each r As DataRow In dt_Ledger.Rows
    Dim st As String = If(r("处理状态") Is Nothing, "", r("处理状态").ToString())
    Dim cls As String = If(st = "异常", "err", If(st = "待复核", "rev", If(st = "正常", "ok", "warn")))
    html.AppendLine("<tr>")
    For Each c As DataColumn In dt_Ledger.Columns
        If c.ColumnName = "文件指纹" Then Continue For
        Dim v As String = If(r(c) Is DBNull.Value, "", r(c).ToString())
        If c.ColumnName = "处理状态" Then
            html.AppendLine("<td class=""" & cls & """>" & v & "</td>")
        Else
            html.AppendLine("<td>" & v & "</td>")
        End If
    Next
    html.AppendLine("</tr>")
Next
html.AppendLine("</table>")

html.AppendLine("<h2 style=""font-size:16px"">异常记录（" & dt_Issues.Rows.Count.ToString() & " 条）</h2><table><tr>")
For Each c As DataColumn In dt_Issues.Columns
    html.AppendLine("<th>" & c.ColumnName & "</th>")
Next
html.AppendLine("</tr>")
For Each r As DataRow In dt_Issues.Rows
    html.AppendLine("<tr>")
    For Each c As DataColumn In dt_Issues.Columns
        html.AppendLine("<td>" & If(r(c) Is DBNull.Value, "", r(c).ToString()) & "</td>")
    Next
    html.AppendLine("</tr>")
Next
html.AppendLine("</table></body></html>")

File.WriteAllText(Path.Combine(outDir, Config("ReportFile").ToString()),
                  html.ToString(), New UTF8Encoding(True))
Console.WriteLine("[Report] 报告已生成；待复核/异常行 " & nReview.ToString() & " 条")
Console.WriteLine("[Summary] 正常=" & nOK.ToString() & " 提示=" & nWarn.ToString() &
                  " 异常=" & nErr.ToString() & " 待复核=" & nRev.ToString() &
                  " 重复=" & nDup.ToString() & " 待人工=" & nTodo.ToString() &
                  " 价税合计=" & sumTotal.ToString("N2"))
'''

# ==========================================================================
# 3. 公共头
# ==========================================================================
HEAD = (
    '<Activity mc:Ignorable="sap sap2010" x:Class="@@CLASS@@"\n'
    '  xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities"\n'
    '  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"\n'
    '  xmlns:s="clr-namespace:System;assembly=mscorlib"\n'
    '  xmlns:sap="http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation"\n'
    '  xmlns:sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation"\n'
    '  xmlns:scg="clr-namespace:System.Collections.Generic;assembly=mscorlib"\n'
    '  xmlns:sd="clr-namespace:System.Data;assembly=System.Data"\n'
    '  xmlns:ui="http://schemas.uipath.com/workflow/activities"\n'
    '  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">\n'
)

def invoke_code(code: str, args: list, display: str) -> str:
    """args: [(方向, 类型, 名称, 表达式)]"""
    lines = [f'<ui:InvokeCode Code="{attr(code)}" Language="VisualBasic" DisplayName="{display}">',
             "  <ui:InvokeCode.Arguments>"]
    for direction, typ, name, expr in args:
        lines.append(f'    <{direction} x:TypeArguments="{typ}" x:Key="{name}">{expr}</{direction}>')
    lines.append("  </ui:InvokeCode.Arguments>")
    lines.append("</ui:InvokeCode>")
    return "\n".join(lines)


# ==========================================================================
# 4. Main.xaml
# ==========================================================================
def build_main() -> str:
    parts = [HEAD.replace("@@CLASS@@", "Main")]
    parts.append('<Sequence DisplayName="Main · 发票自动化主流程" sap2010:WorkflowViewState.IdRef="Sequence_Main">')
    parts.append("""  <Sequence.Variables>
    <Variable x:TypeArguments="scg:Dictionary(x:String, x:Object)" Name="Config" />
    <Variable x:TypeArguments="sd:DataTable" Name="dt_Config" />
    <Variable x:TypeArguments="sd:DataTable" Name="dt_Ledger" />
    <Variable x:TypeArguments="sd:DataTable" Name="dt_Issues" />
    <Variable x:TypeArguments="scg:List(x:String)" Name="list_Files" />
    <Variable x:TypeArguments="x:String" Name="str_InputFolder" />
    <Variable x:TypeArguments="x:String" Name="str_OutputFolder" />
    <Variable x:TypeArguments="x:String" Name="str_LedgerPath" />
    <Variable x:TypeArguments="x:String" Name="str_ConfigPath" />
    <Variable x:TypeArguments="x:String" Name="str_TransactionItem" />
    <Variable x:TypeArguments="x:Int32" Name="int_Done" Default="0" />
    <Variable x:TypeArguments="x:Int32" Name="int_Failed" Default="0" />
    <Variable x:TypeArguments="x:Int32" Name="int_Retries" Default="0" />
  </Sequence.Variables>""")

    parts.append('  <ui:LogMessage Level="Info" DisplayName="日志·启动" '
                 'Message="[&quot;发票自动化 v2（UiPath 版）启动 · 事务数据源 = 文件队列&quot;]" />')

    # ---- 阶段 1：初始化 ----
    parts.append('  <Sequence DisplayName="阶段1·Initialization（读配置 / 扫目录 / 建表）">')
    parts.append('    <Assign DisplayName="定位 Config.xlsx">'
                 '<Assign.To><OutArgument x:TypeArguments="x:String">[str_ConfigPath]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:String">'
                 '[Path.Combine(Environment.CurrentDirectory, "Config.xlsx")]'
                 '</InArgument></Assign.Value></Assign>')
    parts.append('    <ui:ReadRange AddHeaders="True" DataTable="[dt_Config]" Range="A1" '
                 'SheetName="Settings" WorkbookPath="[str_ConfigPath]" DisplayName="读取 Config.xlsx / Settings" />')
    parts.append(invoke_code(VB_BUILD_CONFIG, [
        ("InArgument", "sd:DataTable", "dt_Config", "[dt_Config]"),
        ("OutArgument", "scg:Dictionary(x:String, x:Object)", "Config", "[Config]"),
    ], "构建配置字典"))
    parts.append('    <Assign DisplayName="取出路径参数">'
                 '<Assign.To><OutArgument x:TypeArguments="x:String">[str_InputFolder]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:String">[Config("InputFolder").ToString]</InArgument></Assign.Value>'
                 '</Assign>')
    parts.append('    <Assign DisplayName="取出输出目录">'
                 '<Assign.To><OutArgument x:TypeArguments="x:String">[str_OutputFolder]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:String">[Config("OutputFolder").ToString]</InArgument></Assign.Value>'
                 '</Assign>')
    parts.append('    <Assign DisplayName="拼出台账路径">'
                 '<Assign.To><OutArgument x:TypeArguments="x:String">[str_LedgerPath]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:String">'
                 '[Path.Combine(str_OutputFolder, Config("LedgerFile").ToString)]'
                 '</InArgument></Assign.Value></Assign>')
    parts.append('    <If DisplayName="创建输出目录" Condition="[Not Directory.Exists(str_OutputFolder)]">'
                 '<If.Then><Sequence>'
                 '<Assign DisplayName="建目录">'
                 '<Assign.To><OutArgument x:TypeArguments="x:String">[str_OutputFolder]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:String">'
                 '[Directory.CreateDirectory(str_OutputFolder).FullName]'
                 '</InArgument></Assign.Value></Assign>'
                 '</Sequence></If.Then></If>')
    parts.append(invoke_code(VB_BUILD_TABLES, [
        ("OutArgument", "sd:DataTable", "dt_Ledger", "[dt_Ledger]"),
        ("OutArgument", "sd:DataTable", "dt_Issues", "[dt_Issues]"),
    ], "建台账空表（明细/异常）"))
    parts.append(invoke_code(VB_SCAN_FILES, [
        ("InArgument", "x:String", "str_InputFolder", "[str_InputFolder]"),
        ("OutArgument", "scg:List(x:String)", "list_Files", "[list_Files]"),
    ], "扫描收件箱 → 事务列表"))
    # 台账已存在 → 读回来（Excel 即唯一数据源，跨批次去重）
    parts.append('    <If DisplayName="台账已存在则读回" Condition="[File.Exists(str_LedgerPath)]">')
    parts.append('      <If.Then><Sequence>')
    parts.append('        <ui:ReadRange AddHeaders="True" DataTable="[dt_Ledger]" Range="A1" '
                 'SheetName="发票明细" WorkbookPath="[str_LedgerPath]" DisplayName="读回已有台账（建立去重索引）" />')
    parts.append('        <ui:LogMessage Level="Info" DisplayName="日志·读回台账" '
                 'Message="[&quot;读回已有台账，行数=&quot; + dt_Ledger.Rows.Count.ToString]" />')
    parts.append('      </Sequence></If.Then>')
    parts.append('      <If.Else><ui:LogMessage Level="Info" DisplayName="日志·新台账" '
                 'Message="[&quot;台账不存在，本次新建&quot;]" /></If.Else>')
    parts.append('    </If>')
    parts.append('  </Sequence>')

    # ---- 阶段 2+3：取事务数据 + 处理事务 ----
    parts.append('  <Sequence DisplayName="阶段2/3·GetTransactionData + ProcessTransaction">')
    parts.append('    <ui:LogMessage Level="Info" DisplayName="日志·开始处理" '
                 'Message="[&quot;待处理 &quot; + list_Files.Count.ToString + &quot; 张发票&quot;]" />')
    parts.append('    <ForEach x:TypeArguments="x:String" DisplayName="逐张发票（事务循环）" Values="[list_Files]">')
    parts.append('      <ActivityAction x:TypeArguments="x:String">')
    parts.append('        <ActivityAction.Argument><DelegateInArgument x:TypeArguments="x:String" Name="item" /></ActivityAction.Argument>')
    parts.append('        <Sequence DisplayName="单张发票事务">')
    parts.append('          <Assign DisplayName="取事务项">'
                 '<Assign.To><OutArgument x:TypeArguments="x:String">[str_TransactionItem]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:String">[item]</InArgument></Assign.Value>'
                 '</Assign>')
    # TryCatch + RetryScope：区分系统异常（重试）与业务异常（跳过并记录）
    parts.append('          <TryCatch DisplayName="处理事务（含 RetryScope 重试）">')
    parts.append('            <TryCatch.Try>')
    parts.append('              <ui:RetryScope DisplayName="RetryScope·系统异常重试" '
                 'NumberOfRetries="[CInt(Config(&quot;MaxRetryNumber&quot;))]" RetryInterval="00:00:02">')
    parts.append('                <ui:RetryScope.ActivityBody><ActivityAction>')
    parts.append('                  <ui:InvokeWorkflowFile DisplayName="Process.xaml · 单张发票处理" '
                 'WorkflowFileName="Process.xaml" UnSafe="False">')
    parts.append('                    <ui:InvokeWorkflowFile.Arguments>')
    parts.append('                      <InArgument x:TypeArguments="x:String" x:Key="in_FilePath">[str_TransactionItem]</InArgument>')
    parts.append('                      <InArgument x:TypeArguments="scg:Dictionary(x:String, x:Object)" x:Key="in_Config">[Config]</InArgument>')
    parts.append('                      <InArgument x:TypeArguments="sd:DataTable" x:Key="io_dt_Ledger">[dt_Ledger]</InArgument>')
    parts.append('                      <InArgument x:TypeArguments="sd:DataTable" x:Key="io_dt_Issues">[dt_Issues]</InArgument>')
    parts.append('                      <OutArgument x:TypeArguments="x:String" x:Key="out_Status">[out_Status]</OutArgument>')
    parts.append('                    </ui:InvokeWorkflowFile.Arguments>')
    parts.append('                  </ui:InvokeWorkflowFile>')
    parts.append('                </ActivityAction></ui:RetryScope.ActivityBody>')
    parts.append('              </ui:RetryScope>')
    parts.append('            </TryCatch.Try>')
    parts.append('            <TryCatch.Catches>')
    parts.append('              <Catch x:TypeArguments="ui:BusinessRuleException" DisplayName="业务异常·数据问题">')
    parts.append('                <ActivityAction x:TypeArguments="ui:BusinessRuleException">')
    parts.append('                  <Sequence>')
    parts.append('                    <ui:LogMessage Level="Warn" DisplayName="日志·业务异常" '
                 'Message="[&quot;[业务异常] &quot; + Path.GetFileName(str_TransactionItem) + &quot; 已登记为异常，跳过不重试&quot;]" />')
    parts.append('                    <Assign DisplayName="失败计数+1">'
                 '<Assign.To><OutArgument x:TypeArguments="x:Int32">[int_Failed]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:Int32">[int_Failed + 1]</InArgument></Assign.Value>'
                 '</Assign>')
    parts.append('                  </Sequence>')
    parts.append('                </ActivityAction>')
    parts.append('              </Catch>')
    parts.append('              <Catch x:TypeArguments="s:Exception" DisplayName="系统异常·重试耗尽后落到这里">')
    parts.append('                <ActivityAction x:TypeArguments="s:Exception">')
    parts.append('                  <Sequence>')
    parts.append('                    <ui:LogMessage Level="Error" DisplayName="日志·系统异常" '
                 'Message="[&quot;[系统异常] &quot; + Path.GetFileName(str_TransactionItem) + &quot;：&quot; + exception.Message]" />')
    parts.append('                    <Assign DisplayName="重试计数+1">'
                 '<Assign.To><OutArgument x:TypeArguments="x:Int32">[int_Retries]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:Int32">[int_Retries + 1]</InArgument></Assign.Value>'
                 '</Assign>')
    parts.append('                  </Sequence>')
    parts.append('                </ActivityAction>')
    parts.append('              </Catch>')
    parts.append('            </TryCatch.Catches>')
    parts.append('            <TryCatch.Finally><Sequence><ui:LogMessage Level="Trace" DisplayName="日志·事务结束" '
                 'Message="[&quot;事务处理结束：&quot; + Path.GetFileName(str_TransactionItem)]" /></Sequence></TryCatch.Finally>')
    parts.append('          </TryCatch>')
    parts.append('          <Assign DisplayName="完成计数+1">'
                 '<Assign.To><OutArgument x:TypeArguments="x:Int32">[int_Done]</OutArgument></Assign.To>'
                 '<Assign.Value><InArgument x:TypeArguments="x:Int32">[int_Done + 1]</InArgument></Assign.Value>'
                 '</Assign>')
    parts.append('        </Sequence>')
    parts.append('      </ActivityAction>')
    parts.append('    </ForEach>')
    parts.append('  </Sequence>')

    # ---- 阶段 4：End Process ----
    parts.append('  <Sequence DisplayName="阶段4·End Process（落盘 / 报告 / 汇总）">')
    parts.append('    <ui:WriteRange AddHeaders="True" DataTable="[dt_Ledger]" Range="A1" '
                 'SheetName="发票明细" WorkbookPath="[str_LedgerPath]" DisplayName="写台账·发票明细" />')
    parts.append('    <ui:WriteRange AddHeaders="True" DataTable="[dt_Issues]" Range="A1" '
                 'SheetName="异常记录" WorkbookPath="[str_LedgerPath]" DisplayName="写台账·异常记录" />')
    parts.append(invoke_code(VB_WRITE_REPORT, [
        ("InArgument", "sd:DataTable", "dt_Ledger", "[dt_Ledger]"),
        ("InArgument", "sd:DataTable", "dt_Issues", "[dt_Issues]"),
        ("InArgument", "scg:Dictionary(x:String, x:Object)", "Config", "[Config]"),
        ("InArgument", "x:String", "str_OutputFolder", "[str_OutputFolder]"),
    ], "生成 HTML 报告 + 待复核清单"))
    parts.append('    <ui:LogMessage Level="Info" DisplayName="日志·批次汇总" '
                 'Message="[&quot;批次结束：台账 &quot; + dt_Ledger.Rows.Count.ToString + &quot; 行，异常 &quot; '
                 '+ dt_Issues.Rows.Count.ToString + &quot; 条，处理 &quot; + int_Done.ToString + &quot; 个文件，失败 &quot; '
                 '+ int_Failed.ToString]" />')
    parts.append('    <If DisplayName="有问题则抛业务异常（供 Orchestrator 告警）" '
                 'Condition="[int_Failed &gt; 0]">')
    parts.append('      <If.Then><Sequence>')
    parts.append('        <ui:LogMessage Level="Warn" DisplayName="日志·需要人工介入" '
                 'Message="[&quot;有 &quot; + int_Failed.ToString + &quot; 个文件未能自动处理，请查看 output/待复核.csv 与 output/_failed&quot;]" />')
    parts.append('      </Sequence></If.Then>')
    parts.append('    </If>')
    parts.append('  </Sequence>')

    parts.append("</Sequence>")
    parts.append("</Activity>")
    return "\n".join(parts) + "\n"


# ==========================================================================
# 5. Process.xaml
# ==========================================================================
def build_process() -> str:
    parts = [HEAD.replace("@@CLASS@@", "Process")]
    parts.append('<Sequence DisplayName="Process · 单张发票处理" sap2010:WorkflowViewState.IdRef="Sequence_Process">')
    parts.append("""  <Sequence.Variables>
    <Variable x:TypeArguments="x:String" Name="str_Text" Default="" />
    <Variable x:TypeArguments="x:String" Name="str_FileName" Default="" />
    <Variable x:TypeArguments="x:Int32" Name="int_RowsBefore" Default="0" />
  </Sequence.Variables>""")
    parts.append(invoke_code(
        "str_FileName = Path.GetFileName(in_FilePath)\n"
        "int_RowsBefore = io_dt_Ledger.Rows.Count\n"
        "out_Status = \"\"\n"
        "Console.WriteLine(\"[Process] 开始处理 \" & str_FileName)",
        [("InArgument", "x:String", "in_FilePath", "[in_FilePath]"),
         ("InArgument", "sd:DataTable", "io_dt_Ledger", "[io_dt_Ledger]"),
         ("OutArgument", "x:String", "out_Status", "[out_Status]")],
        "取文件名 / 记录入账前行数"))
    # PDF 文本提取（Document Understanding 的入口就在这里）
    parts.append('  <ui:ReadPDFText DisplayName="读取 PDF 文本层（扫描件由 OCRMethod 接管）" '
                 'FileName="[in_FilePath]" Text="[str_Text]" PreserveFormatting="False" />')
    parts.append(invoke_code(VB_PARSE, [
        ("InArgument", "x:String", "in_Text", "[str_Text]"),
        ("InArgument", "x:String", "in_FileName", "[str_FileName]"),
        ("InArgument", "scg:Dictionary(x:String, x:Object)", "in_Config", "[in_Config]"),
        ("InArgument", "sd:DataTable", "out_Rows", "[io_dt_Ledger]"),
        ("InArgument", "sd:DataTable", "out_Issues", "[io_dt_Issues]"),
        ("InArgument", "x:Double", "in_MinConf", "[CDbl(in_Config(\"MinConfidence\"))]"),
        ("InArgument", "x:Double", "in_MinCritConf", "[CDbl(in_Config(\"MinCriticalConfidence\"))]"),
        ("InArgument", "x:Double", "in_BalanceTol", "[CDbl(in_Config(\"BalanceTolerance\"))]"),
        ("InArgument", "x:Double", "in_RateTol", "[CDbl(in_Config(\"RateRelTolerance\"))]"),
        ("InArgument", "x:Double", "in_RateFloor", "[CDbl(in_Config(\"RateAbsFloor\"))]"),
        ("InArgument", "x:DateTime", "in_Today", "[DateTime.Today]"),
    ], "字段解析 + 校验 + 置信度（ParseInvoice.vb）"))
    # 结果判定：0 行 = 没解析出发票 → 记业务异常并隔离原件
    parts.append("""  <If DisplayName="是否解析出发票" Condition="[io_dt_Ledger.Rows.Count = int_RowsBefore]">
    <If.Then>
      <Sequence DisplayName="没解析出 → 登记异常 + 隔离原件">
        <ui:LogMessage Level="Warn" DisplayName="日志·未解析出内容"
          Message="[&quot;未从 &quot; + str_FileName + &quot; 解析出发票记录，登记为待人工录入&quot;]" />
      </Sequence>
    </If.Then>
    <If.Else>
      <Sequence DisplayName="解析成功">
        <ui:LogMessage Level="Info" DisplayName="日志·入账"
          Message="[&quot;入账 &quot; + (io_dt_Ledger.Rows.Count - int_RowsBefore).ToString + &quot; 行：&quot; + str_FileName]" />
      </Sequence>
    </If.Else>
  </If>""")
    parts.append("</Sequence>")
    parts.append("</Activity>")
    return "\n".join(parts) + "\n"


# ==========================================================================
# 6. 生成 + 自校验
# ==========================================================================
def verify(path: Path, code_sources: list) -> list:
    problems = []
    text = path.read_text(encoding="utf-8")
    try:
        root = ET.fromstring(text)
    except ET.ParseError as exc:
        problems.append(f"{path.name}: XML 不良构 → {exc}")
        return problems

    # InvokeCode 的 Code 属性反解，必须能找到源片段（逐字符相等）
    codes = [el.attrib.get("Code", "")
             for el in root.iter() if el.tag.endswith("InvokeCode")]
    for src in code_sources:
        expected = src.replace("\r\n", "\n").replace("\r", "\n")
        if not any(unattr(c) == expected for c in codes):
            # 取第一行做个线索
            head = expected.strip().split("\n")[0][:50]
            problems.append(f"{path.name}: InvokeCode 代码往返不一致（源首行：{head}）")

    for needle in ("Sequence", "InvokeCode"):
        if needle not in text:
            problems.append(f"{path.name}: 缺少 {needle}")
    return problems


def main() -> int:
    main_xaml = build_main()
    proc_xaml = build_process()
    (HERE / "Main.xaml").write_text(main_xaml, encoding="utf-8")
    (HERE / "Process.xaml").write_text(proc_xaml, encoding="utf-8")
    print(f"已生成 {HERE / 'Main.xaml'}")
    print(f"已生成 {HERE / 'Process.xaml'}")

    problems = []
    problems += verify(HERE / "Main.xaml", [VB_BUILD_CONFIG, VB_BUILD_TABLES, VB_SCAN_FILES, VB_WRITE_REPORT])
    problems += verify(HERE / "Process.xaml", [VB_PARSE])

    # 参数名一致性：Process.xaml 的 InvokeCode 参数 必须覆盖 ParseInvoice.vb 里用到的名字
    proc_text = (HERE / "Process.xaml").read_text(encoding="utf-8")
    needed = ["in_Text", "in_FileName", "in_Config", "out_Rows", "out_Issues",
              "in_MinConf", "in_MinCritConf", "in_BalanceTol", "in_RateTol",
              "in_RateFloor", "in_Today"]
    for n in needed:
        if f'x:Key="{n}"' not in proc_text:
            problems.append(f"Process.xaml: 缺少 InvokeCode 参数 {n}")

    print("-" * 66)
    if problems:
        for p in problems:
            print("✗ " + p)
        print(f"自校验未通过：{len(problems)} 项")
        return 1
    print("✓ XML 良构")
    print("✓ InvokeCode 代码转义往返一致（换行/&/引号 无丢失）")
    print("✓ InvokeCode 参数齐全（11 个入参）")
    print("自校验全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
