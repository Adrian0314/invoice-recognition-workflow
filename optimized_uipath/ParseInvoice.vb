' ##########################################################################
' ParseInvoice.vb —— UiPath InvokeCode 活动的代码体（VisualBasic）
' ##########################################################################
'
' 【重要】本文件不是独立可编译的 .vb 文件，而是 Process.xaml 里
'         「InvokeCode」活动 Code 属性的内容。UiPath 会把这段代码当作
'         一个方法体来编译，所以这里只有语句，没有 Sub/End Sub 外壳。
'
' 由 build_xaml.py 读取本文件 → 转义换行与引号 → 写入 Process.xaml 的
' InvokeCode 节点。改逻辑只改这里，然后重跑 build_xaml.py 即可。
'
' ---------- InvokeCode 活动需要声明这些参数 ----------
'   in_Text         In      String                          发票原始文本
'   in_FileName     In      String                          源文件名
'   in_Config       In      Dictionary(Of String, Object)  配置字典（见 Config.xlsx）
'   out_Rows        In/Out  DataTable                       明细行（一张票一行）
'   out_Issues      In/Out  DataTable                       异常记录
'
' ---------- 与 Python 版的一致性 ----------
' 字段口径、置信度权值、规则编号（R01…R16）与 idle_invoice_bot.py 完全对齐，
' 两个版本处理同一批发票应得到同样的结论。这是本方案可维护性的关键：
' 规则只有一套语义，两边只是运行载体不同。
'
' ---------- 与 v1（invoice_wf 包）相比补了什么 ----------
'   · 逐字段置信度 + 记录级加权置信度 → 低置信度进「待复核」（R 之外的把关）
'   · 4 条交叉校验：税额×金额、数电号码结构、发票代码结构、购销方相似度
'   · 校验结果反哺置信度（先校验再加/减分，形成闭环）
'   · 状态由「正常/异常」两态扩为 4 态，人工只看真正可疑的行
' ##########################################################################

Dim t0 As DateTime = DateTime.Now

' ======================= 0. 小工具 =======================
Dim reWS As New Regex("\s+")
Dim reLineWS As New Regex("[ \t]+")
Dim Flat As Func(Of String, String) = Function(s As String) reWS.Replace(s, "")

' 金额：整数 + 千分位 + 恰好两位小数
Const MoneyPattern As String = "-?\d[\d,]*\.\d{1,2}"
' 名称字段的终止符（防止把相邻单元格内容一起抓进来）
Const NameStop As String = "(?=销售方|购买方|销货方|购货方|统一社会信用代码|纳税人识别号|税号|项目名称|货物或应税|规格型号|单位|数量|单价|金额|税率|征收率|价税合计|合计|备注|开票人|收款人|复核人|地址|开户行|电话|账号|$)"

Dim ToMoney As Func(Of String, Double) =
    Function(s As String)
        If String.IsNullOrEmpty(s) Then Return Double.NaN
        Dim v As String = s.Replace(",", "").Replace("¥", "").Replace("￥", "").Trim()
        Dim d As Double
        If Double.TryParse(v, d) Then Return Math.Round(d, 2)
        Return Double.NaN
    End Function

Dim Norm As Func(Of String, String) =
    Function(s As String)
        Dim x As String = s.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        x = x.Replace(ChrW(160), " "c).Replace(ChrW(12288), " "c) _
             .Replace(ChrW(8203), " "c).Replace(ChrW(65279), " "c)
        Return x.Normalize(System.Text.NormalizationForm.FormKC)
    End Function

' ======================= 1. 取文本、切多张票 =======================
Dim Full As String = Norm(in_Text)
Dim One As String = Flat(Full)                       ' 全压平：跨单元格定位关键词最稳
Dim RawLines() As String = Full.Split(ChrW(10))

' 一个文件可能含多张发票（连打/合并导出）→ 按「发票号码」或标题行切段
Dim Segs As New List(Of String)()
Dim curSeg As New List(Of String)()
Dim reTitle As New Regex("^(电子发票|全电发票|增值税|机动车销售|二手车销售)")
For Each ln As String In RawLines
    Dim L As String = reLineWS.Replace(ln, "").Trim()
    If L = "" Then Continue For
    Dim joined As String = String.Join("", curSeg.ToArray())
    Dim hasCore As Boolean = (joined.Contains("价税合计") OrElse joined.Contains("税额")) _
                             AndAlso joined.Contains("开票日期")
    Dim isTitle As Boolean = reTitle.IsMatch(L) AndAlso L.Contains("发票") _
                             AndAlso Not L.Contains("号码") AndAlso Not L.Contains("代码")
    If curSeg.Count > 0 AndAlso hasCore AndAlso (L.Contains("发票号码") OrElse isTitle) Then
        Segs.Add(String.Join(vbLf, curSeg.ToArray()))
        curSeg.Clear()
    End If
    curSeg.Add(L)
Next
If curSeg.Count > 0 Then Segs.Add(String.Join(vbLf, curSeg.ToArray()))
If Segs.Count = 0 Then Segs.Add(Full)

' ======================= 2. 版式判定 =======================
Dim DetectType As Func(Of String, String) =
    Function(o As String)
        If o.Contains("数电") OrElse o.Contains("全电") Then Return "数电发票"
        If o.Contains("电子发票") AndAlso o.Contains("专用发票") Then Return "数电发票(增值税专用发票)"
        If o.Contains("电子发票") AndAlso o.Contains("普通发票") Then Return "数电发票(普通发票)"
        If o.Contains("增值税电子专用发票") Then Return "电子专用发票"
        If o.Contains("增值税电子普通发票") Then Return "电子普通发票"
        If o.Contains("机动车销售统一发票") Then Return "机动车销售统一发票"
        If o.Contains("二手车销售统一发票") Then Return "二手车销售统一发票"
        If o.Contains("增值税专用发票") Then Return "增值税专用发票"
        If o.Contains("增值税普通发票") Then Return "增值税普通发票"
        If o.Contains("发票") Then Return "电子发票(其他)"
        Return "未知类型"
    End Function

' ======================= 3. 校验位算法（GB 32100-2015）=======================
Const UsccChars As String = "0123456789ABCDEFGHJKLMNPQRTUWXY"
Dim UsccW As Integer() = {1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28}
Dim UsccCheck As Func(Of String, Integer) =          ' 1=通过 0=不通过 -1=不适用
    Function(code As String)
        If code Is Nothing Then Return -1
        code = code.Trim().ToUpper()
        If code.Length <> 18 Then Return -1
        Dim total As Integer = 0
        For i As Integer = 0 To 16
            Dim idx As Integer = UsccChars.IndexOf(code(i))
            If idx < 0 Then Return 0
            total += idx * UsccW(i)
        Next
        Dim c As Integer = 31 - (total Mod 31)
        If c = 31 Then c = 0
        If UsccChars(c) = code(17) Then Return 1
        Return 0
    End Function

Dim ProvCodes As String() = {"11", "12", "13", "14", "15", "21", "22", "23", "31", "32",
    "33", "34", "35", "36", "37", "41", "42", "43", "44", "45", "46", "50", "51", "52",
    "53", "54", "61", "62", "63", "64", "65", "71"}

' ======================= 4. 逐段解析 =======================
Dim Idx As Integer = 0
For Each seg As String In Segs
    Idx += 1
    Dim o As String = Flat(seg)
    Dim conf As New Dictionary(Of String, Double)()
    Dim ev As New Dictionary(Of String, String)()

    Dim FType As String = DetectType(o)
    Dim FCode As String = ""
    Dim FNo As String = ""
    Dim FDate As String = ""
    Dim FBuyer As String = ""
    Dim FBuyerTax As String = ""
    Dim FSeller As String = ""
    Dim FSellerTax As String = ""
    Dim FAmount As Double = Double.NaN
    Dim FTax As Double = Double.NaN
    Dim FTotal As Double = Double.NaN
    Dim FRate As String = ""
    Dim FItem As String = ""
    Dim FCheck As String = ""
    Dim FDrawer As String = ""

    ' --- 发票号码 / 代码 ---
    Dim m As Match = Regex.Match(o, "发票号码[:：]?\s*(\d{8,20})")
    If m.Success Then
        FNo = m.Groups(1).Value : conf("invoice_no") = 0.95 : ev("invoice_no") = m.Value
    End If
    m = Regex.Match(o, "发票代码[:：]?\s*(\d{10,12})")
    If m.Success Then
        FCode = m.Groups(1).Value : conf("invoice_code") = 0.93 : ev("invoice_code") = m.Value
    End If

    ' --- 开票日期 ---
    m = Regex.Match(o, "开票日期[:：]?\s*(\d{4})[年\-/.]?(\d{1,2})[月\-/.]?(\d{1,2})日?")
    If m.Success Then
        Dim y As Integer = CInt(m.Groups(1).Value)
        Dim mo As Integer = CInt(m.Groups(2).Value)
        Dim d As Integer = CInt(m.Groups(3).Value)
        Try
            FDate = New DateTime(y, mo, d).ToString("yyyy-MM-dd")
            conf("issue_date") = 0.95
        Catch
            FDate = y.ToString("0000") & "-" & mo.ToString("00") & "-" & d.ToString("00")
            conf("issue_date") = 0.35
        End Try
        ev("issue_date") = m.Value
    End If

    ' --- 校验码（老版电子发票才有）---
    m = Regex.Match(o, "校验码[:：]?\s*(\d{15,25})")
    If m.Success Then
        FCheck = m.Groups(1).Value : conf("check_code") = 0.9 : ev("check_code") = m.Value
    End If

    ' --- 购销双方：整行视图优先，按行视图兜底 ---
    Dim GetParty As Func(Of String, Boolean, String()) =
        Function(who As String, isBuyer As Boolean)
            Dim aliases As String() =
                If(isBuyer, {"购买方", "购货方", "付款方"}, {"销售方", "销货方", "收款方"})
            Dim nm As String = ""
            Dim tx As String = ""
            For Each al As String In aliases
                If nm <> "" Then Exit For
                Dim mm As Match = Regex.Match(o, al & "(?:信息)?名称[:：]?(.{2,50}?)" & NameStop)
                If mm.Success Then nm = mm.Groups(1).Value
            Next
            If nm = "" Then
                For Each al As String In aliases
                    Dim mm As Match = Regex.Match(o, al & "(?:信息)?[:：]?名称[:：]?(.{2,50}?)" & NameStop)
                    If mm.Success Then
                        nm = mm.Groups(1).Value
                        Exit For
                    End If
                Next
            End If
            If nm = "" Then
                Dim mm As Match = Regex.Match(o, who & "(?:信息)?.{0,120}?(?:统一社会信用代码|纳税人识别号|税号)[/]?[:：]?([0-9A-Z]{15,20})")
                If mm.Success Then tx = mm.Groups(1).Value
            End If
            ' 名称清洗：去掉前缀符号与"（小写）"这类残留
            nm = Regex.Replace(nm, "[:：\s]+$", "")
            nm = Regex.Split(nm, "统一社会信用代码|纳税人识别号|销售方|购买方|销货方|购货方|项目名称|地址|开户行|电话|账号|备注|开票人")(0)
            nm = nm.Trim(" "c, ":"c, "："c, ","c, "，"c, "、"c)
            Return New String() {nm, tx}
        End Function

    Dim bp As String() = GetParty("购买方", True)
    FBuyer = bp(0) : FBuyerTax = bp(1)
    Dim sp As String() = GetParty("销售方", False)
    FSeller = sp(0) : FSellerTax = sp(1)
    If FBuyer <> "" Then conf("buyer_name") = 0.85
    If FBuyerTax <> "" Then conf("buyer_tax") = 0.88
    If FSeller <> "" Then conf("seller_name") = 0.85
    If FSellerTax <> "" Then conf("seller_tax") = 0.88

    ' --- 价税合计 ---
    m = Regex.Match(o, "价税合计[^¥￥]{0,60}?[¥￥]?(" & MoneyPattern & ")")
    If Not m.Success Then m = Regex.Match(o, "[(（]小写[)）][¥￥]?(" & MoneyPattern & ")")
    If Not m.Success Then m = Regex.Match(o, "价税合计[^0-9]{0,60}(" & MoneyPattern & ")")
    If m.Success Then
        FTotal = ToMoney(m.Groups(1).Value)
        If Not Double.IsNaN(FTotal) Then conf("total") = 0.92
        ev("total") = m.Value
    End If

    ' --- 金额 / 税额 ---
    m = Regex.Match(o, "合计[¥￥](" & MoneyPattern & ")[¥￥](" & MoneyPattern & ")")
    If m.Success Then
        FAmount = ToMoney(m.Groups(1).Value)
        FTax = ToMoney(m.Groups(2).Value)
        conf("amount") = 0.9 : conf("tax") = 0.9
        ev("amount") = m.Value : ev("tax") = m.Value
    Else
        m = Regex.Match(o, "合计金额[¥￥]?(" & MoneyPattern & ")")
        If m.Success Then
            FAmount = ToMoney(m.Groups(1).Value)
            If Not Double.IsNaN(FAmount) Then conf("amount") = 0.85
        End If
        m = Regex.Match(o, "(?:合计税额|税额合计)[¥￥]?(" & MoneyPattern & ")")
        If m.Success Then
            FTax = ToMoney(m.Groups(1).Value)
            If Not Double.IsNaN(FTax) Then conf("tax") = 0.85
        End If
    End If
    ' 用价税合计反推缺失项 —— 置信度必须压低：这是算出来的，不是读出来的
    If Not Double.IsNaN(FTotal) AndAlso Not Double.IsNaN(FTax) AndAlso Double.IsNaN(FAmount) Then
        FAmount = Math.Round(FTotal - FTax, 2)
        conf("amount") = 0.6 : ev("amount") = "由 价税合计-税额 推导"
    End If
    If Not Double.IsNaN(FTotal) AndAlso Not Double.IsNaN(FAmount) AndAlso Double.IsNaN(FTax) Then
        FTax = Math.Round(FTotal - FAmount, 2)
        conf("tax") = 0.6 : ev("tax") = "由 价税合计-金额 推导"
    End If

    ' --- 税率：解决"各列粘成 11000.001000.006%"的抽取难题 ---
    Dim rateList As New List(Of String)()
    Dim reRate As New Regex("(?<![\d.])(\d[\d.]*)%")
    Dim reMoneyPrefix As New Regex("^\d[\d,]*\.\d{2}")
    For Each rm As Match In reRate.Matches(o)
        Dim tok As String = rm.Groups(1).Value
        ' 从左往右逐段剥掉"整数.两位小数"（金额固定两位小数），剩下就是税率
        Dim guard As Integer = 0
        While guard < 20
            guard += 1
            Dim mp As Match = reMoneyPrefix.Match(tok)
            If Not mp.Success Then Exit While
            Dim nxt As String = tok.Substring(mp.Length)
            If nxt = "" Then Exit While
            tok = nxt
        End While
        tok = tok.Trim("."c)
        If tok = "" OrElse tok.Length > 5 Then Continue For
        Dim f As Double
        If Not Double.TryParse(tok, f) Then Continue For
        If f < 0 OrElse f > 100 Then Continue For
        Dim s As String = f.ToString("g") & "%"
        If Not rateList.Contains(s) Then rateList.Add(s)
    Next
    If rateList.Count = 0 Then
        If o.Contains("免税") Then rateList.Add("免税")
        If o.Contains("不征税") Then rateList.Add("不征税")
    End If
    FRate = String.Join("/", rateList.ToArray())
    If FRate <> "" Then conf("tax_rate") = 0.85

    ' --- 项目名称（*类别*名称 形式）---
    Dim items As New List(Of String)()
    For Each ln As String In seg.Split(ChrW(10))
        For Each im As Match In Regex.Matches(ln, "\*[^*]{1,20}\*[^*0-9]{0,30}")
            Dim it As String = im.Value.Trim()
            If it.EndsWith("*") Then it = it.Substring(0, it.Length - 1)
            If it <> "" AndAlso Not items.Contains(it) Then items.Add(it)
        Next
    Next
    If items.Count > 0 Then
        FItem = String.Join("; ", items.ToArray())
        conf("item_name") = 0.9
    End If

    ' --- 开票人 ---
    m = Regex.Match(o, "开票人[:：]?([\u4e00-\u9fa5A-Za-z]{2,8})")
    If m.Success Then
        FDrawer = m.Groups(1).Value : conf("drawer") = 0.85
    End If

    ' ======================= 5. 校验（规则编号与 Python 版一致）=======================
    Dim issues As New List(Of String())()   ' {级别, 字段, 说明, 建议}
    Dim AddIssue As Action(Of String, String, String, String) =
        Sub(lv As String, fld As String, msg As String, sug As String)
            issues.Add(New String() {lv, fld, msg, sug})
        End Sub

    ' R01 必填
    If FNo = "" Then AddIssue("严重", "发票号码", "未识别到「发票号码」", "确认 PDF 是否完整原件，必要时人工补录")
    If FDate = "" Then AddIssue("严重", "开票日期", "未识别到「开票日期」", "检查票据是否清晰")
    If Double.IsNaN(FTotal) Then AddIssue("严重", "价税合计", "未识别到「价税合计」", "核对发票「价税合计(小写)」字样")
    If FSeller = "" Then AddIssue("严重", "销售方名称", "未识别到「销售方名称」", "人工补录")
    If FBuyer = "" Then AddIssue("严重", "购买方名称", "未识别到「购买方名称」", "个人抬头可能为空，请确认")

    ' R02/R03 号码
    If FNo <> "" Then
        If Not Regex.IsMatch(FNo, "^\d+$") Then
            AddIssue("严重", "发票号码", "发票号码含非数字字符：" & FNo, "人工核对")
        ElseIf FType.Contains("数电") OrElse FNo.Length = 20 Then
            If FNo.Length <> 20 Then AddIssue("警告", "发票号码", "数电发票号码应为 20 位，当前 " & FNo.Length & " 位", "核对是否被截断")
        ElseIf FNo.Length <> 8 AndAlso FNo.Length <> 9 AndAlso FNo.Length <> 10 AndAlso FNo.Length <> 12 Then
            AddIssue("警告", "发票号码", "发票号码为 " & FNo.Length & " 位，与常见版式不符", "人工核对")
        End If
    End If

    ' R04 日期
    If FDate <> "" Then
        Dim dt As DateTime
        If Not DateTime.TryParse(FDate, dt) Then
            AddIssue("严重", "开票日期", "开票日期无法解析：" & FDate, "人工核对")
        ElseIf dt.Date > in_Today.Date Then
            AddIssue("严重", "开票日期", "开票日期 " & FDate & " 晚于当前日期", "核对是否识别错误")
        ElseIf dt.Year < 2000 Then
            AddIssue("警告", "开票日期", "开票日期 " & FDate & " 过早，疑似识别错误", "人工核对")
        End If
    End If

    ' R05 金额勾稽
    Dim balOK As Boolean = False
    If Not Double.IsNaN(FAmount) AndAlso Not Double.IsNaN(FTax) AndAlso Not Double.IsNaN(FTotal) Then
        Dim diff As Double = Math.Round(FAmount + FTax - FTotal, 2)
        If Math.Abs(diff) > in_BalanceTol Then
            AddIssue("严重", "金额勾稽", "金额(不含税)" & FAmount.ToString("N2") & " + 税额" & FTax.ToString("N2") &
                     " 与价税合计" & FTotal.ToString("N2") & " 相差 " & diff.ToString("N2"),
                     "核对是否漏识别明细行或合计数被遮挡")
        Else
            balOK = True
        End If
    ElseIf Not Double.IsNaN(FTotal) Then
        AddIssue("警告", "金额勾稽", "金额(不含税)或税额缺失，无法完成勾稽校验", "人工补充后核对")
    End If

    ' R06 红字 / 零金额
    If Not Double.IsNaN(FTotal) Then
        If FTotal < 0 Then
            AddIssue("提示", "价税合计", "价税合计为负数（" & FTotal.ToString("N2") & "），属于红字发票", "注意冲减方向")
        ElseIf FTotal = 0 Then
            AddIssue("警告", "价税合计", "价税合计为 0.00", "核对是否为零金额发票或识别异常")
        End If
    End If

    ' R07/R08 税号
    Dim taxCheckPass As Boolean = False
    For Each pair As String() In {New String() {FBuyerTax, "购买方税号"}, New String() {FSellerTax, "销售方税号"}}
        Dim code As String = If(pair(0) Is Nothing, "", pair(0).Trim().ToUpper())
        Dim label As String = pair(1)
        If code = "" Then Continue For
        If Not Regex.IsMatch(code, "^[0-9A-Z]{15,20}$") Then
            AddIssue("严重", label, "税号格式非法：" & code, "人工核对")
        Else
            If code.Length <> 15 AndAlso code.Length <> 17 AndAlso code.Length <> 18 AndAlso code.Length <> 20 Then
                AddIssue("警告", label, "税号长度 " & code.Length & " 位较为罕见", "人工核对")
            End If
            Dim r As Integer = UsccCheck(code)
            If r = 0 Then
                AddIssue("警告", label, "统一社会信用代码校验位不通过：" & code,
                         "疑似识别错误（0/O、1/I 混淆），建议人工核对")
            ElseIf r = 1 Then
                taxCheckPass = True
            End If
        End If
    Next

    ' R09 购销同名
    If FBuyer <> "" AndAlso FSeller <> "" AndAlso FBuyer = FSeller Then
        AddIssue("警告", "购销双方", "购买方与销售方名称完全相同", "核对是否填错，或属于自开自抵情形")
    End If

    ' R10 税率集合
    For Each r As String In rateList
        If Not {"0%", "1%", "1.5%", "3%", "5%", "6%", "9%", "10%", "11%", "13%", "16%", "17%", "免税", "不征税", "***"}.Contains(r) Then
            AddIssue("提示", "税率", "税率 " & r & " 不在常用税率集合内", "确认是否为特殊征收率/减免政策")
            Exit For
        End If
    Next

    ' R11/R12 发票代码 / 销售方税号
    If Not FType.Contains("数电") AndAlso FCode = "" AndAlso FNo <> "" Then
        AddIssue("警告", "发票代码", "未识别到发票代码（非数电发票通常应有 12 位发票代码）", "人工核对")
    End If
    If FSellerTax = "" Then AddIssue("警告", "销售方税号", "销售方纳税人识别号缺失", "人工补录")

    ' R13 税额 与 金额×税率 勾稽（新增）
    Dim rateOK As Boolean = False
    If Not Double.IsNaN(FAmount) AndAlso Not Double.IsNaN(FTax) AndAlso FAmount <> 0 AndAlso rateList.Count = 1 Then
        Dim rp As String = rateList(0)
        If rp.EndsWith("%") Then
            Dim pct As Double
            If Double.TryParse(rp.TrimEnd("%"c), pct) Then
                Dim expect As Double = Math.Abs(FAmount) * pct / 100.0
                Dim tol As Double = Math.Max(in_RateFloor, Math.Abs(FAmount) * in_RateTol)
                If Math.Abs(Math.Abs(FTax) - expect) > tol Then
                    AddIssue("严重", "税率×金额",
                             "按税率 " & rp & " 应有税额 " & expect.ToString("N2") & "，实际税额 " & FTax.ToString("N2"),
                             "税率或金额/税额识别有误，三者需自洽")
                Else
                    rateOK = True
                End If
            End If
        End If
    End If

    ' R14 数电号码结构（新增）
    If FNo.Length = 20 AndAlso Regex.IsMatch(FNo, "^\d+$") Then
        If Not ProvCodes.Contains(FNo.Substring(2, 2)) Then
            AddIssue("警告", "发票号码", "20 位号码第 3-4 位省级区域码 " & FNo.Substring(2, 2) & " 不在编码表内",
                     "疑似号码识别错误")
        End If
        If FDate.Length >= 4 AndAlso Regex.IsMatch(FDate.Substring(0, 4), "^\d{4}$") Then
            If FNo.Substring(0, 2) <> FDate.Substring(2, 2) Then
                AddIssue("警告", "发票号码", "号码前两位 " & FNo.Substring(0, 2) & " 与开票年份 " & FDate.Substring(0, 4) & " 不一致",
                         "号码或开票日期识别有误（数电号码前两位=年份后两位）")
            End If
        End If
    End If

    ' R15 发票代码结构（新增）
    If FCode.Length = 12 AndAlso Regex.IsMatch(FCode, "^\d+$") Then
        If FCode.Substring(0, 1) <> "0" Then
            AddIssue("警告", "发票代码", "12 位发票代码首位应为 0，当前为 " & FCode.Substring(0, 1), "疑似代码识别错误")
        End If
        If FCode.Substring(10, 2) <> "11" AndAlso FCode.Substring(10, 2) <> "13" Then
            AddIssue("提示", "发票代码", "发票代码第 11-12 位票种码为 " & FCode.Substring(10, 2) & "（常见 11/13）",
                     "确认票种码是否为该地区特殊码制")
        End If
    End If

    ' ======================= 6. 置信度闭环 =======================
    ' 校验通过 → 互相印证的字段加分；冲突 → 相关字段打折
    If balOK Then
        For Each k As String In {"amount", "tax", "total"}
            If conf.ContainsKey(k) Then conf(k) = Math.Min(0.99, conf(k) + 0.05)
        Next
    End If
    If taxCheckPass Then
        For Each k As String In {"buyer_tax", "seller_tax"}
            If conf.ContainsKey(k) Then conf(k) = Math.Min(0.99, conf(k) + 0.07)
        Next
    End If
    If rateOK AndAlso conf.ContainsKey("tax_rate") Then conf("tax_rate") = Math.Min(0.99, conf("tax_rate") + 0.03)
    If FNo <> "" AndAlso FDate <> "" AndAlso conf.ContainsKey("invoice_no") Then
        conf("invoice_no") = Math.Min(0.99, conf("invoice_no") + 0.03)
    End If
    For Each it As String() In issues
        If it(0) = "严重" Then
            Dim fld As String = it(1)
            For Each k As String In conf.Keys.ToList()
                If k.Contains(fld) OrElse fld.Contains(k) Then
                    conf(k) = Math.Round(conf(k) * 0.6, 3)
                    Exit For
                End If
            Next
        End If
    Next

    Dim weights As New Dictionary(Of String, Double) From {
        {"invoice_no", 3.0}, {"issue_date", 2.0}, {"total", 3.0},
        {"amount", 2.0}, {"tax", 2.0}, {"seller_name", 2.0}, {"buyer_name", 1.5},
        {"seller_tax", 1.5}, {"buyer_tax", 1.0}, {"invoice_type", 0.5},
        {"invoice_code", 1.0}, {"tax_rate", 1.0}, {"item_name", 0.5}, {"drawer", 0.3}}
    Dim num As Double = 0, den As Double = 0
    For Each kv As KeyValuePair(Of String, Double) In weights
        Dim v As Double = 0
        If conf.ContainsKey(kv.Key) Then v = conf(kv.Key)
        num += v * kv.Value
        den += kv.Value
    Next
    Dim overall As Double = If(den > 0, Math.Round(num / den, 3), 0)

    ' 状态判定：严重 > 待复核 > 提示 > 正常
    Dim Status As String = "正常"
    Dim hasError As Boolean = False, hasWarn As Boolean = False
    For Each it As String() In issues
        If it(0) = "严重" Then hasError = True
        If it(0) = "警告" OrElse it(0) = "提示" Then hasWarn = True
    Next
    Dim lowCrit As Boolean = False
    For Each ck As String In {"invoice_no", "issue_date", "total"}
        Dim cv As Double = 0
        If conf.ContainsKey(ck) Then cv = conf(ck)
        Dim emptyField As Boolean = (ck = "invoice_no" AndAlso FNo = "") OrElse
                                    (ck = "issue_date" AndAlso FDate = "") OrElse
                                    (ck = "total" AndAlso Double.IsNaN(FTotal))
        If emptyField OrElse cv < in_MinCritConf Then lowCrit = True
    Next
    If hasError Then
        Status = "异常"
    ElseIf overall < in_MinConf OrElse lowCrit Then
        Status = "待复核"
    ElseIf hasWarn Then
        Status = "存在提示"
    End If

    ' ======================= 7. 输出明细行 =======================
    Dim displayName As String = in_FileName
    If Segs.Count > 1 Then displayName = in_FileName & " [第" & Idx & "张/共" & Segs.Count & "张]"

    Dim row As DataRow = out_Rows.NewRow()
    row("序号") = out_Rows.Rows.Count + 1
    row("源文件名") = displayName
    row("发票类型") = FType
    row("发票代码") = FCode
    row("发票号码") = FNo
    row("开票日期") = FDate
    row("购买方名称") = FBuyer
    row("购买方税号") = FBuyerTax
    row("销售方名称") = FSeller
    row("销售方税号") = FSellerTax
    If Double.IsNaN(FAmount) Then row("金额(不含税)") = DBNull.Value Else row("金额(不含税)") = FAmount
    If Double.IsNaN(FTax) Then row("税额") = DBNull.Value Else row("税额") = FTax
    If Double.IsNaN(FTotal) Then row("价税合计") = DBNull.Value Else row("价税合计") = FTotal
    row("税率/征收率") = FRate
    row("主要项目/货物名称") = FItem
    row("校验码") = FCheck
    row("开票人") = FDrawer
    row("识别置信度") = overall.ToString("P1")
    row("提取方式") = "pdf-text"
    row("处理状态") = Status
    Dim summary As New List(Of String)()
    For Each it As String() In issues
        summary.Add("[" & it(0) & "] " & it(1) & ": " & it(2))
    Next
    row("异常/提示说明") = String.Join("； ", summary.ToArray())
    row("处理时间") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    row("文件指纹") = in_FileName
    out_Rows.Rows.Add(row)

    ' ======================= 8. 输出异常记录 =======================
    For Each it As String() In issues
        Dim ir As DataRow = out_Issues.NewRow()
        ir("序号") = out_Issues.Rows.Count + 1
        ir("源文件名") = in_FileName
        ir("发票号码") = If(FNo = "", "—", FNo)
        ir("问题级别") = it(0)
        ir("涉及字段") = it(1)
        ir("问题描述") = it(2)
        ir("建议处理") = it(3)
        ir("发生时间") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        out_Issues.Rows.Add(ir)
    Next
Next

' 供审计用：把耗时写到日志（InvokeCode 里可以直接 Console.WriteLine，UiPath 会收进 Output）
Console.WriteLine("[ParseInvoice] 文件名=" & in_FileName & " 段数=" & Segs.Count.ToString() &
                  " 产出行数=" & out_Rows.Rows.Count.ToString() &
                  " 耗时=" & (DateTime.Now - t0).TotalMilliseconds.ToString("F0") & "ms")
