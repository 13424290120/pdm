VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "通用打印"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'这是一个将表格数据导出为EXCEL文件的类，使用的机器上必须安装有EXCEL。
'根据自己的需要可对它进行修改完善。比如，不显示EXCEL界面，直接将EXCEL文件保存到指定位置……
'设计时，必须在引用中选中“Microsoft Excel 11.0 Object library” --此处EXCEL为2003版，其他版本可能不是 11.0
'Download by http://www.codefans.net
Option Explicit
Dim BBGrid As Object
Dim btV, btbtV, ywzjV, ywbtV As String, zzdxV$, bbkdV#, hsV%, dymsV%, dyfxV%, dylxV$
Dim fontsizeV%, tt%

Public Function NotnullN(tt) As Double
If IsNull(tt) Or tt = "" Then
    NotnullN = 0
Else
    NotnullN = tt
End If
End Function

Public Function NotnullC(tt) As String
If IsNull(tt) Then
    NotnullC = ""
Else
    NotnullC = Trim(tt)
End If
End Function

Public Sub Excel打印()
Dim l%, J%, i#, hhV$, k%, m%, N%, hhV1$

On Error Resume Next
Dim ex1 As Excel.Application
Set ex1 = CreateObject("Excel.Application")
'判断机器上是否安装了EXCEL
If ex1 Is Nothing Then
        MsgBox "您的电脑还没有安装Excel，无法将数据导出为EXCEL文件！", vbCritical, "注意"
        ex1.Quit
        Set ex1 = Nothing
End If

ex1.Visible = True
ex1.Workbooks.Add

With BBGrid
hhV1 = ""
 m = 0
 N = 0
If NotnullN(hsV) = 0 Then
        hsV = 0
End If
    For i = 0 To .Cols - 1
        If .ColWidth(i) <> 0 Then
            If NotnullN(bbkdV) = 0 Then
                If (m Mod 26) = 0 And m <> 0 Then
                     hhV1 = Chr(Asc("A") + Int(m / 26) - 1)
                     N = 0
                End If
                If m < 26 Then
                    ex1.Range(Chr(Asc("A") + m) & "3").ColumnWidth = .ColWidth(i) / 35
                Else
                    ex1.Range(hhV1 & Chr(Asc("A") + N) & "3").ColumnWidth = .ColWidth(i) / 35
                    N = N + 1
                End If
            Else
                If (m Mod 26) = 0 And m <> 0 Then
                     hhV1 = Chr(Asc("A") + Int(m / 26) - 1)
                     N = 0
                End If
                If m < 26 Then
                    ex1.Range(Chr(Asc("A") + m) & "3").ColumnWidth = .ColWidth(i) / bbkdV
                Else
                    ex1.Range(hhV1 & Chr(Asc("A") + N) & "3").ColumnWidth = .ColWidth(i) / bbkdV
                    N = N + 1
                End If
            End If
            m = m + 1
        End If
    Next i
    'ex1.Range("a2").ColumnWidth = 35 '定义宽度
    ex1.Range("a1") = btV
    If hhV1 = "" Then
        ex1.Range("a1:" & Chr(Asc("A") + m - 1) & "1").MergeCells = True
    Else
        ex1.Range("a1:" & hhV1 & Chr(Asc("A") + N - 1) & "1").MergeCells = True
    End If
    ex1.Range("a2") = btbtV
    If hhV1 = "" Then
        ex1.Range("a2:" & Chr(Asc("A") + m - 1) & "2").MergeCells = True
    Else
        ex1.Range("a2:" & hhV1 & Chr(Asc("A") + N - 1) & "2").MergeCells = True
    End If
    hhV = ""
    J = 0
    m = 0
    N = 0
    For i = 0 To .Cols - 1
        If .ColWidth(i) <> 0 Then
                If (m Mod 26) = 0 And m <> 0 Then
                     hhV1 = Chr(Asc("A") + Int(m / 26) - 1)
                     N = 0
                End If
                If m < 26 Then
                    ex1.Range(Chr(Asc("A") + m) & "3") = .TextMatrix(hsV, i)
                Else
                    ex1.Range(hhV1 & Chr(Asc("A") + N) & "3") = .TextMatrix(hsV, i)
                    N = N + 1
                End If
            m = m + 1
        End If
    Next i
    
    For i = hsV To .Rows - 1
        
        Select Case Trim(dylxV)
            Case "", "全部"
            Case "未接收"
                If Trim(.TextMatrix(i, 19)) <> "" And i > 0 Then
                    GoTo cw1
                End If
            Case "已接收"
                If Trim(.TextMatrix(i, 19)) = "" And i > 0 Then
                    GoTo cw1
                End If
        End Select
        J = J + 1
        hhV = ""
        l = 0
        For k = 0 To .Cols - 1
            If .ColWidth(k) > 0 Then
                If l Mod 26 = 0 And l > 0 Then
                    If hhV = "" Then
                        hhV = "A"
                    Else
                        hhV = Chr(Asc(hhV) + l)
                    End If
                End If
                If k > 0 Then
                    If InStr(1, .TextMatrix(i, k), ".") Then
                        ex1.Range(hhV & Chr(Asc("A") + l Mod 26) & Trim(Str(J + 2))) = .TextMatrix(i, k)
'                        ex1.Range(hhV & Chr(Asc("A") + l Mod 26) & Trim(Str(j + 2))).NumberFormat = "#0.00"
                    Else
                        If Len(.TextMatrix(i, k)) > 7 Then
                            ex1.Range(hhV & Chr(Asc("A") + l Mod 26) & Trim(Str(J + 2))) = "'" & .TextMatrix(i, k)
                        Else
                            ex1.Range(hhV & Chr(Asc("A") + l Mod 26) & Trim(Str(J + 2))) = .TextMatrix(i, k)
                        End If
                    End If
                Else
                    ex1.Range(hhV & Chr(Asc("A") + l Mod 26) & Trim(Str(J + 2))) = .TextMatrix(i, k)
                End If
                l = l + 1
            End If
        Next k
cw1:
    Next i
End With
ex1.Range("a" & Trim(Str(i + 3))) = ywzjV
If hhV1 = "" Then
    ex1.Range("a" & Trim(Str(i + 3)) & ":" & Chr(Asc("A") + m - 1) & Trim(Str(i + 3))).MergeCells = True
Else
    ex1.Range("a" & Trim(Str(i + 3)) & ":" & hhV1 & Chr(Asc("A") + N - 1) & Trim(Str(i + 3))).MergeCells = True
End If
With ex1.Workbooks(1).Styles.Add(Name:="bookman top border") '定义边框风格1
    .Borders(xlTop).LineStyle = xlDash
    .Borders(xlEdgeLeft).LineStyle = xlDash
    .Borders(xlEdgeRight).LineStyle = xlDash
    .Borders(xlEdgeBottom).LineStyle = xlDash
    If NotnullN(fontsizeV) = 0 Then
        .Font.Size = 9
    Else
        .Font.Size = fontsizeV
    End If
    .Font.Name = "宋体"
End With
With ex1.Workbooks(1).Styles.Add(Name:="bookman top border1") '定义边框风格2
    .Borders(xlEdgeLeft).LineStyle = xlDash
    If NotnullN(fontsizeV) = 0 Then
        .Font.Size = 9
    Else
        .Font.Size = fontsizeV
    End If
    .Font.Name = "宋体"
End With

If NotnullC(zzdxV) = "" Then
    ex1.Worksheets(1).PageSetup.PaperSize = xlPaperA4
Else
    ex1.Worksheets(1).PageSetup.PaperSize = zzdxV
    
End If

If hhV1 = "" Then
    ex1.Worksheets(1).Range("A3:" & Chr(Asc("A") + m - 1) & Trim(Str(i + 2))).Style = "bookman top border" '实现边框风格
Else
     ex1.Worksheets(1).Range("A3:" & hhV1 & Chr(Asc("A") + N - 1) & Trim(Str(i + 2))).Style = "bookman top border" '实现边框风格
End If
If hhV1 = "" Then
    ex1.Worksheets(1).Range(Chr(Asc("A") + m) & "3:" & Chr(Asc("A") + m) & Trim(Str(i + 2))).Style = "bookman top border1"   '实现边框风格1
Else
     ex1.Worksheets(1).Range(hhV1 & Chr(Asc("A") + N) & "3:" & hhV1 & Chr(Asc("A") + N) & Trim(Str(i + 2))).Style = "bookman top border1"   '实现边框风格1
End If
If hhV1 = "" Then
    With ex1.Worksheets(1).Range("A3:" & Chr(Asc("A") + m) & "3")      '定义单元格内居中风格
        .HorizontalAlignment = xlHAlignDistributed
        .AddIndent = True
    End With
Else
    With ex1.Worksheets(1).Range("A3:" & hhV1 & Chr(Asc("A") + N) & "3")      '定义单元格内居中风格
        .HorizontalAlignment = xlHAlignDistributed
        .AddIndent = True
    End With
End If
With ex1.Worksheets(1).Range("A1:b1")      '定义单元格内居中风格
    .HorizontalAlignment = 3
    .AddIndent = True
    .Font.Name = "隶书"
    .Font.Size = 18
End With
With ex1.Worksheets(1).Range("A2:b2")      '定义单元格内居中风格
    .HorizontalAlignment = 3
    .AddIndent = True
    .Font.Name = "宋体"
    .Font.Size = 10
End With
With ex1.Worksheets(1).Range("A" & Trim(Str(i + 3)) & ":b" & Trim(Str(i + 3)))      '定义单元格内居中风格
    .HorizontalAlignment = 3
    .AddIndent = True
    .Font.Name = "隶书"
    .Font.Size = 10
End With

If hhV1 = "" Then
    ex1.Columns("a:" & Chr(Asc("A") + m - 1)).AutoFit '自动调整列宽
Else
     ex1.Columns("a:" & hhV1 & Chr(Asc("A") + N - 1)).AutoFit
End If

'ex1.Worksheets(1).Range("a1:" & Chr(Asc("A") + m) & Trim(Str(i + 2))).NumberFormat = "0.00"  '定义数值类型格式
ex1.ActiveSheet.PageSetup.PrintTitleRows = ex1.ActiveSheet.Rows("1:3").Address '定义表头所用行
If hhV1 = "" Then
    ex1.ActiveSheet.PageSetup.PrintTitleColumns = ex1.ActiveSheet.Columns("A:" & Chr(Asc("A") + m)).Address '定义标题所用列
Else
    ex1.ActiveSheet.PageSetup.PrintTitleColumns = ex1.ActiveSheet.Columns("A:" & hhV1 & Chr(Asc("A") + N)).Address '定义标题所用列
End If

ex1.Worksheets(1).PageSetup.CenterFooter = "&L 第 &P / &N 页  " & ywbtV    '定义页脚
If NotnullN(dymsV) = 0 Then
    dymsV = 0
End If

'Select Case dymsV
'    Case 0
'    Case 1
'        ex1.ActiveSheet.PrintPreview
'    Case 2
'        ex1.ActiveSheet.PrintOut
'End Select

'ex1.Worksheets(1).PageSetup.Orientation = dyfxV

End Sub

Public Property Let 数据表(s As Object)
    Set BBGrid = s
End Property
Public Property Let 数据表初始行(s As Integer)
     hsV = s
End Property

Public Property Let 表头标题(s As String)
     btV = s
End Property

Public Property Let 表头下标题(s As String)
     btbtV = s
End Property

Public Property Let 页尾注解(s As String)
     ywzjV = s
End Property
Public Property Let 页尾标题(s As String)
     ywbtV = s
End Property
Public Property Let 字体大小(s As Integer)
     fontsizeV = s
End Property

Public Property Let 纸张大小(s As String)
     zzdxV = s
End Property

Public Property Let 单元格大小参数(s As Double)
     bbkdV = s
End Property

Public Property Let 打印显示模式(s As Integer)
    dymsV = s
End Property
Public Property Let 打印方向(s As Integer)
    dyfxV = s
End Property
Public Property Let 打印类型(s As String)
    dylxV = s
End Property
'使用示例
'Dim pp As 通用打印
'Set pp = New 通用打印
'pp.数据表 = myGrid
'pp.打印显示模式 = 1
'pp.表头标题 = Me.Caption & "表"
'pp.表头下标题 = "商品代码：" & spmc.Text & Space(30) & "单号：" & dhT.Text
'pp.页尾注解 = "合计   实盘总金额:" & hjxpjeL.Caption & Space(5) & "帐本总金额:" & hjzbjeL.Caption & Space(5) & "盈亏总金额:" & hjykjeL.Caption
'pp.页尾标题 = "&L 第 &P / &N 页  制单人： " & zdrT.Text & Space(20) & "复合人：" & fhrT.Text        '定义页脚
'pp.Excel打印

