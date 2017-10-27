Attribute VB_Name = "GeneralFunc"
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'判断一个文件是否已经打开
Public Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
' 以二进制模式打开指定的文件
' 返回值Long，如执行成功，返回打开文件的句柄。HFILE_ERROR表示出错。会设置GetLastError
'lpPathName -----  String，欲打开文件的名字
'iReadWrite -----  Long，访问模式和共享模式常数的一个组合，如下所示：
'1)  访问模式
'Read
'打开文件，读取其中的内容
'READ_WRITE
'打开文件，对其进行读写
'WRITE
'打开文件，在其中写入内容
'2)  共享模式 (参考OpenFile函数的标志常数表)
'OF_SHARE_COMPAT， OF_SHARE_DENY_NONE， OF_SHARE_DENY_READ，OF_SHARE_DENY_WRITE， OF_SHARE_EXCLUSIVE
Public Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
'关闭指定的文件，请参考CloseHandle函数，了解进一步的情况


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'寻找窗口列表中第一个符合指定条件的顶级窗口（在vb里使用：FindWindow最常见的一个用途是获得ThunderRTMain类的隐藏窗口的句柄；该类是所有运行中vb执行程序的一部分。获得句柄后，可用api函数GetWindowText取得这个窗口的名称；该名也是应用程序的标题）
'返回值Long，找到窗口的句柄。如未找到相符窗口，则返回零。会设置GetLastError
'lpClassName ----  String，指向包含了窗口类名的空中止（C语言）字串的指针；或设为零，表示接收任何类
'lpWindowName ---  String，指向包含了窗口文本（或标签）的空中止（C语言）字串的指针；或设为零，表示接收任何窗口标题
'很少要求同时按类与窗口名搜索。为向自己不准备参数传递一个零，最简便的办法是传递vbNullString常数
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'激活指定的窗口
'返回值Long 前一个活动窗口的句柄
'hwnd -----------Long，待激活窗口的句柄


Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'这个 OpenProcess 函数打开一个已存在的进程对象。
'如成功，返回值为指定进程的打开句柄  如失败，返回值为空，可调用GetLastError获得错误代码。
'dwDesiredAccess是访问进程的权限 bInheritHandle是句柄是否继承进程属性 dwProcessId是进程ＩＤ
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'关闭一个内核对象。其中包括文件、文件映射、进程、线程、安全和同步对象等。
'返回值非零表示成功，零表示失败。会设置GetLastError
'hObject Long，欲关闭的一个对象的句柄
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'WaitForSingleObject(事件句柄, 超时时间)
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'对一个Excel文件打开/关闭的操作
Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private ObjOldWidth As Long  '保存窗体的原始宽度
Private ObjOldHeight As Long '保存窗体的原始高度
Private ObjOldFont As Single '保存窗体的原始字体比

' 判断某文件是否在打开使用中
Function IsFileAlreadyOpen(Filename As String) As Boolean
    Dim hFile     As Long
    Dim lastErr     As Long
    hFile = -1                                 '初始化文件句柄.
    lastErr = 0
    hFile = lOpen(Filename, &H10)
    
    If hFile = -1 Then                         '文件是否能正确打开并可共享
        lastErr = Err.LastDllError
    Else
        lClose (hFile)                     'hFile文件句柄.
    End If
    IsFileAlreadyOpen = (hFile = -1) And (lastErr = 32)
End Function
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'对一个Excel文件关闭的操作
Sub ModulesCloseExcel(strFileName As String)
    
    If IsFileAlreadyOpen(strFileName) Then
        xlBook.Close (True)                      '关闭EXCEL工作薄
        xlApp.Quit                               '关闭EXCEL
        Set xlApp = Nothing                       '释放EXCEL对象
    End If
    
End Sub

'对一个Excel文件打开的操作
Sub ModulesOpenExcel(strFileName As String)               '打开EXCEL过程
    
    If IsFileAlreadyOpen(strFileName) Then
        MsgBox "Excel File is already Openned, Please Close it firstly", vbInformation, "System Info."
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")   '创建EXCEL应用类对象
    xlApp.Visible = True                            '设置EXCEL应用类对象可见
    Set xlBook = xlApp.Workbooks.Open(strFileName)  '打开EXCEL工作薄
    Set xlSheet = xlBook.Worksheets(1)              '打开EXCEL工作表1
    xlSheet.Activate                                '激活EXCEL工作表
    
End Sub
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


'MSHFlexGrid控件导出到Excel
Function ExportFlexDataToExcel(flex As MSFlexGrid, g_CommonDialog As CommonDialog, Optional strHead As String = "")
    
    On Error GoTo ErrHandler
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim Rows As Integer, Cols As Integer
    Dim eRow As Integer, eCol As Integer, fCol As Integer, fRow As Integer
    Dim New_Col As Boolean
    Dim New_Column  As Boolean
    
    g_CommonDialog.CancelError = True
    
    ' 设置标志
    g_CommonDialog.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    g_CommonDialog.Filter = "All Files (*.*)|*.*|Excel Files" & "(*.xls)|*.xls"
    ' 指定缺省的过滤器
    g_CommonDialog.FilterIndex = 2
    ' 显示“打开”对话框  通过使用 CommonDialog 控件的 ShowOpen 和 ShowSave 方法可显示“打开”和“另存为”对话框
    g_CommonDialog.ShowSave
    
    If flex.Rows <= 1 Then
        MsgBox "No Data！", vbInformation, "System Info"
        Exit Function
    End If
    
    MousePointer = vbHourglass   '读写时间较长，需要定义鼠标状态
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    xlApp.Visible = False
    
    With flex
        Rows = .Rows
        Cols = .Cols
        eRow = 1   'Excel行从1开始
        eCol = 1   'Excel列从1开始
        fRow = 0   'MsFlexGird行从0开始
        fCol = 0   'MsFlexGird列从0开始
        xlApp.Cells(1, Int(Cols / 2)).Value = strHead
        For eRow = 2 To Rows + 1
            If flex.RowHeight(fRow) = 0 Then GoTo nextrow           '如果行高为0,表示有隐藏的行,则不输出
            For eCol = 1 To Cols
                fCol = eCol - 1
                xlApp.Cells(eRow, eCol).NumberFormat = "@"             '设置每个单元格数字类型为文本型
                xlApp.Cells(eRow, eCol).Value = .TextMatrix(fRow, fCol)
            Next eCol
nextrow:
            fRow = fRow + 1
        Next eRow
    End With
    
    With xlApp
        .Rows(1).Font.Bold = True
        .Rows(1).Font.Size = 16
        .Cells.Select
        .Columns.AutoFit
        .Cells(1, 1).Select
        '        .Application.Visible = True
    End With
    
    xlBook.SaveAs (g_CommonDialog.Filename)
    xlApp.Application.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.Quit
    Set xlApp = Nothing  '"交还控制给Excel
    Set xlBook = Nothing
    flex.SetFocus
    MousePointer = vbDefault                  '恢复鼠标状态
    MsgBox "Data have been exported into Excel", vbInformation, "System Info"
    Exit Function
    
ErrHandler:
    ' 用户按了“取消”按钮
    If Err.Number <> 32755 Then
        MsgBox "Cancel Data Exporting", vbInformation, "System Info"
    End If
End Function





Function EvaluateExpr(ByVal expr As String) As Single
    '--------------------------------------------------------------------------
    '功能:           字符串表达式的计算函数
    '参数:
    '               [expr]...........................字符串表达式
    '返回值:
    '               [EvaluateExpr]...................计算后的值
    '--------------------------------------------------------------------------
    Const PREC_NONE = 11
    Const PREC_UNARY = 10             'Not actually used.
    Const PREC_POWER = 9
    Const PREC_TIMES = 8
    Const PREC_DIV = 7
    Const PREC_INT_DIV = 6
    Const PREC_MOD = 5
    Const PREC_PLUS = 4
    
    Dim is_unary     As Boolean
    Dim next_unary     As Boolean
    Dim parens     As Integer
    Dim Pos     As Integer
    Dim expr_len     As Integer
    Dim ch     As String
    Dim lexpr     As String
    Dim rexpr     As String
    Dim Value     As String
    Dim status     As Long
    Dim best_pos     As Integer
    Dim best_prec     As Integer
    
    '   删除首尾空格及有效性校验
    expr = Trim$(expr)
    expr_len = Len(expr)
    If expr_len = 0 Then Exit Function
    
    '   If we find + or - now, it is a unary operator.
    is_unary = True
    
    '   So far we have nothing.
    best_prec = PREC_NONE
    
    '   Find the operator with the lowest precedence.
    '   Look for places where there are no open parentheses.
    For Pos = 1 To expr_len
        '   Examine the next character.(检查下一个字符)
        ch = Mid$(expr, Pos, 1)
        
        '   Assume we will not find an operator. In that case the next operator will not be unary.
        next_unary = False
        
        If ch = "   " Then
            'Just skip spaces.
            next_unary = is_unary
        ElseIf ch = "(" Then
            'Increase the open parentheses count.
            parens = parens + 1
            
            '   An operator after "(" is unary.
            next_unary = True
        ElseIf ch = ")" Then
            '   Decrease the open parentheses count.
            parens = parens - 1
            
            '   An operator after ")" is not unary.
            next_unary = False
            
            '   If parens< 0, too many ')'s.
            If parens < 0 Then
                Err.Raise vbObjectError + 1001, "EvaluateExpr", "Too many ) in  '" & expr & "'"
            End If
        ElseIf parens = 0 Then
            '   See if this is an operator.
            If ch = "^" Or ch = "*" Or ch = "/" Or ch = "\" Or ch = "%" Or ch = "+" Or ch = "-" Then
                ' An  operator after an operator is unary.
                next_unary = True
                
                Select Case ch
                Case "^"
                    If best_prec >= PREC_POWER Then
                        best_prec = PREC_POWER
                        best_pos = Pos
                    End If
                Case "*", "/"
                    If best_prec >= PREC_TIMES Then
                        best_prec = PREC_TIMES
                        best_pos = Pos
                    End If
                    
                Case "\"
                    If best_prec >= PREC_INT_DIV Then
                        best_prec = PREC_INT_DIV
                        best_pos = Pos
                    End If
                    
                Case "%"
                    If best_prec >= PREC_MOD Then
                        best_prec = PREC_MOD
                        best_pos = Pos
                    End If
                    
                Case "+", "-"
                    '   Ignore   unary   operators
                    '   for   now.
                    If (Not is_unary) And _
                        best_prec >= PREC_PLUS _
                        Then
                        best_prec = PREC_PLUS
                        best_pos = Pos
                    End If
                End Select
            End If
        End If
        is_unary = next_unary
    Next Pos
    
    '   If   the   parentheses   count   is   not   zero,
    '   there's   a   ')'   missing.
    If parens <> 0 Then
        Err.Raise vbObjectError + 1002, _
        "EvaluateExpr", "Missing  )  in  '" & expr & "'"
    End If
    
    '   Hopefully   we   have   the   operator.
    '   best_prec是最高的运算符
    Dim dblTemp1 As Double, dblTemp2  As Double
    If best_prec < PREC_NONE Then
        lexpr = Left$(expr, best_pos - 1)
        rexpr = Right$(expr, expr_len - best_pos)
        Select Case Mid$(expr, best_pos, 1)
        Case "^"
            EvaluateExpr = EvaluateExpr(lexpr) ^ EvaluateExpr(rexpr)
        Case "*"
            EvaluateExpr = EvaluateExpr(lexpr) * EvaluateExpr(rexpr)
        Case "/"
            dblTemp1 = EvaluateExpr(rexpr)
            dblTemp2 = EvaluateExpr(lexpr)
            If dblTemp1 = 0 Then
                EvaluateExpr = 0
            Else
                EvaluateExpr = dblTemp2 / dblTemp1
            End If
        Case "\"
            EvaluateExpr = EvaluateExpr(lexpr) \ EvaluateExpr(rexpr)
        Case "%"
            EvaluateExpr = EvaluateExpr(lexpr) Mod EvaluateExpr(rexpr)
        Case "+"
            EvaluateExpr = EvaluateExpr(lexpr) + EvaluateExpr(rexpr)
        Case "-"
            EvaluateExpr = EvaluateExpr(lexpr) - EvaluateExpr(rexpr)
        End Select
        Exit Function
    End If
    
    '   If we do not yet have an operator, there are several possibilities:
    '   1.   expr   is   (expr2)   for   some   expr2.
    '   2.   expr   is   -expr2   or   +expr2   for   some   expr2.
    '   3.   expr   is   Fun(expr2)  for  a  function   Fun.  (正玄余玄等函数)
    '   4.   expr   is   a   primitive.
    '   5.   It's a  literal like "3.14159".
    
    '   Look   for   (expr2).
    If Left$(expr, 1) = "(" And Right$(expr, 1) = ")" Then
        '   Remove   the   parentheses.
        EvaluateExpr = EvaluateExpr(Mid$(expr, 2, expr_len - 2))
        Exit Function
    End If
    
    '   Look   for   -expr2.
    If Left$(expr, 1) = "-" Then
        EvaluateExpr = -EvaluateExpr( _
        Right$(expr, expr_len - 1))
        Exit Function
    End If
    
    '   Look   for   +expr2.
    If Left$(expr, 1) = "+" Then
        EvaluateExpr = EvaluateExpr( _
        Right$(expr, expr_len - 1))
        Exit Function
    End If
    
    '   Look   for   Fun(expr2).
    If expr_len > 5 And Right$(expr, 1) = ")" Then
        lexpr = LCase$(Left$(expr, 4))
        rexpr = Mid$(expr, 5, expr_len - 5)
        Select Case lexpr
        Case "sin("
            EvaluateExpr = Sin(EvaluateExpr(rexpr))
            Exit Function
        Case "cos("
            EvaluateExpr = Cos(EvaluateExpr(rexpr))
            Exit Function
        Case "tan("
            EvaluateExpr = Tan(EvaluateExpr(rexpr))
            Exit Function
        Case "sqr("
            EvaluateExpr = Sqr(EvaluateExpr(rexpr))
            Exit Function
        End Select
    End If
    
    ' See if it's a primitive.
    'On Error Resume Next
    'Value = Primitives.Item(expr)
    'status = Err.Number
    'On Error GoTo 0          '禁止当前过程中任何已启动的错误处理程序。
    'If status = 0 Then
    'EvaluateExpr = CSng(Value)
    'Exit Function
    'End If
    
    'It must be a literal like "2.71828".
    On Error Resume Next
    EvaluateExpr = CSng(expr)
    status = Err.Number
    On Error GoTo 0          '禁止当前过程中任何已启动的错误处理程序。
    If status <> 0 Then
        'Err.Raise status, "EvaluateExpr", "Error evaluating '" & expr & "'as a constant."
        MsgBox " The Expression format is wrong ", vbInformation, "System Info."
    End If
End Function

Function IsShow(FormName As String) As Boolean
    '--------------------------------------------------------------------------
    '功能:           判断一个窗口是否已经打开，如果是已经打开的则激活
    '参数:           窗口frm的窗口标题
    '返回值:         Boolean
    '--------------------------------------------------------------------------
    Dim lR     As Long
    lR = FindWindow(vbNullString, FormName)
    If lR <> 0 Then
        IsShow = True
        lR = SetActiveWindow(lR)
        'SendKeys ("{ENTER}")        'SendKeys 语句 将一个或多个按键消息发送到活动窗口，就如同在键盘上进行输入一样。
        
    Else
        IsShow = False
    End If
End Function



'判断文件是否存在
Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
    FileExist = (Dir(Fname) <> "")
End Function
Public Function GetKey(Tmp_File As String, Tmp_Key As String) As String
    Dim File As Long
    '分配文件句柄
    File = FreeFile
    '如果文件不存在则创建一个默认的Setup.ini文件
    If FileExist(Tmp_File) = False Then
        GetKey = ""
        Call WritePrivateProfileString("Setup Information", "Server", "", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "DataBase", " ", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "UserName", " ", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "Password", " ", App.Path + "\Setup.ini")
        Exit Function
    End If
    '读取数据项值
    Open Tmp_File For Input As File
    Do While Not EOF(1)
        Line Input #File, buffer
        If Left(buffer, Len(Tmp_Key)) = Tmp_Key Then
            Pos = InStr(buffer, "=")
            GetKey = Trim(Mid(buffer, Pos + 1))
        End If
    Loop
    Close File
End Function


Public Function FormatNumber6(Num As String) As String
    Dim i As Integer
    Dim Rtn As String
    Rtn = ""
    For i = 1 To 6 - Len(Num)
        Rtn = Rtn & "0"
    Next
    Rtn = Rtn & Num
    FormatNumber6 = Rtn
End Function

Public Function IIsNull(val As Object)
    If IsNull(val) Then
        IIsNull = ""
    Else
        IIsNull = Trim(CStr(val))
    End If
End Function




''在调用ResizeForm前先调用本函数
'Public Sub ResizeInit(FormName As Form)
'  Dim OBJ As Control
'  ObjOldWidth = FormName.ScaleWidth
'  ObjOldHeight = FormName.ScaleHeight
'  ObjOldFont = FormName.Font.Size / ObjOldHeight
'  On Error Resume Next
'
'  For Each OBJ In FormName
'    OBJ.Tag = OBJ.Left & " " & OBJ.Top & " " & OBJ.Width & " " & OBJ.Height & " "
'  Next OBJ
'
'  On Error GoTo 0
'End Sub
'
''按比例改变表单内各元件的大小，
''在调用ReSizeForm前先调用ReSizeInit函数
'Public Sub ResizeForm(FormName As Form)
'
'  Dim Pos(4) As Double
'  Dim i As Long, TempPos As Long, StartPos As Long
'  Dim OBJ As Control
'  Dim ScaleX As Double, ScaleY As Double
'
'  ScaleX = FormName.ScaleWidth / ObjOldWidth
'  '保存窗体宽度缩放比例
'  ScaleY = FormName.ScaleHeight / ObjOldHeight
'  '保存窗体高度缩放比例
'  On Error Resume Next
'
'  For Each OBJ In FormName
'    StartPos = 1
'    For i = 0 To 4
'      '读取控件的原始位置与大小
'      TempPos = InStr(StartPos, OBJ.Tag, " ", vbTextCompare)
'      If TempPos > 0 Then
'        Pos(i) = Mid(OBJ.Tag, StartPos, TempPos - StartPos)
'        StartPos = TempPos + 1
'      Else
'        Pos(i) = 0
'      End If
'
'      '根据控件的原始位置及窗体改变大
'      '小的比例对控件重新定位与改变大小
'      OBJ.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'      OBJ.Font.Size = ObjOldFont * FormName.ScaleHeight
'    Next i
'
'            '根据控件的原始位置及窗体改变大小的比例，对控件重新定位与改变大小
'    If Pos(0) >= 0 Then
'            If TypeName(OBJ) <> "ComboBox" Then
'                OBJ.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'            Else
'                'ComboBox不可改变其高度
'                OBJ.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX
'            End If
'        Else
'            '在SSTAB控件中，如果不在当前选项卡中的控件，其Left属性是要减去75000的
'            If TypeName(OBJ) <> "ComboBox" Then
'                OBJ.Move (Pos(0) + 75000) * ScaleX - 75000, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'            Else
'                'ComboBox不可改变其高度
'                OBJ.Move (Pos(0) + 75000) * ScaleX - 75000, Pos(1) * ScaleY, Pos(2) * ScaleX
'            End If
'        End If
'  Next OBJ
'
'  On Error GoTo 0
'End Sub


'MSFlexGrid   Export   to   MSExcel
Public Function FlexGrd_SaveToExcel(FG As MSFlexGrid, Optional sHeader As String = "", Optional sFooter As String = "", Optional ColumnHeaderFontColorIndex As Long, Optional ColumnHeaderBackColorIndex As Long, Optional CoLogoPicLocation As String, Optional WorkBkBackColorIndex As Long, Optional WorkBkGridColorIndex As Long, Optional AlternateRowColorIndex1 As Long, Optional AlternateRowColorIndex2 As Long, Optional AutoColumnFitter As Boolean)
    
    '     Autofit   columns
    '     Alternating   row   colors   in   excel
    
    Static objExcelDel     As Object
    Static objWorkbookDel     As excel.Workbook
    Static objWorksheetDel     As excel.Worksheet
    Static HeadRange           As excel.Range
    Static NewRange     As excel.Range
    Static GridRange     As Range
    Static PicObject     As excel.ShapeRange
    Dim lRow     As Integer, lCol       As Integer
    Dim i     As Integer, j       As Integer
    Dim C     As Integer
    
    Dim rowOffset     As Long
    Dim TempStr()     As String
    Set objExcelDel = CreateObject("Excel.Application")
    
    If Err.Number <> 0 Then
        Set objExcelDel = New excel.Application
        
        Err.Clear
    End If
    On Error Resume Next
    objExcelDel.Visible = False
    
    If Len(sHeader) > 0 Then
        TempStr = Split(sHeader, vbTab)
        rowOffset = UBound(TempStr) + 1
    End If
    
    
    
    Set objWorkbookDel = objExcelDel.Workbooks.Add
    
    'Turn   off   the   alerts
    objExcelDel.DisplayAlerts = False
    
    'Set   objWorksheet   to   the   remaining   worksheet.
    Set objWorksheetDel = objExcelDel.ActiveSheet
    
    With objWorksheetDel
        
        '   Sheet   Header
        For lRow = 1 To rowOffset
            .PageSetup.CenterHeader = TempStr(lRow - 1)
        Next lRow
        
        '   Get   Column   Headers
        For lRow = 1 To FG.FixedRows
            For lCol = 1 To FG.Cols
                .Cells(4, lCol - 1) = FG.TextMatrix(lRow - 1, lCol - 1)
            Next lCol
        Next lRow
        
        If val(WorkBkBackColorIndex) > 0 Then
            objWorkbookDel.Styles("Normal ").Interior.ColorIndex = WorkBkBackColorIndex
        End If
        'Gridlines   will   not   be   visible   but   you   can   add   that   to   by
        If val(WorkBkGridColorIndex) > 0 Then
            With objWorkbookDel.Styles("Normal ").Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1         '   1   is   black
            End With
            With objWorkbookDel.Styles("Normal ").Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1
            End With
            With objWorkbookDel.Styles("Normal ").Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1
            End With
            With objWorkbookDel.Styles("Normal ").Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1
            End With
        End If
        
        
        Set HeadRange = objWorksheetDel.Range(objWorksheetDel.Cells(4, 1), _
        objWorksheetDel.Cells(4, lCol - 2))
        With HeadRange
            '*****Sets   Column   Header   Back   Color
            If val(ColumnHeaderBackColorIndex) > 0 Then
                .Interior.ColorIndex = ColumnHeaderBackColorIndex
            Else
                '   My   Default   Background   color   for   Column   header   index   change   it   to   what   ever   you   want
                .Interior.ColorIndex = 5
            End If
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = 6
            .Interior.Pattern = xlLightHorizontal
            .Interior.ColorIndex = 20
            .Font.Name = "Rockwell "
            .Font.FontStyle = "Bold "
            .Font.Shadow = True
            '*****   Sets   Column   header   Font   color*****
            If val(ColumnHeaderFontColorIndex) > 0 Then
                .Font.ColorIndex = ColumnHeaderFontColorIndex
            Else
                '   My   Default   Font   color   for   Column   header   index   change   it   to   what   ever   you   want
                .Font.ColorIndex = 2
            End If
            .Font.Bold = True
            '************************************
            'Sets   border   colors   of   header.   You   could   also   add   this
            'to   the   function   but   I   thought   I   was   getting   carried   away
            'as   it   was.
            
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 16         'grey
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 16
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 16
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 16
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1       '   Black
            End With
        End With
        
        HeadRange = Nothing
        ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
        Dim RowCounter     As Integer     '   used   for   all   alternate   row   color
        RowCounter = 0             '   ditto
        '   Dim   ColCounter   As   Integer   '   used   for   all   alternate   row   color
        '   ColCounter   =   0
        Dim G     As Integer     '   ditto
        Dim Alternate     As Boolean       'ditto
        ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
        '   Fill   excel   sheet   with   data
        '   Row   data   from   flexgrid
        For i = 1 To FG.Rows
            
            For j = 0 To FG.Cols
                objWorksheetDel.Cells(i + 4, j) = FG.TextMatrix(i, j)
                objWorksheetDel.Cells(i + 4, j + 1).VerticalAlignment = xlTop
            Next j
            RowCounter = RowCounter + 1
        Next i
        RowCounter = RowCounter - 1             '   Getting   rid   of   extra   row
        ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
        '   Alternate   row   colors   on   Excel   spreadsheet
        If AlternateRowColorIndex1 <> " " And AlternateRowColorIndex2 <> " " Then
            
            G = 0
            Do Until G = RowCounter           '   RowCounter   is   figured   when   row   data   is   taken
                Set NewRange = objWorksheetDel.Range(objWorksheetDel.Cells(G + 5, 1), _
                objWorksheetDel.Cells(G + 5, lCol - 2))
                
                With NewRange
                    If Alternate <> True Then
                        .Interior.ColorIndex = AlternateRowColorIndex1
                        .Borders.ColorIndex = 31
                        'Sets   font   color   either   1   Black   or   2   white   for   row
                        Select Case AlternateRowColorIndex1
                        Case 1, 3, 5, 9, 11, 13, 14, 16, 17, 21, 23, 25
                            .Font.ColorIndex = 2
                        Case Else
                            .Font.ColorIndex = 1
                        End Select
                        Alternate = True
                    Else
                        .Interior.ColorIndex = AlternateRowColorIndex2
                        .Borders.ColorIndex = 31
                        'Sets   font   color   either   1   Black   or   2   white
                        Select Case AlternateRowColorIndex2
                        Case 1, 3, 5, 9, 11, 13, 14, 16, 17, 21, 23, 25
                            .Font.ColorIndex = 2
                        Case Else
                            .Font.ColorIndex = 1
                        End Select
                        Alternate = False
                    End If
                End With
                NewRange = Nothing
                G = G + 1
            Loop
        End If
        ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
        '   Autofit   columns
        If AutoColumnFitter = True Then
            .Columns.AutoFit
        End If
        '******************************************
        
        
        objWorksheetDel.OLEObjects
        
        
        '   Page   Footer
        If Len(sFooter) > 0 Then
            TempStr = Split(sFooter, vbTab)
            For lRow = 0 To UBound(TempStr)
                .PageSetup.CenterFooter = TempStr(lRow)
            Next lRow
        End If
        
    End With
    objExcelDel.Visible = True
    objExcelDel.DisplayAlerts = True
    Set objWorksheetDel = Nothing
    Set objWorkbookDel = Nothing
    Set objExcelDel = Nothing
End Function

'
'Public Function killByName(ByVal procName As String) As Boolean
'    Dim theloop, ret
'    Dim hand As Long
'    Dim proc As PROCESSENTRY32
'    Dim snap As Long
'    Dim exename As String, procId As Long
'    Dim dotExePos As Long
'    snap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0) 'get snapshot handle
'    proc.dwSize = Len(proc)
'    theloop = Process32First(snap, proc) 'first process and return value
'    While theloop <> 0 'next process
'    exename = ""
'
'    exename = LCase(proc.szExeFile)
'    procId = proc.th32ProcessID 'get process ID
'
'    dotExePos = InStr(exename, ".exe")
'    If dotExePos > 0 Then exename = Left(exename, dotExePos + 3)
'
'    If exename = LCase(procName) Then
'    hand = OpenProcess(PROCESS_TERMINATE, True, procId) 'get process handle
'    ret = TerminateProcess(hand, 0) 'end define process
'    End If
'    theloop = Process32Next(snap, proc)
'    Wend
'    CloseHandle snap 'close snapshot handle
'
'    If ret = 1 Then killByName = True Else killByName = False
'End Function

Public Function WriteLOG(ByVal LOGStr As String) As Boolean
On Error GoTo ErrorSet
    
    Dim LOGFilePath As String
    Dim LOGFileName As String
    Dim LOGDate As String
    Dim LOGFileNo As Integer
    
    
    LOGFilePath = App.Path & "\LOG\"
    LOGFileName = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & ".log"
    LOGDate = "[" & Now & "]    "
    LOGFileNo = FreeFile                        'FreeFile 返回一个 Integer，代表下一个可供 Open 语句使用的文件号
    
    If Dir(LOGFilePath, vbDirectory) = "" Then '检查目录是否存在，不存在建立一个LOG目录
        MkDir LOGFilePath
    End If
    
    Open LOGFilePath + LOGFileName For Append Access Write As #LOGFileNo
    
    Print #LOGFileNo, LOGDate & LOGStr
    Close #LOGFileNo
    
    WriteLOG = True
    
    Exit Function
    
ErrorSet:
    WriteLOG = False
    MsgBox "写入日志文件失败！       " & Err.Description & Chr(10) & Chr(13) & "未写入内容：" & LOGStr
    
End Function

Public Function setObjFocus(ByRef o As Object)
    o.SelText
    o.SetFocus
End Function

