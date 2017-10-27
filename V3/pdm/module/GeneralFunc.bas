Attribute VB_Name = "GeneralFunc"
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'�ж�һ���ļ��Ƿ��Ѿ���
Public Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
' �Զ�����ģʽ��ָ�����ļ�
' ����ֵLong����ִ�гɹ������ش��ļ��ľ����HFILE_ERROR��ʾ����������GetLastError
'lpPathName -----  String�������ļ�������
'iReadWrite -----  Long������ģʽ�͹���ģʽ������һ����ϣ�������ʾ��
'1)  ����ģʽ
'Read
'���ļ�����ȡ���е�����
'READ_WRITE
'���ļ���������ж�д
'WRITE
'���ļ���������д������
'2)  ����ģʽ (�ο�OpenFile�����ı�־������)
'OF_SHARE_COMPAT�� OF_SHARE_DENY_NONE�� OF_SHARE_DENY_READ��OF_SHARE_DENY_WRITE�� OF_SHARE_EXCLUSIVE
Public Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
'�ر�ָ�����ļ�����ο�CloseHandle�������˽��һ�������


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Ѱ�Ҵ����б��е�һ������ָ�������Ķ������ڣ���vb��ʹ�ã�FindWindow�����һ����;�ǻ��ThunderRTMain������ش��ڵľ��������������������vbִ�г����һ���֡���þ���󣬿���api����GetWindowTextȡ��������ڵ����ƣ�����Ҳ��Ӧ�ó���ı��⣩
'����ֵLong���ҵ����ڵľ������δ�ҵ�������ڣ��򷵻��㡣������GetLastError
'lpClassName ----  String��ָ������˴��������Ŀ���ֹ��C���ԣ��ִ���ָ�룻����Ϊ�㣬��ʾ�����κ���
'lpWindowName ---  String��ָ������˴����ı������ǩ���Ŀ���ֹ��C���ԣ��ִ���ָ�룻����Ϊ�㣬��ʾ�����κδ��ڱ���
'����Ҫ��ͬʱ�����봰����������Ϊ���Լ���׼����������һ���㣬����İ취�Ǵ���vbNullString����
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'����ָ���Ĵ���
'����ֵLong ǰһ������ڵľ��
'hwnd -----------Long��������ڵľ��


Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'��� OpenProcess ������һ���Ѵ��ڵĽ��̶���
'��ɹ�������ֵΪָ�����̵Ĵ򿪾��  ��ʧ�ܣ�����ֵΪ�գ��ɵ���GetLastError��ô�����롣
'dwDesiredAccess�Ƿ��ʽ��̵�Ȩ�� bInheritHandle�Ǿ���Ƿ�̳н������� dwProcessId�ǽ��̣ɣ�
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'�ر�һ���ں˶������а����ļ����ļ�ӳ�䡢���̡��̡߳���ȫ��ͬ������ȡ�
'����ֵ�����ʾ�ɹ������ʾʧ�ܡ�������GetLastError
'hObject Long�����رյ�һ������ľ��
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'WaitForSingleObject(�¼����, ��ʱʱ��)
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'��һ��Excel�ļ���/�رյĲ���
Public xlApp As excel.Application
Public xlBook As excel.Workbook
Public xlSheet As excel.Worksheet

Private ObjOldWidth As Long  '���洰���ԭʼ���
Private ObjOldHeight As Long '���洰���ԭʼ�߶�
Private ObjOldFont As Single '���洰���ԭʼ�����

' �ж�ĳ�ļ��Ƿ��ڴ�ʹ����
Function IsFileAlreadyOpen(Filename As String) As Boolean
    Dim hFile     As Long
    Dim lastErr     As Long
    hFile = -1                                 '��ʼ���ļ����.
    lastErr = 0
    hFile = lOpen(Filename, &H10)
    
    If hFile = -1 Then                         '�ļ��Ƿ�����ȷ�򿪲��ɹ���
        lastErr = Err.LastDllError
    Else
        lClose (hFile)                     'hFile�ļ����.
    End If
    IsFileAlreadyOpen = (hFile = -1) And (lastErr = 32)
End Function
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'��һ��Excel�ļ��رյĲ���
Sub ModulesCloseExcel(strFileName As String)
    
    If IsFileAlreadyOpen(strFileName) Then
        xlBook.Close (True)                      '�ر�EXCEL������
        xlApp.Quit                               '�ر�EXCEL
        Set xlApp = Nothing                       '�ͷ�EXCEL����
    End If
    
End Sub

'��һ��Excel�ļ��򿪵Ĳ���
Sub ModulesOpenExcel(strFileName As String)               '��EXCEL����
    
    If IsFileAlreadyOpen(strFileName) Then
        MsgBox "Excel File is already Openned, Please Close it firstly", vbInformation, "System Info."
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")   '����EXCELӦ�������
    xlApp.Visible = True                            '����EXCELӦ�������ɼ�
    Set xlBook = xlApp.Workbooks.Open(strFileName)  '��EXCEL������
    Set xlSheet = xlBook.Worksheets(1)              '��EXCEL������1
    xlSheet.Activate                                '����EXCEL������
    
End Sub
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


'MSHFlexGrid�ؼ�������Excel
Function ExportFlexDataToExcel(flex As MSFlexGrid, g_CommonDialog As CommonDialog, Optional strHead As String = "")
    
    On Error GoTo ErrHandler
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim Rows As Integer, Cols As Integer
    Dim eRow As Integer, eCol As Integer, fCol As Integer, fRow As Integer
    Dim New_Col As Boolean
    Dim New_Column  As Boolean
    
    g_CommonDialog.CancelError = True
    
    ' ���ñ�־
    g_CommonDialog.Flags = cdlOFNHideReadOnly
    ' ���ù�����
    g_CommonDialog.Filter = "All Files (*.*)|*.*|Excel Files" & "(*.xls)|*.xls"
    ' ָ��ȱʡ�Ĺ�����
    g_CommonDialog.FilterIndex = 2
    ' ��ʾ���򿪡��Ի���  ͨ��ʹ�� CommonDialog �ؼ��� ShowOpen �� ShowSave ��������ʾ���򿪡��͡����Ϊ���Ի���
    g_CommonDialog.ShowSave
    
    If flex.Rows <= 1 Then
        MsgBox "No Data��", vbInformation, "System Info"
        Exit Function
    End If
    
    MousePointer = vbHourglass   '��дʱ��ϳ�����Ҫ�������״̬
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    xlApp.Visible = False
    
    With flex
        Rows = .Rows
        Cols = .Cols
        eRow = 1   'Excel�д�1��ʼ
        eCol = 1   'Excel�д�1��ʼ
        fRow = 0   'MsFlexGird�д�0��ʼ
        fCol = 0   'MsFlexGird�д�0��ʼ
        xlApp.Cells(1, Int(Cols / 2)).Value = strHead
        For eRow = 2 To Rows + 1
            If flex.RowHeight(fRow) = 0 Then GoTo nextrow           '����и�Ϊ0,��ʾ�����ص���,�����
            For eCol = 1 To Cols
                fCol = eCol - 1
                xlApp.Cells(eRow, eCol).NumberFormat = "@"             '����ÿ����Ԫ����������Ϊ�ı���
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
    Set xlApp = Nothing  '"�������Ƹ�Excel
    Set xlBook = Nothing
    flex.SetFocus
    MousePointer = vbDefault                  '�ָ����״̬
    MsgBox "Data have been exported into Excel", vbInformation, "System Info"
    Exit Function
    
ErrHandler:
    ' �û����ˡ�ȡ������ť
    If Err.Number <> 32755 Then
        MsgBox "Cancel Data Exporting", vbInformation, "System Info"
    End If
End Function





Function EvaluateExpr(ByVal expr As String) As Single
    '--------------------------------------------------------------------------
    '����:           �ַ������ʽ�ļ��㺯��
    '����:
    '               [expr]...........................�ַ������ʽ
    '����ֵ:
    '               [EvaluateExpr]...................������ֵ
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
    
    '   ɾ����β�ո���Ч��У��
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
        '   Examine the next character.(�����һ���ַ�)
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
    '   best_prec����ߵ������
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
    '   3.   expr   is   Fun(expr2)  for  a  function   Fun.  (���������Ⱥ���)
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
    'On Error GoTo 0          '��ֹ��ǰ�������κ��������Ĵ��������
    'If status = 0 Then
    'EvaluateExpr = CSng(Value)
    'Exit Function
    'End If
    
    'It must be a literal like "2.71828".
    On Error Resume Next
    EvaluateExpr = CSng(expr)
    status = Err.Number
    On Error GoTo 0          '��ֹ��ǰ�������κ��������Ĵ��������
    If status <> 0 Then
        'Err.Raise status, "EvaluateExpr", "Error evaluating '" & expr & "'as a constant."
        MsgBox " The Expression format is wrong ", vbInformation, "System Info."
    End If
End Function

Function IsShow(FormName As String) As Boolean
    '--------------------------------------------------------------------------
    '����:           �ж�һ�������Ƿ��Ѿ��򿪣�������Ѿ��򿪵��򼤻�
    '����:           ����frm�Ĵ��ڱ���
    '����ֵ:         Boolean
    '--------------------------------------------------------------------------
    Dim lR     As Long
    lR = FindWindow(vbNullString, FormName)
    If lR <> 0 Then
        IsShow = True
        lR = SetActiveWindow(lR)
        'SendKeys ("{ENTER}")        'SendKeys ��� ��һ������������Ϣ���͵�����ڣ�����ͬ�ڼ����Ͻ�������һ����
        
    Else
        IsShow = False
    End If
End Function



'�ж��ļ��Ƿ����
Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
    FileExist = (Dir(Fname) <> "")
End Function
Public Function GetKey(Tmp_File As String, Tmp_Key As String) As String
    Dim File As Long
    '�����ļ����
    File = FreeFile
    '����ļ��������򴴽�һ��Ĭ�ϵ�Setup.ini�ļ�
    If FileExist(Tmp_File) = False Then
        GetKey = ""
        Call WritePrivateProfileString("Setup Information", "Server", "", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "DataBase", " ", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "UserName", " ", App.Path + "\Setup.ini")
        Call WritePrivateProfileString("Setup Information", "Password", " ", App.Path + "\Setup.ini")
        Exit Function
    End If
    '��ȡ������ֵ
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




''�ڵ���ResizeFormǰ�ȵ��ñ�����
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
''�������ı���ڸ�Ԫ���Ĵ�С��
''�ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
'Public Sub ResizeForm(FormName As Form)
'
'  Dim Pos(4) As Double
'  Dim i As Long, TempPos As Long, StartPos As Long
'  Dim OBJ As Control
'  Dim ScaleX As Double, ScaleY As Double
'
'  ScaleX = FormName.ScaleWidth / ObjOldWidth
'  '���洰�������ű���
'  ScaleY = FormName.ScaleHeight / ObjOldHeight
'  '���洰��߶����ű���
'  On Error Resume Next
'
'  For Each OBJ In FormName
'    StartPos = 1
'    For i = 0 To 4
'      '��ȡ�ؼ���ԭʼλ�����С
'      TempPos = InStr(StartPos, OBJ.Tag, " ", vbTextCompare)
'      If TempPos > 0 Then
'        Pos(i) = Mid(OBJ.Tag, StartPos, TempPos - StartPos)
'        StartPos = TempPos + 1
'      Else
'        Pos(i) = 0
'      End If
'
'      '���ݿؼ���ԭʼλ�ü�����ı��
'      'С�ı����Կؼ����¶�λ��ı��С
'      OBJ.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'      OBJ.Font.Size = ObjOldFont * FormName.ScaleHeight
'    Next i
'
'            '���ݿؼ���ԭʼλ�ü�����ı��С�ı������Կؼ����¶�λ��ı��С
'    If Pos(0) >= 0 Then
'            If TypeName(OBJ) <> "ComboBox" Then
'                OBJ.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'            Else
'                'ComboBox���ɸı���߶�
'                OBJ.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX
'            End If
'        Else
'            '��SSTAB�ؼ��У�������ڵ�ǰѡ��еĿؼ�����Left������Ҫ��ȥ75000��
'            If TypeName(OBJ) <> "ComboBox" Then
'                OBJ.Move (Pos(0) + 75000) * ScaleX - 75000, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'            Else
'                'ComboBox���ɸı���߶�
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
    LOGFileNo = FreeFile                        'FreeFile ����һ�� Integer��������һ���ɹ� Open ���ʹ�õ��ļ���
    
    If Dir(LOGFilePath, vbDirectory) = "" Then '���Ŀ¼�Ƿ���ڣ������ڽ���һ��LOGĿ¼
        MkDir LOGFilePath
    End If
    
    Open LOGFilePath + LOGFileName For Append Access Write As #LOGFileNo
    
    Print #LOGFileNo, LOGDate & LOGStr
    Close #LOGFileNo
    
    WriteLOG = True
    
    Exit Function
    
ErrorSet:
    WriteLOG = False
    MsgBox "д����־�ļ�ʧ�ܣ�       " & Err.Description & Chr(10) & Chr(13) & "δд�����ݣ�" & LOGStr
    
End Function

Public Function setObjFocus(ByRef o As Object)
    o.SelText
    o.SetFocus
End Function

