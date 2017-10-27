Attribute VB_Name = "Const"
'声明写入ini文件的API函数
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFilenchame As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFilenchame As String) As Long

'Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'
'
'Public Type PROCESSENTRY32
'
'    dwSize As Long
'
'    cntUsage As Long
'
'    th32ProcessID As Long
'
'    th32DefaultHeapID As Long
'
'    th32ModuleID As Long
'
'    cntThreads As Long
'
'    th32ParentProcessID As Long
'
'    pcPriClassBase As Long
'
'    dwFlags As Long
'
'    szExeFile As String * 1024
'
'End Type
'
'Const TH32CS_SNAPHEAPLIST = &H1
'
'Const TH32CS_SNAPPROCESS = &H2
'
'Const TH32CS_SNAPTHREAD = &H4
'
'Const TH32CS_SNAPMODULE = &H8
'
'Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
'
'Const TH32CS_INHERIT = &H80000000
'
'
'Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'
'Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Public connString As String

'定义权限变量
Public Try_times As Integer
Public SystemAdmin As String
Public EngineeringSys As String
Public PurchasingSys As String

Public EngineeringApproval As String
Public BOMManagement As String

'定义服务器参数常量
Public Server As String
Public DataBase As String
Public DBUser As String
Public Password As String
Public QueryTableName As String            '定义查询表名的公用字符串
Public PDMUserName As String               '定义登陆PDM用户名的公用字符串
Public PDMUserGroup As String

Public EngineerApprovalRight As Boolean    '定义登陆PDM用户名的是否有EngineerApprovalRight
Public FromForm  As Form
Public FromForm2  As Form
'声明类模块变量
Public MyUsers As New ClsUsers
Public MyDB As New ClsDB
Public MyDatabaseSet As New ClsDatabaseSet

Public MyGlueSupplier As New ClsGlueSupplier
Public MyRFSRFQ As New ClsRFSRFQ
Public MyPJNO As New ClsPJNO
Public MyCPCN As New ClsCPCN
Public MySER As New ClsSER
Public MyCNCSN As New ClsCNCSN
Public MyFinsGd As New ClsFinsGd
Public MySglPrt As New ClsSglPrt

'定义窗体大小调整变量
Public ObjOldWidth     As Long       '保存窗体的原始宽度
Public ObjOldHeight     As Long     '保存窗体的原始高度
Public ObjOldFont     As Single     '保存窗体的原始字体比

'定义数组用来装载信息
Public Arr_Item() As String
Public UsrCtlFind(100) As String

'定义常数用于API OpenProcess和WaitForSingleObject 参见FrmBomAdmin窗口中代码
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = &HFFFFFFFF

'保存执行SQL语句的字符串
Public SqlStmt As String

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'Define Treeview Messages and styles
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETEDITCONTROL As Long = (TV_FIRST + 15)
'Public Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
'Public Const TVM_SELECTITEM As Long = (TV_FIRST + 11)
'
Public Const TVIF_STATE As Long = &H8
'Public Const TVS_TRACKSELECT As Long = &H200&
'Public Const TVS_FULLROWSELECT As Long = &H1000
Public Const TVIS_BOLD As Long = &H10
'
'Public Const TVGN_ROOT As Long = &H0
'Public Const TVGN_NEXT As Long = &H1
Public Const TVGN_CARET As Long = &H9
Public Const EM_LIMITTEXT = &HC5
Public Const WM_VSCROLL = &H115

'Define Treeview Item Structure
Public Const WM_SETREDRAW As Long = &HB
Public Type TVITEM
   mask As Long
   hItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

 Public Type RECT
         left As Long
         top As Long
         right As Long
         bottom As Long
 End Type
 
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'Define TreeView sendmessage and RegisterClipboard
Public Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function RegisterClipboardFormat Lib "USER32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Integer


 Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 Public Declare Function DrawText Lib "USER32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Public Const DT_TOP = &H0
 Public Const DT_VCENTER = &H4
 Public Const DT_BOTTOM = &H8
 Public Const DT_LEFT = &H0
 Public Const DT_CENTER = &H1
 Public Const DT_RIGHT = &H2
 Public Const DT_SINGLELINE = &H20
 Public Const DT_NOPREFIX = &H800
 Public Declare Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 Public Declare Function OffsetRect Lib "USER32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Public Sub LoadSkin(ByRef objForm As Form)
    On Error Resume Next
    Attach objForm.hWnd
    Dim objC As Control
    For Each objC In objForm
        If TypeName(objC) <> "Label" Then
            objC.FontSize = 10
            objC.FontName = "Segoe UI"
        End If
    Next
End Sub

Public Function GetAlign(ByVal Align As AlignmentSettings) As Long
    Select Case Align
    Case flexAlignLeftTop
        GetAlign = DT_LEFT Or DT_TOP
    Case flexAlignLeftCenter
        GetAlign = DT_LEFT Or DT_VCENTER
    Case flexAlignLeftBottom
        GetAlign = DT_LEFT Or DT_BOTTOM
    Case flexAlignCenterTop
        GetAlign = DT_CENTER Or DT_TOP
    Case flexAlignCenterCenter
        GetAlign = DT_CENTER Or DT_VCENTER
    Case flexAlignCenterBottom
        GetAlign = DT_CENTER Or DT_BOTTOM
    Case flexAlignRightTop
        GetAlign = DT_RIGHT Or DT_TOP
    Case flexAlignRightCenter
        GetAlign = DT_RIGHT Or DT_VCENTER
    Case flexAlignRightBottom
        GetAlign = DT_RIGHT Or DT_BOTTOM
    End Select
End Function










