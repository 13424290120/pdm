Attribute VB_Name = "OpnShellExcFile"
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'Define items in open shellExecute File
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
         ByVal lpszOp As String, _
         ByVal lpszFile As String, _
         ByVal lpszParams As String, _
         ByVal lpszDir As String, _
         ByVal FsShowCmd As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const SW_SHOWNORMAL = 1

Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'ShellExecute 查找与指定文件关联在一起的程序的文件名
'Long，非零表示成功，零表示失败。会设置GetLastError
'hwnd -- Long，指定一个窗口的句柄，有时候，Windows程序有必要在创建自己的主窗口前显示一个消息框，一般设置为父窗口句柄
'lpOperation ----  String，文件类型或默认执行程序。指定字串为“Open”将打开lpFlie指定文档；指定字符串为“Print”将打印lpFlie指定文档；指定字符串为“Explore”可浏览lpFlie指定的文件或文件夹，设置为Find可查找lpFlie指定的文件夹下的文件或文件夹。
'lpFile --- String，用关联程序打印或打开一个程序名或文件名，指定程序路径和文件名。
'lpParameters ---  String，如lpszFlie是可执行文件，则这个字串包含传递给执行程序的参数
'lpDirectory ----  String，想使用的完整路径，即默认目录。
'nShowCmd -- Long，定义了如何显示启动程序的常数值。参考ShowWindow函数的nCmdShow参数。

'ShowWindow函数的nCmdShow参数如下
'SW_HIDE              隐藏窗口，活动状态给令一个窗口
'SW_MINIMIZE          最小化窗口，活动状态给令一个窗口
'SW_RESTORE           用原来的大小和位置显示一个窗口，同时令其进入活动状态
'SW_SHOW              用当前的大小和位置显示一个窗口，同时令其进入活动状态
'SW_SHOWMAXIMIZED     最大化窗口，并将其激活
'SW_SHOWMINIMIZED     最小化窗口，并将其激活
'SW_SHOWMINNOACTIVE   最小化一个窗口，同时不改变活动窗口
'SW_SHOWNA            用当前的大小和位置显示一个窗口，不改变活动窗口
'SW_SHOWNOACTIVATE    用最近的大小和位置显示一个窗口，同时不改变活动窗口
'SW_SHOWNORMAL        与SW_RESTORE相同

Function StartDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function

Function OpnFileExist(Filename As String) As Boolean      '给定位置某个文件是否存在
          On Error Resume Next
          
          Dim FileID  As Long
          FileID = FreeFile()        'FreeFile 函数,返回一个 Integer,代表下一个可供 Open 语句使用的文件号。
          
          Open Filename For Input As #FileID
          Close #FileID
          
          OpnFileExist = (Err.Number = 0)
          Err.Clear
End Function

Sub OpnShllExcFile(FilePathandName As String)   '前面不能有private出现，必须是Sub....
    Dim r As Long, msg As String
    r = StartDoc(FilePathandName)         'Change this to a valid path
    If r <= 32 Then
        'There was an error
        Select Case r
            Case SE_ERR_FNF
                msg = "File not found"
            Case SE_ERR_PNF
                msg = "Path not found"
            Case SE_ERR_ACCESSDENIED
                msg = "Access denied"
            Case SE_ERR_OOM
                msg = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                msg = "DLL not found"
            Case SE_ERR_SHARE
                msg = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                msg = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                msg = "DDE Time out"
            Case SE_ERR_DDEFAIL
                msg = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                msg = "DDE busy"
            Case SE_ERR_NOASSOC
                msg = "No association for file extension"
            Case ERROR_BAD_FORMAT
                msg = "Invalid EXE file or error in EXE image"
            Case Else
                msg = "Unknown error"
        End Select
        MsgBox msg, vbInformation, "System Info."
    End If
End Sub

