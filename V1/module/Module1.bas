Attribute VB_Name = "Module1"
Option Explicit

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'Define Items in Database backup/restore utility
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'这个函数可以将hpvSource的memory   copy到hpvDest上去，cbCopy则代表要copy多少个byte。例如想一个Double值存在Memory中的各个byte到底是多少。

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long
'返回值 Long，TRUE（非零）表示成功，否则返回零。会设置GetLastError
'lpBuffer -------  String，随同计算机名载入的字串缓冲区
'nSize ----------  Long，缓冲区的长度。这个变量随同返回计算机名的实际长度载入

'Define Items in Database backup/restore utility
Public GroupName As String    '异地机的工作组名
Public UserName As String     '异地机的用户名
Public Pwd As String          '异地机的密码
Public optflag As Boolean     '标识是否把文件夹放在异地机

Public iFlag As String * 1
Public strConn As String

Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'Define items in open shellExecute File
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
         ByVal lpszOp As String, _
         ByVal lpszFile As String, _
         ByVal lpszParams As String, _
         ByVal lpszDir As String, _
         ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Const SW_SHOWNORMAL = 1

Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Sub BackupDatabase(ByVal BackUpfileName, ByVal DataBaseName, Optional ByVal IsAddBackup As Boolean = False)
Dim ConcStr, sql, iReturn$

        sql = "backup   database   [" & DataBaseName & "]" & vbCrLf & "to   disk='" & BackUpfileName & "'" & vbCrLf & _
        "with   description='" & "zj-backup   at:" & Date & "(" & Time & ")'" & vbCrLf & _
        IIf(IsAddBackup, "", ",init")
        TransactSQL sql

End Sub



  '*************************************************************************
  '**模   块   名：Rstoredatabase
  '**描         述：恢复数据库,返回出错信息,正常恢复,返回""
  '**调         用：RStoredatabase   "备份文件名","数据库名"
  '**参数说明：
  '**                     DataBasePath     恢复后的数据库存放目录
  '**                     BackupNumber     是从那个备份号恢复
  '**                     ReplaceExist     指定是否覆盖已经存在的数据
  '**说         明：引用Microsoft   ActiveX   Data   Objects   2.x   Library
 
  '*************************************************************************
Sub RestoreDatabase(ByVal BackUpfileName, ByVal DataBaseName _
, Optional ByVal DataBasePath = "", Optional ByVal BackupNumber = 1 _
, Optional ByVal ReplaceExist As Boolean = False)

Dim ConcStr, sql, i, MsgInfo
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim cn As New Connection

On Error GoTo ErrHandel

        
        sql = "restore   filelistonly   from   disk='" & BackUpfileName & "'" & vbCrLf & _
        "with   file=" & BackupNumber
        Call Connet
        cn.Open strConn
        rs.Open sql, cn
        If Len(Trim(DataBaseName)) = 0 Then    '取得备份数据库的逻辑名称
            DataBaseName = rs.Fields(0)
        End If

        If DataBasePath = "" Then
            sql = "select   filename   from   master..sysfiles"
            Set rs1 = TransactSQL(sql)
            sql = rs1.Fields(0)
            DataBasePath = Left(sql, InStrRev(sql, "\") - 1) '   InStrRev(SQL, "\")返回字符串数量,49
            rs1.Close
        End If
            
        If ReplaceExist = False Then
            sql = "select   1   from   master..sysdatabases     where   name='" & DataBaseName & "'"
            Set rs1 = TransactSQL(sql)
            If rs1.EOF = False Then
                MsgBox "数据库已经存在!", vbInformation, "提示"
                rs1.Close
                GoTo ErrExit
            End If
            rs1.Close
        End If
        

        sql = "select   spid   from   master..sysprocesses   where   dbid=db_id('" & DataBaseName & "')"
        Set rs1 = TransactSQL(sql)
        While rs1.EOF = False
            sql = "kill   " & rs1(0)
            TransactSQL (sql)
            rs1.MoveNext
        Wend
        rs1.Close
                                  
        sql = "restore   database   [" & DataBaseName & "]" & vbCrLf & _
        "from   disk='" & BackUpfileName & "'" & vbCrLf & _
        "with   file=" & BackupNumber & vbCrLf

        With rs
            While Not .EOF
                sql = sql & ",move   '" & rs("LogicalName") & "' to  '" & DataBasePath & "\" & rs("LogicalName") & "'" & vbCrLf
                .MoveNext
            Wend
            .Close
        End With
        sql = sql & IIf(ReplaceExist, ",replace", "")
        TransactSQL (sql)
        If iFlag = 1 Then
            MsgBox "恭喜你,数据库恢复成功了!", , "你真聪明"
        Else
            MsgBox "不好意思,数据库恢复失败了!注意:如果从一台机恢复到另一台机,请共享存放恢复文件的文件夹!", , "努力学习"
        End If

ErrExit:
    Exit Sub
ErrHandel:
    MsgBox Err.Description, vbExclamation, "错误"
    GoTo ErrExit:
    
End Sub

Sub Connet()
    If FrmServerBkup.OptCheck(0).Value = True Then
       strConn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=" & Trim(FrmServerBkup.CobServer.Text)
    Else
       strConn = "Provider=SQLOLEDB.1;Password='" & FrmServerBkup.TxtPassword.Text & "';Persist Security Info=False;User ID=" & Trim(FrmServerBkup.TxtName.Text) & ";Data Source=" & Trim(FrmServerBkup.CobServer.Text)
    End If
End Sub


Function GetIPAddress(ByVal HostName As String) As String '给定机器名,返回本机Ip地址
Dim lpHost      As Long
Dim HOST        As HOSTENT     'HOSTENT是一个Type语句定义的嵌套数据结构，见上面
Dim dwIPAddr    As Long
Dim tmpIPAddr() As Byte        'Byte 变量存储为单精度型,范围在 0 至 255 之间,在存储二进制数据时很有用。
Dim i           As Integer
Dim sIPAddr     As String

On Error GoTo ErrHandel

    lpHost = gethostbyname(HostName)  '本模块Modulel中有Declare Function gethostbyname Lib "WSOCK32.DLL...."
    
    CopyMemory HOST, lpHost, Len(HOST)  '本模块Modulel中有Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ......
    
    CopyMemory dwIPAddr, HOST.hAddrList, 4   '本模块Modulel中有Declare Sub CopyMemory Lib "kernel32" Alias......
    
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
    
    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    
ErrExit:
    Exit Function
ErrHandel:
    MsgBox Err.Description, vbExclamation, "错误"
    GetIPAddress = "不能检测到此计算机IP!"
    GoTo ErrExit:
End Function



Function TransactSQL(ByVal sql As String) As ADODB.Recordset
Dim strArray() As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
 
On Error GoTo TransactSQL_Error
    strArray = Split(sql)
    Call Connet
    cn.CommandTimeout = 0
    cn.Open strConn
    If StrComp(UCase$(strArray(0)), "select", vbTextCompare) = 0 Then
        rs.Open Trim$(sql), cn, 1, 1
        Set TransactSQL = rs
        iFlag = 1
    Else
        cn.Execute sql
        iFlag = 1
    End If
TransactSQL_Exit:
    Set rs = Nothing
    Set cn = Nothing
    Exit Function
TransactSQL_Error:
    MsgBox "查询错误：" & Err.Number & ": " & Err.Description
    iFlag = 2
    Resume TransactSQL_Exit
End Function


Sub HandleFile(optflag As Boolean, flag As Boolean)
Dim sql As String

    If optflag = True Then  '检测是否存在网络磁盘,如果有,则断开
       sql = "master..xp_cmdshell 'NET USE K: /DELETE /YES'"
       TransactSQL (sql)
       optflag = False
    End If
    If Len(Dir(App.Path & "\share.bat")) > 0 Then     '如果在本机操作,连接异地的服务器,而把备
        Kill App.Path & "\share.bat"                  '份文件放在本机时,因为需要共享文件夹,所
        Open App.Path & "\share.bat" For Output As #1 '以当存放好备份文件时则断开共享.
        Print #1, "net share sharefile /delete"
        Close #1
        Shell App.Path & "\share.bat"
        Kill App.Path & "\share.bat"
    End If
    flag = False
End Sub


Sub CheckServer(filename As String, flag As Boolean)
On Error GoTo ErrHandle:
Dim sql As String
Dim FilePath As String
Dim vstr As String
        
        vstr = IIf(FrmServerBkup.Tag = "1", "备份", "恢复")
        If Left(filename, 2) = "\\" Then
        '在异地服务器里备份到异地备份机里
        '在本地服务器里备份到异地备份机里
            FilePath = Left(filename, InStrRev(filename, "\") - 1)
            FrmSaveOption.Label4.Caption = "你将数据库的" & vstr & "文件存放在" & FilePath & "上,请输入这台机的登录信息!"
            FrmSaveOption.Show 1
            If FrmServerBkup.Tag = "1" Then
                FrmServerBkup.MousePointer = vbHourglass
            Else
                FrmOption.MousePointer = vbHourglass
            End If
            sql = "master..xp_cmdshell 'net use k: " & FilePath & " " & Pwd & " /user:" & GroupName & "\" & UserName & "'"
        Else
            '在异地服务器里备份到本地备份机里
            If StrComp(Trim(FrmServerBkup.LabServerName), Trim(FrmServerBkup.LabComputer.Caption), 1) <> 0 Then
                FilePath = Left(filename, InStrRev(filename, "\") - 1)
                Open App.Path & "\share.bat" For Output As #1
                Print #1, "net share ShareFile=" & FilePath; ""
                Close #1
                Shell App.Path & "\share.bat"
                FrmSaveOption.Label4.Caption = "你将数据库的" & vstr & "文件存放在" & FilePath & "上,请输入这台机的登录信息!"
                FrmSaveOption.Show 1
                If FrmServerBkup.Tag = "1" Then
                    FrmServerBkup.MousePointer = vbHourglass
                Else
                    FrmOption.MousePointer = vbHourglass
                End If
                sql = "master..xp_cmdshell 'net use k: \\" & FrmServerBkup.LabComputer & "\sharefile " & Pwd & " /user:" & GroupName & "\" & UserName & "'"
            Else
                '在本地服务器里备份到本地备份机里
                GoTo HandelPoint
            End If
        End If
        
        If optflag = False Then flag = False: GoTo ExitPoint
        TransactSQL sql
        filename = Right(filename, Len(filename) - InStrRev(filename, "\"))
        filename = "k:\" & filename
        
HandelPoint:
        flag = True
ExitPoint:
    Exit Sub
ErrHandle:
    MsgBox "Error: " & Err.Description & "On Sub CheckServer!", vbExclamation, "错误"
    GoTo ExitPoint
End Sub

