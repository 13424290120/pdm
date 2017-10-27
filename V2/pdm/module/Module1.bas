Attribute VB_Name = "Module1"
Option Explicit

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'Define Items in Database backup/restore utility
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'����������Խ�hpvSource��memory   copy��hpvDest��ȥ��cbCopy�����Ҫcopy���ٸ�byte��������һ��Doubleֵ����Memory�еĸ���byte�����Ƕ��١�

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long
'����ֵ Long��TRUE�����㣩��ʾ�ɹ������򷵻��㡣������GetLastError
'lpBuffer -------  String����ͬ�������������ִ�������
'nSize ----------  Long���������ĳ��ȡ����������ͬ���ؼ��������ʵ�ʳ�������

'Define Items in Database backup/restore utility
Public GroupName As String    '��ػ��Ĺ�������
Public UserName As String     '��ػ����û���
Public Pwd As String          '��ػ�������
Public optflag As Boolean     '��ʶ�Ƿ���ļ��з�����ػ�

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
  '**ģ   ��   ����Rstoredatabase
  '**��         �����ָ����ݿ�,���س�����Ϣ,�����ָ�,����""
  '**��         �ã�RStoredatabase   "�����ļ���","���ݿ���"
  '**����˵����
  '**                     DataBasePath     �ָ�������ݿ���Ŀ¼
  '**                     BackupNumber     �Ǵ��Ǹ����ݺŻָ�
  '**                     ReplaceExist     ָ���Ƿ񸲸��Ѿ����ڵ�����
  '**˵         ��������Microsoft   ActiveX   Data   Objects   2.x   Library
 
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
        If Len(Trim(DataBaseName)) = 0 Then    'ȡ�ñ������ݿ���߼�����
            DataBaseName = rs.Fields(0)
        End If

        If DataBasePath = "" Then
            sql = "select   filename   from   master..sysfiles"
            Set rs1 = TransactSQL(sql)
            sql = rs1.Fields(0)
            DataBasePath = Left(sql, InStrRev(sql, "\") - 1) '   InStrRev(SQL, "\")�����ַ�������,49
            rs1.Close
        End If
            
        If ReplaceExist = False Then
            sql = "select   1   from   master..sysdatabases     where   name='" & DataBaseName & "'"
            Set rs1 = TransactSQL(sql)
            If rs1.EOF = False Then
                MsgBox "���ݿ��Ѿ�����!", vbInformation, "��ʾ"
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
            MsgBox "��ϲ��,���ݿ�ָ��ɹ���!", , "�������"
        Else
            MsgBox "������˼,���ݿ�ָ�ʧ����!ע��:�����һ̨���ָ�����һ̨��,�빲���Żָ��ļ����ļ���!", , "Ŭ��ѧϰ"
        End If

ErrExit:
    Exit Sub
ErrHandel:
    MsgBox Err.Description, vbExclamation, "����"
    GoTo ErrExit:
    
End Sub

Sub Connet()
    If FrmServerBkup.OptCheck(0).Value = True Then
       strConn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=" & Trim(FrmServerBkup.CobServer.Text)
    Else
       strConn = "Provider=SQLOLEDB.1;Password='" & FrmServerBkup.TxtPassword.Text & "';Persist Security Info=False;User ID=" & Trim(FrmServerBkup.TxtName.Text) & ";Data Source=" & Trim(FrmServerBkup.CobServer.Text)
    End If
End Sub


Function GetIPAddress(ByVal HostName As String) As String '����������,���ر���Ip��ַ
Dim lpHost      As Long
Dim HOST        As HOSTENT     'HOSTENT��һ��Type��䶨���Ƕ�����ݽṹ��������
Dim dwIPAddr    As Long
Dim tmpIPAddr() As Byte        'Byte �����洢Ϊ��������,��Χ�� 0 �� 255 ֮��,�ڴ洢����������ʱ�����á�
Dim i           As Integer
Dim sIPAddr     As String

On Error GoTo ErrHandel

    lpHost = gethostbyname(HostName)  '��ģ��Modulel����Declare Function gethostbyname Lib "WSOCK32.DLL...."
    
    CopyMemory HOST, lpHost, Len(HOST)  '��ģ��Modulel����Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ......
    
    CopyMemory dwIPAddr, HOST.hAddrList, 4   '��ģ��Modulel����Declare Sub CopyMemory Lib "kernel32" Alias......
    
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
    
    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    
ErrExit:
    Exit Function
ErrHandel:
    MsgBox Err.Description, vbExclamation, "����"
    GetIPAddress = "���ܼ�⵽�˼����IP!"
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
    MsgBox "��ѯ����" & Err.Number & ": " & Err.Description
    iFlag = 2
    Resume TransactSQL_Exit
End Function


Sub HandleFile(optflag As Boolean, flag As Boolean)
Dim sql As String

    If optflag = True Then  '����Ƿ�����������,�����,��Ͽ�
       sql = "master..xp_cmdshell 'NET USE K: /DELETE /YES'"
       TransactSQL (sql)
       optflag = False
    End If
    If Len(Dir(App.Path & "\share.bat")) > 0 Then     '����ڱ�������,������صķ�����,���ѱ�
        Kill App.Path & "\share.bat"                  '���ļ����ڱ���ʱ,��Ϊ��Ҫ�����ļ���,��
        Open App.Path & "\share.bat" For Output As #1 '�Ե���źñ����ļ�ʱ��Ͽ�����.
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
        
        vstr = IIf(FrmServerBkup.Tag = "1", "����", "�ָ�")
        If Left(filename, 2) = "\\" Then
        '����ط������ﱸ�ݵ���ر��ݻ���
        '�ڱ��ط������ﱸ�ݵ���ر��ݻ���
            FilePath = Left(filename, InStrRev(filename, "\") - 1)
            FrmSaveOption.Label4.Caption = "�㽫���ݿ��" & vstr & "�ļ������" & FilePath & "��,��������̨���ĵ�¼��Ϣ!"
            FrmSaveOption.Show 1
            If FrmServerBkup.Tag = "1" Then
                FrmServerBkup.MousePointer = vbHourglass
            Else
                FrmOption.MousePointer = vbHourglass
            End If
            sql = "master..xp_cmdshell 'net use k: " & FilePath & " " & Pwd & " /user:" & GroupName & "\" & UserName & "'"
        Else
            '����ط������ﱸ�ݵ����ر��ݻ���
            If StrComp(Trim(FrmServerBkup.LabServerName), Trim(FrmServerBkup.LabComputer.Caption), 1) <> 0 Then
                FilePath = Left(filename, InStrRev(filename, "\") - 1)
                Open App.Path & "\share.bat" For Output As #1
                Print #1, "net share ShareFile=" & FilePath; ""
                Close #1
                Shell App.Path & "\share.bat"
                FrmSaveOption.Label4.Caption = "�㽫���ݿ��" & vstr & "�ļ������" & FilePath & "��,��������̨���ĵ�¼��Ϣ!"
                FrmSaveOption.Show 1
                If FrmServerBkup.Tag = "1" Then
                    FrmServerBkup.MousePointer = vbHourglass
                Else
                    FrmOption.MousePointer = vbHourglass
                End If
                sql = "master..xp_cmdshell 'net use k: \\" & FrmServerBkup.LabComputer & "\sharefile " & Pwd & " /user:" & GroupName & "\" & UserName & "'"
            Else
                '�ڱ��ط������ﱸ�ݵ����ر��ݻ���
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
    MsgBox "Error: " & Err.Description & "On Sub CheckServer!", vbExclamation, "����"
    GoTo ExitPoint
End Sub

