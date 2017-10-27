Attribute VB_Name = "DtbsBkpRst"
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


Sub BackupDatabase(ByVal BackUpfileName, ByVal DataBaseName, Optional ByVal IsAddBackup As Boolean = False)
Dim ConcStr, sql, iReturn$

        sql = "backup   database   [" & DataBaseName & "]" & vbCrLf & "to   disk='" & BackUpfileName & "'" & vbCrLf & _
        "with   description='" & "zj-backup   at:" & Date & "(" & Time & ")'" & vbCrLf & IIf(IsAddBackup, "", ",init")
        
        TransactSQL sql

End Sub



  '*************************************************************************
  '**ģ   ��   ����Rstoredatabase
  '**��         �����ָ����ݿ�,���س�����Ϣ,�����ָ�,����""
  '**��         �ã�RStoredatabase   "�����ļ���","���ݿ���"
  '**����˵����
  '**                     DataBasePath     �ָ�������ݿ���Ŀ¼
  '**                     BackupNumber     �Ǵ��ĸ����ݺŻָ�
  '**                     ReplaceExist     ָ���Ƿ񸲸��Ѿ����ڵ�����
  '**˵         ��������Microsoft   ActiveX   Data   Objects   2.x   Library
 
  '*************************************************************************
Sub RestoreDatabase(ByVal BackUpfileName, ByVal DataBaseName _
, Optional ByVal DataBasePath = "", Optional ByVal BackupNumber = 1 _
, Optional ByVal ReplaceExist As Boolean = False)

Dim ConcStr, sql, I, MsgInfo
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
            MsgBox "��ϲ��,���ݿ�ָ��ɹ���!"
        Else
            MsgBox "������˼,���ݿ�ָ�ʧ����!ע��:�����һ̨���ָ�����һ̨��,�빲���Żָ��ļ����ļ���!"
        End If

ErrExit:
    Exit Sub
ErrHandel:
    MsgBox Err.Description, vbExclamation, "����"
    GoTo ErrExit:
    
End Sub

Sub Connet()
    If FrmServerBkup.OptCheck(0).Value = True Then        'OptCheck�Ǽ����windows��ʽ���ӻ���SQL�������û�����
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
Dim I           As Integer
Dim sIPAddr     As String

On Error GoTo ErrHandel

    lpHost = gethostbyname(HostName)  '��ģ��Modulel����Declare Function gethostbyname Lib "WSOCK32.DLL...."
    
    CopyMemory HOST, lpHost, Len(HOST)  '��ģ��Modulel����Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ......
    
    CopyMemory dwIPAddr, HOST.hAddrList, 4   '��ģ��Modulel����Declare Sub CopyMemory Lib "kernel32" Alias......
    
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
    
    For I = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(I) & "."
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
    strArray = Split(sql)          'Split��������һ���±���㿪ʼ��һά����. delimiter����������ԣ���ʹ�ÿո��ַ�(" ")��Ϊ�ָ�����
    Call Connet   'Call Connet�����Ǿ��������cn.Open strConn��strConn������(��windows��ʽ���ӻ���SQL�������û�����)
    cn.CommandTimeout = 0
    cn.Open strConn
    If StrComp(UCase$(strArray(0)), "SELECT", vbTextCompare) = 0 Then
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
Dim pId As Long, pHnd As Long ' �ֱ����� Process Id �� Process Handle ����
Dim sql As String
    'FrmSaveOption�����л᷵��optflag = True or False
    If optflag = True Then  '����Ƿ�����������,�����,��Ͽ�
       'net use ��windows�ڲ�������������ǽ�������͹�����Դ���ӻ�Ͽ�
       sql = "master..xp_cmdshell 'NET USE K: /DELETE /YES'"
       TransactSQL (sql)
       optflag = False
    End If
    If Dir("D:\Temp", vbDirectory) = "" Then MkDir "D:\Temp"           '�ж�һ��Ŀ¼�Ƿ���ڲ���������
    If Len(Dir("D:\Temp\Share.bat")) > 0 Then                     '����ڱ�������,������صķ������ѱ����ļ����ڱ���ʱ,
        Kill "D:\Temp\Share.bat"                                  '��Ϊ��Ҫ�����ļ���,���Ե���źñ����ļ�ʱ��Ͽ�����.
    End If
        Open "D:\Temp\Share.bat" For Output As #1
        Print #1, "net share ShareFile /delete /Y"   '/Y����Ϊִ�����delete����ʱ����ʾ����Y��N��ȷ��һ��
        Close #1
        'Shell "D:\Temp" & "\Share.bat"
        pId = Shell("D:\Temp\Share.bat", 0)        ' Shell ���� Process Id
        pHnd = OpenProcess(SYNCHRONIZE, 0, pId)    ' ȡ�� Process Handle
        If pHnd <> 0 Then
        Call WaitForSingleObject(pHnd, INFINITE)   ' INFINITE�����޵ȴ�ֱ�����̽���
        Call CloseHandle(pHnd)
        End If
        Kill "D:\Temp\Share.bat"
    flag = False
End Sub


Sub CheckServer(FileName As String, flag As Boolean)
On Error GoTo ErrHandle:
Dim sql As String
Dim FilePath As String
Dim vstr As String
Dim remot2local As Boolean
        
        vstr = IIf(FrmServerBkup.Tag = "1", "����", "�ָ�")     'IIf �������ݱ��ʽ��ֵ���������������е�����һ����IIf(expr, truepart, falsepart)
        If Left(FileName, 2) = "\\" Then
        '����ط������ﱸ�ݵ���ر��ݻ���
        '�ڱ��ط������ﱸ�ݵ���ر��ݻ���
            FilePath = Left(FileName, InStrRev(FileName, "\") - 1)
            FrmSaveOption.Label4.Caption = "�㽫���ݿ��" & vstr & "�ļ������" & FilePath & "��,��������̨���ĵ�¼��Ϣ!"
            FrmSaveOption.Show 1
            If FrmServerBkup.Tag = "1" Then
                FrmServerBkup.MousePointer = vbHourglass
            Else
                FrmOption.MousePointer = vbHourglass
            End If
            'net use ��windows�ڲ�������������ǽ�������͹�����Դ���ӻ�Ͽ�
            sql = "master..xp_cmdshell 'net use k: " & FilePath & " " & Pwd & " /user:" & GroupName & "\" & UserName & "'"
        Else
            '����ط������ﱸ�ݵ����ر��ݻ���
            If StrComp(Trim(FrmServerBkup.LabServerName), Trim(FrmServerBkup.LabComputer.Caption), 1) <> 0 Then          '1 ��ѡ����,ִ��һ������ԭ�ĵıȽϡ�
                FilePath = Left(FileName, InStrRev(FileName, "\") - 1)    'InStrRev��������һ���ַ�������һ���ַ����г��ֵ�λ�ã����ַ�����ĩβ����
                If Dir("D:\Temp", vbDirectory) = "" Then MkDir "D:\Temp"  '�ж�һ��Ŀ¼�Ƿ���ڲ���������
                Open "D:\Temp" & "\Share.bat" For Output As #1
                Print #1, "net share ShareFile=" & FilePath; ""          '��(D:\temp...)Ŀ¼�²���һ��bat�ļ���д�� net share ShareFile(�����ļ��й�����) = ·��
                Close #1
                Shell "D:\Temp" & "\Share.bat"
                FrmSaveOption.Label4.Caption = "�㽫���ݿ��" & vstr & "�ļ������" & FilePath & "��,��������̨���ĵ�¼��Ϣ!"
                FrmSaveOption.Show 1
                If FrmServerBkup.Tag = "1" Then
                    FrmServerBkup.MousePointer = vbHourglass
                Else
                    FrmOption.MousePointer = vbHourglass
                End If
                'net use ��windows�ڲ�������������ǽ�������͹�����Դ���ӻ�Ͽ�
                sql = "master..xp_cmdshell 'net use k: \\" & FrmServerBkup.LabComputer & "\ShareFile " & Pwd & " /user:" & GroupName & "\" & UserName & "'"
                remot2local = True
            Else
                '�ڱ��ط������ﱸ�ݵ����ر��ݻ���
                GoTo HandelPoint
            End If
        End If
        'FrmSaveOption�����л᷵��optflag = True or False
        If optflag = False Then flag = False: GoTo ExitPoint
        TransactSQL sql
        FileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
        'FileName = "k:\" & FileName   'ԭ�������ȡ��
        If remot2local Then
        FileName = "\\" & FrmServerBkup.LabComputer & "\ShareFile\ " & FileName
        Else
        FileName = "\\" & FilePath & "\" & FileName
        End If
        
HandelPoint:
        flag = True
ExitPoint:
    Exit Sub
ErrHandle:
    MsgBox "Error: " & Err.Description & "On Sub CheckServer!", vbExclamation, "����"
    GoTo ExitPoint
End Sub

