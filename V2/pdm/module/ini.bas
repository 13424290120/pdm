Attribute VB_Name = "ini"
'����д��ini�ļ���API����
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFilenchame As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFilenchame As String) As Long
Sub main()
  'ServerName�������������������
  Dim ServerName As String
  Dim UserName As String
  Dim PasswordName As String
  '��Setup.ini�ж�ȡ������������
  ServerName = GetKey(App.Path + "\Setup.ini", "Server")
  UserName = GetKey(App.Path + "\Setup.ini", "User")
  PasswordName = GetKey(App.Path + "\Setup.ini", "Password")
  '�����ȡ���ɹ����˳�
  If ServerName = "" Then
    MsgBox "Setup.ini��ʽ����ȷ������������"
    End
  End If


Server = ServerName
User = UserName
Password = PasswordName
'�������ݻ���������
DataEnvironmentItem.Item.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(User) + ";pwd=" + Trim(Password) + ";database=ERP"
'��ʾ������
FrmManEngineer.Show

End Sub

'�ж��ļ��Ƿ����
Function FileExist(Fname As String) As Boolean
  On Local Error Resume Next
  FileExist = (Dir(Fname) <> "")
End Function
'��ȡini�ļ���������ֵ
Public Function GetKey(Tmp_File As String, Tmp_Key As String) As String
  Dim File As Long
  '�����ļ����
  File = FreeFile
  '����ļ��������򴴽�һ��Ĭ�ϵ�Setup.ini�ļ�
  If FileExist(Tmp_File) = False Then
    GetKey = ""
    Call WritePrivateProfileString("Setup Information", "Server Name ", " NtServer", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "UserName ", " ", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "PasswordName ", " ", App.Path + "\Setup.ini")
    Exit Function
  End If
  '��ȡ������ֵ
  Open Tmp_File For Input As File
    Do While Not EOF(1)
      Line Input #File, buffer
      If Left(buffer, Len(Tmp_Key)) = Tmp_Key Then
        pos = InStr(buffer, "=")
        GetKey = Trim(Mid(buffer, pos + 1))
      End If
    Loop
  Close File
End Function
