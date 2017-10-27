Attribute VB_Name = "ini"
'声明写入ini文件的API函数
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFilenchame As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFilenchame As String) As Long
Sub main()
  'ServerName用来保存服务器的名字
  Dim ServerName As String
  Dim UserName As String
  Dim PasswordName As String
  '从Setup.ini中读取服务器的名字
  ServerName = GetKey(App.Path + "\Setup.ini", "Server")
  UserName = GetKey(App.Path + "\Setup.ini", "User")
  PasswordName = GetKey(App.Path + "\Setup.ini", "Password")
  '如果读取不成功，退出
  If ServerName = "" Then
    MsgBox "Setup.ini格式不正确，请重新设置"
    End
  End If


Server = ServerName
User = UserName
Password = PasswordName
'设置数据环境器参数
DataEnvironmentItem.Item.ConnectionString = "driver={SQL Server};server=" + Trim(Server) + ";uid=" + Trim(User) + ";pwd=" + Trim(Password) + ";database=ERP"
'显示主窗体
FrmManEngineer.Show

End Sub

'判断文件是否存在
Function FileExist(Fname As String) As Boolean
  On Local Error Resume Next
  FileExist = (Dir(Fname) <> "")
End Function
'读取ini文件的数据项值
Public Function GetKey(Tmp_File As String, Tmp_Key As String) As String
  Dim File As Long
  '分配文件句柄
  File = FreeFile
  '如果文件不存在则创建一个默认的Setup.ini文件
  If FileExist(Tmp_File) = False Then
    GetKey = ""
    Call WritePrivateProfileString("Setup Information", "Server Name ", " NtServer", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "UserName ", " ", App.Path + "\Setup.ini")
    Call WritePrivateProfileString("Setup Information", "PasswordName ", " ", App.Path + "\Setup.ini")
    Exit Function
  End If
  '读取数据项值
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
