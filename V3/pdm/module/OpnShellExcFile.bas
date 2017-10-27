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

'ShellExecute ������ָ���ļ�������һ��ĳ�����ļ���
'Long�������ʾ�ɹ������ʾʧ�ܡ�������GetLastError
'hwnd -- Long��ָ��һ�����ڵľ������ʱ��Windows�����б�Ҫ�ڴ����Լ���������ǰ��ʾһ����Ϣ��һ������Ϊ�����ھ��
'lpOperation ----  String���ļ����ͻ�Ĭ��ִ�г���ָ���ִ�Ϊ��Open������lpFlieָ���ĵ���ָ���ַ���Ϊ��Print������ӡlpFlieָ���ĵ���ָ���ַ���Ϊ��Explore�������lpFlieָ�����ļ����ļ��У�����ΪFind�ɲ���lpFlieָ�����ļ����µ��ļ����ļ��С�
'lpFile --- String���ù��������ӡ���һ�����������ļ�����ָ������·�����ļ�����
'lpParameters ---  String����lpszFlie�ǿ�ִ���ļ���������ִ��������ݸ�ִ�г���Ĳ���
'lpDirectory ----  String����ʹ�õ�����·������Ĭ��Ŀ¼��
'nShowCmd -- Long�������������ʾ��������ĳ���ֵ���ο�ShowWindow������nCmdShow������

'ShowWindow������nCmdShow��������
'SW_HIDE              ���ش��ڣ��״̬����һ������
'SW_MINIMIZE          ��С�����ڣ��״̬����һ������
'SW_RESTORE           ��ԭ���Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
'SW_SHOW              �õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
'SW_SHOWMAXIMIZED     ��󻯴��ڣ������伤��
'SW_SHOWMINIMIZED     ��С�����ڣ������伤��
'SW_SHOWMINNOACTIVE   ��С��һ�����ڣ�ͬʱ���ı�����
'SW_SHOWNA            �õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ����ı�����
'SW_SHOWNOACTIVATE    ������Ĵ�С��λ����ʾһ�����ڣ�ͬʱ���ı�����
'SW_SHOWNORMAL        ��SW_RESTORE��ͬ

Function StartDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function

Function OpnFileExist(Filename As String) As Boolean      '����λ��ĳ���ļ��Ƿ����
          On Error Resume Next
          
          Dim FileID  As Long
          FileID = FreeFile()        'FreeFile ����,����һ�� Integer,������һ���ɹ� Open ���ʹ�õ��ļ��š�
          
          Open Filename For Input As #FileID
          Close #FileID
          
          OpnFileExist = (Err.Number = 0)
          Err.Clear
End Function

Sub OpnShllExcFile(FilePathandName As String)   'ǰ�治����private���֣�������Sub....
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

