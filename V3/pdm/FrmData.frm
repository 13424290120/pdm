VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmServerBkup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ݿⱸ����ָ�"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "FrmData.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   7650
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "��    ��"
      Height          =   495
      Left            =   6060
      TabIndex        =   22
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "��   ��"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "��ԭ���ݿ�"
      Height          =   495
      Left            =   2100
      TabIndex        =   12
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton CmdBackUp 
      Caption         =   "�������ݿ�"
      Enabled         =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   11
      Top             =   5160
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Height          =   3675
      Left            =   180
      TabIndex        =   0
      Top             =   1320
      Width           =   7275
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   60
         Top             =   780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox CobDatabase 
         Height          =   300
         Left            =   1920
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox TxtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox TxtName 
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Top             =   2460
         Width           =   3615
      End
      Begin VB.OptionButton OptCheck 
         Caption         =   "ʹ��SQL�����֤"
         Height          =   315
         Index           =   1
         Left            =   660
         TabIndex        =   4
         Top             =   1920
         Value           =   -1  'True
         Width           =   5655
      End
      Begin VB.OptionButton OptCheck 
         Caption         =   "ʹ��Windows�����֤"
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   3
         Top             =   1560
         Width           =   5175
      End
      Begin VB.ComboBox CobServer 
         Height          =   300
         Left            =   2580
         TabIndex        =   2
         Text            =   "(local)"
         Top             =   1140
         Width           =   3435
      End
      Begin VB.Label LabServerIP 
         Height          =   315
         Left            =   5460
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "���ݿ������IP:"
         Height          =   255
         Left            =   4020
         TabIndex        =   20
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label LabServerName 
         Height          =   255
         Left            =   2340
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "���ݿ����������:"
         Height          =   255
         Left            =   660
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LabIp 
         Height          =   315
         Left            =   5460
         TabIndex        =   17
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "������IP��ַ:"
         Height          =   315
         Left            =   4020
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LabComputer 
         Height          =   315
         Left            =   2340
         TabIndex        =   15
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "���ؼ��������:"
         Height          =   255
         Left            =   660
         TabIndex        =   14
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "���ݿ�:"
         Height          =   315
         Left            =   660
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "����:"
         Height          =   315
         Left            =   660
         TabIndex        =   6
         Top             =   2880
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "�û���:"
         Height          =   255
         Left            =   660
         TabIndex        =   5
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "ѡ�������:"
         Height          =   255
         Left            =   660
         TabIndex        =   1
         Top             =   1140
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   2340
      Picture         =   "FrmData.frx":16542
      Top             =   0
      Width           =   5370
   End
End
Attribute VB_Name = "FrmServerBkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub CmdBackUp_Click()   '�������ݿ�
Dim flag As Boolean
Dim filename As String
Dim sql As String
On Error GoTo ErrHandle:

    CommonDialog1.CancelError = True    '�ж��Ƿ�ȡ������
    CommonDialog1.DialogTitle = "ѡ����Ҫ������ݿ�ĵط�"
    CommonDialog1.Filter = "���ݿ��ļ�(*.MsDat)|*.MsDat"
    CommonDialog1.ShowSave
    MousePointer = vbHourglass
    filename = CommonDialog1.filename
    FrmMain.Tag = "1"       '��ʶ�����������CheckServer����
    Call CheckServer(filename, flag)
    If flag = False Then MsgBox "���ݿ�û�б���!", vbExclamation, "����": GoTo ExitPoint
    Call BackupDatabase(filename, CobDatabase.Text)
    If iFlag = 1 Then
        MsgBox "��ϲ��,���ݿⱸ�ݳɹ���!", , "�������"
    Else
        MsgBox "������˼,���ݿⱸ��ʧ����!ע��:�����һ̨�����ݵ���һ̨��,�������¼����Ҫ��ű����ļ�����̨������ȷ����!", , "Ŭ��ѧϰ"
    End If
    Call HandleFile(optflag, flag)

ExitPoint:
    MousePointer = vbDefault
    Exit Sub
ErrHandle:
    GoTo ExitPoint
End Sub

Private Sub CmdExit_Click()
    End
End Sub

Private Sub CmdRestore_Click()   '��ԭ���ݿ�
Dim SQLServer As New SQLDMO.SQLServer
On Error GoTo ErrHandle:

    MousePointer = vbHourglass
     If OptCheck(1).Value = True Then
        If Len(TxtName.Text) > 0 Then
            SQLServer.Connect CobServer.Text, TxtName.Text, TxtPassword.Text
        Else
            MsgBox "���������ݿ��û���������!", vbInformation, "��ʾ"
            GoTo ErrExit
        End If
     Else
        SQLServer.Connect CobServer.Text
     End If
     FrmOption.Show 1
    
ErrExit:
    MousePointer = vbDefault
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbExclamation, "����"
    GoTo ErrExit
End Sub

Private Sub CobDatabase_Click()
    CmdBackUp.Enabled = True
End Sub

Sub CobDatabase_DropDown()   '���������е����ݿ�

Dim SQLServer As New SQLDMO.SQLServer
Dim i As Integer
On Error GoTo ErrHandle:

    MousePointer = vbHourglass
    CobDatabase.Clear
    If OptCheck(0).Value = True Then
        SQLServer.Connect CobServer.Text
    Else
        SQLServer.Connect CobServer.Text, TxtName.Text, TxtPassword.Text
    End If
    
'    SQLServer.AutoReConnect
    '�г����е����ݿ�
    For i = 1 To SQLServer.Databases.Count
        CobDatabase.AddItem SQLServer.Databases.Item(i).Name
    Next
    
    CmdBackUp.Enabled = False
    
ErrExit:
    MousePointer = vbDefault
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbExclamation, "����"
    GoTo ErrExit
End Sub


Sub LocalInfo()    'ȡ�ñ�������,�ͷ��ظ�����������Ip��ַ

Dim Name  As String, Length As Long
'************�ñ�������*****************************
    Length = 255
    Name = String(Length, 0)
    GetComputerName Name, Length
    Name = Left(Name, Length)
    LabComputer.Caption = Name
        
 '****************��������Ip��ַ************************
   LabIp.Caption = GetIPAddress(Name)

End Sub


Private Sub CobDatabase_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CobServer_Click()
    LabServerName.Caption = IIf(StrComp("(local)", Trim(CobServer.Text), 1) = 0, Trim(LabComputer.Caption), Trim(CobServer.Text))
    LabServerIP.Caption = GetIPAddress(Trim(LabServerName.Caption))
    CmdBackUp.Enabled = False
End Sub

Private Sub CobServer_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub Command1_Click()
    MsgBox "�������˲���,������ʵ�־������ڵ�SQL���ݿⱸ��!���ò���,��E-MAIL����ָ��!" + Chr(10) + Chr(13) _
           & "�����㲥���Ӵ�ѧ�������������" + Chr(10) + Chr(13) _
           & "���������:����" + Chr(10) + Chr(13) _
           & "E-MAIL:andylau700@163.com"
           
End Sub

Private Sub OptCheck_Click(Index As Integer)
    Select Case Index
        Case 0
            strConn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=" & Trim(CobServer.Text)
            TxtName.Enabled = False
            TxtPassword.Enabled = False
            Label2.Enabled = False
            Label3.Enabled = False
        Case 1
            strConn = "Provider=SQLOLEDB.1;Password='" & TxtPassword.Text & "';Persist Security Info=False;User ID=" & TxtName.Text & ";Data Source=" & Trim(CobServer.Text)
            TxtName.Enabled = True
            TxtPassword.Enabled = True
            Label2.Enabled = True
            Label3.Enabled = True
    End Select
End Sub

