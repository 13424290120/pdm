VERSION 5.00
Begin VB.Form FrmUsersEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System User Edit ϵͳ�û���Ϣ"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUsersEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8280
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox txtUserGroup 
      Height          =   390
      ItemData        =   "FrmUsersEdit.frx":08CA
      Left            =   4950
      List            =   "FrmUsersEdit.frx":08E9
      TabIndex        =   13
      Top             =   1500
      Width           =   2910
   End
   Begin VB.TextBox txtTitle 
      Height          =   390
      Left            =   4950
      TabIndex        =   11
      Top             =   660
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ȩ��"
      Height          =   945
      Left            =   4920
      TabIndex        =   8
      Top             =   1980
      Width           =   2925
      Begin VB.ComboBox cmbGroup 
         Height          =   390
         ItemData        =   "FrmUsersEdit.frx":0963
         Left            =   150
         List            =   "FrmUsersEdit.frx":0973
         TabIndex        =   9
         Text            =   "cmbGroup"
         Top             =   330
         Width           =   2715
      End
   End
   Begin VB.TextBox TxtPassword2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2340
      Width           =   2295
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1380
      Width           =   2295
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "���"
      Height          =   435
      Left            =   4950
      TabIndex        =   12
      Top             =   1200
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "ְ��"
      Height          =   435
      Left            =   4980
      TabIndex        =   10
      Top             =   360
      Width           =   2865
   End
   Begin VB.Label LblOK 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUsersEdit.frx":099C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      MouseIcon       =   "FrmUsersEdit.frx":09AB
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3300
      Width           =   675
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   720
      Picture         =   "FrmUsersEdit.frx":0CB5
      Top             =   3360
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUsersEdit.frx":10D1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3180
      MouseIcon       =   "FrmUsersEdit.frx":10E5
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3300
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2640
      Picture         =   "FrmUsersEdit.frx":13EF
      Top             =   3360
      Width           =   300
   End
   Begin VB.Label LblPwdAgain 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUsersEdit.frx":180B
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   300
      MouseIcon       =   "FrmUsersEdit.frx":182B
      TabIndex        =   2
      Top             =   2280
      Width           =   1755
   End
   Begin VB.Label LblPwd 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUsersEdit.frx":1B35
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   300
      MouseIcon       =   "FrmUsersEdit.frx":1B55
      TabIndex        =   1
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label LbUser 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUsersEdit.frx":1E5F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   420
      MouseIcon       =   "FrmUsersEdit.frx":1E78
      TabIndex        =   0
      Top             =   300
      Width           =   1455
   End
End
Attribute VB_Name = "FrmUsersEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modify As Boolean '����Modify�����û������ڽ��洫�ݹ�������Ϣ
Public OriName As String  '����OriName������޸��û���Ϣ���û���


Private Sub LblCancel_Click()
Unload Me
End Sub
Private Function Check() As Boolean
'����Ƿ���д�û���
  If Trim(TxtName) = "" Then
    MsgBox "Please input UserName" + vbCrLf + "�������û���", vbInformation, "System Info."
    TxtName.SetFocus
    Check = False
    Exit Function
  End If
  
  '����Ƿ���д����
If Trim(TxtPassword) = "" Then
    MsgBox "Please input PassWord" + vbCrLf + "����������", vbInformation, "System Info."
    TxtPassword.SetFocus
    Check = False
    Exit Function
   End If
   
   '����Ƿ���дȷ������
If Trim(TxtPassword2) = "" Then
    MsgBox "Please input PassWord again" + vbCrLf + "������ȷ������", vbInformation, "System Info."
    TxtPassword2.SetFocus
    Check = False
    Exit Function
   End If
   
   '���������ȷ�������Ƿ���ͬ
If Trim(TxtPassword2) <> Trim(TxtPassword) Then
    MsgBox "Please Confirm PassWord, check and re-input" + vbCrLf + "�����������벻�ϣ�����������", vbInformation, "System Info."
    TxtPassword.Text = ""
    TxtPassword2.Text = ""
    TxtPassword.SetFocus
    Check = False
    Exit Function
   End If
   
   '�������λ���Ƿ����6λ
If Len(Trim(TxtPassword2)) < 6 Then
    MsgBox "PassWord Length needs 6 Char.at least" + vbCrLf + "����С��6λ,����������", vbInformation, "System Info."
    TxtPassword.Text = ""
    TxtPassword2.Text = ""
    TxtPassword.SetFocus
    Check = False
    Exit Function
   End If
   
   If Trim(cmbGroup.Text) = "cmbGroup" Then
    MsgBox "Please choose the user privillege.", vbInformation, "System Info"
    cmbGroup.SetFocus
    Check = False
    Exit Function
  End If
   
   '���������ⶼͨ������CheckΪ��
   Check = True
End Function
Private Sub lblOk_Click()
    
   '�ж�Ҫ�༭��Ϣ�Ƿ�����
   If Check = False Then 'Check���������涨�壬����û�������ȺϷ���
   '��������������ò����Ϲ涨����������
    Exit Sub
   End If
   
   
With MyUsers
'����ģ��ClsUsers��MyUsers����Ĳ�����ֵ
.Name = TxtName.Text
.Password = TxtPassword.Text
.UserGroup = Trim(txtUserGroup.Text)
.UserTitle = Trim(txtTitle.Text)
.GrantGroup = Trim(cmbGroup.Text)

    '�жϲ�������ӻ����޸�
    If Modify = False Then        '�ж�Ϊ��Ӳ���
    '�жϸ��û����Ƿ��Ѿ�����ʹ��
                If .In_DB(TxtName.Text) = True Then '����Ѿ����ڣ���ģ��ClsUsers��MyUsers�����In_DB����
                   MsgBox "User is Existing, Please Reset" + vbCrLf + "�û��Ѿ����ڣ�����������", vbInformation, "System Info."
                   TxtName.SetFocus '������������Ǽ���������������ַ�
                   TxtName.SelStart = 0
                   TxtName.SelLength = Len(TxtName)
                   Exit Sub
                Else                         '���������
                   .Insert                   'ִ����Ӳ���
                    MsgBox "Successful Add!" + vbCrLf + "��ӳɹ�!", vbInformation, "System Info."
                End If
    Else  '�ж�Ϊ�޸Ĳ���
     .Update (OriName)                      '�洢�޸ĺ�ļ�¼����ģ��ClsUsers��MyUsers�����Update�ӹ���
     MsgBox "Successful Modify!" + vbCrLf + "�޸ĳɹ�!", vbInformation, "System Info."
    End If
End With
Unload Me

End Sub

