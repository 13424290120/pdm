VERSION 5.00
Begin VB.Form FrmChangePwd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password ��������"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChangePwd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6870
   StartUpPosition =   2  '��Ļ����
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
      Left            =   3750
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
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
      Left            =   3750
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox TxtName 
      Enabled         =   0   'False
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
      Left            =   3750
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label LblOK 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OK ȷ ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1770
      MouseIcon       =   "FrmChangePwd.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3270
      Width           =   1020
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   1275
      Picture         =   "FrmChangePwd.frx":0BD4
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel ȡ ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4365
      MouseIcon       =   "FrmChangePwd.frx":0FF0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3270
      Width           =   1440
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3765
      Picture         =   "FrmChangePwd.frx":12FA
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label LblPassword2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password ȷ������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   660
      MouseIcon       =   "FrmChangePwd.frx":1716
      TabIndex        =   5
      Top             =   2100
      Width           =   2820
   End
   Begin VB.Label LblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password ��    ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MouseIcon       =   "FrmChangePwd.frx":1A20
      TabIndex        =   4
      Top             =   1425
      Width           =   2640
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName �� �� ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   810
      MouseIcon       =   "FrmChangePwd.frx":1D2A
      TabIndex        =   3
      Top             =   720
      Width           =   2670
   End
End
Attribute VB_Name = "FrmChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function Check() As Boolean

    If Trim(TxtPassword) = "" Then
        MsgBox "Please input password" + vbCrLf + "����������", vbInformation, "System Info."
        TxtPassword.SetFocus
        Check = False
        Exit Function
    End If
    If Trim(TxtPassword2) = "" Then
        MsgBox "Please input password 2" + vbCrLf + "������ȷ������", vbInformation, "System Info."
        TxtPassword2.SetFocus
        Check = False
        Exit Function
    End If
    If Trim(TxtPassword2) <> Trim(TxtPassword) Then
        MsgBox "Password is not same in twice, please re-input" + vbCrLf + "�����������벻�ϣ�����������", vbInformation, "System Info."
        TxtPassword.Text = ""
        TxtPassword2.Text = ""
        TxtPassword.SetFocus
        Check = False
        Exit Function
    End If
    If Len(Trim(TxtPassword2)) < 6 Then
        MsgBox "Password length need 6 letter/number at least,please re-input" + vbCrLf + "����С��6λ,����������", vbInformation, "System Info."
        TxtPassword.Text = ""
        TxtPassword2.Text = ""
        TxtPassword.SetFocus
        Check = False
        Exit Function
    End If
    Check = True
End Function

Private Sub Form_Load()
    'Load Skin & Format Control
    'LoadSkin Me
    
    '''Call ResizeInit(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub LblCancel_Click()
    Unload Me
End Sub

Private Sub lblOk_Click()
    
    '�ж�Ҫ�༭��Ϣ�Ƿ�����
    If Check = False Then
        Exit Sub
    End If
    
    With MyUsers
        .Name = Trim(TxtName.Text)
        .Password = Trim(TxtPassword.Text)
        
        
     .UpdatePwd (Trim(TxtName.Text))                      '�洢�޸ĺ�ļ�¼����ģ��ClsUsers��MyUsers�����Update�ӹ���
     MsgBox "Successful Modify!" + vbCrLf + "�޸ĳɹ�!", vbInformation, "System Info."
        
    End With
    Unload Me
    FrmLogin.Show 0
End Sub


