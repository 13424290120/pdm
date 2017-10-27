VERSION 5.00
Begin VB.Form FrmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register a new User 用户注册"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6780
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3645
      TabIndex        =   2
      Top             =   750
      Width           =   2295
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3645
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1470
      Width           =   2295
   End
   Begin VB.TextBox TxtPassword2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3645
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2190
      Width           =   2295
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName 用 户 名"
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
      Left            =   705
      MouseIcon       =   "FrmRegister.frx":08CA
      TabIndex        =   7
      Top             =   750
      Width           =   2670
   End
   Begin VB.Label LblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Pass Word 密    码"
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
      Left            =   735
      MouseIcon       =   "FrmRegister.frx":0BD4
      TabIndex        =   6
      Top             =   1455
      Width           =   2640
   End
   Begin VB.Label LblPassword2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pwd Twice确认密码"
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
      Left            =   555
      MouseIcon       =   "FrmRegister.frx":0EDE
      TabIndex        =   5
      Top             =   2190
      Width           =   2820
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3660
      Picture         =   "FrmRegister.frx":11E8
      Top             =   3270
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel 取 消"
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
      Left            =   4260
      MouseIcon       =   "FrmRegister.frx":1604
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3300
      Width           =   1440
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   1170
      Picture         =   "FrmRegister.frx":190E
      Top             =   3270
      Width           =   300
   End
   Begin VB.Label LblOK 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OK 确 定"
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
      Left            =   1665
      MouseIcon       =   "FrmRegister.frx":1D2A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3300
      Width           =   1020
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LblCancel_Click()
    Unload Me
End Sub

Private Sub lblOk_Click()
    
    '判断要编辑信息是否完整
    If Check = False Then
        Exit Sub
    End If
    
    With MyUsers
        .Name = TxtName.Text
        .Password = TxtPassword.Text
        
        
        '判断采购员ID是否已经存在
        If .In_DB(TxtName.Text) = True Then
            MsgBox "UserName exist, Please reset" + vbCrLf + "用户名已经存在，请重新设置", vbInformation, "System Info."
            TxtName.SetFocus
            TxtName.SelStart = 0
            TxtName.SelLength = Len(TxtName)
            Exit Sub
        Else
            .Insert '添加
            MsgBox "UserName is registered, Wait Administrator to distribute access right" + vbCrLf + "添加成功,请等待管理员给您分配权限", vbInformation, "System Info."
        End If
        
    End With
    Unload Me
End Sub


Private Function Check() As Boolean
    If Trim(TxtName) = "" Then
        MsgBox "Please input UserName" + vbCrLf + "请输入用户名", vbInformation, "System Info."
        TxtName.SetFocus
        Check = False
        Exit Function
    End If
    If Trim(TxtPassword) = "" Then
        MsgBox "Please input password" + vbCrLf + "请输入密码", vbInformation, "System Info."
        TxtPassword.SetFocus
        Check = False
        Exit Function
    End If
    If Trim(TxtPassword2) = "" Then
        MsgBox "Please input password 2" + vbCrLf + "请输入确认密码", vbInformation, "System Info."
        TxtPassword2.SetFocus
        Check = False
        Exit Function
    End If
    If Trim(TxtPassword2) <> Trim(TxtPassword) Then
        MsgBox "Password is not same in twice, please re-input" + vbCrLf + "两次输入密码不合，请重新输入", vbInformation, "System Info."
        TxtPassword.Text = ""
        TxtPassword2.Text = ""
        TxtPassword.SetFocus
        Check = False
        Exit Function
    End If
    If Len(Trim(TxtPassword2)) < 6 Then
        MsgBox "Password length need 6 letter/number at least,please re-input" + vbCrLf + "密码小于6位,请重新设置", vbInformation, "System Info."
        TxtPassword.Text = ""
        TxtPassword2.Text = ""
        TxtPassword.SetFocus
        Check = False
        Exit Function
    End If
    Check = True
End Function
