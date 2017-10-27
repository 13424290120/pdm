VERSION 5.00
Begin VB.Form FrmUsersEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System User Edit 系统用户信息"
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
   StartUpPosition =   2  '屏幕中心
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
      Caption         =   "权限"
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
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2340
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
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1380
      Width           =   2295
   End
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
      Left            =   2340
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "组别"
      Height          =   435
      Left            =   4950
      TabIndex        =   12
      Top             =   1200
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "职务"
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
Public Modify As Boolean '变量Modify保存用户管理窗口界面传递过来的信息
Public OriName As String  '变量OriName保存待修改用户信息的用户名


Private Sub LblCancel_Click()
Unload Me
End Sub
Private Function Check() As Boolean
'检查是否填写用户名
  If Trim(TxtName) = "" Then
    MsgBox "Please input UserName" + vbCrLf + "请输入用户名", vbInformation, "System Info."
    TxtName.SetFocus
    Check = False
    Exit Function
  End If
  
  '检查是否填写密码
If Trim(TxtPassword) = "" Then
    MsgBox "Please input PassWord" + vbCrLf + "请输入密码", vbInformation, "System Info."
    TxtPassword.SetFocus
    Check = False
    Exit Function
   End If
   
   '检查是否填写确认密码
If Trim(TxtPassword2) = "" Then
    MsgBox "Please input PassWord again" + vbCrLf + "请输入确认密码", vbInformation, "System Info."
    TxtPassword2.SetFocus
    Check = False
    Exit Function
   End If
   
   '检查密码与确认密码是否相同
If Trim(TxtPassword2) <> Trim(TxtPassword) Then
    MsgBox "Please Confirm PassWord, check and re-input" + vbCrLf + "两次输入密码不合，请重新输入", vbInformation, "System Info."
    TxtPassword.Text = ""
    TxtPassword2.Text = ""
    TxtPassword.SetFocus
    Check = False
    Exit Function
   End If
   
   '检查密码位数是否大于6位
If Len(Trim(TxtPassword2)) < 6 Then
    MsgBox "PassWord Length needs 6 Char.at least" + vbCrLf + "密码小于6位,请重新设置", vbInformation, "System Info."
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
   
   '如果上述检测都通过则标记Check为真
   Check = True
End Function
Private Sub lblOk_Click()
    
   '判断要编辑信息是否完整
   If Check = False Then 'Check函数在上面定义，检查用户名密码等合法性
   '如果不完整或设置不符合规定则跳出函数
    Exit Sub
   End If
   
   
With MyUsers
'给类模块ClsUsers的MyUsers对象的参数赋值
.Name = TxtName.Text
.Password = TxtPassword.Text
.UserGroup = Trim(txtUserGroup.Text)
.UserTitle = Trim(txtTitle.Text)
.GrantGroup = Trim(cmbGroup.Text)

    '判断操作是添加还是修改
    If Modify = False Then        '判断为添加操作
    '判断该用户名是否已经有人使用
                If .In_DB(TxtName.Text) = True Then '如果已经存在，类模块ClsUsers的MyUsers对象的In_DB函数
                   MsgBox "User is Existing, Please Reset" + vbCrLf + "用户已经存在，请重新设置", vbInformation, "System Info."
                   TxtName.SetFocus '以下三条语句是加亮输入框中所有字符
                   TxtName.SelStart = 0
                   TxtName.SelLength = Len(TxtName)
                   Exit Sub
                Else                         '如果不存在
                   .Insert                   '执行添加操作
                    MsgBox "Successful Add!" + vbCrLf + "添加成功!", vbInformation, "System Info."
                End If
    Else  '判断为修改操作
     .Update (OriName)                      '存储修改后的纪录，类模块ClsUsers的MyUsers对象的Update子过程
     MsgBox "Successful Modify!" + vbCrLf + "修改成功!", vbInformation, "System Info."
    End If
End With
Unload Me

End Sub

