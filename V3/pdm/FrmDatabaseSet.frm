VERSION 5.00
Begin VB.Form FrmDatabaseSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dbase ConfigCheck 数据库配置检查"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDatabaseSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   12330
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdFindServer 
      Caption         =   "列出所有SQL服务器"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   4830
      Width           =   2040
   End
   Begin VB.ListBox LstServer 
      Height          =   3210
      Left            =   345
      TabIndex        =   8
      Top             =   1260
      Width           =   3135
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   7785
      TabIndex        =   5
      Top             =   1275
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All 全选"
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
      Left            =   7065
      TabIndex        =   4
      Top             =   555
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   3705
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1275
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   $"FrmDatabaseSet.frx":08CA
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
      Left            =   8055
      TabIndex        =   11
      Top             =   5670
      Width           =   4125
   End
   Begin VB.Label Label3 
      Caption         =   $"FrmDatabaseSet.frx":0912
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
      Left            =   570
      TabIndex        =   10
      Top             =   375
      Width           =   2730
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   4545
      Picture         =   "FrmDatabaseSet.frx":0941
      Top             =   5655
      Width           =   300
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmDatabaseSet.frx":0D5D
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
      Left            =   5145
      MouseIcon       =   "FrmDatabaseSet.frx":0D8C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5655
      Width           =   2835
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmDatabaseSet.frx":1096
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
      Left            =   9345
      TabIndex        =   6
      Top             =   375
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmDatabaseSet.frx":10B1
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
      Left            =   3705
      TabIndex        =   3
      Top             =   375
      Width           =   3195
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   9225
      Picture         =   "FrmDatabaseSet.frx":10E6
      Top             =   4875
      Width           =   300
   End
   Begin VB.Label LblCancel 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmDatabaseSet.frx":1502
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
      Left            =   9825
      MouseIcon       =   "FrmDatabaseSet.frx":1512
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4875
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   4545
      Picture         =   "FrmDatabaseSet.frx":181C
      Top             =   4875
      Width           =   300
   End
   Begin VB.Label LblOK 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmDatabaseSet.frx":1C38
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
      Left            =   5145
      MouseIcon       =   "FrmDatabaseSet.frx":1C51
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4815
      Width           =   1575
   End
End
Attribute VB_Name = "FrmDatabaseSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tabNum As Integer


Private Sub CmdFindServer_Click()
Dim oAppliction As New SQLDMO.Application
Dim oNameList As SQLDMO.NameList
Dim I As Integer

On Error Resume Next

Screen.MousePointer = vbHourglass


'列出可以使用的SQL Server实例
Set oNameList = oAppliction.ListAvailableSQLServers

    '如果有可以利用的SQL Server实例
    If oNameList.Count >= 1 Then
            '添加到下拉列表框中
            With LstServer
                .Clear
                For I = 1 To oNameList.Count
                    .AddItem oNameList.Item(I)
                Next I
            End With
    End If

'断开连接，并释放对象
oAppliction.Quit
Set oAppliction = Nothing


Screen.MousePointer = vbDefault
End Sub


Private Sub Check1_Click() 'Check1是全选按钮
Dim I As Integer
I = 0
If Check1.Value = 1 Then  '如果钩上选择框
    Do Until I = tabNum        '######这里根据实际存在的表个数填写,见 Private Sub Form_Load()
    List1.Selected(I) = True
    I = I + 1
    Loop
Else                      '如果不钩选择框
    Do Until I = tabNum        '######这里根据实际存在的表个数填写
    List1.Selected(I) = False
    I = I + 1
    Loop
End If
End Sub


Private Sub Form_Load()
'为List1装载条目,条目为所有使用的表名字
List1.AddItem "CNCSN"
List1.AddItem "CPCN"
List1.AddItem "FinsGd"
List1.AddItem "GlueSupplier"
List1.AddItem "PJNO"
List1.AddItem "RFSRFQ"
List1.AddItem "SER"
List1.AddItem "SglPrt"
List1.AddItem "Users"
List1.AddItem "BOMOrigData"
List1.AddItem "BOMSubmitApprove"
List1.AddItem "StdPrtLibStructr"
tabNum = 12
End Sub

Private Sub LblCancel_Click()
Unload Me
End Sub

Private Sub LblDelete_Click()
'弹出确认，提示用户
If MsgBox("All Info. in selected form will be deleted, Sure?" + vbCrLf + "是否清空所选表内信息?", vbYesNo + vbDefaultButton2, "Confirm Again 再次确认") = vbYes Then
   
'循环调用Delete函数初始化数据表信息
Dim I As Integer
I = 0
    With MyDatabaseSet
     Do Until I = tabNum         '######这里根据实际存在的表个数填写
        If List1.Selected(I) = True Then
         .Delete (List1.List(I))
        End If
      I = I + 1
     Loop
    End With
End If
End Sub

Private Sub lblOk_Click()   '检验表存在和状态
'清除先前的检测结果
List2.Clear

'循环调用In_DB检查选中的数据表是否存在
Dim I As Integer
I = 0
With MyDatabaseSet         '在Const模块中有Public MyDatabaseSet As New ClsDatabaseSet
Do Until I = tabNum        '######这里根据实际存在的表个数填写,见 Private Sub Form_Load()
If List1.Selected(I) = True Then
   If .In_DB(List1.List(I)) = True Then    '如果存在
     List2.AddItem "Form 表  " + List1.List(I) + " OK 配置正常"
   Else                                    '不存在
     List2.AddItem "Form 表" + List1.List(I) + "NOK 不存在"
   End If
End If
I = I + 1
Loop
End With
End Sub














