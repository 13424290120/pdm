VERSION 5.00
Begin VB.Form FrmCode 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "ERP"
   ClientHeight    =   1095
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   2100
      TabIndex        =   4
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox txtNewCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      MaxLength       =   12
      TabIndex        =   3
      Top             =   660
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "No"
      Height          =   375
      Left            =   3540
      TabIndex        =   1
      Top             =   660
      Width           =   885
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes"
      Height          =   375
      Left            =   2670
      TabIndex        =   0
      Top             =   660
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Would you like to copy the entire item & its childs?   Please input the new code for copied item!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "FrmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    FrmBOMAdmin.Enabled = True
    FrmBOMAdmin.mnuCopy = True
    FrmBOMAdmin.mnuPaste = False
    FrmBOMAdmin.mnuUncopy = False
    FrmBOMAdmin.IsCopy = False
    FrmBOMAdmin.CopyNodeSource = ""
    Unload Me
End Sub

Private Sub Command1_Click()
    FrmSglPrtEdit.Show 1, FrmCode
End Sub

Private Sub Form_Load()
'Load Skin & Format Control
LoadSkin Me
txtNewCode.Text = FrmBOMAdmin.CopyNodeSource
End Sub

Private Sub OKButton_Click()
    '新编码必须是12位数字
    If (Len(Trim(txtNewCode.Text)) <> 12 Or Not IsNumeric(txtNewCode.Text)) Then
        MsgBox ("The code MUST be made up of 12 numeric.")
    Else
        '新编码必须是从来没有出现过得
        Dim Conn As New ADODB.Connection
        Dim strSql As String
        Conn.Open connString
        Dim rs As New ADODB.Recordset
        Set rs.ActiveConnection = Conn
        strSql = "Select * from SglPrt where SglPrtIndex=" & Left(txtNewCode.Text, 11) & "0"
        rs.Open strSql, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            MsgBox "The new code is not existing, please first apply the code", vbCritical, "ERP"
            Exit Sub
        Else
            Call FrmBOMAdmin.CopyNodeData(Left(txtNewCode.Text, 11) & CStr(rs("SglPrtVer")))
            FrmBOMAdmin.Enabled = True
            FrmBOMAdmin.mnuCopy = True
            FrmBOMAdmin.mnuPaste = False
            FrmBOMAdmin.mnuUncopy = False
            FrmBOMAdmin.CopyNodeSource = ""
            FrmBOMAdmin.IsCopy = False
            Me.Hide
        End If
    End If
End Sub

Private Sub txtNewCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OKButton_Click
End Sub
