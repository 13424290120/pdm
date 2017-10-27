VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Info. ÏµÍ³°ïÖú"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7695
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   $"FrmHelp.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1335
      TabIndex        =   1
      Top             =   2730
      Width           =   5145
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmHelp.frx":091F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   570
      TabIndex        =   0
      Top             =   645
      Width           =   6690
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

Dim OpnFilePathName As String
OpnFilePathName = App.Path + "\D&M PSS PDM Software Handbook.ppt"
OpnShllExcFile (OpnFilePathName)

End Sub



Private Sub Form_Load()
'Load Skin & Format Control
LoadSkin Me
End Sub
