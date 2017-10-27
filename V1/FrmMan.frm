VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmMan 
   BackColor       =   &H8000000C&
   Caption         =   "D&M PSS PDM系统"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "FrmMan.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "登陆用户："
            TextSave        =   "登陆用户："
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   "2011-12-6"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "23:26"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11190
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":1A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":3516
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":3DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":46CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":4FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMan.frx":5886
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1349
      ButtonWidth     =   1614
      ButtonHeight    =   1349
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sys Admin"
            Key             =   "SystemAdmin"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eng.Sys"
            Key             =   "EngineeringSys"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pur.Sys"
            Key             =   "PurchasingSys"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sale.Sys"
            Key             =   "SaleLogisticSys"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Manu.Sys"
            Key             =   "ManufactureSys"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "HR.Sys"
            Key             =   "HumanResrcSys"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Qlty.Sys"
            Key             =   "QualityClientSys"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setting"
            Key             =   "Ini"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "End"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu LogIn 
      Caption         =   "LogIn"
   End
   Begin VB.Menu Action 
      Caption         =   "Action"
      NegotiatePosition=   1  'Left
      Begin VB.Menu SystemAdmin 
         Caption         =   "SystemAdmin"
      End
      Begin VB.Menu EngineeringSys 
         Caption         =   "EngineeringSys"
      End
      Begin VB.Menu PurchasingSys 
         Caption         =   "PurchasingSys"
      End
      Begin VB.Menu SaleLogisticSys 
         Caption         =   "SaleLogisticSys"
      End
      Begin VB.Menu ManufactureSys 
         Caption         =   "ManufactureSys"
      End
      Begin VB.Menu HumanResrcSys 
         Caption         =   "HumanResrcSys"
      End
      Begin VB.Menu QualityClientSys 
         Caption         =   "QualityClientSys"
      End
   End
   Begin VB.Menu ABout 
      Caption         =   "About"
   End
   Begin VB.Menu Exit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "FrmMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'hWnd         - the Window handle of the parent form.
'szApp        - the information you want to appear in the caption
'               of the About box, usually the application's title.
'szOtherStuff - any additional message you want to display.
'hIcon        - the handle to the icon you want displayed in the upper
'               left corner.
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
        (ByVal hWnd As Long, _
         ByVal szApp As String, _
         ByVal szOtherStuff As String, _
         ByVal hIcon As Long) As Long

Private Sub MDIForm_Load()
'Load Skin & Format Control
LoadSkin Me
    LogIn.Enabled = True
    
    
    Server = GetKey(App.Path + "\Setup.ini", "Server")
    DataBase = GetKey(App.Path + "\Setup.ini", "DataBase")
    DBUser = GetKey(App.Path + "\Setup.ini", "UserName")
    Password = GetKey(App.Path + "\Setup.ini", "Password")
  
  
    SystemAdmin.Enabled = False
    EngineeringSys.Enabled = False
    PurchasingSys.Enabled = False
    SaleLogisticSys.Enabled = False
    ManufactureSys.Enabled = False
    HumanResrcSys.Enabled = False
    QualityClientSys.Enabled = False
    
    Toolbar1.Buttons.Item(1).Enabled = False
    Toolbar1.Buttons.Item(3).Enabled = False
    Toolbar1.Buttons.Item(5).Enabled = False
    Toolbar1.Buttons.Item(7).Enabled = False
    Toolbar1.Buttons.Item(9).Enabled = False
    Toolbar1.Buttons.Item(11).Enabled = False
    Toolbar1.Buttons.Item(13).Enabled = False
    FrmLogin.Show
    FrmMan.Hide
    
    
End Sub

Private Sub EXIT_Click()
End
End Sub

Private Sub ABout_Click()
    Call ShellAbout(Me.hWnd, "D&M PSS PDM", "Copy right @ Jason Feng, D&M PSS APAC" + vbCrLf + "All Rights Reserved, July 2009", Me.Icon)
End Sub

Private Sub LogIn_Click()
    FrmLogin.Show 1
End Sub

Private Sub SystemAdmin_Click()
    Toolbar1.Enabled = False
    Action.Enabled = False
    FrmSystemAdmin.Show
    FrmMan.Hide
End Sub

Private Sub EngineeringSys_Click()
    Toolbar1.Enabled = False
    Action.Enabled = False
    FrmEngineeringSys.Show
    FrmMan.Hide
End Sub
Private Sub PurchasingSys_Click()
    Toolbar1.Enabled = False
    Action.Enabled = False
    FrmPurchasingSys.Show
    FrmMan.Hide
End Sub

Private Sub SaleLogisticSys_Click()
    'Action.Enabled = False
    'Toolbar1.Enabled = False
    'FrmManSales.Show
    FrmThanks.Show
End Sub

Private Sub ManufactureSys_Click()
    'Action.Enabled = False
    'Toolbar1.Enabled = False
    'FrmManStocks.Show
    FrmThanks.Show
End Sub

Private Sub HumanResrcSys_Click()
    'Toolbar1.Enabled = False
    'Action.Enabled = False
    'FrmManManpower.Show
    FrmThanks.Show
End Sub

Private Sub QualityClientSys_Click()
    'Toolbar1.Enabled = False
    'Action.Enabled = False
    'FrmManClient.Show
    FrmThanks.Show
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button) '判断图片菜单条中哪个按钮被点击
    Dim nfile As Long
    Select Case Button.Key
        Case Is = "SystemAdmin"
            Toolbar1.Enabled = False
            Action.Enabled = False
            FrmSystemAdmin.Show
            FrmMan.Hide
            
        Case Is = "EngineeringSys"
            Action.Enabled = False
            Toolbar1.Enabled = False
            FrmEngineeringSys.Show
            FrmMan.Hide
            
        Case Is = "PurchasingSys"
            Action.Enabled = False
            Toolbar1.Enabled = False
            FrmPurchasingSys.Show
            FrmMan.Hide
            
        Case Is = "SaleLogisticSys"
            'Action.Enabled = False
            'Toolbar1.Enabled = False
            'FrmManSales.Show
            FrmThanks.Show
            
        Case Is = "ManufactureSys"
            'Action.Enabled = False
            'Toolbar1.Enabled = False
            'FrmManStocks.Show
            FrmThanks.Show
            
        Case Is = "HumanResrcSys"
            'Action.Enabled = False
            'Toolbar1.Enabled = False
            'FrmManManpower.Show
            FrmThanks.Show
            
        Case Is = "QualityClientSys"
            'Action.Enabled = False
            'Toolbar1.Enabled = False
            'FrmManClient.Show
            FrmThanks.Show
            
        Case Is = "Help"
            FrmHelp.Show 1
            
            
        Case Is = "End"
        End
End Select

End Sub
