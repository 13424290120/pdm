VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Form and control Resizer
'Mike Du Bois
'mikedb@aug.com
'June 1 1999
'Version 1
'The standard textbox does not have a hdc property
'we are using the hdc from the form to read text metrics
'Every control's font is then set to the same font as the form
'You can use the FontChanged event to reset any controls that
'need a different font attribute for example to bold or italics after the form
'and controls are resized
'
'If the form will be larger than the screen then an error occures
'and the form is NOT resized and an error message is returned.

Event FontChanged()
'Use this event
'To reset the fonts on any controls where you want to have control
'rather than the user


'The GetOldFontInfo function reads the current height and
'average character width using the hdc of the form
'Returns "OK" on success or an error msg.
'
'The SetForm loops through all the controls and
'resets the height, width and length of the controls
'Retunrs OK or an error if the form will not fit on the screen

'The LeftMargin var is used to set the left margin on the form
'So that when resetting the controls the left margins remain constant with
'refernece to the container.
'If the default value is used then the distance from the left of the
'form to all controls will change in proportion to the change in font size
'By setting this value you can keep the distance for the first row
'of controls constant.  The rest of the controls will then be moved
'to accommdate the change in the font size.


Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type
'Declare API Call
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

'Private vars
Private m_OldFontHeight As Long
Private m_OldFontAvWidth As Long
Private m_OldFontName As String
Private m_oldFontBold As Boolean
Private m_oldFontItalic As Boolean
Private m_OldFontUnderlined As Boolean
Private m_OldFontStrikeThru As Boolean
Private m_OldFontSize As Long
Private m_LeftMargin As Long
Private m_OldTextmetrics As TEXTMETRIC
Private m_HasBeenInitialized As Boolean
Private m_DefaultWindowState As Integer
Private m_DefaultFontHeight As Long
Private m_DefaultFontAvWidth As Long
Private m_DefaultFontName As String
Private m_DefaultFontBold As Boolean
Private m_DefaultFontItalic As Boolean
Private m_DefaultFontUnderlined As Boolean
Private m_DefaultFontStrikeThru As Boolean
Private m_DefaultFontSize As Long
Private m_DefaultFormWidth As Long
Private m_defaultFormHeight As Long
'Constants
Private Const errmsg = "An error occured when trying to change the font"




Public Function GetOldFontInfo(frm As Form) As String
'Returns OK on success
'An error message on failure

Dim Textmetrics As TEXTMETRIC
Dim hdc As Long
Dim msg As String
'Get the HDC
hdc = frm.hdc

'Check to see if class was initialized
If Not m_HasBeenInitialized Then 'If not then init
   msg = Init(frm)
   If msg <> "OK" Then
    GetOldFontInfo = msg
    Exit Function
   End If
End If


'Get the current font values
With frm
    m_OldFontName = .FontName
    m_oldFontBold = .FontBold
    m_oldFontItalic = .FontItalic
    m_OldFontUnderlined = .FontUnderline
    m_OldFontStrikeThru = .FontStrikethru
    m_OldFontSize = .FontSize
End With

'Get the text height and average width of the font
If GetTextMetrics(hdc, Textmetrics) <> 0 Then 'success
    m_OldFontHeight = Textmetrics.tmHeight
    m_OldFontAvWidth = Textmetrics.tmAveCharWidth
    m_OldTextmetrics = Textmetrics
    GetOldFontInfo = "OK"
    
Else
    GetOldFontInfo = errmsg
    
End If

'Get the default values

End Function
Public Function SetForm(ByRef frm As Form, Optional ByVal LeftMargin As Long = 0) As String
Dim Textmetrics As TEXTMETRIC
Dim NewFontHeight As Long
Dim NewFontAvWidth As Long
Dim RatioOfChangeHeight As Double
Dim RatioOfChangeWidth As Double
Dim ctl As Control
Dim Max As Long
Dim msg As String
'Returns OK on success
'On error returns err msg


On Error GoTo SetFormErr
'Get the new font info
m_LeftMargin = LeftMargin

'If the default window state is maximized then
'we can not change the size of the form
'But we may be able to change the size of the controls

If m_DefaultWindowState = vbMaximized Or frm.WindowState = vbMaximized Then
    msg = FormIsMaximized(frm)
    SetForm = msg
    Exit Function
End If

'the default window state is not maximized so we can resize the form as well as the
'controls

If GetTextMetrics(frm.hdc, Textmetrics) <> 0 Then 'success
    With Textmetrics
        NewFontHeight = .tmHeight
        NewFontAvWidth = .tmAveCharWidth
    End With
    
    
'compute the ratio of change
    RatioOfChangeHeight = NewFontHeight / m_OldFontHeight
    RatioOfChangeWidth = NewFontAvWidth / m_OldFontAvWidth
    
'Check to see if we have room to resize
    If Not Resizable(frm, RatioOfChangeHeight, RatioOfChangeWidth) Then
        SetForm = errmsg & vbLf & "The font you chose is too large"
        Exit Function
    Else
 'resize and center form
        
            frm.Height = frm.Height * RatioOfChangeHeight
            frm.Width = frm.Width * RatioOfChangeWidth
            'Center the form
            frm.Left = (Screen.Width - frm.Width) / 2
            frm.Top = (Screen.Height - frm.Height) / 2
    
    End If

  'Resize the controls
     ResizeControls frm, RatioOfChangeHeight, RatioOfChangeWidth, NewFontHeight
  'all is well so
     SetForm = "OK"
    
        
  'Fire fontchanged event
        RaiseEvent FontChanged
        Exit Function
Else 'We have an error
    SetForm = errmsg
    Exit Function
End If


SetFormErr:

msg = "Error number: " & CStr(Err.Number) & Err.Description
ReSetForm frm
SetForm = errmsg & vbLf & msg




End Function

Public Sub ReSetForm(frm As Form)
'Reset to form to the old values
'Note if we get here none of the controls have been changed
If Not frm.WindowState = vbMaximized Then

With frm
    .FontName = m_DefaultFontName
     .FontBold = m_DefaultFontBold
     .FontItalic = m_DefaultFontItalic
     .FontSize = m_DefaultFontSize
     .FontStrikethru = m_DefaultFontStrikeThru
     .FontUnderline = m_DefaultFontUnderlined
     .Height = m_defaultFormHeight
     .Width = m_DefaultFormWidth
End With

    SetForm frm, m_LeftMargin
    
End If

End Sub

Private Function Resizable(frm As Form, RatioOfHeight As Double, RatioOfWidth As Double) As Boolean
Dim ctl As Control
Dim MaxLength As Double
Dim MaxHeight As Double
Dim TempLength As Double
Dim TempHeight As Double

On Error Resume Next
    For Each ctl In frm.Controls
        TempLength = (ctl.Left * RatioOfWidth) + (ctl.Width * RatioOfWidth)
    'Get the maxlenght
        If TempLength > MaxLength Then MaxLength = TempLength
    'if wider than screen then we can not resize
    'return false and exit
        If MaxLength > Screen.Width Then
            Resizable = False
            Exit Function
        End If
     'Get the maxheight
        TempHeight = (ctl.Top * RatioOfHeight) + (ctl.Height * RatioOfHeight)
        If TempHeight > MaxHeight Then MaxHeight = TempHeight
      'If taller than screen then we can not resixe
      'return false and exit
         If MaxHeight > Screen.Height Then
            Resizable = False
            Exit Function
          End If
      Next
     
      
      Resizable = True
          
     

End Function

Private Sub Class_Initialize()
m_LeftMargin = 0 'default value
m_HasBeenInitialized = False

End Sub

Public Function Init(frm As Form) As String
 On Error GoTo InitErr
 m_HasBeenInitialized = True
    With frm
        m_DefaultFontName = .FontName
        m_DefaultFontBold = .FontBold
        m_DefaultFontItalic = .FontItalic
        m_DefaultFontSize = .FontSize
        m_DefaultFontStrikeThru = .FontStrikethru
        m_DefaultFontUnderlined = .FontUnderline
        m_defaultFormHeight = .Height
        m_DefaultFormWidth = .Width
        m_DefaultWindowState = .WindowState
    End With
Init = "OK"
GetOldFontInfo frm
Exit Function

InitErr:

Init = CStr(Err.Number) & "   " & Err.Description

End Function

Private Function FormIsMaximized(frm As Form) As String
Dim Textmetrics As TEXTMETRIC
Dim NewFontHeight As Long
Dim NewFontAvWidth As Long
Dim RatioOfChangeHeight As Double
Dim RatioOfChangeWidth As Double
Dim ctl As Control
Dim Max As Long
'Returns OK on success
'Returns err msg on failure
On Error GoTo FormIsMaxedErr
If GetTextMetrics(frm.hdc, Textmetrics) <> 0 Then 'success
    With Textmetrics
        NewFontHeight = .tmHeight
        NewFontAvWidth = .tmAveCharWidth
    End With
    
    
'compute the ratio of change
    RatioOfChangeHeight = NewFontHeight / m_OldFontHeight
    RatioOfChangeWidth = NewFontAvWidth / m_OldFontAvWidth
    
    If Not Resizable(frm, RatioOfChangeHeight, RatioOfChangeWidth) Then 'we can not make any changes
        FormIsMaximized = errmsg & vbLf & "The font is too large"
        Exit Function
        
    Else 'We can resize the controls
          ResizeControls frm, RatioOfChangeHeight, RatioOfChangeWidth, NewFontHeight
    End If
    
    FormIsMaximized = "OK"
    Exit Function
    'and we are out of here

Else 'call to gettextmetrics failed

    FormIsMaximized = "The call to GetTextMetrics failed"
    Exit Function
End If

Exit Function
FormIsMaxedErr:
FormIsMaximized = CStr(Err.Number) & "  " & Err.Description

End Function

Private Function ResizeControls(frm As Form, _
RatioOfChangeHeight As Double, _
RatioOfChangeWidth As Double, _
NewFontHeight As Long)
Dim ctl As Control
  
    On Error Resume Next 'some controls may not have a font property
'Note that when using typeof you can not use a Select statement
'and must use IF..Then logic
        For Each ctl In frm.Controls
        'change the font information
        'some controls may not have a font property
        'so resume next on error
               With frm
                    ctl.FontSize = .FontSize
                    ctl.FontBold = .FontBold
                    ctl.FontName = .FontName
                    ctl.FontItalic = .FontItalic
                    ctl.FontUnderline = .FontUnderline
                    ctl.fontstrickethru = .FontStrikethru
                End With
               
              'set the controls height and width
              
              'You could add logic here to undo some of global
              'changes for a class of control. For example set the
              'font to bold for all commandbuttons
              'This would make the class less reusable unless
              'this is a programming convention that you always
              'use.  It would be better to make such changes in
              'the font changed event.
              
              'Set command buttons
              If TypeOf ctl Is CommandButton Then
                   ctl.Height = ctl.Height * RatioOfChangeHeight
                   ctl.Width = ctl.Width * RatioOfChangeWidth
                   ctl.Top = ctl.Top * RatioOfChangeHeight
              End If
              
              
              'set listbox
              If TypeOf ctl Is ListBox Then
                   ctl.Height = ctl.eight * RatioOfChangeHeight
                   ctl.Width = ctl.Width * RatioOfChangeWidth
                   ctl.Top = ctl.Top * RatioOfChangeHeight
              End If
              
              'set combo box
              If TypeOf ctl Is ComboBox Then
                  ctl.Height = ctl.Height * RatioOfChangeHeight
                  ctl.Width = ctl.Width * RatioOfChangeWidth
                  ctl.Top = ctl.Top * RatioOfChangeHeight
              End If
              
              'set option box
              If TypeOf ctl Is OptionButton Then
                   ctl.Height = ctl.Height * RatioOfChangeHeight
                   ctl.Width = ctl.Width * RatioOfChangeWidth
                   ctl.Top = ctl.Top * RatioOfChangeHeight
              End If
              
              'set check box
              If TypeOf ctl Is CheckBox Then
                   ctl.Height = ctl.Height * RatioOfChangeHeight
                   ctl.Width = ctl.Width * RatioOfChangeWidth
                   ctl.Top = ctl.Top * RatioOfChangeHeight
              End If
              'Set picture box
              If TypeOf ctl Is PictureBox Then
                    ctl.Height = ctl.Height * RatioOfChangeHeight
                    ctl.Width = ctl.Width * RatioOfChangeWidth
                    ctl.Top = ctl.Top * RatioOfChangeHeight
              End If
              
              ' set Frames
               If TypeOf ctl Is Frame Then
                  
                  
                    ctl.Height = ctl.Height * RatioOfChangeHeight
                    ctl.Width = ctl.Width * RatioOfChangeWidth
                    ctl.Top = ctl.Top * RatioOfChangeHeight
               End If
               
              'Set text box
              If TypeOf ctl Is TextBox Then
                                                                            
                ctl.Height = NewFontHeight
                ctl.Width = ctl.Width * RatioOfChangeWidth
                ctl.Top = ctl.Top * RatioOfChangeHeight
               End If
               
               'Set labels
               
               If TypeOf ctl Is Label Then
                    
                    ctl.Width = ctl.Width * RatioOfChangeWidth
                    ctl.Height = ctl.Height * RatioOfChangeHeight
                    ctl.Top = ctl.Top * RatioOfChangeHeight
               End If
               
               'Set Masked edit contol
               'comment out these lines in there is no maskedbox
'***********************************************************************
               If TypeOf ctl Is MaskEdBox Then
                    
                    ctl.Width = ctl.Width * RatioOfChangeWidth
                    ctl.Height = ctl.Height * RatioOfChangeHeight
                    ctl.Top = ctl.Top * RatioOfChangeHeight
               
               End If
'*************************************************************************
        '
        'Set the left margin for controls
        '
        If ctl.Left <> m_LeftMargin Then
            ctl.Left = ctl.Left * RatioOfChangeWidth
        Else
            'LEFT OF CONTROL IS UNCHANGED
        End If
        Next
End Function
