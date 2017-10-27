VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const PS_SOLID = 0

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_NOPREFIX = &H800
Private Const DT_SINGLELINE = 32
Private Const DT_CALCRECT = 1024
Private Const DT_END_ELLIPSIS = 32768
Private Const DT_MODIFYSTRING = 65536
Private Const DT_WORD_ELLIPSIS = 262144

Private m_oFG As MSFlexGrid
Private m_oDest As Object

Private m_lMarginLeft As Long
Private m_lMarginTop As Long
Private m_lMarginRight As Long
Private m_lMarginBottom As Long
Private m_lDataLeft As Long
Private m_lDataTop As Long
Private m_lDestWidth As Long
Private m_lDestHeight As Long
Private m_dXFactor As Double
Private m_dYFactor As Double
Private m_bPrinter As Boolean
Private m_lDrawWidth As Long
Private m_fntHeader As StdFont
Private m_fntFooter As StdFont
Private m_lPageCnt As Long
Private m_bPrintBackColor As Boolean
Private m_bPrintGridLines As Boolean
Private m_bCountOnly As Boolean
Private m_lPrinterOrientation As Long
Private m_lPageHeigth As Long

Public Enum pgOrientation
  pgoPortrait = vbPRORPortrait
  pgoLandscape = vbPRORLandscape
End Enum

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Event NewPage(PageNumber As Long)

Public Property Let PrinterOrientation(ByVal NewOrientation As pgOrientation)
  m_lPrinterOrientation = NewOrientation
  Printer.Orientation = NewOrientation
End Property

Public Property Get PrinterOrientation() As pgOrientation
  PrinterOrientation = m_lPrinterOrientation
End Property

Public Property Let MarginLeft(ByVal NewMarginLeft As Long)
  m_lMarginLeft = NewMarginLeft
End Property

Public Property Get MarginLeft() As Long
  MarginLeft = m_lMarginLeft
End Property

Public Property Let MarginRight(ByVal NewMarginRight As Long)
  m_lMarginRight = NewMarginRight
End Property

Public Property Get MarginRight() As Long
  MarginRight = m_lMarginRight
End Property

Public Property Let MarginTop(ByVal NewMarginTop As Long)
  m_lMarginTop = NewMarginTop
End Property

Public Property Get MarginTop() As Long
  MarginTop = m_lMarginTop
End Property

Public Property Let MarginBottom(ByVal NewMarginBottom As Long)
  m_lMarginBottom = NewMarginBottom
End Property

Public Property Get MarginBottom() As Long
  MarginBottom = m_lMarginBottom
End Property

Public Property Let PrintBackColor(ByVal NewPrintBackColor As Boolean)
  m_bPrintBackColor = NewPrintBackColor
End Property

Public Property Get PrintBackColor() As Boolean
  PrintBackColor = m_bPrintBackColor
End Property

Public Property Let PrintGridLines(ByVal NewPrintGridLines As Boolean)
  m_bPrintGridLines = NewPrintGridLines
End Property

Public Property Get PrintGridLines() As Boolean
  PrintGridLines = m_bPrintGridLines
End Property

Public Property Get NumberOfPages() As Long
  m_bCountOnly = True
  m_lPageCnt = -1
  PrintGrid
  NumberOfPages = m_lPageCnt
  m_bCountOnly = False
End Property

Public Property Set Source(NewGrid As MSFlexGrid)
  Set m_oFG = NewGrid
End Property

Public Property Set Destination(NewDest As Object)
  If TypeOf NewDest Is PictureBox Then
    Set m_oDest = NewDest
    m_bPrinter = False
    m_dYFactor = 1
    m_dXFactor = 1
  ElseIf TypeOf NewDest Is Printer Then
    Set m_oDest = NewDest
    m_bPrinter = True
    m_dXFactor = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
    m_dYFactor = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
  Else
    Set m_oDest = Nothing
    m_bPrinter = False
    Exit Property
  End If
  
End Property

'Public Property Let HeaderFontName(ByVal NewFontName As String)
'  m_fntHeader.Name = NewFontName
'End Property
'
'Public Property Let HeaderFontSize(ByVal NewFontSize As Double)
'  m_fntHeader.Size = NewFontSize
'End Property
'
'Public Property Let HeaderFontBold(ByVal NewFontBold As Boolean)
'  m_fntHeader.Bold = NewFontBold
'End Property
'
'Public Property Let HeaderText(ByVal NewHeader As String)
'
'End Property
'
'Public Property Let FooterFontName(ByVal NewFontName As String)
'  m_fntFooter.Name = NewFontName
'End Property
'
'Public Property Let FooterFontSize(ByVal NewFontSize As Double)
'  m_fntFooter.Size = NewFontSize
'End Property
'
'Public Property Let FooterFontBold(ByVal NewFontBold As Boolean)
'  m_fntFooter.Bold = NewFontBold
'End Property
'
'Public Property Let FooterText(ByVal NewHeader As String)
'
'End Property

Public Function PrintGrid() As Long
  Dim lPrevScaleMode As Long
  Dim bPrevRedraw As Boolean
  Dim lRow As Long, lCol As Long
  
  If m_oDest Is Nothing Then Exit Function
  If m_oFG Is Nothing Then Exit Function

  If Not m_bCountOnly Then
    If m_bPrinter Then
      Printer.Print "":
    Else
      m_oDest.Visible = False
      m_oDest.AutoRedraw = True
    End If
  End If
  
  If Not m_bPrinter Then
    m_oDest.Width = Printer.Width
    m_oDest.Height = Printer.Height
  End If
  
  lPrevScaleMode = m_oDest.ScaleMode
  m_oDest.ScaleMode = vbPixels
  
  If Not m_bPrinter Then
    m_lPageHeigth = m_oDest.Height
  End If
  
  m_lDataTop = m_lMarginTop
  m_lDataLeft = m_lMarginLeft
  
  With m_oFG
    bPrevRedraw = .Redraw
    .Redraw = False
    
    m_lDrawWidth = .GridLineWidth * m_dXFactor
    
    For lRow = 0 To .FixedRows - 1
      m_lDataTop = m_lDataTop + m_oDest.ScaleY(.RowHeight(lRow), vbTwips, vbPixels)
    Next lRow
    
    For lCol = 0 To .FixedCols - 1
      m_lDataLeft = m_lDataLeft + m_oDest.ScaleX(.ColWidth(lCol), vbTwips, vbPixels)
    Next lCol
  End With

  m_lDestWidth = m_oDest.ScaleWidth - m_lMarginRight
  m_lDestHeight = m_oDest.ScaleHeight - m_lMarginBottom
  
  ' start the actual print
  m_lPageCnt = 1
  If Not m_bCountOnly Then
    RaiseEvent NewPage(m_lPageCnt)
  End If
  doPrint 0, m_oFG.Rows - 1, 0, m_oFG.Cols - 1
  '
  
  If Not m_bCountOnly Then
    If m_bPrinter Then
      Printer.EndDoc
    Else
      m_oDest.Refresh
      m_oDest.Visible = True
    End If
  End If
  m_oDest.ScaleMode = lPrevScaleMode
  m_oFG.Redraw = bPrevRedraw
  PrintGrid = m_lPageCnt
End Function

Private Sub doPrint(ByVal lStartRow As Long, ByVal lEndRow As Long, ByVal lStartCol As Long, ByVal lEndCol As Long, Optional ByVal bColumns As Boolean = False)
  Dim bNewColPage As Boolean, bNewRowPage As Boolean
  Dim lCol As Long, lRow As Long
  Dim lCellWidth As Long, lCellHeight As Long
  Dim lCurrentX As Double, lCurrentY As Double
  Dim lNewStartCol As Long, lNewStartRow As Long
  
  With m_oFG
    lCurrentX = m_lDataLeft
    
    For lCol = lStartCol To lEndCol
      
      If lCol >= .FixedCols Then
        lCellWidth = m_oDest.ScaleX(.ColWidth(lCol), vbTwips, vbPixels)
        
        If lCurrentX + lCellWidth > m_lDestWidth Then
          bNewColPage = True
          lNewStartCol = lCol
          Exit For
        End If
        
        lCurrentY = m_lDataTop
        
        For lRow = lStartRow To lEndRow
          lCellHeight = m_oDest.ScaleY(.RowHeight(lRow), vbTwips, vbPixels)
          
          If lRow >= .FixedRows Then
            
            If lCurrentY + lCellHeight > m_lDestHeight Then
              If Not bColumns Then
                bNewRowPage = True
                lNewStartRow = lRow
              End If
              Exit For
            End If
            
            If Not m_bCountOnly Then
              printCell lRow, lCol, lCurrentX, lCurrentY, lCellWidth, lCellHeight
            End If
            
            lCurrentY = lCurrentY + lCellHeight
          End If
        Next lRow
        
        lCurrentX = lCurrentX + lCellWidth
      End If
    Next lCol
  End With
  
  If Not m_bCountOnly Then
    doPrintFixed lStartRow, lEndRow, lStartCol, lEndCol
  End If
  
  If bNewColPage Then
    m_lPageCnt = m_lPageCnt + 1
    If Not m_bCountOnly Then
      RaiseEvent NewPage(m_lPageCnt)
      If m_bPrinter Then
        Printer.NewPage
      End If
    End If
    doPrint lStartRow, lEndRow, lNewStartCol, lEndCol, True
  End If
  
  'If bColumns Then Exit Sub
  
  If bNewRowPage Then
    m_lPageCnt = m_lPageCnt + 1
    If Not m_bCountOnly Then
      RaiseEvent NewPage(m_lPageCnt)
      If m_bPrinter Then
        Printer.NewPage
      End If
    End If
    doPrint lNewStartRow, lEndRow, lStartCol, lEndCol, False
  End If
End Sub
    
   
Private Sub doPrintFixed(ByVal lStartRow As Long, ByVal lEndRow As Long, ByVal lStartCol As Long, ByVal lEndCol As Long)
  Dim lCol As Long, lRow As Long
  Dim lCellWidth As Long, lCellHeight As Long
  Dim lCurrentX As Double, lCurrentY As Double
  
  With m_oFG
    lCurrentX = m_lMarginLeft
    
    ' print only fixedcolumns
    For lCol = 0 To .FixedCols - 1
      lCellWidth = m_oDest.ScaleX(.ColWidth(lCol), vbTwips, vbPixels)
      
      lCurrentY = m_lMarginTop
      
      For lRow = 0 To .FixedRows - 1
        lCellHeight = m_oDest.ScaleY(.RowHeight(lRow), vbTwips, vbPixels)
        printCell lRow, lCol, lCurrentX, lCurrentY, lCellWidth, lCellHeight
        lCurrentY = lCurrentY + lCellHeight
      Next lRow
      
      For lRow = lStartRow To lEndRow
        If lRow >= .FixedRows Then
          lCellHeight = m_oDest.ScaleY(.RowHeight(lRow), vbTwips, vbPixels)
          
          If lCurrentY + lCellHeight > m_lDestHeight Then Exit For
          
          printCell lRow, lCol, lCurrentX, lCurrentY, lCellWidth, lCellHeight
          lCurrentY = lCurrentY + lCellHeight
        End If
      Next lRow
      
      lCurrentX = lCurrentX + lCellWidth
    Next lCol
      
    ' print only fixedrows
    lCurrentY = m_lMarginTop
    For lRow = 0 To .FixedRows - 1
      lCellHeight = m_oDest.ScaleY(.RowHeight(lRow), vbTwips, vbPixels)
      
      lCurrentX = m_lDataLeft
      For lCol = lStartCol To lEndCol
        If lCol >= .FixedCols Then
          lCellWidth = m_oDest.ScaleX(.ColWidth(lCol), vbTwips, vbPixels)
          If lCurrentX + lCellWidth > m_lDestWidth Then Exit For
          printCell lRow, lCol, lCurrentX, lCurrentY, lCellWidth, lCellHeight
          lCurrentX = lCurrentX + lCellWidth
        End If
      Next lCol
      lCurrentY = lCurrentY + lCellHeight
    Next lRow
  End With
End Sub
    
Private Sub printCell(ByVal lRow As Long, ByVal lCol As Long, lCurrentX As Double, ByVal lCurrentY As Double, ByVal lCellWidth, ByVal lCellHeight)
  Dim lCellAlignment As Long
  Dim textRECT As RECT, bottomRECT As RECT, textFormat As Long
  Dim lSucces As Long
  Dim lGridColor As Long, lBackColor As Long
  Dim lNewPen As Long, lOldPen As Long
  Dim lNewBrush As Long, lOldBrush As Long
    
  With m_oFG
    .Row = lRow
    .Col = lCol
    If lCol < .FixedCols Or lRow < .FixedRows Then
      lBackColor = .BackColorFixed
      lGridColor = .GridColorFixed
    Else
      lBackColor = .BackColor
      lGridColor = .GridColor
    End If
    If .CellBackColor <> 0 Then lBackColor = .CellBackColor
    If lBackColor < 0 Then
      lBackColor = lBackColor And Not &H80000000 ' remove the high bit
      lBackColor = GetSysColor(lBackColor)
    End If
    If lGridColor < 0 Then
      lGridColor = lGridColor And Not &H80000000 ' remove the high bit
      lGridColor = GetSysColor(lGridColor)
    End If
    m_oDest.FontName = .CellFontName
    m_oDest.FontBold = .CellFontBold
    m_oDest.FontSize = .CellFontSize
    m_oDest.FontItalic = .CellFontItalic
    m_oDest.FontStrikethru = .CellFontStrikeThrough
    m_oDest.FontUnderline = .CellFontUnderline
    
    If .WordWrap Then
      textFormat = DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING
    Else
      textFormat = DT_SINGLELINE Or DT_NOPREFIX
    End If

    lCellAlignment = .CellAlignment
    If lCellAlignment = 0 Then lCellAlignment = .ColAlignment(lCol)
    Select Case lCellAlignment
      Case flexAlignLeftBottom
        textFormat = textFormat Or DT_LEFT
        textFormat = textFormat Or DT_BOTTOM
      Case flexAlignLeftCenter
        textFormat = textFormat Or DT_LEFT
        textFormat = textFormat Or DT_VCENTER
      Case flexAlignLeftTop
        textFormat = textFormat Or DT_LEFT
        textFormat = textFormat Or DT_TOP
      Case flexAlignCenterBottom
        textFormat = textFormat Or DT_CENTER
        textFormat = textFormat Or DT_BOTTOM
      Case flexAlignCenterCenter
        textFormat = textFormat Or DT_CENTER
        textFormat = textFormat Or DT_VCENTER
      Case flexAlignCenterTop
        textFormat = textFormat Or DT_CENTER
        textFormat = textFormat Or DT_TOP
      Case flexAlignRightBottom
        textFormat = textFormat Or DT_RIGHT
        textFormat = textFormat Or DT_BOTTOM
      Case flexAlignRightCenter
        textFormat = textFormat Or DT_RIGHT
        textFormat = textFormat Or DT_VCENTER
      Case flexAlignRightTop
        textFormat = textFormat Or DT_RIGHT
        textFormat = textFormat Or DT_TOP
      Case flexAlignGeneral
        If IsNumeric(.TextMatrix(lRow, lCol)) Or IsDate(.TextMatrix(lRow, lCol)) Then
          textFormat = textFormat Or DT_RIGHT
        Else
          textFormat = textFormat Or DT_LEFT
        End If
        textFormat = textFormat Or DT_VCENTER
    End Select
    
    If m_bPrintBackColor Then
      lNewBrush = CreateSolidBrush(lBackColor)
      lOldBrush = SelectObject(m_oDest.hDC, lNewBrush)
    End If
    
    If m_bPrintGridLines Then
      lNewPen = CreatePen(PS_SOLID, m_lDrawWidth, lGridColor)
      lOldPen = SelectObject(m_oDest.hDC, lNewPen)
    
      With textRECT
        .Top = lCurrentY
        .Left = lCurrentX
        .Bottom = lCurrentY + lCellHeight
        .Right = lCurrentX + lCellWidth
        RoundRect m_oDest.hDC, .Left, .Top, .Right + m_lDrawWidth, .Bottom + m_lDrawWidth, 0, 0
      End With
      
      SelectObject m_oDest.hDC, lOldPen
      DeleteObject lNewPen
    End If
    
    If m_bPrintBackColor Then
      SelectObject m_oDest.hDC, lOldBrush
      DeleteObject lNewBrush
    End If
    
    With textRECT
      .Top = lCurrentY + 3 * m_dYFactor
      .Left = lCurrentX + 3 * m_dXFactor
      .Bottom = lCurrentY + lCellHeight - 3 * m_dYFactor
      .Right = lCurrentX + lCellWidth - 3 * m_dXFactor
    End With
    
    SetTextColor m_oDest.hDC, .CellForeColor
    
    If .WordWrap And (lCellAlignment <> flexAlignRightTop Or lCellAlignment <> flexAlignCenterTop) Then
      bottomRECT = textRECT
      With bottomRECT
        .Bottom = DrawText(m_oDest.hDC, m_oFG.TextMatrix(lRow, lCol), Len(m_oFG.TextMatrix(lRow, lCol)), bottomRECT, DT_CALCRECT Or textFormat)
        .Left = 0
        .Top = 0
        .Right = textRECT.Right - textRECT.Left
      End With
      AlignRECT textRECT, bottomRECT, lCellAlignment
      lSucces = DrawText(m_oDest.hDC, .TextMatrix(lRow, lCol), Len(.TextMatrix(lRow, lCol)), bottomRECT, textFormat)
    Else
      lSucces = DrawText(m_oDest.hDC, .TextMatrix(lRow, lCol), Len(.TextMatrix(lRow, lCol)), textRECT, textFormat)
    End If
  End With
End Sub

' To justify the actual celltext in the complete cell
Private Sub AlignRECT(CellRect As RECT, textRECT As RECT, ByVal Alignment As Integer)
  Dim lCellTextWidth As Long, lCellTextHeight As Long
  Dim lCellWidth As Long, lCellHeight As Long
  
  lCellTextWidth = textRECT.Right - textRECT.Left
  lCellTextHeight = textRECT.Bottom - textRECT.Top
  
  lCellWidth = CellRect.Right - CellRect.Left
  lCellHeight = CellRect.Bottom - CellRect.Top

  If lCellTextHeight > lCellHeight Then
    textRECT.Top = CellRect.Top
    textRECT.Bottom = CellRect.Bottom
    lCellTextHeight = textRECT.Bottom - textRECT.Top
  End If

  If lCellTextWidth > lCellWidth Then
    textRECT.Left = CellRect.Left
    textRECT.Right = CellRect.Right
    lCellTextWidth = textRECT.Right - textRECT.Left
  End If

  With textRECT
    Select Case Alignment
      Case flexAlignLeftTop   'Left top
        .Left = CellRect.Left
        .Top = CellRect.Top
        .Right = .Left + lCellTextWidth
        .Bottom = .Top + lCellTextHeight
      Case flexAlignLeftCenter   'Left Center
        .Left = CellRect.Left
        .Top = ((lCellHeight - lCellTextHeight) / 2) + CellRect.Top
        .Right = .Left + lCellTextWidth
        .Bottom = .Top + lCellTextHeight
      Case flexAlignLeftBottom   'Left bottom
        .Top = CellRect.Bottom - lCellTextHeight
        .Left = CellRect.Left
        .Right = .Left + lCellTextWidth
        .Bottom = .Top + lCellTextHeight
      Case flexAlignCenterTop 'Center Top
        .Left = ((lCellWidth - lCellTextWidth) / 2) + CellRect.Left
        .Top = CellRect.Top
        .Right = .Left + lCellTextWidth
        .Bottom = .Top + lCellTextHeight
      Case flexAlignCenterCenter, flexAlignGeneral    'Center Center
        .Top = ((lCellHeight - lCellTextHeight) / 2) + CellRect.Top
        .Left = ((lCellWidth - lCellTextWidth) / 2) + CellRect.Left
        .Right = .Left + lCellTextWidth
        .Bottom = .Top + lCellTextHeight
      Case flexAlignCenterBottom   'Center Bottom
        .Top = CellRect.Bottom - lCellTextHeight
        .Left = ((lCellWidth - lCellTextWidth) / 2) + CellRect.Left
        .Right = .Left + lCellTextWidth
        .Bottom = .Top + lCellTextHeight
      Case flexAlignRightTop   'Right Top
        .Left = CellRect.Right - lCellTextWidth
      End Select
      End With
End Sub
      
