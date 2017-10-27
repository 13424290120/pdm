VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

End Function

Public Sub ResizeControl(inControl As Control, pfrmIn As Form)

  On Error Resume Next
  Dim i As Long
  Dim widthfactor As Single, heightfactor As Single
  Dim minFactor As Single
  Dim yRatio, xRatio, lTop, lLeft, lWidth, lHeight As Long
  yRatio = PerHeight(pfrmIn)
  xRatio = PerWidth(pfrmIn)
  i = FindControl(inControl, pfrmIn.Name)

  If inControl.Left < 0 Then
    lLeft = CLng(((ControlRecord(i).Left * xRatio) / 100) - 75000)
  Else
    lLeft = CLng((ControlRecord(i).Left * xRatio) / 100)
  End If

  lTop = CLng((ControlRecord(i).Top * yRatio) / 100)
  lWidth = CLng((ControlRecord(i).Width * xRatio) / 100)
  lHeight = CLng((ControlRecord(i).Height * yRatio) / 100)
  If TypeOf inControl Is Line Then

    If inControl.X1 < 0 Then
      inControl.X1 = CLng(((ControlRecord(i).Left * xRatio) / 100) - 75000)
    Else
      inControl.X1 = CLng((ControlRecord(i).Left * xRatio) / 100)
    End If

    inControl.Y1 = CLng((ControlRecord(i).Top * yRatio) / 100)
    If inControl.X2 < 0 Then
      inControl.X2 = CLng(((ControlRecord(i).Width * xRatio) / 100) - 75000)
    Else
      inControl.X2 = CLng((ControlRecord(i).Width * xRatio) / 100)
    End If

    inControl.Y2 = CLng((ControlRecord(i).Height * yRatio) / 100)
  Else
    inControl.Move lLeft, lTop, lWidth, lHeight
    inControl.Move lLeft, lTop, lWidth
    inControl.Move lLeft, lTop
  End If

End Sub

Public Sub ResizeForm(pfrmIn As Form)

  Dim FormControl As Control
  Dim isVisible As Boolean
  Dim StartX, StartY, MaxX, MaxY As Long
  Dim bNew As Boolean

  If Not bRunning Then
    bRunning = True

    If FindForm(pfrmIn) < 0 Then
      bNew = True
    Else
      bNew = False
    End If
    If pfrmIn.Top < 30000 Then
      isVisible = pfrmIn.Visible
      On Error Resume Next
      If Not pfrmIn.MDIChild Then
        On Error GoTo 0
        ' ' pfrmIn.Visible = False
      Else

        If bNew Then
          StartY = pfrmIn.Height
          StartX = pfrmIn.Width
          On Error Resume Next
          For Each FormControl In pfrmIn
            If FormControl.Left + FormControl.Width + 200 > MaxX Then
              MaxX = FormControl.Left + FormControl.Width + 200
            End If

            If FormControl.Top + FormControl.Height + 500 > MaxY Then
              MaxY = FormControl.Top + FormControl.Height + 500
            End If

            If FormControl.X1 + 200 > MaxX Then
              MaxX = FormControl.X1 + 200
            End If

            If FormControl.Y1 + 500 > MaxY Then
              MaxY = FormControl.Y1 + 500
            End If

            If FormControl.X2 + 200 > MaxX Then
              MaxX = FormControl.X2 + 200
            End If

            If FormControl.Y2 + 500 > MaxY Then
              MaxY = FormControl.Y2 + 500
            End If

          Next FormControl

          On Error GoTo 0
          pfrmIn.Height = MaxY
          pfrmIn.Width = MaxX
        End If

        On Error GoTo 0
      End If

      For Each FormControl In pfrmIn
        ResizeControl FormControl, pfrmIn
      Next FormControl

      On Error Resume Next

      If Not pfrmIn.MDIChild Then
        On Error GoTo 0
        pfrmIn.Visible = isVisible
      Else

        If bNew Then
        pfrmIn.Height = StartY
        pfrmIn.Width = StartX

        For Each FormControl In pfrmIn
          ResizeControl FormControl, pfrmIn
        Next FormControl

      End If
    End If
    On Error GoTo 0
  End If
  bRunning = False
End If

End Sub

Public Sub SaveFormPosition(pfrmIn As Form)

  Dim i As Long

  If MaxForm > 0 Then

    For i = 0 To (MaxForm - 1)

      If FormRecord(i).Name = pfrmIn.Name Then

        FormRecord(i).Top = pfrmIn.Top
        FormRecord(i).Left = pfrmIn.Left
        FormRecord(i).Height = pfrmIn.Height
        FormRecord(i).Width = pfrmIn.Width
        Exit Sub
      End If
    Next i

    AddForm (pfrmIn)
  End If
End Sub

Public Sub RestoreFormPosition(pfrmIn As Form)

  Dim i As Long
  If MaxForm > 0 Then
    For i = 0 To (MaxForm - 1)
      If FormRecord(i).Name = pfrmIn.Name Then
        If FormRecord(i).Top < 0 Then
          pfrmIn.WindowState = 2
        ElseIf FormRecord(i).Top < 30000 Then
          pfrmIn.WindowState = 0
          pfrmIn.Move FormRecord(i).Left, FormRecord(i).Top, FormRecord(i).Width, FormRecord(i).Height
        Else
          pfrmIn.WindowState = 1
        End If
          Exit Sub
      End If
    Next i
  End If
End Sub

Public Sub Resize_ALL(Form_Name As Form)

  Dim OBJ As Object
  For Each OBJ In Form_Name
    ResizeControl OBJ, Form_Name
  Next OBJ
End Sub

Public Sub DragForm(frm As Form)

  On Local Error Resume Next
  Call ReleaseCapture
  Call SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, 2, 0)

End Sub
