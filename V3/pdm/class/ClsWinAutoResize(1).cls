VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Resizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Dim m_buff(200) As ad
Dim m_obj As Form
Private Type ad
    name As String
    obj As Object
    left As Long    '居左
    top As Long     '居上
    right As Long   '居右
    button As Long '居下
    width As Long   '宽
    height As Long '高
End Type
'在这里面对控件初始位置进行定位
Private Function Init()
Dim i0 As Integer
Dim tmp As Object
Dim rongqi As Object
Dim n0 As Integer
    On Error Resume Next
    For Each tmp In m_obj.Controls
        
        Set rongqi = m_obj
        Set m_buff(i0).obj = tmp
        m_buff(i0).name = tmp.name & "#" & GetCtlIndex(tmp)
        m_buff(i0).left = tmp.left
        m_buff(i0).top = tmp.top
        m_buff(i0).width = tmp.width
        m_buff(i0).height = tmp.height
        
        Set rongqi = tmp.Container
        m_buff(i0).right = rongqi.width - tmp.left - tmp.width
        m_buff(i0).button = rongqi.height - tmp.top - tmp.height
        i0 = i0 + 1
err:
    Next
End Function
Private Function GetCtlIndex(obj As Object) As Long
    On Error Resume Next
    GetCtlIndex = -1
    GetCtlIndex = obj.Index
End Function

'调整位置
Public Function resize(obj As Form)
On Error Resume Next
Dim tmp As Object
Dim n0 As Long
Dim str0 As String
    If obj.WindowState = 1 Then Exit Function
    If m_obj Is Nothing Then
        Set m_obj = obj
        Call Init
        Exit Function
    End If
    For Each tmp In m_obj.Controls
        n0 = -1
        n0 = FindAD(tmp.name & "#" & GetCtlIndex(tmp))
        If n0 <> -1 Then
            '在这里对控件的TAB进行分析，然后调整控件位置
            Call ExeStr(tmp, n0)
        End If
    Next
End Function
Private Function FindAD(name As String) As Long
Dim i0 As Integer
    If InStr(1, name, "#") = 0 Then
        name = name & "#-1"
    End If
    FindAD = -1
    For i0 = 0 To UBound(m_buff)
        If m_buff(i0).name = name Then
            FindAD = i0
            Exit Function
        End If
    Next
    Exit Function
err:
End Function
'根据字符串进行重定义
Private Function ExeStr(obj As Object, id As Long)
Dim rongqi As Object
Dim str0 As String
Dim n0 As Integer
    str0 = obj.Tag
    Set rongqi = m_obj
    Set rongqi = obj.Container
        If InStr(1, str0, "left") <> 0 Then
            obj.left = m_buff(id).left
            obj.width = rongqi.width - m_buff(id).right - m_buff(id).left
    If InStr(1, str0, "right") <> 0 Then
        Else
            obj.left = rongqi.width - m_buff(id).right - m_buff(id).width
        End If
    End If
    If InStr(1, str0, "button") <> 0 Then
        If InStr(1, str0, "top") <> 0 Then
            obj.top = m_buff(id).top
            obj.height = rongqi.height - m_buff(id).button - m_buff(id).top
        Else
            obj.top = rongqi.height - m_buff(id).button - m_buff(id).height
        End If
    End If
    '如果这个容器是TAB控件，那么就把这个控件的index设置为1
    
End Function

Private Sub Class_Initialize()
    '清空数组
End Sub


