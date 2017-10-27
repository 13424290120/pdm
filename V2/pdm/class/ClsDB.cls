VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����ģ������Ͳ����ݿ�ֱ�Ӳ�����
'����ģ��������������ɣ������ݿ⣬�ر����ݿ⣬ִ�����ݿ�SQL���

Option Explicit
'���ݿ�����
Public dbConn As New ADODB.Connection

'������ǰ���ݿ������״̬:False��δ���ӣ�True��������
Public dbStatus As Boolean

Public Function openDB() As Boolean  '�����ݿ⺯������True/False�߼�ֵ
    dbConn.Open connString
    If Err.Number <> 0 Then MsgBox Err.Description: GoTo ErrHandler
    openDB = True
    Exit Function

ErrHandler:
    openDB = False
    
End Function

Public Function closeDB() As Integer   '�ر����ݿ�
    dbConn.Close

    dbStatus = False
End Function

'ִ�����ݿ�SQL��䷵��True/False�߼�ֵ
'����ģ���ں���sqlStr����¼��Sourcing��SQL��䣻queryResult��¼�������趨��д�߼�����ҳ�ʼֵΪ��(��д)
Public Function exeSql(sqlStr As String, queryResult As ADODB.Recordset, Optional enableWrite As Boolean = True) As Boolean
    On Error GoTo ErrHandler:
    Dim ret
    Set queryResult = New ADODB.Recordset
    
    If enableWrite Then    '���ݿ�SQL����п�д��ֻ��֮��
        queryResult.Open sqlStr, dbConn, adOpenStatic, adLockOptimistic
    Else
        queryResult.Open sqlStr, dbConn, adOpenStatic, adLockReadOnly
    End If
    
    exeSql = True    '���ر���ģ�麯��ֵTrueִ�гɹ�
    Exit Function
    
ErrHandler:
    MsgBox Err.Description
    exeSql = False    '����Falseִֵ��ʧ��
    Exit Function
End Function