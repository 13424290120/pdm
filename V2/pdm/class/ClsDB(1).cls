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
'本类模块是最低层数据库直接操作层
'本类模块由三个函数组成：打开数据库，关闭数据库，执行数据库SQL语句

Option Explicit
'数据库连接
Public dbConn As New ADODB.Connection

'表明当前数据库的连接状态:False－未连接；True－已连接
Public dbStatus As Boolean

Public Function openDB() As Boolean  '打开数据库函数返回True/False逻辑值
    dbConn.Open connString
    If Err.Number <> 0 Then MsgBox Err.Description: GoTo ErrHandler
    openDB = True
    Exit Function

ErrHandler:
    openDB = False
    
End Function

Public Function closeDB() As Integer   '关闭数据库
    dbConn.Close

    dbStatus = False
End Function

'执行数据库SQL语句返回True/False逻辑值
'本类模块内函数sqlStr：记录集Sourcing的SQL语句；queryResult记录集名；设定可写逻辑标记且初始值为真(可写)
Public Function exeSql(sqlStr As String, queryResult As ADODB.Recordset, Optional enableWrite As Boolean = True) As Boolean
    On Error GoTo ErrHandler:
    Dim ret
    Set queryResult = New ADODB.Recordset
    
    If enableWrite Then    '数据库SQL语句有可写和只读之分
        queryResult.Open sqlStr, dbConn, adOpenStatic, adLockOptimistic
    Else
        queryResult.Open sqlStr, dbConn, adOpenStatic, adLockReadOnly
    End If
    
    exeSql = True    '返回本类模块函数值True执行成功
    Exit Function
    
ErrHandler:
    MsgBox Err.Description
    exeSql = False    '返回False值执行失败
    Exit Function
End Function
