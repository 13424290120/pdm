VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDatabaseSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'本ClsDatabaseSet类主要是封装对数据库表的检测和初始化
'In_DB函数用于检查某个数据表是否存在于数据库当中，传递参数是表名
'函数将传递进来的table参数作为SQL语句中的表名
Public Function In_DB(ByVal Table As String) As Boolean
  '连接数据库
  MyDB.openDB
  
  '设置查询语句
  SqlStmt = "SELECT * FROM " + Trim(Table) + ""
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsTable As ADODB.Recordset
  In_DB = MyDB.exeSql(SqlStmt, rsTable, False)
  
  '断开与数据库的连接
  MyDB.closeDB
  
End Function



'Delete函数用于消除指定数据表的内容
Public Sub Delete(ByVal Table As String)
   '连接数据库
  MyDB.openDB
  
  '设置Delete语句
  SqlStmt = "Delete FROM " + Trim(Table) + ""
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsTable As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsTable, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
