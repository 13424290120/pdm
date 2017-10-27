VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGlueSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'本类模块是针对GlueSupplier表中的调用创建
'GlueSupplier表中有以下的字段需要定义

Public Glue12NC As String
Public SupplierName As String
Public SupplierPN As String
'GlueSupplier类模块中的insert子过程，用于添加一个记录
Public Sub Insert()
  
  '连接数据库,另外一个类模块在被引用
  MyDB.openDB
  
  '设置Insert语句,Insert中的各字符串变量Glue12NC，SupplierName，SupplierPN在调用此类模块的窗口中赋值
  SqlStmt = "INSERT INTO GlueSupplier (Glue12NC, SupplierName,SupplierPN)" _
     + " Values('" + Trim(Glue12NC) + "','" + Trim(SupplierName) + "','" + Trim(SupplierPN) + "')"
    
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'GlueSupplier类模块中的Update子过程，用于更新一个记录
Public Sub Update(ByVal OriGlue12NC As String)
   
  '连接数据库
  MyDB.openDB
  
   '设置Update语句,Update中的各字符串变量Glue12NC，SupplierName，SupplierPN在调用此类模块的窗口中赋值
   ' update语句有点特别之处是还要一个源记录的主键位置OriGlue12NC
   SqlStmt = "Update GlueSupplier Set Glue12NC='" + Trim(Glue12NC) + "',SupplierName='" + Trim(SupplierName) + "',SupplierPN='" + Trim(SupplierPN) + "' WHERE Glue12NC='" + Trim(OriGlue12NC) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'GlueSupplier类模块中的Delete子过程，用于删除一个记录
Public Sub Delete(ByVal TempGlue12NC As String)
   '连接数据库
  MyDB.openDB
  
  '设置Delete语句，Delete中的字符串变量Glue12NC在调用此类模块的窗口中赋值
  'Delete语句有点特别之处是只要一个记录的主键位置Glue12NC
  SqlStmt = "Delete FROM GlueSupplier WHERE Glue12NC='" + Trim(TempGlue12NC) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempGlue12NC As String) As Boolean
  '连接数据库
  MyDB.openDB
  
  '设置查询语句，In_DB中的字符串变量TempGlue12NC在调用此类模块的窗口中赋值
  '判断一个记录主键是否存在用该主键值在记录集中SELECT
  SqlStmt = "SELECT Glue12NC FROM GlueSupplier WHERE Glue12NC='" + Trim(TempGlue12NC) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, False)
  
  '判断是否存在
  If rsSupplier.RecordCount <> 0 Then
    In_DB = True
  Else
    In_DB = False
  End If
  
  '断开与数据库的连接
  MyDB.closeDB
  
End Function


