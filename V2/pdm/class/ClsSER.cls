VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'本类模块是针对对应的表中的调用创建
'表中有以下的字段需要定义

Public SERIndex As String                          '############按对应表中字段修改
Public Applicant As String                          '############按对应表中字段修改
Public CAorA As String                             '############按对应表中字段修改
Public Description As String                        '############按对应表中字段修改
Public IDSO As String                               '############按对应表中字段修改
Public OpnDate As Date                              '############按对应表中字段修改
Public ClosDate As Date                             '############按对应表中字段修改
Public PJNOIndex As String                             '############按对应表中字段修改
Public PjtName As String                            '############按对应表中字段修改
Public FinsGdNO As String                             '############按对应表中字段修改
Public SglPrtNO As String                             '############按对应表中字段修改
Public CommtNote As String                             '############按对应表中字段修改



'SER类模块中的insert子过程，用于添加一个记录
Public Sub Insert()
  
  '连接数据库,另外一个类模块在被引用
  MyDB.openDB
  
  '#################设置Insert语句,Insert中的各字符串变量SERIndex,Applicant,CAorA,Description在调用此类模块的窗口中赋值
  SqlStmt = "INSERT INTO SER (SERIndex,Applicant,CAorA,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,FinsGdNO,SglPrtNO,CommtNote)" _
     + " Values('" + Trim(SERIndex) + "','" + Trim(Applicant) + "','" + Trim(CAorA) + "','" + Trim(Description) + "','" + Trim(IDSO) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" _
      + Trim(PJNOIndex) + "','" + Trim(PjtName) + "','" + Trim(FinsGdNO) + "','" + Trim(SglPrtNO) + "','" + Trim(CommtNote) + "')"
    
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                    '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)       '###########变量改成对应的表名字
  
  
  '增加一个触发，当Status=A，直接将Single Part NEW转为OLD
  SqlStmt = "UPDATE SglPrt SET NewOldStatus='OLD' WHERE SglPrtIndex=" & Trim(SglPrtNO)
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)
  
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'SER类模块中的Update子过程，用于更新一个记录
Public Sub Update(ByVal OriSERIndex As String)
   
  '连接数据库
  MyDB.openDB
  
   '#################设置Update语句,Update中的各字符串变量SERIndex,Applicant,CAorA,Description...在调用此类模块的窗口中赋值
   ' update语句有点特别之处是还要一个源记录的主键位置OriSERIndex
   SqlStmt = "Update SER Set SERIndex='" + Trim(SERIndex) + "',Applicant='" + Trim(Applicant) + "',CAorA='" + Trim(CAorA) + "',Description='" + Trim(Description) + "',IDSO='" + Trim(IDSO) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + _
            "',PJNOIndex='" + Trim(PJNOIndex) + "',PjtName='" + Trim(PjtName) + "',FinsGdNO='" + Trim(FinsGdNO) + "',SglPrtNO='" + Trim(SglPrtNO) + "',CommtNote='" + Trim(CommtNote) + "' WHERE SERIndex='" + Trim(OriSERIndex) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                      '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)         '###########变量改成对应的表名字
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'SER类模块中的Delete子过程，用于删除一个记录
Public Sub Delete(ByVal TempSERIndex As String)             '###########变量改成对应的表字段名
   '连接数据库
  MyDB.openDB
  
  '设置Delete语句，Delete中的字符串变量SERIndex在调用此类模块的窗口中赋值
  'Delete语句有点特别之处是只要一个记录的主键位置SERIndex
  SqlStmt = "Delete FROM SER WHERE SERIndex='" + Trim(TempSERIndex) + "'"       '###########变量改成对应的表字段名 还有表名字
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                         '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)            '###########变量改成对应的表名字
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempSERIndex As String) As Boolean                 '###########变量改成对应的表字段名
  '连接数据库
  MyDB.openDB
  
  '设置查询语句，In_DB中的字符串变量TempSERIndex在调用此类模块的窗口中赋值
  '判断一个记录主键是否存在用该主键值在记录集中SELECT
  SqlStmt = "SELECT SERIndex FROM SER WHERE SERIndex='" + Trim(TempSERIndex) + "'"      '###########变量改成对应的表字段名 还有表名字
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                         '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsSER, False)           '###########变量改成对应的表名字
  
  '判断是否存在
  If rsSER.RecordCount <> 0 Then                         '###########变量改成对应的表名字
    In_DB = True
  Else
    In_DB = False
  End If
  
  '断开与数据库的连接
  MyDB.closeDB
  
End Function










