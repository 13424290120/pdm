VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFinsGd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'本类模块是针对对应的表中的调用创建
'表中有以下的字段需要定义

Public FinsGdIndex As String                          '############按对应表中字段修改
Public Applicant As String                          '############按对应表中字段修改
Public ProductLine As String                             '############按对应表中字段修改
Public Description As String                        '############按对应表中字段修改
Public IDSO As String                               '############按对应表中字段修改
Public OpnDate As Date                              '############按对应表中字段修改
Public ClosDate As Date                             '############按对应表中字段修改
Public PJNOIndex As String                             '############按对应表中字段修改
Public PjtName As String                            '############按对应表中字段修改
Public ItemType As String                             '############按对应表中字段修改
Public Location As String                             '############按对应表中字段修改
Public CommtNote As String                             '############按对应表中字段修改



'FinsGd类模块中的insert子过程，用于添加一个记录
Public Sub Insert()
  
  '连接数据库,另外一个类模块在被引用
  MyDB.openDB
  
  '#################设置Insert语句,Insert中的各字符串变量FinsGdIndex,Applicant,ProductLine,Description在调用此类模块的窗口中赋值
  SqlStmt = "INSERT INTO FinsGd (FinsGdIndex,Applicant,ProductLine,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,ItemType,Location,CommtNote)" _
     + " Values('" + Trim(FinsGdIndex) + "','" + Trim(Applicant) + "','" + Trim(ProductLine) + "','" + Trim(Description) + "','" + Trim(IDSO) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" _
      + Trim(PJNOIndex) + "','" + Trim(PjtName) + "','" + Trim(ItemType) + "','" + Trim(Location) + "','" + Trim(CommtNote) + "')"
    
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                    '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, True)       '###########变量改成对应的表名字
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'FinsGd类模块中的Update子过程，用于更新一个记录
Public Sub Update(ByVal OriFinsGdIndex As String)
   
  '连接数据库
  MyDB.openDB
  
   '#################设置Update语句,Update中的各字符串变量FinsGdIndex,Applicant,ProductLine,Description...在调用此类模块的窗口中赋值
   ' update语句有点特别之处是还要一个源记录的主键位置OriFinsGdIndex
   SqlStmt = "Update FinsGd Set FinsGdIndex='" + Trim(FinsGdIndex) + "',Applicant='" + Trim(Applicant) + "',ProductLine='" + Trim(ProductLine) + "',Description='" + Trim(Description) + "',IDSO='" + Trim(IDSO) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + _
            "',PJNOIndex='" + Trim(PJNOIndex) + "',PjtName='" + Trim(PjtName) + "',ItemType='" + Trim(ItemType) + "',Location='" + Trim(Location) + "',CommtNote='" + Trim(CommtNote) + "' WHERE FinsGdIndex='" + Trim(OriFinsGdIndex) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                      '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, True)         '###########变量改成对应的表名字
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'FinsGd类模块中的Delete子过程，用于删除一个记录
Public Sub Delete(ByVal TempFinsGdIndex As String)             '###########变量改成对应的表字段名
   '连接数据库
  MyDB.openDB
  
  '设置Delete语句，Delete中的字符串变量FinsGdIndex在调用此类模块的窗口中赋值
  'Delete语句有点特别之处是只要一个记录的主键位置FinsGdIndex
  SqlStmt = "Delete FROM FinsGd WHERE FinsGdIndex='" + Trim(TempFinsGdIndex) + "'"       '###########变量改成对应的表字段名 还有表名字
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                         '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, True)            '###########变量改成对应的表名字
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempFinsGdIndex As String) As Boolean                 '###########变量改成对应的表字段名
  '连接数据库
  MyDB.openDB
  
  '设置查询语句，In_DB中的字符串变量TempFinsGdIndex在调用此类模块的窗口中赋值
  '判断一个记录主键是否存在用该主键值在记录集中SELECT
  SqlStmt = "SELECT FinsGdIndex FROM FinsGd WHERE FinsGdIndex='" + Trim(TempFinsGdIndex) + "'"      '###########变量改成对应的表字段名 还有表名字
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                         '###########变量改成对应的表名字
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, False)           '###########变量改成对应的表名字
  
  '判断是否存在
  If rsFinsGd.RecordCount <> 0 Then                         '###########变量改成对应的表名字
    In_DB = True
  Else
    In_DB = False
  End If
  
  '断开与数据库的连接
  MyDB.closeDB
  
End Function












