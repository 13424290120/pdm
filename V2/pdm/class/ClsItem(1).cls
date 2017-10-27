VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Model As String
Public Item As String
Public Drawing As String
Public Name As String
Public Specification As String
Public Color As String
Public Unit As String
Public MB As String
Public IT As String
Public ST As String
Public Inspect As String
Public Capacity As String
Public NWeight As String
Public RWeight As String
Public PackWeight As String
Public PackVolume As String
Public PackType As String
Public PackQty As String
Public ProLow As String
Public ProHigh As String
Public Valid As String
Public AlarmDays As String
Public Remark As String






Public Sub Insert()
  '连接数据库
  MyDB.openDB
  
  '设置Insert语句
  SqlStmt = "INSERT INTO Item (Model,Item,Drawing,Name,Specification,Color,Unit,MB,IT,ST,Inspect,Capacity,NWeight,RWeight,PackWeight,PackVolume,PackType,PackQty,ProLow,ProHigh,Valid,AlarmDays,Remark)" _
  + " Values('" + Trim(Model) + "','" + Trim(Item) + "','" + Trim(Drawing) _
  + "','" + Trim(Name) + "','" + Trim(Specification) + "','" + Trim(Color) _
  + "','" + Trim(Unit) + "','" + Trim(MB) + "','" + Trim(IT) + "','" + Trim(ST) _
  + "','" + Trim(Inspect) + "','" + Trim(Capacity) + "','" + Trim(NWeight) _
  + "','" + Trim(RWeight) + "','" + Trim(PackWeight) + "','" + Trim(PackVolume) _
  + "','" + Trim(PackType) + "','" + Trim(PackQty) + "','" + Trim(ProLow) + "','" + Trim(ProHigh) _
  + "','" + Trim(Valid) + "','" + Trim(AlarmDays) + "','" + Trim(Remark) + "')"

  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsItem As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsItem, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Sub Update(ByVal OriItem As String)
   
  '连接数据库
  MyDB.openDB
  
   '设置Update语句
  SqlStmt = "Update Item Set Model='" + Trim(Model) + "',Item='" + Trim(Item) + "',Drawing='" + Trim(Drawing) _
  + "',Name='" + Trim(Name) + "',Specification='" + Trim(Specification) + "',Color='" + Trim(Color) _
  + "',Unit='" + Trim(Unit) + "',MB='" + Trim(MB) + "',IT='" + Trim(IT) + "',ST='" + Trim(ST) _
  + "',Inspect='" + Trim(Inspect) + "',Capacity='" + Trim(Capacity) + "',NWeight='" + Trim(NWeight) _
  + "',RWeight='" + Trim(RWeight) + "',PackWeight='" + Trim(PackWeight) + "',PackVolume='" + Trim(PackVolume) _
  + "',PackType='" + Trim(PackType) + "',PackQty='" + Trim(PackQty) _
  + "',ProLow='" + Trim(ProLow) + "',ProHigh='" + Trim(ProHigh) _
  + "',Valid='" + Trim(Valid) _
  + "',AlarmDays='" + Trim(AlarmDays) _
  + "',Remark='" + Trim(Remark) + "'WHERE Item='" + Trim(OriItem) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsItem As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsItem, True)
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Sub Delete(ByVal TempItem As String)
   '连接数据库
  MyDB.openDB
  
  '设置Delete语句删除Item表中的信息
  SqlStmt = "Delete FROM Item WHERE Item='" + Trim(TempItem) + "'"
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsItem As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsItem, True)
  
  '设置Delete语句删除Parent表中的信息
  SqlStmt = "Delete FROM Parent WHERE ParentItem='" + Trim(TempItem) + "'"
  '执行SQL语句
  Isexe = MyDB.exeSql(SqlStmt, rsItem, True)
  
  '设置Delete语句删除Child表中的信息
  SqlStmt = "Delete FROM Child WHERE ChildItem='" + Trim(TempItem) + "' or ParentItem='" + Trim(TempItem) + "' "
  '执行SQL语句
  Isexe = MyDB.exeSql(SqlStmt, rsItem, True)
  
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempItem As String) As Boolean
  '连接数据库
  MyDB.openDB
  
  '设置查询语句
  SqlStmt = "SELECT Item FROM Item WHERE Item='" + Trim(TempItem) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsItem As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsItem, False)
  
  '判断是否存在
  If rsItem.RecordCount <> 0 Then
    In_DB = True
  Else
    In_DB = False
  End If
  
  '断开与数据库的连接
  MyDB.closeDB
  
End Function

