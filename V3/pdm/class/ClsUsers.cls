VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'本ClsUser类主要是封装对Users表的各种操作，以下定义Users表的各个字段
Public Name As String
Public Password As String
Public UserGroup As String
Public UserTitle As String
Public GrantGroup As String



Public Function GetInfo(ByVal TmpName As String, ByVal TmpPassWord As String) As Boolean
    
    '连接数据库
    MyDB.openDB
    
    SqlStmt = "SELECT * FROM Users WHERE Name='" + Trim(TmpName) + "'"
    
    '执行SQL语句
    Dim Isexe As Boolean
    Dim rsUsers As ADODB.Recordset
    Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
    
    If Trim(rsUsers.Fields("Password")) <> TmpPassWord Then
        GetInfo = False
        
        '断开与数据库的连接
        MyDB.closeDB
        Exit Function
        
    End If
    GetInfo = True
    
    '断开与数据库的连接
    MyDB.closeDB
End Function


Public Sub Insert()
  
  '连接数据库
  MyDB.openDB 'MyDB是声明类模块ClsDB的一个对象，openDB是类模块ClsDB中一个函数
  
  '设置Insert语句 Insert into 表名 （字段1，字段2，...）Values （字段1值，字段2值，...）例如 Values （1,'市场部','开发部'，''）
  'SqlStmt = "INSERT INTO Users (Name, Password,EngineeringSys,PurchasingSys,SaleLogisticSys,ManufactureSys,HumanResrcSys,QualityClientSys,SystemAdmin,EngineeringApproval)" _
     + " Values('" + Trim(Name) + "','" + Trim(Password) + "','" + Trim(EngineeringSys) + "','" + Trim(PurchasingSys) + "','" + Trim(SaleLogisticSys) _
     + "','" + Trim(ManufactureSys) + "','" + Trim(HumanResrcSys) + "','" + Trim(QualityClientSys) + "','" + Trim(SystemAdmin) + "','" + Trim(EngineeringApproval) + "')"
  SqlStmt = "INSERT INTO Users (Name, Password,UserGroup,UserTitle,GrantGroup) Values('" + Trim(Name) + "','" + Trim(Password) + "','" + Trim(UserGroup) + "','" + Trim(UserTitle) + "','" + Trim(GrantGroup) + "')"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True) '执行SQL成功或者失败标记给逻辑变量Isexe
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Sub Update(ByVal OriName As String)
   
  '连接数据库（参考上面Insert）
  MyDB.openDB
  
   '设置Update语句 Update 表名 SET 字段1 = 表达式1 ，字段2 = 表达式2...
'   SqlStmt = "Update Users Set Name='" + Trim(Name) + "',Password='" + Trim(Password) _
'   + "',EngineeringSys='" + Trim(EngineeringSys) _
'   + "',PurchasingSys='" + Trim(PurchasingSys) + "',SaleLogisticSys='" + Trim(SaleLogisticSys) _
'   + "',ManufactureSys='" + Trim(ManufactureSys) + "',HumanResrcSys='" + Trim(HumanResrcSys) _
'   + "',QualityClientSys='" + Trim(QualityClientSys) + "',SystemAdmin='" + Trim(SystemAdmin) _
'   + "',EngineeringApproval='" + Trim(EngineeringApproval) _
'   + "' WHERE Name='" + Trim(OriName) + "'"
    SqlStmt = "Update Users Set Name='" + Trim(Name) + "',Password='" + Trim(Password) _
    + "',UserGroup='" + Trim(UserGroup) + "',UserTitle='" + Trim(UserTitle) + "',GrantGroup='" + Trim(GrantGroup) _
    + "' WHERE Name='" + Trim(OriName) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
Public Sub UpdatePwd(ByVal OriName As String)            '只用于更新password其它字段不变
   
  '连接数据库（参考上面Insert）
  MyDB.openDB
  
   '设置Update语句 Update 表名 SET 字段1 = 表达式1 ，字段2 = 表达式2...
   SqlStmt = "Update Users Set Name='" + Trim(Name) + "',Password='" + Trim(Password) + "' WHERE Name='" + Trim(OriName) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub

Public Sub Delete(ByVal TempName As String)
   '连接数据库
  MyDB.openDB
  
  '设置Delete语句
  SqlStmt = "Delete FROM Users WHERE Name='" + Trim(TempName) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
  
  '断开与数据库的连接
  MyDB.closeDB
End Sub
'In_DB函数用于检查要添加的用户是否存在于数据库当中，传递参数是用户名
Public Function In_DB(ByVal TempName As String) As Boolean
  '连接数据库
  MyDB.openDB
  
  '设置查询语句
  SqlStmt = "SELECT Name FROM Users WHERE Name='" + Trim(TempName) + "'"
  
  '执行SQL语句
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, False)
  
  '判断是否存在
  If rsUsers.RecordCount <> 0 Then  'RecordCount属性返回记录集中的记录数
    In_DB = True
  Else
    In_DB = False
  End If
  
  '断开与数据库的连接
  MyDB.closeDB
  
End Function






