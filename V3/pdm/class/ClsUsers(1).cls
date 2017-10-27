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
'��ClsUser����Ҫ�Ƿ�װ��Users���ĸ��ֲ��������¶���Users���ĸ����ֶ�
Public Name As String
Public Password As String
Public UserGroup As String
Public UserTitle As String
Public GrantGroup As String



Public Function GetInfo(ByVal TmpName As String, ByVal TmpPassWord As String) As Boolean
    
    '�������ݿ�
    MyDB.openDB
    
    SqlStmt = "SELECT * FROM Users WHERE Name='" + Trim(TmpName) + "'"
    
    'ִ��SQL���
    Dim Isexe As Boolean
    Dim rsUsers As ADODB.Recordset
    Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
    
    If Trim(rsUsers.Fields("Password")) <> TmpPassWord Then
        GetInfo = False
        
        '�Ͽ������ݿ������
        MyDB.closeDB
        Exit Function
        
    End If
    GetInfo = True
    
    '�Ͽ������ݿ������
    MyDB.closeDB
End Function


Public Sub Insert()
  
  '�������ݿ�
  MyDB.openDB 'MyDB��������ģ��ClsDB��һ������openDB����ģ��ClsDB��һ������
  
  '����Insert��� Insert into ���� ���ֶ�1���ֶ�2��...��Values ���ֶ�1ֵ���ֶ�2ֵ��...������ Values ��1,'�г���','������'��''��
  'SqlStmt = "INSERT INTO Users (Name, Password,EngineeringSys,PurchasingSys,SaleLogisticSys,ManufactureSys,HumanResrcSys,QualityClientSys,SystemAdmin,EngineeringApproval)" _
     + " Values('" + Trim(Name) + "','" + Trim(Password) + "','" + Trim(EngineeringSys) + "','" + Trim(PurchasingSys) + "','" + Trim(SaleLogisticSys) _
     + "','" + Trim(ManufactureSys) + "','" + Trim(HumanResrcSys) + "','" + Trim(QualityClientSys) + "','" + Trim(SystemAdmin) + "','" + Trim(EngineeringApproval) + "')"
  SqlStmt = "INSERT INTO Users (Name, Password,UserGroup,UserTitle,GrantGroup) Values('" + Trim(Name) + "','" + Trim(Password) + "','" + Trim(UserGroup) + "','" + Trim(UserTitle) + "','" + Trim(GrantGroup) + "')"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True) 'ִ��SQL�ɹ�����ʧ�ܱ�Ǹ��߼�����Isexe
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Sub Update(ByVal OriName As String)
   
  '�������ݿ⣨�ο�����Insert��
  MyDB.openDB
  
   '����Update��� Update ���� SET �ֶ�1 = ����ʽ1 ���ֶ�2 = ����ʽ2...
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
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Sub UpdatePwd(ByVal OriName As String)            'ֻ���ڸ���password�����ֶβ���
   
  '�������ݿ⣨�ο�����Insert��
  MyDB.openDB
  
   '����Update��� Update ���� SET �ֶ�1 = ����ʽ1 ���ֶ�2 = ����ʽ2...
   SqlStmt = "Update Users Set Name='" + Trim(Name) + "',Password='" + Trim(Password) + "' WHERE Name='" + Trim(OriName) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub

Public Sub Delete(ByVal TempName As String)
   '�������ݿ�
  MyDB.openDB
  
  '����Delete���
  SqlStmt = "Delete FROM Users WHERE Name='" + Trim(TempName) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'In_DB�������ڼ��Ҫ���ӵ��û��Ƿ���������ݿ⵱�У����ݲ������û���
Public Function In_DB(ByVal TempName As String) As Boolean
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ���
  SqlStmt = "SELECT Name FROM Users WHERE Name='" + Trim(TempName) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsUsers As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsUsers, False)
  
  '�ж��Ƿ����
  If rsUsers.RecordCount <> 0 Then  'RecordCount���Է��ؼ�¼���еļ�¼��
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function





