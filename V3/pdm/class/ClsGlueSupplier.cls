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
'����ģ�������GlueSupplier���еĵ��ô���
'GlueSupplier���������µ��ֶ���Ҫ����

Public Glue12NC As String
Public SupplierName As String
Public SupplierPN As String
'GlueSupplier��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '����Insert���,Insert�еĸ��ַ�������Glue12NC��SupplierName��SupplierPN�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO GlueSupplier (Glue12NC, SupplierName,SupplierPN)" _
     + " Values('" + Trim(Glue12NC) + "','" + Trim(SupplierName) + "','" + Trim(SupplierPN) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'GlueSupplier��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriGlue12NC As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '����Update���,Update�еĸ��ַ�������Glue12NC��SupplierName��SupplierPN�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriGlue12NC
   SqlStmt = "Update GlueSupplier Set Glue12NC='" + Trim(Glue12NC) + "',SupplierName='" + Trim(SupplierName) + "',SupplierPN='" + Trim(SupplierPN) + "' WHERE Glue12NC='" + Trim(OriGlue12NC) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'GlueSupplier��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempGlue12NC As String)
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������Glue12NC�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��Glue12NC
  SqlStmt = "Delete FROM GlueSupplier WHERE Glue12NC='" + Trim(TempGlue12NC) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempGlue12NC As String) As Boolean
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempGlue12NC�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT Glue12NC FROM GlueSupplier WHERE Glue12NC='" + Trim(TempGlue12NC) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSupplier As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsSupplier, False)
  
  '�ж��Ƿ����
  If rsSupplier.RecordCount <> 0 Then
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function

