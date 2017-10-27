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
'��ClsDatabaseSet����Ҫ�Ƿ�װ�����ݿ���ļ��ͳ�ʼ��
'In_DB�������ڼ��ĳ�����ݱ��Ƿ���������ݿ⵱�У����ݲ����Ǳ���
'���������ݽ�����table������ΪSQL����еı���
Public Function In_DB(ByVal Table As String) As Boolean
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ���
  SqlStmt = "SELECT * FROM " + Trim(Table) + ""
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsTable As ADODB.Recordset
  In_DB = MyDB.exeSql(SqlStmt, rsTable, False)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function



'Delete������������ָ�����ݱ�������
Public Sub Delete(ByVal Table As String)
   '�������ݿ�
  MyDB.openDB
  
  '����Delete���
  SqlStmt = "Delete FROM " + Trim(Table) + ""
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsTable As ADODB.Recordset
  Isexe = MyDB.exeSql(SqlStmt, rsTable, True)
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub