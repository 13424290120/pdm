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
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public SERIndex As String                          '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public CAorA As String                             '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public IDSO As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�
Public PJNOIndex As String                             '############����Ӧ�����ֶ��޸�
Public PjtName As String                            '############����Ӧ�����ֶ��޸�
Public FinsGdNO As String                             '############����Ӧ�����ֶ��޸�
Public SglPrtNO As String                             '############����Ӧ�����ֶ��޸�
Public CommtNote As String                             '############����Ӧ�����ֶ��޸�



'SER��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������SERIndex,Applicant,CAorA,Description�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO SER (SERIndex,Applicant,CAorA,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,FinsGdNO,SglPrtNO,CommtNote)" _
     + " Values('" + Trim(SERIndex) + "','" + Trim(Applicant) + "','" + Trim(CAorA) + "','" + Trim(Description) + "','" + Trim(IDSO) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" _
      + Trim(PJNOIndex) + "','" + Trim(PjtName) + "','" + Trim(FinsGdNO) + "','" + Trim(SglPrtNO) + "','" + Trim(CommtNote) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  
  '����һ����������Status=A��ֱ�ӽ�Single Part NEWתΪOLD
  SqlStmt = "UPDATE SglPrt SET NewOldStatus='OLD' WHERE SglPrtIndex=" & Trim(SglPrtNO)
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)
  
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'SER��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriSERIndex As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '#################����Update���,Update�еĸ��ַ�������SERIndex,Applicant,CAorA,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriSERIndex
   SqlStmt = "Update SER Set SERIndex='" + Trim(SERIndex) + "',Applicant='" + Trim(Applicant) + "',CAorA='" + Trim(CAorA) + "',Description='" + Trim(Description) + "',IDSO='" + Trim(IDSO) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + _
            "',PJNOIndex='" + Trim(PJNOIndex) + "',PjtName='" + Trim(PjtName) + "',FinsGdNO='" + Trim(FinsGdNO) + "',SglPrtNO='" + Trim(SglPrtNO) + "',CommtNote='" + Trim(CommtNote) + "' WHERE SERIndex='" + Trim(OriSERIndex) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                      '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)         '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'SER��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempSERIndex As String)             '###########�����ĳɶ�Ӧ�ı��ֶ���
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������SERIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��SERIndex
  SqlStmt = "Delete FROM SER WHERE SERIndex='" + Trim(TempSERIndex) + "'"       '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsSER, True)            '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempSERIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempSERIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT SERIndex FROM SER WHERE SERIndex='" + Trim(TempSERIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSER As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsSER, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsSER.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function









