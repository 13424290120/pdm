VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRFSRFQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public RFSRFQIndex As String                        '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public Leader As String                             '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public IDSQ As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�

'RFSRFQ��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������RFSRFQIndex,Applicant,Leader,Description�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO RFSRFQ (RFSRFQIndex,Applicant,Leader,Description,IDSQ,OpnDate,ClosDate)" _
     + " Values('" + Trim(RFSRFQIndex) + "','" + Trim(Applicant) + "','" + Trim(Leader) + "','" + Trim(Description) + "','" + Trim(IDSQ) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsRFSRFQ As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsRFSRFQ, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'RFSRFQ��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriRFSRFQIndex As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '#################����Update���,Update�еĸ��ַ�������RFSRFQIndex,Applicant,Leader,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriRFSRFQIndex
   SqlStmt = "Update RFSRFQ Set RFSRFQIndex='" + Trim(RFSRFQIndex) + "',Applicant='" + Trim(Applicant) + "',Leader='" + Trim(Leader) + "',Description='" + Trim(Description) + "',IDSQ='" + Trim(IDSQ) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + "' WHERE RFSRFQIndex='" + Trim(OriRFSRFQIndex) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsRFSRFQ As ADODB.Recordset                      '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsRFSRFQ, True)         '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'RFSRFQ��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempRFSRFQIndex As String)             '###########�����ĳɶ�Ӧ�ı��ֶ���
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������RFSRFQIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��RFSRFQIndex
  SqlStmt = "Delete FROM RFSRFQ WHERE RFSRFQIndex='" + Trim(TempRFSRFQIndex) + "'"       '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsRFSRFQ As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsRFSRFQ, True)            '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempRFSRFQIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempRFSRFQIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT RFSRFQIndex FROM RFSRFQ WHERE RFSRFQIndex='" + Trim(TempRFSRFQIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsRFSRFQ As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsRFSRFQ, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsRFSRFQ.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function



