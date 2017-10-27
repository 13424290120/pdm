VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPJNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public PJNOIndex As String                          '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public Leader As String                             '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public IDSQ As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�

'PJNO��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������PJNOIndex,Applicant,Leader,Description�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO PJNO (PJNOIndex,Applicant,Leader,Description,IDSQ,OpnDate,ClosDate)" _
     + " Values('" + Trim(PJNOIndex) + "','" + Trim(Applicant) + "','" + Trim(Leader) + "','" + Trim(Description) + "','" + Trim(IDSQ) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsPJNO As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsPJNO, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'PJNO��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriPJNOIndex As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '#################����Update���,Update�еĸ��ַ�������PJNOIndex,Applicant,Leader,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriPJNOIndex
   SqlStmt = "Update PJNO Set PJNOIndex='" + Trim(PJNOIndex) + "',Applicant='" + Trim(Applicant) + "',Leader='" + Trim(Leader) + "',Description='" + Trim(Description) + "',IDSQ='" + Trim(IDSQ) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + "' WHERE PJNOIndex='" + Trim(OriPJNOIndex) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsPJNO As ADODB.Recordset                      '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsPJNO, True)         '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'PJNO��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempPJNOIndex As String)             '###########�����ĳɶ�Ӧ�ı��ֶ���
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������PJNOIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��PJNOIndex
  SqlStmt = "Delete FROM PJNO WHERE PJNOIndex='" + Trim(TempPJNOIndex) + "'"       '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsPJNO As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsPJNO, True)            '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempPJNOIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempPJNOIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT PJNOIndex FROM PJNO WHERE PJNOIndex='" + Trim(TempPJNOIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsPJNO As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsPJNO, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsPJNO.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function





