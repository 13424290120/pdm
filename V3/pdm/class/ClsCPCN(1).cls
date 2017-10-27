VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCPCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public CPCNIndex As String                          '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public CPCNMP As String                             '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public IDSO As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�
Public PJNOIndex As String                             '############����Ӧ�����ֶ��޸�
Public PJTName As String                            '############����Ӧ�����ֶ��޸�
Public FinsGdNO As String                             '############����Ӧ�����ֶ��޸�
Public SglPrtNO As String                             '############����Ӧ�����ֶ��޸�
Public CommtNote As String                             '############����Ӧ�����ֶ��޸�
Public Reason As String




'CPCN��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������CPCNIndex,Applicant,CPCNMP,Description�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO CPCN (CPCNIndex,Applicant,CPCNMP,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,FinsGdNO,SglPrtNO,CommtNote,Reason)" _
     + " Values('" + Trim(CPCNIndex) + "','" + Trim(Applicant) + "','" + Trim(CPCNMP) + "','" + Trim(Description) + "','" + Trim(IDSO) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" _
      + Trim(PJNOIndex) + "','" + Trim(PJTName) + "','" + Trim(FinsGdNO) + "','" + Trim(SglPrtNO) + "','" + Trim(CommtNote) + "','" + Trim(Reason) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCPCN As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCPCN, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'CPCN��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriCPCNIndex As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '#################����Update���,Update�еĸ��ַ�������CPCNIndex,Applicant,CPCNMP,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriCPCNIndex
   SqlStmt = "Update CPCN Set CPCNIndex='" + Trim(CPCNIndex) + "',Applicant='" + Trim(Applicant) + "',CPCNMP='" + Trim(CPCNMP) + "',Description='" + Trim(Description) + "',IDSO='" + Trim(IDSO) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + _
            "',PJNOIndex='" + Trim(PJNOIndex) + "',PjtName='" + Trim(PJTName) + "',FinsGdNO='" + Trim(FinsGdNO) + "',SglPrtNO='" + Trim(SglPrtNO) + "',CommtNote='" + Trim(CommtNote) + "', Reason ='" + Trim(Reason) + "' WHERE CPCNIndex='" + Trim(OriCPCNIndex) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCPCN As ADODB.Recordset                      '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCPCN, True)         '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'CPCN��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempCPCNIndex As String)             '###########�����ĳɶ�Ӧ�ı��ֶ���
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������CPCNIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��CPCNIndex
  SqlStmt = "Delete FROM CPCN WHERE CPCNIndex='" + Trim(TempCPCNIndex) + "'"       '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCPCN As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCPCN, True)            '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempCPCNIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempCPCNIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT CPCNIndex FROM CPCN WHERE CPCNIndex='" + Trim(TempCPCNIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCPCN As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCPCN, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsCPCN.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function







