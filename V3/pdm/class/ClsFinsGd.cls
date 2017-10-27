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
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public FinsGdIndex As String                          '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public ProductLine As String                             '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public IDSO As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�
Public PJNOIndex As String                             '############����Ӧ�����ֶ��޸�
Public PjtName As String                            '############����Ӧ�����ֶ��޸�
Public ItemType As String                             '############����Ӧ�����ֶ��޸�
Public Location As String                             '############����Ӧ�����ֶ��޸�
Public CommtNote As String                             '############����Ӧ�����ֶ��޸�



'FinsGd��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������FinsGdIndex,Applicant,ProductLine,Description�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO FinsGd (FinsGdIndex,Applicant,ProductLine,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,ItemType,Location,CommtNote)" _
     + " Values('" + Trim(FinsGdIndex) + "','" + Trim(Applicant) + "','" + Trim(ProductLine) + "','" + Trim(Description) + "','" + Trim(IDSO) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" _
      + Trim(PJNOIndex) + "','" + Trim(PjtName) + "','" + Trim(ItemType) + "','" + Trim(Location) + "','" + Trim(CommtNote) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'FinsGd��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriFinsGdIndex As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '#################����Update���,Update�еĸ��ַ�������FinsGdIndex,Applicant,ProductLine,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriFinsGdIndex
   SqlStmt = "Update FinsGd Set FinsGdIndex='" + Trim(FinsGdIndex) + "',Applicant='" + Trim(Applicant) + "',ProductLine='" + Trim(ProductLine) + "',Description='" + Trim(Description) + "',IDSO='" + Trim(IDSO) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + _
            "',PJNOIndex='" + Trim(PJNOIndex) + "',PjtName='" + Trim(PjtName) + "',ItemType='" + Trim(ItemType) + "',Location='" + Trim(Location) + "',CommtNote='" + Trim(CommtNote) + "' WHERE FinsGdIndex='" + Trim(OriFinsGdIndex) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                      '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, True)         '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'FinsGd��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempFinsGdIndex As String)             '###########�����ĳɶ�Ӧ�ı��ֶ���
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������FinsGdIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��FinsGdIndex
  SqlStmt = "Delete FROM FinsGd WHERE FinsGdIndex='" + Trim(TempFinsGdIndex) + "'"       '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, True)            '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempFinsGdIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempFinsGdIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT FinsGdIndex FROM FinsGd WHERE FinsGdIndex='" + Trim(TempFinsGdIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsFinsGd As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsFinsGd, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsFinsGd.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function











