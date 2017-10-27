VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCNCSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public CNCSNIndex As String                          '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public CPCNMP As String                             '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public IDSO As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�
Public PJNOIndex As String                             '############����Ӧ�����ֶ��޸�
Public PjtName As String                            '############����Ӧ�����ֶ��޸�
Public FinsGdNO As String                             '############����Ӧ�����ֶ��޸�
Public SglPrtNO As String                             '############����Ӧ�����ֶ��޸�
Public CommtNote As String                             '############����Ӧ�����ֶ��޸�



'CNCSN��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������CNCSNIndex,Applicant,CPCNMP,Description�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO CNCSN (CNCSNIndex,Applicant,CPCNMP,Description,IDSO,OpnDate,ClosDate,PJNOIndex,PjtName,FinsGdNO,SglPrtNO,CommtNote)" _
     + " Values('" + Trim(CNCSNIndex) + "','" + Trim(Applicant) + "','" + Trim(CPCNMP) + "','" + Trim(Description) + "','" + Trim(IDSO) + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" _
      + Trim(PJNOIndex) + "','" + Trim(PjtName) + "','" + Trim(FinsGdNO) + "','" + Trim(SglPrtNO) + "','" + Trim(CommtNote) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCNCSN As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCNCSN, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'CNCSN��ģ���е�Update�ӹ��̣����ڸ���һ����¼
Public Sub Update(ByVal OriCNCSNIndex As String)
   
  '�������ݿ�
  MyDB.openDB
  
   '#################����Update���,Update�еĸ��ַ�������CNCSNIndex,Applicant,CPCNMP,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
   ' update����е��ر�֮���ǻ�Ҫһ��Դ��¼������λ��OriCNCSNIndex
   SqlStmt = "Update CNCSN Set CNCSNIndex='" + Trim(CNCSNIndex) + "',Applicant='" + Trim(Applicant) + "',CPCNMP='" + Trim(CPCNMP) + "',Description='" + Trim(Description) + "',IDSO='" + Trim(IDSO) + "',OpnDate='" + Trim(OpnDate) + "',ClosDate='" + Trim(ClosDate) + _
            "',PJNOIndex='" + Trim(PJNOIndex) + "',PjtName='" + Trim(PjtName) + "',FinsGdNO='" + Trim(FinsGdNO) + "',SglPrtNO='" + Trim(SglPrtNO) + "',CommtNote='" + Trim(CommtNote) + "' WHERE CNCSNIndex='" + Trim(OriCNCSNIndex) + "'"
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCNCSN As ADODB.Recordset                      '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCNCSN, True)         '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
'CNCSN��ģ���е�Delete�ӹ��̣�����ɾ��һ����¼
Public Sub Delete(ByVal TempCNCSNIndex As String)             '###########�����ĳɶ�Ӧ�ı��ֶ���
   '�������ݿ�
  MyDB.openDB
  
  '����Delete��䣬Delete�е��ַ�������CNCSNIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  'Delete����е��ر�֮����ֻҪһ����¼������λ��CNCSNIndex
  SqlStmt = "Delete FROM CNCSN WHERE CNCSNIndex='" + Trim(TempCNCSNIndex) + "'"       '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCNCSN As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCNCSN, True)            '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempCNCSNIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempCNCSNIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT CNCSNIndex FROM CNCSN WHERE CNCSNIndex='" + Trim(TempCNCSNIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsCNCSN As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsCNCSN, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsCNCSN.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function









