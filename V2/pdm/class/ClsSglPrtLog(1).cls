VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSglPrtLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����ģ������Զ�Ӧ�ı��еĵ��ô���
'���������µ��ֶ���Ҫ����

Public SglPrtIndex As String                          '############����Ӧ�����ֶ��޸�
Public Seq As Integer
Public SglPrtVer As String                           '############����Ӧ�����ֶ��޸�
Public PrtUnit As String                          '############����Ӧ�����ֶ��޸�
Public Description As String                        '############����Ӧ�����ֶ��޸�
Public Applicant As String                          '############����Ӧ�����ֶ��޸�
Public ProductLine As String                             '############����Ӧ�����ֶ��޸�
Public IDSO As String                               '############����Ӧ�����ֶ��޸�
Public NewOldStatus As String                               '############����Ӧ�����ֶ��޸�
Public OpnDate As Date                              '############����Ӧ�����ֶ��޸�
Public ClosDate As Date                             '############����Ӧ�����ֶ��޸�
Public PJNOIndex As String                             '############����Ӧ�����ֶ��޸�
Public PjtName As String                            '############����Ӧ�����ֶ��޸�
Public ItemType As String                             '############����Ӧ�����ֶ��޸�
Public Location As String                             '############����Ӧ�����ֶ��޸�
Public CommtNote As String                             '############����Ӧ�����ֶ��޸�



'SglPrt��ģ���е�insert�ӹ��̣���������һ����¼
Public Sub Insert()
  
  '�������ݿ�,����һ����ģ���ڱ�����
  MyDB.openDB
  
  '#################����Insert���,Insert�еĸ��ַ�������SglPrtIndex,Applicant,ProductLine,Description...�ڵ��ô���ģ��Ĵ����и�ֵ
  SqlStmt = "INSERT INTO SglPrt (SglPrtIndex,Seq,SglPrtVer,PrtUnit,Description,Applicant,ProductLine,IDSO,NewOldStatus,OpnDate,ClosDate,PJNOIndex,PjtName,ItemType,Location,CommtNote)" _
     + " Values('" + Trim(SglPrtIndex) + "'," + Seq + ",'" + Trim(SglPrtVer) + "','" + Trim(PrtUnit) + "','" + Trim(Description) + "','" + Trim(Applicant) + "','" + Trim(ProductLine) + "','" + Trim(IDSO) + "','" + Trim(NewOldStatus) _
      + "','" + Trim(OpnDate) + "','" + Trim(ClosDate) + "','" + Trim(PJNOIndex) + "','" + Trim(PjtName) + "','" + Trim(ItemType) + "','" + Trim(Location) + "','" + Trim(CommtNote) + "')"
    
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSglPrt As ADODB.Recordset                    '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsSglPrt, True)       '###########�����ĳɶ�Ӧ�ı�����
  
  '�Ͽ������ݿ������
  MyDB.closeDB
End Sub
Public Function In_DB(ByVal TempSglPrtIndex As String) As Boolean                 '###########�����ĳɶ�Ӧ�ı��ֶ���
  '�������ݿ�
  MyDB.openDB
  
  '���ò�ѯ��䣬In_DB�е��ַ�������TempSglPrtIndex�ڵ��ô���ģ��Ĵ����и�ֵ
  '�ж�һ����¼�����Ƿ�����ø�����ֵ�ڼ�¼����SELECT
  SqlStmt = "SELECT SglPrtIndex FROM SglPrt WHERE SglPrtIndex='" + Trim(TempSglPrtIndex) + "'"      '###########�����ĳɶ�Ӧ�ı��ֶ��� ���б�����
  
  'ִ��SQL���
  Dim Isexe As Boolean
  Dim rsSglPrt As ADODB.Recordset                         '###########�����ĳɶ�Ӧ�ı�����
  Isexe = MyDB.exeSql(SqlStmt, rsSglPrt, False)           '###########�����ĳɶ�Ӧ�ı�����
  
  '�ж��Ƿ����
  If rsSglPrt.RecordCount <> 0 Then                         '###########�����ĳɶ�Ӧ�ı�����
    In_DB = True
  Else
    In_DB = False
  End If
  
  '�Ͽ������ݿ������
  MyDB.closeDB
  
End Function















