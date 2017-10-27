Attribute VB_Name = "ImageFunc"
Const ChunkSize As Long = 100
Const BlockSize As Long = 100
Const TempFile As String = "tempfile.tmp"

Dim byteData() As Byte '�������ݿ�����
Dim DiskFile As String 'ͼ���ļ���
Dim NumBlocks As Long '�������ݿ����
Dim FileLength As Long '��ʶ�ļ�����
Dim LeftOver As Long '����ʣ���ֽڳ���
Dim SourceFile As Long '���������ļ���
Dim byteChunk() As Byte
Dim i As Long '����ѭ������

Public Sub ShowImage(Image1 As Image, _
                     Adodc1 As Adodc)
    Erase byteChunk()
    FieldSize = Adodc1.Recordset.Fields(2).ActualSize
    If FieldSize <= 0 Then
        Image1.Picture = LoadPicture("")
        Exit Sub
    End If
    '�ṩһ����δʹ�õ��ļ���
    SourceFile = FreeFile
    '���ļ�
    Open TempFile For Binary Access Write As SourceFile
    '�������ݿ�
    NumBlocks = FieldSize \ BlockSize
    LeftOver = FieldSize Mod BlockSize '�õ�ʣ���ֽ���
    '�ֿ��ȡͼ�����ݣ���д�뵽�ļ���
    If LeftOver <> 0 Then
        ReDim byteChunk(LeftOver)
        byteChunk() = Adodc1.Recordset.Fields(2).GetChunk(LeftOver)
        Put SourceFile, , byteChunk()
    End If
    For i = 1 To NumBlocks
        ReDim byteChunk(BlockSize)
        byteChunk() = Adodc1.Recordset.Fields(2).GetChunk(BlockSize)
        Put SourceFile, , byteChunk()
    Next i
    Close SourceFile
    '���ļ�װ�뵽Image1�ؼ���
    Image1.Picture = LoadPicture(TempFile)
    'ɾ����ʱ�ļ�
    Kill (TempFile)
End Sub

Public Sub SaveImage(ByVal ImageFile As String, _
                     Adodc1 As Adodc)
    
    If Adodc1.Recordset.BOF = True Or Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
    If ImageFile = "" Then
        Exit Sub
    End If
    '�ṩһ����δʹ�õ��ļ���
    SourceFile = FreeFile
    '���ļ�
    Open ImageFile For Binary Access Read As SourceFile
    '�õ��ļ�����
    FileLength = LOF(SourceFile)
    '�ж��ļ��Ƿ����
    If FileLength = 0 Then
        Close SourceFile
        MsgBox DiskFile & "�����ݻ򲻴���!"
    Else
        NumBlocks = FileLength \ BlockSize '�õ����ݿ�ĸ���
        LeftOver = FileLength Mod BlockSize '�õ�ʣ���ֽ���
        Adodc1.Recordset.Fields(2).Value = Null
        ReDim byteData(BlockSize) '���¶������ݿ�Ĵ�С
        For i = 1 To NumBlocks
            Get SourceFile, , byteData() '�����ڴ����
            Adodc1.Recordset.Fields(2).AppendChunk byteData() 'д��FLD
        Next i
        ReDim byteData(LeftOver) '���¶������ݿ�Ĵ�С
        Get SourceFile, , byteData() '�����ڴ����
        Adodc1.Recordset.Fields(2).AppendChunk byteData() 'д��FLD
        Close SourceFile '�ر�Դ�ļ�
        Adodc1.Recordset.Update
    End If
End Sub
