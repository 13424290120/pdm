Attribute VB_Name = "ImageFunc"
Const ChunkSize As Long = 100
Const BlockSize As Long = 100
Const TempFile As String = "tempfile.tmp"

Dim byteData() As Byte '定义数据块数组
Dim DiskFile As String '图像文件名
Dim NumBlocks As Long '定义数据块个数
Dim FileLength As Long '标识文件长度
Dim LeftOver As Long '定义剩余字节长度
Dim SourceFile As Long '定义自由文件号
Dim byteChunk() As Byte
Dim i As Long '定义循环变量

Public Sub ShowImage(Image1 As Image, _
                     Adodc1 As Adodc)
    Erase byteChunk()
    FieldSize = Adodc1.Recordset.Fields(2).ActualSize
    If FieldSize <= 0 Then
        Image1.Picture = LoadPicture("")
        Exit Sub
    End If
    '提供一个尚未使用的文件号
    SourceFile = FreeFile
    '打开文件
    Open TempFile For Binary Access Write As SourceFile
    '计算数据块
    NumBlocks = FieldSize \ BlockSize
    LeftOver = FieldSize Mod BlockSize '得到剩余字节数
    '分块读取图像数据，并写入到文件中
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
    '将文件装入到Image1控件中
    Image1.Picture = LoadPicture(TempFile)
    '删除临时文件
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
    '提供一个尚未使用的文件号
    SourceFile = FreeFile
    '打开文件
    Open ImageFile For Binary Access Read As SourceFile
    '得到文件长度
    FileLength = LOF(SourceFile)
    '判断文件是否存在
    If FileLength = 0 Then
        Close SourceFile
        MsgBox DiskFile & "无内容或不存在!"
    Else
        NumBlocks = FileLength \ BlockSize '得到数据块的个数
        LeftOver = FileLength Mod BlockSize '得到剩余字节数
        Adodc1.Recordset.Fields(2).Value = Null
        ReDim byteData(BlockSize) '重新定义数据块的大小
        For i = 1 To NumBlocks
            Get SourceFile, , byteData() '读到内存块中
            Adodc1.Recordset.Fields(2).AppendChunk byteData() '写入FLD
        Next i
        ReDim byteData(LeftOver) '重新定义数据块的大小
        Get SourceFile, , byteData() '读到内存块中
        Adodc1.Recordset.Fields(2).AppendChunk byteData() '写入FLD
        Close SourceFile '关闭源文件
        Adodc1.Recordset.Update
    End If
End Sub
