作者：付杨
链接：https://www.zhihu.com/question/40974557/answer/145193012
来源：知乎
著作权归作者所有。商业转载请联系作者获得授权，非商业转载请注明出处。

Private Function ToHexString(ByRef buf() As Byte) As String
    Dim i As Long, j As Long
    Dim nlen As Long
    Dim tmpHex As String
    Dim HexStr As String
    Dim tmpbuf() As Byte
    nlen = (UBound(buf) + 1) * 2
    ReDim tmpbuf(nlen - 1)
    j = 0
    For i = 0 To UBound(buf)
        HexStr = Hex(buf(i))
        If Len(HexStr) = 1 Then HexStr = "0" & HexStr
        tmpbuf(j) = Asc(Mid(HexStr, 1, 1))
        j = j + 1
        tmpbuf(j) = Asc(Mid(HexStr, 2, 1))
        j = j + 1
    Next
    ToHexString = StrConv(tmpbuf, vbUnicode)
End Function

Private Sub PostFile(ByVal PUrl As String, ByVal PFile As String)
Dim PostData, Boundary As String
Dim Upload_File  As String
Dim Http As Object
Dim fn As Integer
Dim fbuf() As Byte
Upload_File = PFile
'------------------打开Adodb.stream 流读取二进制文件------------------
fn = FreeFile()
ReDim fbuf(FileLen(Upload_File) - 1)
Open Upload_File For Binary As #fn
Get #fn, , fbuf
Close #fn
'-----------------构造POST数据 ----------------------
Boundary = "----WebKitFormBoundary1iVXNONaGEDOCghI"
PostData = "--" & Boundary & vbCrLf
PostData = PostData & "Content-Disposition: form-data; name=file; filename=F:\Work\E盘\mydata\VBSource\FrontClient2012 For SQL SERVER\20170208000018010000.jpg; payje=4.9; paytype=; payxsdbh=20170208000018010000; payxssj=15:53:40; payfdbh=0000;" & vbCrLf
PostData = PostData & "Content-Type: application/x-jpg" & vbCrLf
PostData = PostData & "" & vbCrLf
PostData = PostData & ToHexString(fbuf) & vbCrLf     '写入文件二进制内容PostData = PostData & "--" & Boundary & vbCrLf'---------------发送数据包-------------------------------------
Set Http = CreateObject("Msxml2.XMLHTTP")
Http.Open "POST", PUrl, True
Http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & Boundary
Http.send PostData
End Sub



