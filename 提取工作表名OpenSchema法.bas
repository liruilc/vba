Attribute VB_Name = "提取工作表名OpenSchema法"
Sub 提取工作表名OpenSchema法()

Dim arr, n&, i&, j&, s$

Dim myPath$, myFile$

Dim cnn As Object, rs As Object

myPath = ThisWorkbook.path & "\"

myFile = Dir(myPath & "*.xls")

n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '计算文件个数，减1不包括自身

ReDim arr(1 To 1000, 1 To n) '定义arr,最大工作表数1000

Do While myFile <> ""

If myFile <> ThisWorkbook.Name Then '不等于本工作簿执行

    j = j + 1

 i = 1

  arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1)  '去后辍
  Set cnn = CreateObject("ADODB.Connection")

  cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & myPath & myFile

Set rs = cnn.OpenSchema(20) 'Set rs = cnn.OpenSchema(adSchemaTables)，创建数据表记录集

  Do Until rs.EOF
  If rs.Fields("TABLE_TYPE") = "TABLE" Then

  i = i + 1

  s = Replace(rs("TABLE_NAME").Value, "'", "")    '去除"’"(数字工作表）

 If Right(s, 1) = "$" Then arr(i, j) = Left(s, Len(s) - 1) '去除$号

  End If

rs.MoveNext

Loop

End If

 myFile = Dir

Loop

rs.Close

cnn.Close

Set rs = Nothing

Set cnn = Nothing

Range("A1").Resize(i, j) = arr  '输出

End Sub
