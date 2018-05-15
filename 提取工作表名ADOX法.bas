Attribute VB_Name = "提取工作表名ADOX法"
Sub 提取工作表名ADOX法()

Dim cat, MyTable As Object

 Dim n&, i&, j&, s$

 Dim myPath$, myFile$

myPath = ThisWorkbook.path & "\"

myFile = Dir(myPath & "*.xls")

 n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '计算文件个数，减1不包括自身

ReDim arr(1 To 1000, 1 To n)

Do While myFile <> ""

If myFile <> ThisWorkbook.Name Then '不等于本工作簿执行

  j = j + 1

i = 1

 arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1)  '去后辍

 Set cat = CreateObject("ADOX.Catalog")

 'cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & myPath & myFile
cat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & myPath & myFile
 For Each MyTable In cat.Tables

If MyTable.Type = "TABLE" Then

  s = Replace(MyTable.Name, "'", "")

  If Right(s, 1) = "$" Then

  i = i + 1

arr(i, j) = Left(s, Len(s) - 1)

End If

End If

Next

End If

myFile = Dir

Loop

Set cat = Nothing

Set MyTable = Nothing

Range("A1").Resize(i, j) = arr  '输出

End Sub
