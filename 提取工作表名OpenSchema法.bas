Attribute VB_Name = "��ȡ��������OpenSchema��"
Sub ��ȡ��������OpenSchema��()

Dim arr, n&, i&, j&, s$

Dim myPath$, myFile$

Dim cnn As Object, rs As Object

myPath = ThisWorkbook.path & "\"

myFile = Dir(myPath & "*.xls")

n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '�����ļ���������1����������

ReDim arr(1 To 1000, 1 To n) '����arr,���������1000

Do While myFile <> ""

If myFile <> ThisWorkbook.Name Then '�����ڱ�������ִ��

    j = j + 1

 i = 1

  arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1)  'ȥ���
  Set cnn = CreateObject("ADODB.Connection")

  cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & myPath & myFile

Set rs = cnn.OpenSchema(20) 'Set rs = cnn.OpenSchema(adSchemaTables)���������ݱ��¼��

  Do Until rs.EOF
  If rs.Fields("TABLE_TYPE") = "TABLE" Then

  i = i + 1

  s = Replace(rs("TABLE_NAME").Value, "'", "")    'ȥ��"��"(���ֹ�����

 If Right(s, 1) = "$" Then arr(i, j) = Left(s, Len(s) - 1) 'ȥ��$��

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

Range("A1").Resize(i, j) = arr  '���

End Sub
