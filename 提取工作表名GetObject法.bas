Attribute VB_Name = "��ȡ��������GetObject��"
Sub ��ȡ��������GetObject��()
Dim cat As Object, MyTable As Object
Dim n&, i&, j&, s$
Dim myPath$, myFile$
Application.ScreenUpdating = False '��ˢ��
myPath = ThisWorkbook.path & "\"
myFile = Dir(myPath & "*.xls")
n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '�����ļ���������1����������
ReDim arr(1 To 100, 1 To n)
Do While myFile <> ""
If myFile <> ThisWorkbook.Name Then '�����ڱ�������ִ��
j = j + 1
   i = 1
  arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1)  'ȥ���
With GetObject(myPath & myFile) 'ʹ�� GetObject �������Է����ļ�
For i = 1 To .Worksheets.Count  '�����ļ��Ĺ�������
arr(i + 1, j) = .Worksheets(i).Name
 Next
  End With
   GetObject(myPath & myFile).Close '�ر�
End If
myFile = Dir
Loop
Application.ScreenUpdating = True   '�Զ�����
 Range("A1").Resize(i, j) = arr  '���
End Sub
