Attribute VB_Name = "��ȡ��������Open��"
Sub ��ȡ��������Open��()

Dim arr

 Dim n&, i&, j&, s$

Dim wb As Workbook, sht As Worksheet, wbk As Workbook

Dim myPath$, myFile$

Application.ScreenUpdating = False '��ˢ�� Application.Calculation = xlManual '������

Set wbk = ThisWorkbook

 myPath = ThisWorkbook.path & "\"

myFile = Dir(myPath & "*.xls")

 n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '�����ļ���������1����������

ReDim arr(1 To 1000, 1 To n)

 Do While myFile <> ""

If myFile <> wbk.Name Then

  j = j + 1

  i = 1

 arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1) 'ȥ���

 Set wb = Workbooks.Open(myPath & "\" & myFile)   '�򿪹�����

 For Each sht In wb.Sheets '����������

   i = i + 1

  arr(i, j) = sht.Name

  Next

 wb.Close

End If

 myFile = Dir

 Loop

 wbk.ActiveSheet.Range("A1").Resize(i, j) = arr '���

Application.Calculation = xlAutomatic 'ˢ�� Application.ScreenUpdating = True   '�Զ�����

End Sub
