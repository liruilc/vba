Attribute VB_Name = "提取工作表名Open法"
Sub 提取工作表名Open法()

Dim arr

 Dim n&, i&, j&, s$

Dim wb As Workbook, sht As Worksheet, wbk As Workbook

Dim myPath$, myFile$

Application.ScreenUpdating = False '禁刷新 Application.Calculation = xlManual '禁计算

Set wbk = ThisWorkbook

 myPath = ThisWorkbook.path & "\"

myFile = Dir(myPath & "*.xls")

 n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '计算文件个数，减1不包括自身

ReDim arr(1 To 1000, 1 To n)

 Do While myFile <> ""

If myFile <> wbk.Name Then

  j = j + 1

  i = 1

 arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1) '去后辍

 Set wb = Workbooks.Open(myPath & "\" & myFile)   '打开工作簿

 For Each sht In wb.Sheets '遍历工作表

   i = i + 1

  arr(i, j) = sht.Name

  Next

 wb.Close

End If

 myFile = Dir

 Loop

 wbk.ActiveSheet.Range("A1").Resize(i, j) = arr '输出

Application.Calculation = xlAutomatic '刷新 Application.ScreenUpdating = True   '自动计算

End Sub
