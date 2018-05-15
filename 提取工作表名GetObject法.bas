Attribute VB_Name = "提取工作表名GetObject法"
Sub 提取工作表名GetObject法()
Dim cat As Object, MyTable As Object
Dim n&, i&, j&, s$
Dim myPath$, myFile$
Application.ScreenUpdating = False '禁刷新
myPath = ThisWorkbook.path & "\"
myFile = Dir(myPath & "*.xls")
n = CreateObject("Scripting.FileSystemObject").getfolder(myPath).Files.Count - 1  '计算文件个数，减1不包括自身
ReDim arr(1 To 100, 1 To n)
Do While myFile <> ""
If myFile <> ThisWorkbook.Name Then '不等于本工作簿执行
j = j + 1
   i = 1
  arr(1, j) = Left(myFile, InStrRev(myFile, ".") - 1)  '去后辍
With GetObject(myPath & myFile) '使用 GetObject 函数可以访问文件
For i = 1 To .Worksheets.Count  '遍历文件的工作表数
arr(i + 1, j) = .Worksheets(i).Name
 Next
  End With
   GetObject(myPath & myFile).Close '关闭
End If
myFile = Dir
Loop
Application.ScreenUpdating = True   '自动计算
 Range("A1").Resize(i, j) = arr  '输出
End Sub
