Attribute VB_Name = "获取指定路径下的文件夹名称"
Sub GetFolderName() '获取指定路径下的文件夹名称
Dim fs As Object
Dim afile
afile = "c:\windows\system32"
n = 2
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.getfolder(afile)
For Each fd In f.subfolders
Cells(1, 1) = afile
Cells(n, 1) = fd.Name
n = n + 1
Next
Set f = Nothing
Set fs = Nothing
End Sub
