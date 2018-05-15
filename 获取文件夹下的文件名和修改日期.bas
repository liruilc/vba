Attribute VB_Name = "获取文件夹下的文件名和修改日期"
Sub FileTest()
   
    '获取文件名和修改日期
    Dim path, file, i
    path = "c:\windows\system32\"
    '只是文件名 不是文件对象
    file = Dir(path & "*.*")
    
    With ThisWorkbook.Worksheets(1)
        .Cells(1, 1) = "文件名"
        .Cells(1, 2) = "修改日期"
        i = 2
        Do While file <> "" And file <> ThisWorkbook.Name
            .Cells(i, 1) = file
            .Cells(i, 2) = FileDateTime(path & file)
            i = i + 1
            '一定要 不然死循环
            file = Dir
        Loop

    End With
    
    '自动适应列宽
   ' Columns("A:B").AutoFit
    
End Sub

Sub SelectFolder()

    '选择单一文件
    'www.okexcel.com.cn
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
        'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
          'MsgBox "您选择的文件夹是：" & .SelectedItems(1), vbOKOnly + vbInformation, "智能Excel"
          Cells(1, 1) = .SelectedItems(1)
        End If
    End With
End Sub
