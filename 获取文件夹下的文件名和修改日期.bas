Attribute VB_Name = "��ȡ�ļ����µ��ļ������޸�����"
Sub FileTest()
   
    '��ȡ�ļ������޸�����
    Dim path, file, i
    path = "c:\windows\system32\"
    'ֻ���ļ��� �����ļ�����
    file = Dir(path & "*.*")
    
    With ThisWorkbook.Worksheets(1)
        .Cells(1, 1) = "�ļ���"
        .Cells(1, 2) = "�޸�����"
        i = 2
        Do While file <> "" And file <> ThisWorkbook.Name
            .Cells(i, 1) = file
            .Cells(i, 2) = FileDateTime(path & file)
            i = i + 1
            'һ��Ҫ ��Ȼ��ѭ��
            file = Dir
        Loop

    End With
    
    '�Զ���Ӧ�п�
   ' Columns("A:B").AutoFit
    
End Sub

Sub SelectFolder()

    'ѡ��һ�ļ�
    'www.okexcel.com.cn
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
        'FileDialog ����� Show ������ʾ�Ի��򣬲��ҷ��� -1��������� OK���� 0��������� Cancel����
          'MsgBox "��ѡ����ļ����ǣ�" & .SelectedItems(1), vbOKOnly + vbInformation, "����Excel"
          Cells(1, 1) = .SelectedItems(1)
        End If
    End With
End Sub
