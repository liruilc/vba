VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Macroinstructions 
   Caption         =   "Macroinstruction"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   OleObjectBlob   =   "Macroinstructions.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "Macroinstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
       ActiveWindow.SmallScroll Down:=-21
       ChDir "D:\my file\BOM backup browsers\Datebase Update"
             Workbooks.Open Filename:= _
             "D:\my file\BOM backup browsers\Datebase Update\Assembly Daily Plan.xls"
                 Range("B2:J100").Select
                 Selection.ClearContents
                 Range("B2").Select
            For k = 1 To 2
                  ActiveWindow.ActivateNext
                  Range("A2").Select
              Exit For
                  Range("B2").Select
            Next
  With fuzhi
'----------------------------------------------------------------------
    ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Select   'ѡ�й�������������������½ǵ�Ԫ��
'----------------------------------------------------------------------
        Dim x, y, z As Byte
        Dim rng, w, ii As Range  '����I�ı���
        Set rng = Cells(Rows.Count, 1).End(xlUp)
              w = rng.Row
              z = 100
              x = 0
              y = 0
      Set ii = [ActiveCell].Offset(x - (z - w), y - 4) 'I�Ķ�ֵ̬Ϊ��ǰ������ĵ�ǰ���Ԫ���ַ ;
               Range("a2", ii).Select                   'ѡ��ָ����Ԫ������,"Activecell"Ϊ��ǰ�������л�ĵ�Ԫ���ַ; _
                                                          "[��ǰ���Ԫ��].offset(����������,����������)
                 Selection.Copy                              '�������ݶ�̬ѡ��ʹ�õ�Ԫ������
      Set ii = Nothing                                                                              'ii����ֵ�ÿ�
      Set rng = Nothing
   End With
            With zhantie
                  Windows("Assembly Daily Plan.xls").Activate
                                    Range("B2").Select
                                      Selection.PasteSpecial Paste:= _
                                                                       xlPasteValues, _
                                                         Operation:= _
                                                                       xlNone                        '����ѡ����ճ��
                                                                                                    'ճ������Ϊ��ֵ (xlPasteValues)��ֵ
                                                                                                    '����ѡ��Ϊ��ֵ���ޣ�xlNone
                                                             Range("A2").Select
                                                         ActiveCell.FormulaR1C1 = "1"
                                                             Range("A3").Select
                                                         ActiveCell.FormulaR1C1 = "2"
                                                          Range("A2:A3").Select
                                                 Selection.AutoFill Destination:=Range("A2:A100"), Type:=xlFillDefault
                                            ActiveWorkbook.Save
             End With
                       With r
                          Cells(Cells.Rows.Count, 2).End(xlUp).Offset(1, 0).Select
                                Dim hh As Range                                                       '����hh�ı���
                                Set hh = [ActiveCell]                                                 'I�Ķ�ֵ̬Ϊ��ǰ������ĵ�ǰ���Ԫ���ַ ;
                                       Range("A100", hh).Select
                                           Selection.ClearContents
                                           Range("a1").Select
                                Set hh = Nothing
                      End With
       ActiveWorkbook.Save
With gx
            For l = 1 To 2
                  ActiveWindow.ActivateNext
                  Range("A2").Select
              Exit For
                  Range("B2").Select
            Next
ActiveWorkbook.FollowHyperlink "D:\my file\BOM backup browsers\Warehouse material list.mdb", NewWindow:=True
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.QueryTables.Add(Connection:=Array( _
        "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=Admin;Data Source=D:\my file\BOM backup browsers\Warehouse material list.md" _
        , _
        "b;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Passwo" _
        , _
        "rd="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transa" _
        , _
        "ctions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Do" _
        , _
        "n't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False" _
        ), Destination:=Range("A1"))
        .CommandType = xlCmdTable
        .CommandText = Array("All material list")
        .Name = "Warehouse material list"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = _
        "D:\my file\BOM backup browsers\Warehouse material list.mdb"
        .Refresh BackgroundQuery:=False
    End With
    d = Date + 1
    ActiveWorkbook.Save
    ChDir "D:\my file\BOM backup browsers\lswj\�ռƻ�\12"
    ActiveWorkbook.SaveAs Filename:="D:\my file\BOM backup browsers\lswj\�ռƻ�\12" & "\" & Format(d, "yyyy.mm.dd") & "�ƻ��ֽ�.xls", _
        FileFormat:=xlNormal, Password:="", WriteResPassword:="" ', _
       ReadOnlyRecommended:=False, CreateBackup:=False
     'Kill ("D:\my file\BOM backup browsers\Datebase Update" & "\" & Format(d, "mm-dd") & " assembly plan.xls")
     End With
     With ������ݱ�������S
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "All_material_list"
    Sheets(Array("Sheet1", "Sheet2", "Sheet3")).Select
    Sheets("Sheet3").Activate
    ActiveWindow.SelectedSheets.Delete
    ActiveWorkbook.Save
With qinglie
  Application.ScreenUpdating = False '�ر���Ļˢ��
   Range("A:B,D:F,L:L,N:P").Select
    Range("N1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Application.ScreenUpdating = True '����Ļˢ��
    Range("A:H").Select
        ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
        "All_material_list!a:h").CreatePivotTable TableDestination:="", TableName _
        :="PivotTable1", DefaultVersion:=xlPivotTableVersion10
    ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
    ActiveSheet.Cells(3, 1).Select
    Range("A3").Select
    End With
       With ActiveSheet.PivotTables("PivotTable1").PivotFields("MLFB")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("����������")
        .Orientation = xlColumnField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("�߱�����")
        .Orientation = xlColumnField
        .Position = 3
    End With
    Range("C6").Select
    Selection.Delete
    Range("C5").Select
    Selection.Delete
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BOM comp")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("comp description")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Material Qty"), "Count of Material Qty", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material type")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .Orientation = xlRowField
        .Position = 4
    End With
    Range("B10").Select
    Selection.Delete
    Range("B9").Select
    Selection.Delete
    Range("D8").Select
    Selection.Delete
    Range("E20").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Material Qty"). _
        Function = xlSum
    Range("A3").Select
    With jia
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .PivotItems("BLL").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("����").Visible = False
        .PivotItems("С����").Visible = False
        .PivotItems("PH").Visible = False
        '.PivotItems("С����").Visible = False
        '.PivotItems("С����").Visible = False
        '.PivotItems("С����").Visible = False
        '.PivotItems("С����").Visible = False
        Sheets("Sheet5").Select
       Sheets("Sheet5").Name = "Sheet1"
    End With
    Columns("B:B").ColumnWidth = 25
    Cells.Select
    Selection.ColumnWidth = 7#
    Range("A3").Select
    Cells.Select
    Selection.Copy
    ActiveWorkbook.Save
    Windows("Assembly Daily Plan.xls").Activate
    ActiveWindow.Close                                          '�Ҽ��ر�"Assembly Daily Plan.xls"�˹�����
    End With
  End With
End Sub
'Option Explicit
Private Sub CommandButton2_Click()
    Dim aaa As Shape
    Dim b, c As Integer
    Dim dtif1, dpdf1, dxls1, dtif2, dpdf2, dxls2, dtif3, dpdf3, tj1, tj2, tj3, _
    tj4, tj5, tj6, tj7, pd1, pd2, pd3, pd4, pd5, pd6 As String
    Dim d As Range
    With ActiveSheet
        For Each aaa In .Shapes
            If aaa.Type = 13 Then
                aaa.Delete
            End If
        Next
             Range("A3").Select   '^_____________________________________________________________
    ActiveSheet.Pictures.Insert("D:\qita\p\fangdajing.gif").Select
    Selection.ShapeRange.LockAspectRatio = msoTrue
    Selection.ShapeRange.Height = 20 '�߶�
    Selection.ShapeRange.Width = 17.5 '���
    Selection.ShapeRange.Rotation = 85#
    Selection.ShapeRange.IncrementLeft 58#
    Selection.ShapeRange.IncrementTop 2
    Selection.OnAction = "macroinstruction.xls!fangda.fangda"
        Range("A3:A11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .AddIndent = False
        .ReadingOrder = xlContext
    End With
        'Rows("5:45").Select
        'Selection.RowHeight = 12.5
        'Columns("AJ:AU").Select
        'Selection.ColumnWidth = 3.5
        'Range("A1").Select         'v___________________________________________________________
        For b = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            For c = 1 To 1 Step 2
                dtif1 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_1" & ".tif"
                dpdf1 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_1" & ".pdf"
                dxls1 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_1" & ".xls"
                dtif2 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_2" & ".tif"
                dpdf2 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_2" & ".pdf"
                dxls2 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_2" & ".xls"
                dtif3 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_3" & ".tif"
                dpdf3 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_3" & ".pdf"
                    tj1 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_1_.jpg"
                    tj2 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_2_.jpg"
                    tj3 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_1_.tif"
                    tj4 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_2_.tif"
                    tj5 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_3_.tif"
                    tj6 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_4_.tif"
                    tj7 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_5_.tif"
                    pd1 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_1_.pdf"
                    pd2 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_2_.pdf"
                    pd3 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_3_.pdf"
                    pd4 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_4_.pdf"
                    pd5 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_5_.pdf"
                    pd6 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "_6_.pdf"
                  If Dir(dtif1) <> "" Then
                      Set aaa = .Shapes.AddPicture(dtif1, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 3) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dpdf1) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 4), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_1" & ".pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 44) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dxls1) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 5), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_1" & ".xls", ScreenTip:="XLS�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 45) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dtif2) <> "" Then
                      Set aaa = .Shapes.AddPicture(dtif2, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 6) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dpdf2) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 7), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_2" & ".pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 47) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dxls2) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 8), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_2" & ".xls", ScreenTip:="XLS�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 48) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dtif3) <> "" Then
                      Set aaa = .Shapes.AddPicture(dtif3, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 9) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 49) = "" '��ͼƬ��ʾ������
                     End If
                    If Dir(dpdf3) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 10), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_3" & ".pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                      End With
                     Else
                   .Cells(b, c + 50) = "" '��ͼƬ��ʾ������
                   End If
                  If Dir(tj1) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj1, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(tj2) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj2, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(tj3) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj3, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(tj4) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj4, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(tj5) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj5, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(tj6) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj6, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(tj7) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj7, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) 'ͼƬ��ʾ������
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '����λ��
                      .Left = d.Left + 1   '���λ��
                      .Width = d.Width - 1.5   '���
                      .Height = d.Height - 1.5  '�߶�
                      .TopLeftCell = ""   '���Ͻǵ�Ԫ��
                    End With
                      Else
                   .Cells(b, c + 43) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(pd1) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 40), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_1_.pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(pd2) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 41), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_2_.pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(pd3) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 42), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_3_.pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(pd4) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 43), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_4_.pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(pd5) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 44), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_5_.pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                  If Dir(pd6) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 45), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_6_.pdf", ScreenTip:="PDF�ļ�"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '��ͼƬ��ʾ������
                     End If
                Next
              Next
         End With
       Set aaa = Nothing
    Set d = Nothing
End Sub

Private Sub P_Click()
'
' Macro1 Macro
' Macro recorded 6/25/2011 by Z002TE2Z;jingrui.li@siemens.com
'

'
    Columns("B:K").Select
    Range("B2").Activate
    Columns("B:K").EntireColumn.AutoFit
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNA(LEFT(R[1]C,13)),"""",LEFT(R[1]C,13))"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(R[-2]C,Sheet1!C[-2]:C[-1],2,0)),""������"",VLOOKUP(R[-2]C,Sheet1!C[-2]:C[-1],2,0))"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(R[-2]C,Sheet1!C[-2]:C[-1],2,0)),""<������>"",VLOOKUP(R[-2]C,Sheet1!C[-2]:C[-1],2,0))"
    Range("K6").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("L17").Select
    Sheets("Sheet1").Select
    Sheets.Add
    Cells.Select
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "list"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "sap No"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "mlfb"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "qty"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "mlfb"
    Range("A1:D1").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("A2:A4").Select
    Selection.AutoFill Destination:=Range("A2:A61"), Type:=xlFillDefault
    Range("A2:A61").Select
    ActiveWindow.SmallScroll Down:=-45
    Range("C13").Select
    ActiveWorkbook.Names.Add Name:="list", RefersToR1C1:="=Sheet2!R2C1:R61C1"
    ActiveWorkbook.Names.Add Name:="list", RefersToR1C1:="=Sheet2!R2C1:R61C1"
    Sheets("picking list for assembly").Select
    Range("L11:N11").Select
    Range("N11").Activate
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    Range("L12:N12").Select
    Range("N12").Activate
    With Selection.Interior
        .ColorIndex = 4
        .Pattern = xlSolid
    End With
    Range("L11:N12").Select
    Range("N12").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("L11:N11").Select
    Range("N11").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Range("N12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=list"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("N12").Select
    Sheets("Sheet2").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],Sheet1!R2C1:R65536C2,2,0)"
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(RC[-2],Sheet1!R2C1:R65536C2,2,0)),""����������"",VLOOKUP(RC[-2],Sheet1!R2C1:R65536C2,2,0))"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D61"), Type:=xlFillDefault
    Range("D2:D61").Select
    Range("G67").Select
    ActiveWindow.SmallScroll Down:=-42
    Columns("D:D").EntireColumn.AutoFit
    Columns("A:D").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("H8").Select
    ActiveWindow.SmallScroll Down:=-42
    Sheets("picking list for assembly").Select
    Range("L12").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],Sheet2!R2C1:R61C2,2,0)"
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(RC[2],Sheet2!R2C1:R61C2,2,0)),""���õ�ַ������"",VLOOKUP(RC[2],Sheet2!R2C1:R61C2,2,0))"
    Range("M19").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "list"
    Range("L11:N12").Select
    Range("N12").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("M12").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Sheet2!R2C2:R61C3,2,0)"
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(RC[-1],Sheet2!R2C2:R61C3,2,0)),""��ֵ"",VLOOKUP(RC[-1],Sheet2!R2C2:R61C3,2,0))"
    Range("O20").Select
    Sheets("Sheet1").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveWindow.ScrollRow = 13527
    ActiveWindow.ScrollRow = 1
    Range("A2").Select
    Sheets("Sheet2").Select
    ActiveWindow.SmallScroll Down:=-36
    Range("B2").Select
    Sheets("picking list for assembly").Select
    Range("N13").Select
    Columns("I:I").EntireColumn.AutoFit
    Sheets("picking list for assembly").Select
    Range("N13").Select
End Sub

Private Sub ������ʹ��ʱ������S_Click()
Range("A2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.AutoFill Destination:=Range("A2:D3"), Type:=xlFillDefault
    Range("A2:D3").Select
    Range("A2:D2").Select
    ActiveCell.FormulaR1C1 = "�� �� ��"
    With ActiveCell.Characters(Start:=1, Length:=3).Font
        .Name = "����"
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A2:D2").Select
    With Selection.Font
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection.Font
        .Name = "��������"
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A3:D3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()+1"
    Range("A3:D3").Select
    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A2:D3").Select
    Selection.Font.Bold = True
    Range("A3:D3").Select
    With Selection.Font
        .Name = "��������"
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A2:D3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("G15").Select
        Range("A2:D3").Select
    Selection.Copy
    Range("A2:D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:D2").Select
    Application.CutCopyMode = False '
End Sub

Private Sub �����Ʒͼʾs_Click()
            Rows("2:2").Select
    Selection.RowHeight = 70
    Rows("3:3").Select
    Selection.RowHeight = 120
    Range("A3").Select
    ActiveSheet.Buttons.Add(0, 99, 75, 19).Select
    Selection.OnAction = "macroinstruction.xls!InsertPicturecp.InsertPicturecp"
    Selection.ShapeRange.IncrementLeft 2.5
    Selection.ShapeRange.IncrementTop 3#
    Selection.ShapeRange.IncrementTop 2.25
  ' ActiveSheet.Shapes("Button 4").Select
    Selection.Characters.Text = "��Ʒͼʾ"
    With Selection.Characters(Start:=1, Length:=6).Font
        .Name = "������"
        .FontStyle = "bold"
        .Size = 12
        .Underline = xlUnderlineStyleSingle
        .ColorIndex = 0
    End With
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
    End With
     Range("E5").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub ��������ͼʾs_Click()
            Rows("2:2").Select
    Selection.RowHeight = 70
    Rows("3:3").Select
    Selection.RowHeight = 120
    Range("A3").Select
    ActiveSheet.Buttons.Add(0, 80, 75, 19).Select
    Selection.OnAction = "macroinstruction.xls!InsertPicture.InsertPicture"
    Selection.ShapeRange.IncrementLeft 2.5
    Selection.ShapeRange.IncrementTop 3#
    Selection.ShapeRange.IncrementTop 2.25
  ' ActiveSheet.Shapes("Button 4").Select
    Selection.Characters.Text = "�ϼ�ͼʾ"
    With Selection.Characters(Start:=1, Length:=6).Font
        .Name = "������"
        .FontStyle = "bold"
        .Size = 12
        .Underline = xlUnderlineStyleSingle
        .ColorIndex = 0
    End With
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
    End With
     Range("E5").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub ��ֺϲ�����ϢS_Click()
    Dim s As String
    Dim t As String
    Dim o As String
    UnmergeCells.Show
End Sub

Private Sub ��ʼ��������S_Click()
'
' Macro4 Macro
' Macro recorded 6/23/2011 by Z002TE2Z
'

'
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub

Private Sub ��ӡS_Click()
Dim A()
Dim b As Integer
Dim c As Integer
b = Worksheets.Count
c = (b - 2)
If 1 > c Then Exit Sub
ReDim A(1 To c)
For i = 1 To c
A(i) = i
Next
Worksheets(A()).Select
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets(1).Select
End Sub

Private Sub ����A1��Ԫ��S_Click()
'
' Macro3 Macro
' Macro recorded 6/23/2011 by Z002TE2Z
'

'
    Range("A1").Select
End Sub

Private Sub �ϲ���Ԫ��ͬ��S_Click()
   Dim LRow As Long
   Dim JRow As Long
   MergeCells.Show
   Application.DisplayAlerts = True
End Sub

Private Sub �ƻ��༭S_Click()
    With firstline
        Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
    End With
    With first
      Dim arr, rng As Range, i&, temp$
        arr = Range("a1:a" & Range("a65536").End(xlDown).Row)
         For i = 1 To UBound(arr)
          temp = arr(i, 1)
            If InStr(temp, "JCV") = 0 And ((InStr(temp, "ate") = 0 And InStr(temp, "A5E") = 0) And InStr(temp, "�ز�") = 0) Then
           If rng Is Nothing Then Set rng = Cells(i, 1) Else Set rng = Union(rng, Cells(i, 1))
          End If
         Next
       If Not rng Is Nothing Then rng.EntireRow.Delete
     End With
    With secondly
      Range("L1").Select
      Selection.EntireColumn.Delete
    End With
    With thirdly
       For i = 2 To Range("a65536").End(xlUp).Row
           s = Cells(i, 11).Value
         cnt = Cells(i, 11).MergeArea.Count
         Cells(i, 11).UnMerge
         Range(Cells(i, 11), Cells(i + cnt - 1, 11)).Value = s
         i = i + cnt - 1
       Next
    End With
    With mm
     For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 13).Value
        cnt = Cells(i, 13).MergeArea.Count
        Cells(i, 13).UnMerge
       Range(Cells(i, 13), Cells(i + cnt - 1, 13)).Value = s
        i = i + cnt - 1
    Next
    End With
    With fourthly
        Range("C2").Select
    ActiveCell.FormulaR1C1 = "=RC[1]&RC[9]&RC[6]"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C100"), Type:=xlFillDefault
    Range("C2:C100").Select
    ActiveWindow.SmallScroll Down:=-90
    Range("C2").Select
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 70
    Range("C2:C100").Select
    Selection.Copy
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    End With
    With fifthly
        Range("B:D,L:L").Select
        Range("L1").Activate
        Selection.Delete Shift:=xlToLeft
        Range("A1").Select
    End With
End Sub

Private Sub �ƻ�����S_Click()
'
' Macro3 Macro
' Macro recorded 10/21/2011 by Z002TE2Z
'

'
    ActiveWorkbook.SaveAs Filename:="D:\Users\Z002TE2Z\Desktop\assembly plan.xls" _
        , FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.SmallScroll Down:=-6
    Range("A2:H100").Select
    ActiveWindow.SmallScroll Down:=-87
    Selection.Copy
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "JCV3311219571"
    Range("A2").Select
    ChDir "D:\my file\BOM backup browsers\Datebase Update"
    Workbooks.Open Filename:= _
        "D:\my file\BOM backup browsers\Datebase Update\Assembly Daily Plan.xls"
    Windows("assembly plan.xls").Activate
    Range("A2:H100").Select
    Selection.Copy
    Windows("Assembly Daily Plan.xls").Activate
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2").Select
        Application.CutCopyMode = False
    ActiveWorkbook.Save
        ActiveWorkbook.FollowHyperlink Address:= _
        "D:\my file\BOM backup browsers\Warehouse material list1.mdb", NewWindow:= _
        False, AddHistory:=True
    Windows("assembly plan.xls").Activate
    ActiveWindow.Close
    Windows("Assembly Daily Plan.xls").Activate
    ActiveWindow.Close
    ChDir "D:\my file\BOM backup browsers\Daily Material List"
    Workbooks.Open Filename:= _
        "D:\my file\BOM backup browsers\Daily Material List\All material list.XLS"
    Windows("macroinstruction.xls").Activate
    Windows("All material list.XLS").Activate
    Rows("1:1").EntireRow.AutoFit
End Sub

Private Sub ������ݱ�հ���S_Click()
    Dim firstRow As Long, i As Long
        firstRow = ActiveSheet.UsedRange.Row
        LastRow = firstRow + ActiveSheet.UsedRange.Rows.Count - 1
        For i = LastRow To firstRow Step -1
            If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then
               Rows(i).Delete
        End If
    Next
End Sub

Private Sub ������ݱ�������S_Click()
With ������ݱ�������S
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "All material list"
    Sheets(Array("Sheet1", "Sheet2", "Sheet3")).Select
    Sheets("Sheet3").Activate
    ActiveWindow.SelectedSheets.Delete
    ActiveWorkbook.Save
'With qinglie
  Application.ScreenUpdating = False '�ر���Ļˢ��
   Range("A:B,D:F,L:L,N:P").Select
    Range("N1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Application.ScreenUpdating = True '����Ļˢ��
    With qinglie
    Range("A:H").Select
        ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
        "All_material_list!a:h").CreatePivotTable TableDestination:="", TableName _
        :="PivotTable1", DefaultVersion:=xlPivotTableVersion10
    ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
    ActiveSheet.Cells(3, 1).Select
    Range("A3").Select
    End With
       With ActiveSheet.PivotTables("PivotTable1").PivotFields("MLFB")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("����������")
        .Orientation = xlColumnField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("�߱�����")
        .Orientation = xlColumnField
        .Position = 3
    End With
    Range("C6").Select
    Selection.Delete
    Range("C5").Select
    Selection.Delete
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BOM comp")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("comp description")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Material Qty"), "Count of Material Qty", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material type")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .Orientation = xlRowField
        .Position = 4
    End With
    Range("B10").Select
    Selection.Delete
    Range("B9").Select
    Selection.Delete
    Range("D8").Select
    Selection.Delete
    Range("E20").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Material Qty"). _
        Function = xlSum
    Range("A3").Select
    With jia
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .PivotItems("BLL").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("����").Visible = False
        .PivotItems("С����").Visible = False
        .PivotItems("PH").Visible = False
        '.PivotItems("С����").Visible = False
        '.PivotItems("С����").Visible = False
        '.PivotItems("С����").Visible = False
        '.PivotItems("С����").Visible = False
    End With
    Columns("B:B").ColumnWidth = 25
    Cells.Select
    Selection.ColumnWidth = 7#
    Range("A3").Select
    Cells.Select
    Selection.Copy
    ActiveWorkbook.Save
    End With
    End With
    'With xinjian
        'Sheets("Sheet1").Select
    'Sheets.Add
    'Range("A1").Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       ' :=False, Transpose:=False
    'Range("A1").Select
    'Sheets("Sheet1").Select
    'Range("A1").Select
    'Application.CutCopyMode = False
        'Range("A1").Select
    'ActiveWorkbook.Save
    'End With
End Sub

Private Sub ���͸�ӱ��е�BlankS_Click()
'
' Macro3 Macro
' Macro recorded 6/23/2011 by Z002TE2Z
'

'
    Cells.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Private Sub �������������S_Click()
'
' Macro1 Macro
' Macro recorded 6/23/2011 by Z002TE2Z
'

'
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B12").Select
End Sub

Private Sub ���ô�ӡ�뱸ע��ϢS_Click()
'
' Macro1 Macro
' Macro recorded 7/28/2011 by Z002TE2Z
'

'
    ActiveSheet.PageSetup.CenterFooterPicture.Filename = _
        "D:\my file\BOM backup browsers\lswj\��ע.bmp"
    With ActiveSheet.PageSetup.CenterFooterPicture
        .Height = 40.5
        .Width = 459
    End With
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "&T&D"
        .LeftFooter = ""
        .CenterFooter = "&G"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.15)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
End Sub

Private Sub ˿ӡ�������S_Click()
'
' Macro4 Macro
' Macro recorded 9/29/2011 by Z002TE2Z
'

'
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .PivotItems("BLL").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    Columns("B:B").ColumnWidth = 25
    Cells.Select
    Selection.ColumnWidth = 7#
    Range("A3").Select
End Sub

Private Sub ��ӱ��쳣��¼��S_Click()
'
' Macro3 Macro
' Macro recorded 7/28/2011 by Z002TE2Z
'

'
    Range("A23").Select
    ActiveSheet.Pictures.Insert("D:\my file\BOM backup browsers\lswj\�ӱ��.bmp"). _
        Select
    ActiveWindow.Zoom = 70
    ActiveWindow.Zoom = 85
    ActiveWindow.Zoom = 100
    ActiveWindow.Zoom = 115
    ActiveWindow.Zoom = 130
    ActiveWindow.Zoom = 115
    Range("M33").Select
End Sub

Private Sub ����½�������S_Click()
Dim f, g As Integer
g = Worksheets.Count
f = (g - 1)
Sheets(f).Select
Selection.Copy
    Sheets.Add
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Range("A3:bz80").Select
    Range("E6").Activate
    Selection.Copy
    Worksheets.Add after:=Sheets(g)
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Transpose:=True
            Range("D7:D100").Select
    Range("B7:DD100").Sort Key1:=Range("D7"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
   Range("A3:bz80").Select
   Range("D7").Activate
   Selection.Copy
   Sheets(f).Select
   Range("A3").Select
   Selection.PasteSpecial Paste:=xlPasteValues, Transpose:=True
   Sheets(f + 2).Delete
   Range("a1").Select
   ActiveWorkbook.Save
   With dise
        ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Select
        Dim OO As Range  '����I�ı���
      Set OO = [ActiveCell].Offset(-1, -1) 'I�Ķ�ֵ̬Ϊ��ǰ������ĵ�ǰ���Ԫ���ַ ;
               Range("E7", OO).Select
                   'Selection.Font.Bold = True   '(���õ�Ԫ������Ϊ����)
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    Selection.Font.ColorIndex = 0
ActiveSheet.Cells.SpecialCells(xlCellTypeBlanks, xlLogical).Select
Selection.Interior.ColorIndex = xlNone
Range("a1").Select
Set OO = Nothing
End With
Sheets(f + 1).Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    Selection.Copy
End Sub

Private Sub ͸�ӱ��ͷ�޸�S_Click()
'
' Macro3 Macro
' Macro recorded 8/4/2011 by Z002TE2Z
'

'
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Material Qty"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "������"
    With ActiveCell.Characters(Start:=1, Length:=3).Font
        .Name = "����"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "�߱�"
    With ActiveCell.Characters(Start:=1, Length:=2).Font
        .Name = "����"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Rows("1:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp
    Range("H11").Select
End Sub

Private Sub ͸�ӱ�ṹת��S_Click()
'
' Macro1 Macro
' Macro recorded 6/23/2011 by Z002TE2Z
'

'
    Rows("4:6").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = -90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A6:D6").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("E22").Select
End Sub

Private Sub ͸�ӱ�˿ӡ��������S_Click()
    Range("D6:C100").Select
    Range("A6:cc100").Sort Key1:=Range("D6"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
End Sub

Private Sub ͸�ӱ�ָ��������S_Click()
    Range("C6:C100").Select
    Range("A6:cc100").Sort Key1:=Range("C6"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
End Sub


Private Sub ѡ����ճ��S_Click()
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Private Sub ����͸�Ӳ���S_Click()
   With ActiveSheet.PivotTables("PivotTable1").PivotFields("MLFB")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("����������")
        .Orientation = xlColumnField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("�߱�����")
        .Orientation = xlColumnField
        .Position = 3
    End With
    Range("C6").Select
    Selection.Delete
    Range("C5").Select
    Selection.Delete
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BOM comp")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("comp description")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Material Qty"), "Count of Material Qty", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material type")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .Orientation = xlRowField
        .Position = 4
    End With
    Range("B10").Select
    Selection.Delete
    Range("B9").Select
    Selection.Delete
    Range("D8").Select
    Selection.Delete
    Range("E20").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Material Qty"). _
        Function = xlSum
    Range("A3").Select
    With jia
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("��ע")
        .PivotItems("BLL").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("����").Visible = False
        .PivotItems("С����").Visible = False
    End With
    Columns("B:B").ColumnWidth = 25
    Cells.Select
    Selection.ColumnWidth = 7#
    Range("A3").Select
    Cells.Select
    Selection.Copy
    End With
    With xinjian
        Sheets("Sheet1").Select
    Sheets.Add
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Sheets("Sheet1").Select
    Range("A1").Select
    Application.CutCopyMode = False
        Range("A1").Select
    ActiveWorkbook.Save
    End With
End Sub
Private Sub ���������С��8S_Click()
  Dim b()
   Dim d, c As Integer
   c = Worksheets.Count
   d = (c - 2)
   With E
   For x = 1 To 2
       m = ActiveSheet.Name
         If m <> Sheets(1).Name Then
         With j
If 1 > d Then Exit Sub
ReDim b(1 To d)
For i = 1 To d
b(i) = i
Next
Worksheets(b()).Select
With ll

    Range("E2:BO4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = -90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        Range("A1").Select
        End With
            With pm
             Sheets(1).Select
                 Range("D4:C100").Select
    Range("A4:cc100").Sort Key1:=Range("d4"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = """&""����,Regular""&9&U�˱�ֻ�����ϲο�&""Arial,Regular""&11&U"""
        .CenterHeader = ""
        .RightHeader = "&T&D"
        .LeftFooter = ""
        .CenterFooter = "&G"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.15)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    ActiveWindow.SmallScroll Down:=-6
    Rows("5:38").Select
    Selection.RowHeight = 12
    ActiveWindow.SmallScroll Down:=-21
    Range("A1").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("E:AH").Select
    Selection.ColumnWidth = 5.43
    Range("A1").Select
    End With
            Range("A1").Select
          MsgBox "�˳�1"
          End With
     Else
            For u = 1 To d
            With OO
Cells.Select
With Selection.Font
.Size = 8                        '<Range("a1").Select>
End With
'___________________________________________________________________________________

    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 8
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With


'-----------------------------------------------------------------------------------
With ��ṹת��
    Rows("4:6").Select
       With Selection
          .Orientation = -90
       End With
    Range("a6:d6").Select
       With Selection
          .Orientation = 0
       End With
End With
'-----------------------------------------------------------------------------------
With ���blank
Cells.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End With
'----------------------------------------------------------------------------------------------
With ��ͷ�޸�
    'Range("A1:BC40").Select
   ' Range("BC40").Activate
    'Selection.AutoFormat Format:=xlRangeAutoFormatClassic1, Number:=True, Font _
        :=True, Alignment:=True, Border:=True, Pattern:=True, Width:=True
    'With Selection.Interior
     '   .ColorIndex = 43
     '   .Pattern = xlSolid
    'End With
    Range("C6:C100").Select
    Range("A6:cc100").Sort Key1:=Range("C6"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Material Qty"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "������"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "�߱�"
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    End With
  End With
'---------------------------------------------------------------------------------------------------------
With ʱ������
Range("A2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.AutoFill Destination:=Range("A2:D3"), Type:=xlFillDefault
    Range("A2:D3").Select
    Range("A2:D2").Select
    ActiveCell.FormulaR1C1 = "�� �� ��"
    With ActiveCell.Characters(Start:=1, Length:=3).Font
        .Name = "����"
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A2:D2").Select
    With Selection.Font
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection.Font
        .Name = "��������"
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A3:D3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()+1"
    Range("A3:D3").Select
    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A2:D3").Select
    Selection.Font.Bold = True
    Range("A3:D3").Select
    With Selection.Font
        .Name = "��������"
        .Size = 30
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A2:D3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("G15").Select
        Range("A2:D3").Select
    Selection.Copy
    Range("A2:D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:D2").Select
    Application.CutCopyMode = False
    End With
'----------------------------------------------------------------------------------------------------------------------------
With ��ӡ�뱸ע
ActiveSheet.PageSetup.CenterFooterPicture.Filename = _
        "D:\my file\BOM backup browsers\lswj\��ע.bmp"
    With ActiveSheet.PageSetup.CenterFooterPicture
        .Height = 40.5
        .Width = 459
    End With
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = """&""����,Regular""&9&U�˱�ֻ�����ϲο�&""Arial,Regular""&11&U"""  '"�˱�ֻ���ڲ����ϲο�"
        .CenterHeader = ""
        .RightHeader = "&T&D"
        .LeftFooter = ""
        .CenterFooter = "&G"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.15)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
      With fuzhe
'----------------------------------------------------------------------
    ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Select   'ѡ�й�������������������½ǵ�Ԫ��
'----------------------------------------------------------------------
        Dim r, t, n As Byte
        Dim rno, o, pp As Range  '����I�ı���
        Set rng = Cells(Rows.Count, 1).End(xlUp)
              o = rng.Row
              r = 100
              t = 0
              n = 0
      Set pp = [ActiveCell].Offset(t - 0, n + 2) 'I�Ķ�ֵ̬Ϊ��ǰ������ĵ�ǰ���Ԫ���ַ ;
               Range("a1", pp).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    Columns("A:D").Select
    Columns("A:D").EntireColumn.AutoFit
        Rows("3:3").Select
    Selection.RowHeight = 100
    Range("A1").Select
Set pp = Nothing
Set rno = Nothing
End With
    End With
'---------------------------------------------------------------------------------------------------------------------------
    With m
    ActiveSheet.Next.Select
    End With
            Next u
        End If
   Next x
   End With
   With l
   Sheets(1).Select
   End With
End Sub

