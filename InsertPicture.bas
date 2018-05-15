Attribute VB_Name = "InsertPicture"
Option Explicit
Sub InsertPicture()
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
    Selection.ShapeRange.Height = 20 '高度
    Selection.ShapeRange.Width = 17.5 '宽度
    Selection.ShapeRange.Rotation = 85#
    Selection.ShapeRange.IncrementLeft 58#
    Selection.ShapeRange.IncrementTop 2
    Selection.OnAction = "macroinstruction.xls!fangda.fangda"
        Range("AJ5:AU60").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .AddIndent = False
        .ReadingOrder = xlContext
    End With
        Rows("5:45").Select
        Selection.RowHeight = 12.5
        Columns("AJ:AU").Select
        Selection.ColumnWidth = 3.5
        Range("A1").Select         'v___________________________________________________________
        For b = 5 To .Cells(.Rows.Count, 1).End(xlUp).Row
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
                      Set d = .Cells(b, c + 35) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                    If Dir(dpdf1) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 35), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_1" & ".pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 44) = "" '无图片显示列设置
                     End If
                    If Dir(dxls1) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 35), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_1" & ".xls", ScreenTip:="XLS文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 45) = "" '无图片显示列设置
                     End If
                    If Dir(dtif2) <> "" Then
                      Set aaa = .Shapes.AddPicture(dtif2, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 36) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                    If Dir(dpdf2) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 35), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_2" & ".pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 47) = "" '无图片显示列设置
                     End If
                    If Dir(dxls2) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 35), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_2" & ".xls", ScreenTip:="XLS文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 48) = "" '无图片显示列设置
                     End If
                    If Dir(dtif3) <> "" Then
                      Set aaa = .Shapes.AddPicture(dtif3, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 37) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 49) = "" '无图片显示列设置
                     End If
                    If Dir(dpdf3) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 35), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_3" & ".pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                      End With
                     Else
                   .Cells(b, c + 50) = "" '无图片显示列设置
                   End If
                  If Dir(tj1) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj1, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(tj2) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj2, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(tj3) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj3, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(tj4) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj4, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(tj5) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj5, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(tj6) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj6, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(tj7) <> "" Then
                      Set aaa = .Shapes.AddPicture(tj7, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 46) '图片显示列设置
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(b, c + 43) = "" '无图片显示列设置
                     End If
                  If Dir(pd1) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 40), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_1_.pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                  If Dir(pd2) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 41), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_2_.pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                  If Dir(pd3) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 42), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_3_.pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                  If Dir(pd4) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 43), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_4_.pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                  If Dir(pd5) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 44), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_5_.pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                  If Dir(pd6) <> "" Then
                      ActiveSheet.Hyperlinks.Add Anchor:=.Cells(b, c + 45), _
                      Address:="\qita\p\" & .Cells(b, c).Text & "_6_.pdf", ScreenTip:="PDF文件"
                    With aaa
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1
                      .Left = d.Left + 1
                      .Width = d.Width - 1.5
                      .Height = d.Height - 1.5
                      .TopLeftCell = ""
                    End With
                      Else
                   .Cells(b, c + 46) = "" '无图片显示列设置
                     End If
                Next
              Next
         End With
    Set aaa = Nothing
    Set d = Nothing
End Sub

