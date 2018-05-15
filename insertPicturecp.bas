Attribute VB_Name = "insertPicturecp"
Option Explicit
Sub InsertPicturecp()
    Dim ccc As Shape
    Dim b, c As Integer
    Dim cp1, cp2, cp3, cp4, cp5, cp6, cp7, cp8, cp9, cp10, cp11, cp12, cp13, cp14, cp15, cp16, cp17, cp18, cp19, cp20, cp21, cp22 As String
    Dim d As Range
    With ActiveSheet
        For Each ccc In .Shapes
            If ccc.Type = 13 Then
                ccc.Delete
            End If
        Next
        For b = 5 To .Range("IV2").End(xlToLeft).Column
            For c = 2 To 2 Step 2
           cp1 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Azio-Module-picture" & ".JPG"
           cp2 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Azio-outside-carton-carton" & ".JPG"
           cp3 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Dragon-Picture" & ".JPG"
           cp4 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Electrium-Picture" & ".JPG"
           cp5 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Floor-Socket" & ".JPG"
           cp6 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Foreign" & ".JPG"
           cp7 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Germany" & ".JPG"
           cp8 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Ocero-Picture" & ".JPG"
           cp9 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Opaque-Picture" & ".JPG"
           cp10 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "TP-label-picture" & ".JPG"
           cp11 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Valere-Label-Picture" & ".JPG"
           cp12 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Vega-India-Picture" & ".JPG"
           cp13 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Vega-outside-carton-colour-carton" & ".JPG"
           cp14 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "VEIV-Picture" & ".JPG"
           cp15 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Opaque-Complete" & ".JPG"
           cp16 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Opaque-Cover-Plate" & ".JPG"
           cp17 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Opaque-Modular-Selling" & ".JPG"
           cp18 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Valere-Aluminum-Cover-Picture" & ".JPG"
           cp19 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Valere-Glass-Cover-Picture" & ".JPG"
           cp20 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Valere-Modular-Picture" & ".JPG"
           cp21 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Valere-Plastic-Cover-Picture" & ".JPG"
           cp22 = ThisWorkbook.path & "\" & .Cells(b, c).Text & "Valere-Wood-Cover-Picture" & ".JPG"
                    If Dir(cp1) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp1, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp2) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp2, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 25) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp3) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp3, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp4) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp4, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp5) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp5, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp6) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp6, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp7) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp7, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp8) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp8, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 25) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp9) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp9, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 25) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp10) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp10, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp11) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp11, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp12) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp12, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp13) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp13, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 25) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp14) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp14, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp15) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp15, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp16) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp16, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp17) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp17, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp18) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp18, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp19) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp19, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp20) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp20, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 26) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp21) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp21, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                    If Dir(cp22) <> "" Then
                        Set ccc = .Shapes.AddPicture(cp22, False, True, 6, 6, 6, 6)
                      Set d = .Cells(b, c + 27) '图片显示列设置
                    With ccc
                    .LockAspectRatio = msoFalse
                      .Top = d.Top + 1   '顶端位置
                      .Left = d.Left + 1   '左侧位置
                      .Width = d.Width - 1.5   '宽度
                      .Height = d.Height - 1.5  '高度
                      .TopLeftCell = ""   '左上角单元格
                    End With
                      Else
                   .Cells(c, b + 28) = "" '无图片显示列设置
                     End If
                  Next
              Next
         End With
    Set ccc = Nothing
    Set d = Nothing
End Sub
