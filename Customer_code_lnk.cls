VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Sub Customer_code_lnk()
Dim WshShell As Object, oShellLink As Object
Dim fil, ifil, aa, aaa, bb, bbb, ccc As String
Dim iwsq, twsNo, lrow, lcol, i, ii, iii As Long
Dim f, zfd
    Sheets(2).Select
    lrow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Row
    Sheets(1).Select
    lrow1 = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Row

        zfd = "D:\产品结构匹配\客户代码\"
        For i = 1 To lrow
        Sheets(2).Select
            aa = Cells(i, 1)
            bb = Mid(Cells(i, 1), 1, 1)
               
            If (bb Like "[A-Z]" = False) Then
                 f = Dir(zfd & "\" & aa, vbDirectory)
                   If f = "" Then
                      MkDir (zfd & aa)
                    End If
'--------------------------------------------------------------------------------------------------------------
                   Sheets(1).Select
                   For ii = 2 To lrow1
                   aaa = Cells(ii, 3)
                   bbb = Cells(ii, 27)
                   ccc = Cells(ii, 9)
                      If Cells(ii, 6) = aa Then
                         f = Dir(zfd & "\" & aa & "\" & ccc, vbDirectory)
                             If f = "" Then
                                 MkDir (zfd & "\" & aa & "\" & ccc)
                             End If
                         Set WshShell = CreateObject("WScript.Shell")
                         Set oShellLink = WshShell.CreateShortcut(zfd & "\" & aa & "\" & ccc & "\" & aaa & ".cvs.lnk")
                         oShellLink.TargetPath = bbb  '指定快捷方式的目标文件
                         oShellLink.WindowStyle = 1          '窗口状态/运行方式
                         oShellLink.Description = "这是一个快捷方式"
                         oShellLink.WorkingDirectory = zfd & aa & "\" & ccc
                         oShellLink.Save    '保存
                         Set WshShell = Nothing
                         Set oShellLink = Nothing
                      End If
                  Next ii
'--------------------------------------------------------------------------------------------------------------
              Else
                 f = Dir(zfd & "\" & bb, vbDirectory)
                   If f = "" Then
                      MkDir (zfd & bb)
                    End If
                     f = Dir(zfd & "\" & bb & "\" & aa, vbDirectory)
                      If f = "" Then
                         MkDir (zfd & bb & "\" & aa)
                      End If
'--------------------------------------------------------------------------------------------------------------
                   Sheets(1).Select
                   For ii = 2 To lrow1
                   aaa = Cells(ii, 3)
                   bbb = Cells(ii, 27)
                   ccc = Cells(ii, 9)
                      If Cells(ii, 6) = aa Then
                         f = Dir(zfd & "\" & bb & "\" & aa & "\" & ccc, vbDirectory)
                              If f = "" Then
                                 MkDir (zfd & "\" & bb & "\" & aa & "\" & ccc)
                             End If
                         Set WshShell = CreateObject("WScript.Shell")
                         Set oShellLink = WshShell.CreateShortcut(zfd & bb & "\" & aa & "\" & ccc & "\" & aaa & ".cvs.lnk")
                         oShellLink.TargetPath = bbb  '指定快捷方式的目标文件
                         oShellLink.WindowStyle = 1          '窗口状态/运行方式
                         oShellLink.Description = "这是一个快捷方式"
                         oShellLink.WorkingDirectory = zfd & "\" & bb & "\" & aa & "\" & ccc
                         oShellLink.Save    '保存
                         Set WshShell = Nothing
                         Set oShellLink = Nothing
                      End If
                  Next ii
'--------------------------------------------------------------------------------------------------------------
            End If
        
        Next i

End Sub
