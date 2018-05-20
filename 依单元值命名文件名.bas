Attribute VB_Name = "依单元值命名文件名"
Option Explicit
Public Sub 单元值命名文件名()
'#----------------------------------------------------------------------------------指定处理目录路径---1
Static pathna As String
 With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
        pathna = .SelectedItems(1) 'pathna路径变量名
        End If
 End With
'#----------------------------------------------------------------------------------------------------1
'##------------------------------------------------------------------------------目录文件夹数量与路径---2
Dim fs As Object: Dim n, f, fd, zf, zfd, file, afile, nfil, infil, fff: n = 2 '"n"'"n"起始数变量，

Set fs = CreateObject("Scripting.FileSystemObject") '变量fs为系统目录对象'"fs"
    Set f = fs.getfolder(pathna) '返回与指定路径中的文件夹相对应的 Folder对象'"f"
        For Each fd In f.subfolders '"fd"
            Set zf = fs.getfolder(pathna & "\" & fd.Name) '"zf"
                For Each zfd In zf.subfolders '"zdf"
                    afile = zf & "\" & Dir(zfd & "*.*") '"afile"
'###------------------------------------------------------------------遍历文件夹及子目录下的EXCEL文件---3
    Dim oWB, noWB As Workbook: Dim oWK As Worksheet: Dim sFPath As String
    Application.DisplayAlerts = False
                           Set oWB = Excel.Workbooks.Open(afile) 'oWB 打开工作簿路径下的工作簿
                           
                           With oWB
                               Set oWK = .Worksheets(1) 'oWK
                                   With oWK
'####-----------------------------------------------------------------------每工作表的指定列的值的字典---4
    Dim d As Object: Dim str As Variant: Dim strKey, fil, ifil, TextStream As String
    Dim iwsq, twsNo, lrow, i, ii, iii As Long
    
                                  iwsq = Worksheets.Count
                                 twsNo = ActiveSheet.Index
                                 Set d = CreateObject("scripting.dictionary")
                                  lrow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Row
                                   str = Range("I1:I" & lrow + 1 - 1)
                                       For i = 2 To lrow
                                      strKey = CStr(str(i, 1))
                                            If Not d.exists(strKey) Then
                                                   d.Add strKey, strKey
                                            End If
                                       Next i
                             ActiveWorkbook.Worksheets.Add after:=Worksheets(iwsq)
                               Cells(1, 1).Resize(UBound(d.keys) + 1, 1) = Application.Transpose(d.keys)
'#####-----------------------------------------------------在当前工作簿子目录新建文件夹并用字典数值命名---5
    ActiveWorkbook.Save
                                        For ii = 1 To d.Count
                                           fil = Cells(ii, 1)
                                            f = Dir(zfd & "\" & fil, vbDirectory)   '判断是否已经存在
                                              If f = "" Then
                                                 MkDir (zfd & "\" & fil)   '如果不存在就建立
                                              End If

     Sheets(twsNo).Select '激活初始工作表
'######--------------------------------------------------在当前工作簿子子目录新建cvs文件并用字典数值命名---6
                                         For iii = 2 To lrow
                                           ifil = Cells(iii, 3)
                        f = Dir(zfd & "\" & fil & "\" & ifil & ".csv", vbDirectory)   '判断是否已经存在
                                             nfil = zfd & "\" & fil & "\" & ifil & ".csv"
                                              If f = "" Then
                                                   If Cells(iii, 9) = fil Then
                                                      With fs
                                                           On Error Resume Next
                                                           fff = fs.CreateTextFile(nfil, False)
                                                      End With
                                                   End If
                                              End If
                                          Next iii
'######-------------------------------------------------------------------------------------------------6
                                         Sheets(twsNo + 1).Select
                                        Next ii
'#####--------------------------------------------------------------------------------------------------5
'####---------------------------------------------------------------------------------------------------4
                                   End With
                                   Sheets(twsNo + 1).Delete
                                  ActiveWorkbook.Save
                               .Close
                           End With
'###----------------------------------------------------------------------------------------------------3
                    Application.DisplayAlerts = True
                    n = n + 1
                Next
            Set zf = Nothing
            n = n + 1
        Next
    Set f = Nothing
Set fs = Nothing
'##-----------------------------------------------------------------------------------------------------2
End Sub
