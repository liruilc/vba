Attribute VB_Name = "����Ԫֵ�����ļ���"
Option Explicit
Public Sub ��Ԫֵ�����ļ���()
'#----------------------------------------------------------------------------------ָ������Ŀ¼·��---1
Static pathna As String
 With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
        pathna = .SelectedItems(1) 'pathna·��������
        End If
 End With
'#----------------------------------------------------------------------------------------------------1
'##------------------------------------------------------------------------------Ŀ¼�ļ���������·��---2
Dim fs As Object: Dim n, f, fd, zf, zfd, file, afile, nfil, infil, fff: n = 2 '"n"'"n"��ʼ��������

Set fs = CreateObject("Scripting.FileSystemObject") '����fsΪϵͳĿ¼����'"fs"
    Set f = fs.getfolder(pathna) '������ָ��·���е��ļ������Ӧ�� Folder����'"f"
        For Each fd In f.subfolders '"fd"
            Set zf = fs.getfolder(pathna & "\" & fd.Name) '"zf"
                For Each zfd In zf.subfolders '"zdf"
                    afile = zf & "\" & Dir(zfd & "*.*") '"afile"
'###------------------------------------------------------------------�����ļ��м���Ŀ¼�µ�EXCEL�ļ�---3
    Dim oWB, noWB As Workbook: Dim oWK As Worksheet: Dim sFPath As String
    Application.DisplayAlerts = False
                           Set oWB = Excel.Workbooks.Open(afile) 'oWB �򿪹�����·���µĹ�����
                           
                           With oWB
                               Set oWK = .Worksheets(1) 'oWK
                                   With oWK
'####-----------------------------------------------------------------------ÿ�������ָ���е�ֵ���ֵ�---4
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
'#####-----------------------------------------------------�ڵ�ǰ��������Ŀ¼�½��ļ��в����ֵ���ֵ����---5
    ActiveWorkbook.Save
                                        For ii = 1 To d.Count
                                           fil = Cells(ii, 1)
                                            f = Dir(zfd & "\" & fil, vbDirectory)   '�ж��Ƿ��Ѿ�����
                                              If f = "" Then
                                                 MkDir (zfd & "\" & fil)   '��������ھͽ���
                                              End If

     Sheets(twsNo).Select '�����ʼ������
'######--------------------------------------------------�ڵ�ǰ����������Ŀ¼�½�cvs�ļ������ֵ���ֵ����---6
                                         For iii = 2 To lrow
                                           ifil = Cells(iii, 3)
                        f = Dir(zfd & "\" & fil & "\" & ifil & ".csv", vbDirectory)   '�ж��Ƿ��Ѿ�����
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
