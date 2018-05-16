VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExtractAtts()

  Dim Excel As Excel.Application

  Dim ExcelSheet As Object

  Dim ExcelWorkbook As Object



  Dim RowNum As Integer

  Dim Header As Boolean

  Dim elem As AcadEntity 'IAcadEntity界面表示AutoCAD实体的默认界面。

  Dim Array1 As Variant '数组变量的名称返回一个LONG数据类型

  Dim Count As Integer



  ' Launch Excel.'启动EXCEL

  Set Excel = New Excel.Application



  ' Create a new workbook and find the active sheet.

  '创建一个新的工作簿并找到活动工作表。

  Set ExcelWorkbook = Excel.Workbooks.Add

  Set ExcelSheet = Excel.ActiveSheet

  ExcelWorkbook.SaveAs "Attribute.xlsx"



  RowNum = 1

  Header = False



  ' Iterate through model space finding

  ' all block references.

  '遍历模型空间查找

  '所有块参考。

  For Each elem In ThisDrawing.ModelSpace

    With elem

      ' When a block reference has been found,

      ' check it for attributes

      '当找到块参考时，

      '检查它的属性

      If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then

        If .HasAttributes Then

          ' Get the attributes

          '获取属性

          Array1 = .GetAttributes



          ' Copy the Tagstrings for the

          ' Attributes into Excel

          '复制标签串

          '属性到Excel

          For Count = LBound(Array1) To UBound(Array1) '返回一个 Long 型数据(数组变量的名称[,返回哪一维的下界]值为指定数组维可用的最小下标)

                                                       ' 到 返回一个 Long 型数据（数组变量的名称[,返回哪一维的上界]值为指定的数组维可用的最大下标。）

            If Header = False Then

              If StrComp(Array1(Count).EntityName, "AcDbAttribute", 1) = 0 Then

                ExcelSheet.Cells(RowNum, Count + 1).Value = Array1(Count).TagString

              End If

            End If

          Next Count



          RowNum = RowNum + 1

          For Count = LBound(Array1) To UBound(Array1)

            ExcelSheet.Cells(RowNum, Count + 1).Value = Array1(Count).TextString

          Next Count



          Header = True

        End If

      End If

    End With

  Next elem

  Excel.Application.Quit

End Sub
