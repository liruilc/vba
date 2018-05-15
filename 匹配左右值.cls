VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "匹配左右值"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub lxa()

       ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, xlLogical).Select

End Sub

Public Sub cx()

       Dim sname As String

           sname = ActiveSheet.Name
           Application.ScreenUpdating = False

           Worksheets("客户代码表").Activate '数据源工作表激活

           ActiveWorkbook.Save

Call lxa

       Dim rng, w, rng1, w1 As Range

       Dim rowss, rowss1 As Byte

           Set rng = Cells(Rows.Count, 2).End(xlUp)

                 w = rng.Row

           Set rng = Nothing

               Worksheets(sname).Activate  '查询工作表激活

               Dim rngg, ww As String

                   Set rngg = Cells(Rows.Count, 1).End(xlUp)

                         ww = rngg.Row

                   Set rngg = Nothing

For rowss = 2 To ww '查询工作表行

    For rowss1 = 2 To w '数据源工作表行

        If Cells(rowss, 1).Value = Worksheets("客户代码表").Cells(rowss1, 2).Value Then

           Cells(rowss, 2).Value = Worksheets("客户代码表").Cells(rowss1, 1).Value

           Cells(rowss, 3).Value = Worksheets("客户代码表").Cells(rowss1, 3).Value

        End If

    Next rowss1

  Next rowss
  Application.ScreenUpdating = True

End Sub
