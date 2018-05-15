Attribute VB_Name = "UNMergeCell"
Sub UNMergeCell()
Dim s As String
Dim t As String
Dim o As String
UnmergeCells.Show
  '  For i = 2 To Range("a65536").End(xlUp).Row
    '    s = Cells(i, 1).Value
    '    cnt = Cells(i, 1).MergeArea.Count
     '   Cells(i, 1).UnMerge
     '   Range(Cells(i, 1), Cells(i + cnt - 1, 1)).Value = s
     '    t = Cells(i, 2).Value
     '   cnt = Cells(i, 2).MergeArea.Count
     '   Cells(i, 2).UnMerge
     '   Range(Cells(i, 2), Cells(i + cnt - 1, 2)).Value = t
     '    o = Cells(i, 3).Value
      '  cnt = Cells(i, 3).MergeArea.Count
     '   Cells(i, 3).UnMerge
      '  Range(Cells(i, 3), Cells(i + cnt - 1, 3)).Value = o
      '  i = i + cnt - 1
   ' Next
End Sub
