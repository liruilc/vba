VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnmergeCells 
   Caption         =   "UserForm1"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   OleObjectBlob   =   "UnmergeCells.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UnmergeCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim s As String
Private Sub A_Click()
    For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 1).Value
        cnt = Cells(i, 1).MergeArea.Count
        Cells(i, 1).UnMerge
       Range(Cells(i, 1), Cells(i + cnt - 1, 1)).Value = s
        i = i + cnt - 1
    Next
End Sub

Private Sub CommandButton10_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 12).Value
        cnt = Cells(i, 12).MergeArea.Count
        Cells(i, 12).UnMerge
       Range(Cells(i, 12), Cells(i + cnt - 1, 12)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton11_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 10).Value
        cnt = Cells(i, 10).MergeArea.Count
        Cells(i, 10).UnMerge
       Range(Cells(i, 10), Cells(i + cnt - 1, 10)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton12_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 11).Value
        cnt = Cells(i, 11).MergeArea.Count
        Cells(i, 11).UnMerge
       Range(Cells(i, 11), Cells(i + cnt - 1, 11)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton13_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 13).Value
        cnt = Cells(i, 13).MergeArea.Count
        Cells(i, 13).UnMerge
       Range(Cells(i, 13), Cells(i + cnt - 1, 13)).Value = s
        i = i + cnt - 1
    Next

End Sub


Private Sub CommandButton14_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 14).Value
        cnt = Cells(i, 14).MergeArea.Count
        Cells(i, 14).UnMerge
       Range(Cells(i, 14), Cells(i + cnt - 1, 14)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton15_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 15).Value
        cnt = Cells(i, 15).MergeArea.Count
        Cells(i, 15).UnMerge
       Range(Cells(i, 15), Cells(i + cnt - 1, 15)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton16_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 16).Value
        cnt = Cells(i, 16).MergeArea.Count
        Cells(i, 16).UnMerge
       Range(Cells(i, 16), Cells(i + cnt - 1, 16)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton17_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 17).Value
        cnt = Cells(i, 17).MergeArea.Count
        Cells(i, 17).UnMerge
       Range(Cells(i, 17), Cells(i + cnt - 1, 17)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton18_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 18).Value
        cnt = Cells(i, 18).MergeArea.Count
        Cells(i, 18).UnMerge
       Range(Cells(i, 18), Cells(i + cnt - 1, 18)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton19_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 19).Value
        cnt = Cells(i, 19).MergeArea.Count
        Cells(i, 19).UnMerge
       Range(Cells(i, 19), Cells(i + cnt - 1, 19)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton2_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 2).Value
        cnt = Cells(i, 2).MergeArea.Count
        Cells(i, 2).UnMerge
       Range(Cells(i, 2), Cells(i + cnt - 1, 2)).Value = s
        i = i + cnt - 1
    Next
End Sub

Private Sub CommandButton20_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 20).Value
        cnt = Cells(i, 20).MergeArea.Count
        Cells(i, 20).UnMerge
       Range(Cells(i, 20), Cells(i + cnt - 1, 20)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton21_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 21).Value
        cnt = Cells(i, 21).MergeArea.Count
        Cells(i, 21).UnMerge
       Range(Cells(i, 21), Cells(i + cnt - 1, 21)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton22_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 22).Value
        cnt = Cells(i, 22).MergeArea.Count
        Cells(i, 22).UnMerge
       Range(Cells(i, 22), Cells(i + cnt - 1, 22)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton23_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 23).Value
        cnt = Cells(i, 23).MergeArea.Count
        Cells(i, 23).UnMerge
       Range(Cells(i, 23), Cells(i + cnt - 1, 23)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton24_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 24).Value
        cnt = Cells(i, 24).MergeArea.Count
        Cells(i, 24).UnMerge
       Range(Cells(i, 24), Cells(i + cnt - 1, 24)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton25_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 25).Value
        cnt = Cells(i, 25).MergeArea.Count
        Cells(i, 25).UnMerge
       Range(Cells(i, 25), Cells(i + cnt - 1, 25)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton26_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 26).Value
        cnt = Cells(i, 26).MergeArea.Count
        Cells(i, 26).UnMerge
       Range(Cells(i, 26), Cells(i + cnt - 1, 26)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton3_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 3).Value
        cnt = Cells(i, 3).MergeArea.Count
        Cells(i, 3).UnMerge
       Range(Cells(i, 3), Cells(i + cnt - 1, 3)).Value = s
        i = i + cnt - 1
    Next
End Sub

Private Sub CommandButton4_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 4).Value
        cnt = Cells(i, 4).MergeArea.Count
        Cells(i, 4).UnMerge
       Range(Cells(i, 4), Cells(i + cnt - 1, 4)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton5_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 5).Value
        cnt = Cells(i, 5).MergeArea.Count
        Cells(i, 5).UnMerge
       Range(Cells(i, 5), Cells(i + cnt - 1, 5)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton6_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 6).Value
        cnt = Cells(i, 6).MergeArea.Count
        Cells(i, 6).UnMerge
       Range(Cells(i, 6), Cells(i + cnt - 1, 6)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton7_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 7).Value
        cnt = Cells(i, 7).MergeArea.Count
        Cells(i, 7).UnMerge
       Range(Cells(i, 7), Cells(i + cnt - 1, 7)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton8_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 8).Value
        cnt = Cells(i, 8).MergeArea.Count
        Cells(i, 8).UnMerge
       Range(Cells(i, 8), Cells(i + cnt - 1, 8)).Value = s
        i = i + cnt - 1
    Next

End Sub

Private Sub CommandButton9_Click()
 For i = 2 To Range("a65536").End(xlUp).Row
        s = Cells(i, 9).Value
        cnt = Cells(i, 9).MergeArea.Count
        Cells(i, 9).UnMerge
       Range(Cells(i, 9), Cells(i + cnt - 1, 9)).Value = s
        i = i + cnt - 1
    Next

End Sub
Sub CheckBlanRow()
    Dim firstRow As Long, i As Long
        firstRow = ActiveSheet.UsedRange.Row
        LastRow = firstRow + ActiveSheet.UsedRange.Rows.Count - 1
        For i = LastRow To firstRow Step -1
            If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then
               Rows(i).Delete
        End If
    Next
End Sub
