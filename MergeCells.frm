VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MergeCells 
   Caption         =   "UserForm1"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   OleObjectBlob   =   "MergeCells.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "MergeCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub A_Click()
  Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         JRow = .Range("B65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 1).Value = .Cells(i - 1, 1).Value Then
                 .Range(.Cells(i - 1, 1), .Cells(i, 1)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton10_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 12).Value = .Cells(i - 1, 12).Value Then
                 .Range(.Cells(i - 1, 12), .Cells(i, 12)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton11_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 10).Value = .Cells(i - 1, 10).Value Then
                 .Range(.Cells(i - 1, 10), .Cells(i, 10)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton12_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 11).Value = .Cells(i - 1, 11).Value Then
                 .Range(.Cells(i - 1, 11), .Cells(i, 11)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton13_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 13).Value = .Cells(i - 1, 13).Value Then
                 .Range(.Cells(i - 1, 13), .Cells(i, 13)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton14_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 14).Value = .Cells(i - 1, 14).Value Then
                 .Range(.Cells(i - 1, 14), .Cells(i, 14)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton15_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 15).Value = .Cells(i - 1, 15).Value Then
                 .Range(.Cells(i - 1, 15), .Cells(i, 15)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton16_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 16).Value = .Cells(i - 1, 16).Value Then
                 .Range(.Cells(i - 1, 16), .Cells(i, 16)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton17_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 17).Value = .Cells(i - 1, 17).Value Then
                 .Range(.Cells(i - 1, 17), .Cells(i, 17)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton18_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 18).Value = .Cells(i - 1, 18).Value Then
                 .Range(.Cells(i - 1, 18), .Cells(i, 18)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton19_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 19).Value = .Cells(i - 1, 19).Value Then
                 .Range(.Cells(i - 1, 19), .Cells(i, 19)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton2_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         JRow = .Range("B65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 2).Value = .Cells(i - 1, 2).Value Then
                 .Range(.Cells(i - 1, 2), .Cells(i, 2)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton20_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 20).Value = .Cells(i - 1, 20).Value Then
                 .Range(.Cells(i - 1, 20), .Cells(i, 20)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton21_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 21).Value = .Cells(i - 1, 21).Value Then
                 .Range(.Cells(i - 1, 21), .Cells(i, 21)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton22_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 22).Value = .Cells(i - 1, 22).Value Then
                 .Range(.Cells(i - 1, 22), .Cells(i, 22)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton23_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 23).Value = .Cells(i - 1, 23).Value Then
                 .Range(.Cells(i - 1, 23), .Cells(i, 23)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton24_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 24).Value = .Cells(i - 1, 24).Value Then
                 .Range(.Cells(i - 1, 24), .Cells(i, 24)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton25_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 25).Value = .Cells(i - 1, 25).Value Then
                 .Range(.Cells(i - 1, 25), .Cells(i, 25)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton26_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 26).Value = .Cells(i - 1, 26).Value Then
                 .Range(.Cells(i - 1, 26), .Cells(i, 26)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton3_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 3).Value = .Cells(i - 1, 3).Value Then
                 .Range(.Cells(i - 1, 3), .Cells(i, 3)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton4_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 4).Value = .Cells(i - 1, 4).Value Then
                 .Range(.Cells(i - 1, 4), .Cells(i, 4)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton5_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 5).Value = .Cells(i - 1, 5).Value Then
                 .Range(.Cells(i - 1, 5), .Cells(i, 5)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton6_Click()
Application.DisplayAlerts = False
     With ActiveSheet
   LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 6).Value = .Cells(i - 1, 6).Value Then
                 .Range(.Cells(i - 1, 6), .Cells(i, 6)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton7_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 7).Value = .Cells(i - 1, 7).Value Then
                 .Range(.Cells(i - 1, 7), .Cells(i, 7)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton8_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 8).Value = .Cells(i - 1, 8).Value Then
                 .Range(.Cells(i - 1, 8), .Cells(i, 8)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub

Private Sub CommandButton9_Click()
Application.DisplayAlerts = False
     With ActiveSheet
         LRow = .Range("A65536").End(xlUp).Row
         For i = LRow To 2 Step -1
             If .Cells(i, 9).Value = .Cells(i - 1, 9).Value Then
                 .Range(.Cells(i - 1, 9), .Cells(i, 9)).Merge
             End If
         Next

     End With
     Application.DisplayAlerts = True
End Sub
