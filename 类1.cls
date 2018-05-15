VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "类1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub cfz()

 Dim i&, Myr&, arr

Dim d, k, t

Set d = CreateObject("Scripting.Dictionary")

Myr = Sheet1.[a65536].End(xlUp).Row

arr = Sheet1.Range("a1:g" & Myr)

For i = 2 To UBound(arr)

d(arr(i, 3)) = d(arr(i, 3)) + 1

Next

k = d.keys

t = d.items

Sheet2.Activate

 [a2].Resize(d.Count, 1) = Application.Transpose(k)

[b2].Resize(d.Count, 1) = Application.Transpose(t)

[a1].Resize(1, 2) = Array("姓名", "重复个数")

Set d = Nothing

End Sub
