Attribute VB_Name = "fangda"
Option Explicit
Sub fangda()
Selection.ShapeRange.IncrementLeft 50#
Selection.ShapeRange.ScaleWidth 10, msoFalse, msoScaleFromTopLeft
Selection.ShapeRange.ScaleHeight 15.1, msoFalse, msoScaleFromBottomRight
Selection.ShapeRange.PictureFormat.Brightness = 0.45
End Sub
