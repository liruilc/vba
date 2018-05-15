Attribute VB_Name = "活动工作表个数"

Sub asp_s() 'Option Explicit
asp = ActiveWorkbook.Sheets.Count '当前活动工作簿工作表个数为
End Sub

Sub 查找工作表选中() 'Public
Dim asp As Byte
Dim asps As Byte
Dim aspn As Integer
'Call asp_s
asp = ActiveWorkbook.Sheets.Count
For asps = 1 To asp
     
    aspn = Sheets(asps).Name
    'ActiveSheet("" & "asps" & "").Select
    MsgBox "当前活动工作簿工作表个数为" & asp
    MsgBox "当前活动工作簿工作表名称为" & aspn
Next asps
End Sub
