Attribute VB_Name = "����������"

Sub asp_s() 'Option Explicit
asp = ActiveWorkbook.Sheets.Count '��ǰ����������������Ϊ
End Sub

Sub ���ҹ�����ѡ��() 'Public
Dim asp As Byte
Dim asps As Byte
Dim aspn As Integer
'Call asp_s
asp = ActiveWorkbook.Sheets.Count
For asps = 1 To asp
     
    aspn = Sheets(asps).Name
    'ActiveSheet("" & "asps" & "").Select
    MsgBox "��ǰ����������������Ϊ" & asp
    MsgBox "��ǰ�����������������Ϊ" & aspn
Next asps
End Sub
