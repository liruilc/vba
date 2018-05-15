Attribute VB_Name = "取消后四位字符"
Sub Macro2()
Attribute Macro2.VB_Description = "Macro recorded 12/7/2011 by Z002TE2Z"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
' Macro recorded 12/7/2011 by Z002TE2Z

    Columns("A:A").Selection.Copy
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Selection.Replace What:=".JPG", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A1").Select 'jingrui.li
End Sub
