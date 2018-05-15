Attribute VB_Name = "all_框体S"
Sub all_框体()
Call 开始结束流程框
Call 步
Call 连接器线符
Call 进出yes_no
Call documents
Call decisions
Call links
End Sub
Sub 删除流程框按钮()

ActiveSheet.Shapes.Range(Array("Button 1", "Button 2", "Button 3", "Button 4", "Button 5", "Button 6", "Button 7")).Delete

End Sub

Sub startend()
Attribute startend.VB_Description = "Macro recorded 1/17/2009 by Max D. Christolear"
Attribute startend.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 95, 5, 85, 30).Select '左边距离 上边距离 宽度 高度
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset25
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "开始 / 结束"
    Cells(1, 1).Select
End Sub
Sub 开始结束流程框() '1
ActiveSheet.Buttons.Add(0, 15, 85, 15).Select
Selection.OnAction = "all_框体S.startend"
Selection.Characters.Text = "开始 / 结束"
Call startend
End Sub
Sub step()
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 95, 45, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset23
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "步"
    Cells(1, 1).Select
End Sub
Sub 步() '2
ActiveSheet.Buttons.Add(0, 55, 50, 15).Select
Selection.OnAction = "all_框体S.step"
Selection.Characters.Text = "步"
Call step
End Sub
Sub connector()
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 137, 87, 137, 117).Select '长度；
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
    'Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "连接器 线 符"
    Cells(1, 1).Select
    
End Sub
Sub 连接器线符() '3
ActiveSheet.Buttons.Add(0, 95, 85, 15).Select
Selection.OnAction = "all_框体S.connector"
Selection.Characters.Text = "连接器 线 符"
Call connector
End Sub
Sub inout()
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartDecision, 95, 125, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset27
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "进出  yes / no"
    Cells(1, 1).Select
End Sub
Sub 进出yes_no() '4
ActiveSheet.Buttons.Add(0, 135, 85, 15).Select
Selection.OnAction = "all_框体S.进出yes_no"
Selection.Characters.Text = "进出  yes / no"
Call inout
End Sub
Sub document()
    ActiveSheet.Shapes.AddShape(msoShapeOval, 95, 165, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset26
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "文件"
    Cells(1, 1).Select
End Sub
Sub documents() '5
ActiveSheet.Buttons.Add(0, 175, 85, 15).Select
Selection.OnAction = "all_框体S.document"
Selection.Characters.Text = "文件"
Call document
End Sub
Sub decision()
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartData, 95, 205, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset23
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "decision"
    Cells(1, 1).Select
End Sub
Sub decisions() '6
ActiveSheet.Buttons.Add(0, 215, 85, 15).Select
Selection.OnAction = "all_框体S.decision"
Selection.Characters.Text = "决定"
Call decision
End Sub
Sub link()
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartDocument, 95, 245, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset22
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "link"
    Cells(1, 1).Select
End Sub
Sub links() '7
ActiveSheet.Buttons.Add(0, 255, 85, 15).Select
Selection.OnAction = "all_框体S.link"
Selection.Characters.Text = "链接"
Call link
End Sub


