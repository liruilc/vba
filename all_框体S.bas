Attribute VB_Name = "all_����S"
Sub all_����()
Call ��ʼ�������̿�
Call ��
Call �������߷�
Call ����yes_no
Call documents
Call decisions
Call links
End Sub
Sub ɾ�����̿�ť()

ActiveSheet.Shapes.Range(Array("Button 1", "Button 2", "Button 3", "Button 4", "Button 5", "Button 6", "Button 7")).Delete

End Sub

Sub startend()
Attribute startend.VB_Description = "Macro recorded 1/17/2009 by Max D. Christolear"
Attribute startend.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 95, 5, 85, 30).Select '��߾��� �ϱ߾��� ��� �߶�
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset25
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "��ʼ / ����"
    Cells(1, 1).Select
End Sub
Sub ��ʼ�������̿�() '1
ActiveSheet.Buttons.Add(0, 15, 85, 15).Select
Selection.OnAction = "all_����S.startend"
Selection.Characters.Text = "��ʼ / ����"
Call startend
End Sub
Sub step()
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 95, 45, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset23
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "��"
    Cells(1, 1).Select
End Sub
Sub ��() '2
ActiveSheet.Buttons.Add(0, 55, 50, 15).Select
Selection.OnAction = "all_����S.step"
Selection.Characters.Text = "��"
Call step
End Sub
Sub connector()
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 137, 87, 137, 117).Select '���ȣ�
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
    'Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "������ �� ��"
    Cells(1, 1).Select
    
End Sub
Sub �������߷�() '3
ActiveSheet.Buttons.Add(0, 95, 85, 15).Select
Selection.OnAction = "all_����S.connector"
Selection.Characters.Text = "������ �� ��"
Call connector
End Sub
Sub inout()
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartDecision, 95, 125, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset27
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "����  yes / no"
    Cells(1, 1).Select
End Sub
Sub ����yes_no() '4
ActiveSheet.Buttons.Add(0, 135, 85, 15).Select
Selection.OnAction = "all_����S.����yes_no"
Selection.Characters.Text = "����  yes / no"
Call inout
End Sub
Sub document()
    ActiveSheet.Shapes.AddShape(msoShapeOval, 95, 165, 85, 30).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset26
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "�ļ�"
    Cells(1, 1).Select
End Sub
Sub documents() '5
ActiveSheet.Buttons.Add(0, 175, 85, 15).Select
Selection.OnAction = "all_����S.document"
Selection.Characters.Text = "�ļ�"
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
Selection.OnAction = "all_����S.decision"
Selection.Characters.Text = "����"
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
Selection.OnAction = "all_����S.link"
Selection.Characters.Text = "����"
Call link
End Sub


