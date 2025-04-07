Sub GenerateBoxLayoutFromItemList()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim shapeTop As Double: shapeTop = 30      ' Start Y Position
    Dim shapeLeft As Double: shapeLeft = 30    ' Start X Position
    Dim boxWidth As Double: boxWidth = 50
    Dim boxHeight As Double: boxHeight = 30
    Dim hGap As Double: hGap = 10              ' Horizontal gap between boxes
    Dim vGap As Double: vGap = 10              ' Vertical gap between rows
    Dim groupGapX As Double: groupGapX = 220   ' Gap between left/right groups
    Dim groupGapY As Double: groupGapY = 100   ' Gap between group rows
    Dim columnCount As Long: columnCount = 2   ' How many columns per row (like left/right group)
    
    Dim i As Long
    Dim itemName As String
    Dim totalItems As Long
    totalItems = WorksheetFunction.CountA(ws.Range("A:A"))
    
    Dim currentCol As Long: currentCol = 0
    Dim currentRow As Long: currentRow = 0

    For i = 1 To totalItems
        itemName = ws.Cells(i, 1).Value
        
        Dim baseTop As Double
        Dim baseLeft As Double
        
        baseLeft = shapeLeft + (currentCol * groupGapX)
        baseTop = shapeTop + (currentRow * groupGapY)
        
        ' Draw Title Box
        Dim titleBox As Shape
        Set titleBox = ws.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=baseLeft, Top:=baseTop, Width:=boxWidth * 3 + hGap * 2, Height:=18)
        titleBox.TextFrame2.TextRange.Text = itemName
        titleBox.TextFrame2.TextRange.Font.Size = 9
        titleBox.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        titleBox.Line.Weight = 1

        ' Draw 6 Boxes: 2 rows x 3 columns
        Dim r As Long, c As Long
        For r = 0 To 1
            For c = 0 To 2
                ws.Shapes.AddShape( _
                    Type:=msoShapeRectangle, _
                    Left:=baseLeft + c * (boxWidth + hGap), _
                    Top:=baseTop + 20 + r * (boxHeight + vGap), _
                    Width:=boxWidth, Height:=boxHeight).Line.Weight = 1
            Next c
        Next r

        ' Manage column & row index
        currentCol = currentCol + 1
        If currentCol >= columnCount Then
            currentCol = 0
            currentRow = currentRow + 1
        End If
    Next i

End Sub
