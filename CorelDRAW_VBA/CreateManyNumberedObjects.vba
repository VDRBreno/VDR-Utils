'Necessário ter o objeto texto centralizado a um objeto de referencia e com alinhamento no centro
Sub CreateManyNumberedObjects(ByVal TargetFinalNumber As String, ByVal TargetGapHorizontal As String, ByVal TargetGapVertical As String, ByVal TargetColumnsQuantity As String)
    
    ActiveDocument.Unit = cdrCentimeter

    'Verifica se o usuário não selecionou exatamente dois itens: um texto e um objeto
    If ActiveSelection.Shapes.Count <> 2 Then
        MsgBox "Selecione APENAS 2 objetos, sendo um texto e um objeto"
        Exit Sub
    End If
    
    Dim Text As Shape
    Dim Object As Shape
    
    'Atribui o texto e o objeto selecionados
    If ActiveSelection.Shapes(1).Type = cdrTextShape Then
        Set Text = ActiveSelection.Shapes(1)
        Set Object = ActiveSelection.Shapes(2)
    ElseIf ActiveSelection.Shapes(2).Type = cdrTextShape Then
        Set Text = ActiveSelection.Shapes(2)
        Set Object = ActiveSelection.Shapes(1)
    Else
        MsgBox "Selecione dois objetos, sendo um deles um texto!"
        Exit Sub
    End If
    
    'Verifica se não é texto (contendo apenas números)
    If Not IsNumeric(Text.Text.Story) Then
        MsgBox "O objeto texto deve conter apenas números"
        Exit Sub
    End If
    
    Dim MainTextShapeName As String
    MainTextShapeName = "MainTextShape"
    Text.Name = MainTextShapeName
    
    'Solicita os parâmetros
    Dim FinalNumber As Long
    Dim GapHorizontal As Double
    Dim GapVertical As Double
    Dim ColumnQuantity As Long
    Dim InputValue As String
        
    If TargetFinalNumber = "" Or TargetGapHorizontal = "" Or TargetGapVertical = "" Or TargetColumnsQuantity = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    ElseIf Not IsNumeric(CLng(TargetFinalNumber)) Or Not IsNumeric(CDbl(TargetGapHorizontal)) Or Not IsNumeric(CDbl(TargetGapVertical)) Or Not IsNumeric(CLng(TargetColumnsQuantity)) Then
        MsgBox "Valor precisa ser um número"
        Exit Sub
    End If
    
    FinalNumber = CLng(TargetFinalNumber)
    If FinalNumber < Int(Text.Text.Story) Then
        MsgBox "Valor menor que o número inicial"
        Exit Sub
    End If
    
    GapHorizontal = CDbl(TargetGapHorizontal)
    GapVertical = CDbl(TargetGapVertical)
    ColumnQuantity = CDbl(TargetColumnsQuantity)
    
    Dim xInitial As Double
    Dim yInitial As Double
    xInitial = Object.PositionX
    yInitial = Object.PositionY - ActiveDocument.ToUnits(GapVertical, cdrCentimeter)
    
    Dim i As Long
    Dim CurrentNumber As Long
    CurrentNumber = Int(Text.Text.Story)
        
    Dim RowIndex As Long
    Dim ColumnIndex As Long
    RowIndex = 0
    ColumnIndex = 0
    
    For i = CurrentNumber To FinalNumber
        Dim PositionX As Double
        Dim PositionY As Double
        
        PositionX = xInitial + RowIndex * ActiveDocument.ToUnits(GapHorizontal, cdrCentimeter) + (RowIndex + 1) * Object.SizeWidth
        PositionY = yInitial - ColumnIndex * ActiveDocument.ToUnits(GapVertical, cdrCentimeter) - (ColumnIndex + 1) * Object.SizeHeight
        
        Dim DuplicatedShapes As Shape
        Set DuplicatedShapes = ActiveSelection.Duplicate
        DuplicatedShapes.PositionX = PositionX
        DuplicatedShapes.PositionY = PositionY
        
        Dim NewText As Shape
        If DuplicatedShapes.Shapes(1).Type = cdrTextShape Then
            Set NewText = DuplicatedShapes.Shapes(1)
        Else
            Set NewText = DuplicatedShapes.Shapes(2)
        End If
        NewText.Text.Story = i
    
        RowIndex = RowIndex + 1
        
        If RowIndex = ColumnQuantity Then
            RowIndex = 0
            ColumnIndex = ColumnIndex + 1
        End If
    Next i
    
    Dim OriginalText As Shape
    Set OriginalText = ActivePage.FindShape(MainTextShapeName)
    OriginalText.Text.Story = Int(FinalNumber) + 1
    
End Sub
