'Para funcionar corretamente é necessário ter o objeto texto centralizado a um objeto de referencia e com alinhamento no centro
Sub Call_Form_CreateManyNumberedObjects()

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

    Call Form_CreateManyNumberedObjects.Show
    
End Sub

Sub Main_CreateManyNumberedObjects(ByVal TargetFinalNumber As String, ByVal TargetGapHorizontal As String, ByVal TargetGapVertical As String, ByVal TargetColumnsQuantity As String)
    
    ActiveDocument.Unit = cdrCentimeter
    
    Dim Text As Shape
    Dim Object As Shape
    
    'Atribui o texto e o objeto selecionados
    If ActiveSelection.Shapes(1).Type = cdrTextShape Then
        Set Text = ActiveSelection.Shapes(1)
        Set Object = ActiveSelection.Shapes(2)
    ElseIf ActiveSelection.Shapes(2).Type = cdrTextShape Then
        Set Text = ActiveSelection.Shapes(2)
        Set Object = ActiveSelection.Shapes(1)
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
    Dim MinimumDigits As Long
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
    MinimumDigits = Len(Text.Text.Story)
    
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
        Dim NewTextLabel As String
        
        PositionX = xInitial + RowIndex * ActiveDocument.ToUnits(GapHorizontal, cdrCentimeter) + (RowIndex + 1) * Object.SizeWidth
        PositionY = yInitial - ColumnIndex * ActiveDocument.ToUnits(GapVertical, cdrCentimeter) - (ColumnIndex + 1) * Object.SizeHeight
        
        Dim DuplicatedShapes As Shape
        Set DuplicatedShapes = ActiveSelection.Duplicate
        
        'Após a primeira execução do loop, os objetos duplicados são agrupados, necessário desagrupar antes
        If Not i = CurrentNumber Then
            Call DuplicatedShapes.Ungroup
        End If
        NewTextLabel = FormatWithZeros(i, MinimumDigits)
        
        DuplicatedShapes.PositionX = PositionX
        DuplicatedShapes.PositionY = PositionY
        
        Dim NewText As Shape
        If DuplicatedShapes.Shapes(1).Type = cdrTextShape Then
            Set NewText = DuplicatedShapes.Shapes(1)
        Else
            Set NewText = DuplicatedShapes.Shapes(2)
        End If
        NewText.Text.Story = NewTextLabel
        NewText.Name = NewTextLabel
        
        Dim DuplicatedGroup As Shape
        Set DuplicatedGroup = DuplicatedShapes.Group
        DuplicatedGroup.Name = NewTextLabel
    
        RowIndex = RowIndex + 1
        
        If RowIndex = ColumnQuantity Then
            RowIndex = 0
            ColumnIndex = ColumnIndex + 1
        End If
    Next i
    
    Dim OriginalText As Shape
    Set OriginalText = ActivePage.FindShape(MainTextShapeName)
    OriginalText.Text.Story = FormatWithZeros(Int(FinalNumber) + 1, MinimumDigits)
    
End Sub

Function FormatWithZeros(ByVal TargetValue As String, ByVal MinimumDigits As Long) As String
    
    If Len(TargetValue) >= MinimumDigits Then
        FormatWithZeros = TargetValue
        Exit Function
    Else
        FormatWithZeros = String(MinimumDigits - Len(TargetValue), "0") & TargetValue
    End If

End Function

