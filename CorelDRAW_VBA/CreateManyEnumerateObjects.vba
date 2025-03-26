Sub CreateManyEnumerateObjects()
    
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
    
    'Solicita os parâmetros
    Dim FinalNumber As Long
    Dim GapHorizontal As Double
    Dim GapVertical As Double
    Dim ColumnQuantity As Long
    Dim InputValue As String
        
    InputValue = InputBox("Digite até qual número: ", "Até Qual Número")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    ElseIf Not IsNumeric(Int(InputValue)) Then
        MsgBox "Valor precisa ser um número"
        Exit Sub
    End If
    FinalNumber = CLng(InputValue)
    If FinalNumber < Int(Text.Text.Story) Then
        MsgBox "Valor menor que o número inicial"
        Exit Sub
    End If
    
    InputValue = InputBox("Digite o espaçamento horizontal ( em cm ( 1,24 ) ):")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    End If
    GapHorizontal = CDbl(InputValue)
    
    InputValue = InputBox("Digite o espaçamento vertical ( em cm ( 1,24 ) ):")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    End If
    GapVertical = CDbl(InputValue)
    
    InputValue = InputBox("Digite a quantidade de itens por linha:")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    End If
    ColumnQuantity = CDbl(InputValue)
    
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
        
        'Duplicar Objeto
        Object.Duplicate
        Object.PositionX = PositionX
        Object.PositionY = PositionY
        
        'Duplicar Texto
        Dim NewText As Shape
        Set NewText = Text.Duplicate
        NewText.Text.Story = i
        Call NewText.AlignToShape(cdrAlignHCenter, Object)
        Call NewText.AlignToShape(cdrAlignVCenter, Object)
    
        RowIndex = RowIndex + 1
        
        If RowIndex = ColumnQuantity Then
            RowIndex = 0
            ColumnIndex = ColumnIndex + 1
        End If
    Next i
    
    Text.Text.Story = FinalNumber

End Sub


Sub CreateManyEnumerateObjectsDuplicated()
    
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
    
    'Solicita os parâmetros
    Dim FinalNumber As Long
    Dim GapHorizontal As Double
    Dim GapVertical As Double
    Dim ColumnQuantity As Long
    Dim InputValue As String
        
    InputValue = InputBox("Digite até qual número: ", "Até Qual Número")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    ElseIf Not IsNumeric(Int(InputValue)) Then
        MsgBox "Valor precisa ser um número"
        Exit Sub
    End If
    FinalNumber = CLng(InputValue)
    If FinalNumber < Int(Text.Text.Story) Then
        MsgBox "Valor menor que o número inicial"
        Exit Sub
    End If
    
    InputValue = InputBox("Digite o espaçamento horizontal ( em cm ( 1,24 ) ):")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    End If
    GapHorizontal = CDbl(InputValue)
    
    InputValue = InputBox("Digite o espaçamento vertical ( em cm ( 1,24 ) ):")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    End If
    GapVertical = CDbl(InputValue)
    
    InputValue = InputBox("Digite a quantidade de itens por linha:")
    If InputValue = "" Then
        MsgBox "Valor inválido"
        Exit Sub
    End If
    ColumnQuantity = CDbl(InputValue)
    
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
        Dim j As Long
        For j = 1 To 2
        
            Dim PositionX As Double
            Dim PositionY As Double
            
            PositionX = xInitial + RowIndex * ActiveDocument.ToUnits(GapHorizontal, cdrCentimeter) + (RowIndex + 1) * Object.SizeWidth
            PositionY = yInitial - ColumnIndex * ActiveDocument.ToUnits(GapVertical, cdrCentimeter) - (ColumnIndex + 1) * Object.SizeHeight
            
            'Duplicar Objeto
            Object.Duplicate
            Object.PositionX = PositionX
            Object.PositionY = PositionY
            
            'Duplicar Texto
            Dim NewText As Shape
            Set NewText = Text.Duplicate
            NewText.Text.Story = i
            Call NewText.AlignToShape(cdrAlignHCenter, Object)
            Call NewText.AlignToShape(cdrAlignVCenter, Object)
        
            RowIndex = RowIndex + 1
            
            If RowIndex = ColumnQuantity Then
                RowIndex = 0
                ColumnIndex = ColumnIndex + 1
            End If
        
        Next j
    Next i
    
    Text.Text.Story = FinalNumber

End Sub
