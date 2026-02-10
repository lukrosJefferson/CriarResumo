Sub FormulasUltimaLinha()
    Dim ws As Worksheet
    Dim celula As Range
    Dim primeiraCelula As Range
    Dim Contador As String
    Dim CPrimeira As String
    Dim CSegunda As String
    Dim CTerceira As String
    Dim LugarFormula As String
    
    Set ws = ThisWorkbook.Sheets("RESUMO") ' Defina sua planilha
    Set primeiraCelula = ws.Range("B3") ' Ponto de partida (ex: Coluna A)
    Contador = 0
    
    ' Encontra a última célula preenchida na coluna A
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Começa a procurar a partir da linha seguinte à última preenchida
    For Each celula In ws.Range(primeiraCelula.Offset(1, 0), ws.Cells(ultimaLinha + 1, "B"))
        If IsEmpty(celula) Or celula.Value = "" Then ' Verifica se está vazia
            If Contador = 0 Then
                CPrimeira = celula.Offset(-1, 0).Address
                Contador = Contador + 1
            ElseIf Contador = 1 Then
                CSegunda = celula.Offset(-1, 0).Address
                Contador = Contador + 1    
            Else
                CTerceira = celula.Offset(-1, 0).Address
                LugarFormula = celula.Offset(1, 0).Address
            End If    
        End If
    Next celula
    
    ws.Range(LugarFormula).FormulaLocal = "=" & CPrimeira & "+" & CSegunda & "+" & CTerceira
    
End Sub
