Sub FormulasUltimaLinha()
    Dim ws As Worksheet
    Dim celula As Range
    Dim primeiraCelula As Range
    Dim Contador As String
    Dim CPrimeira As String
    Dim CSegunda As String
    Dim CTerceira As String
    Dim LugarFormula As String
    Dim ultimaLinha As Long

    Set ws = ThisWorkbook.Sheets("RESUMO") ' Defina sua planilha
    
    Dim Colunas(0 To 3) As String
    Dim IDColunas As Integer
    
    Colunas(0) = "B"
    Colunas(1) = "C"
    Colunas(2) = "E"
    Colunas(3) = "F"
    
    For IDColunas = LBound(Colunas) To UBound(Colunas)
    
        Set primeiraCelula = ws.Range(Colunas(IDColunas) & "3") ' Ponto de partida (ex: Coluna A)
        Contador = 0
        
        ' Encontra a última célula preenchida na coluna A
        ultimaLinha = ws.Cells(ws.Rows.Count, Colunas(IDColunas)).End(xlUp).Row
        
        ' Começa a procurar a partir da linha seguinte à última preenchida
        For Each celula In ws.Range(primeiraCelula.Offset(1, 0), ws.Cells(ultimaLinha + 1, Colunas(IDColunas)))
            If IsEmpty(celula) Or celula.Value = "" Then ' Verifica se está vazia
                If Contador = 0 Then
                    CPrimeira = celula.Offset(-1, 0).Address
                ElseIf Contador = 2 Then
                    CSegunda = celula.Offset(-1, 0).Address    
                ElseIf Contador = 4 Then 
                    CTerceira = celula.Offset(-1, 0).Address
                    LugarFormula = celula.Offset(1, 0).Address
                End If
                Contador = Contador + 1
            End If
        Next celula
        
        ws.Range(LugarFormula).FormulaLocal = "=" & CPrimeira & "+" & CSegunda & "+" & CTerceira
    Next IDColunas
End Sub
