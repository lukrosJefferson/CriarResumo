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
    Dim NumeroLinha

    Set ws = ThisWorkbook.Sheets("RESUMO") ' Defina sua planilha
    
    Dim Colunas(0 To 39) As String
    Dim IDColunas As Integer
    
    Colunas(0)  = "B"
    Colunas(1)  = "C"
    Colunas(2)  = "E"
    Colunas(3)  = "F"
    Colunas(4)  = "G"
    Colunas(5)  = "H"
    Colunas(6)  = "I"
    Colunas(7)  = "K"
    Colunas(8)  = "L"
    Colunas(9)  = "M"
    Colunas(10) = "N"
    Colunas(11) = "O"
    Colunas(12) = "Q"
    Colunas(13) = "R"
    Colunas(14) = "S"
    Colunas(15) = "T"
    Colunas(16) = "U"
    Colunas(17) = "W"
    Colunas(18) = "X"
    Colunas(19) = "Y"
    Colunas(20) = "Z"
    Colunas(21) = "AA"
    Colunas(22) = "AC"
    Colunas(23) = "AD"
    Colunas(24) = "AE"
    Colunas(25) = "AF"
    Colunas(26) = "AG"
    Colunas(27) = "AI"
    Colunas(28) = "AJ"
    Colunas(29) = "AK"
    Colunas(30) = "AL"
    Colunas(31) = "AM"
    Colunas(32) = "AN"
    Colunas(33) = "AO"
    Colunas(34) = "AP"
    Colunas(35) = "AQ"
    Colunas(36) = "AR"
    Colunas(37) = "AS"
    Colunas(38) = "AT"
    Colunas(39) = "AU"
    
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
                    NumeroLinha  = celula.Offset(1, 0).Row
                End If
                Contador = Contador + 1
            End If
        Next celula
        
        If Colunas(IDColunas) = "G" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(F" & NumeroLinha & "/B" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "M" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(L" & NumeroLinha & "/H" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "S" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(R" & NumeroLinha & "/N" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "Y" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(X" & NumeroLinha & "/T" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "AE" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(AD" & NumeroLinha & "/Z" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "AK" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(AJ" & NumeroLinha & "/AF" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "AN" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(AM" & NumeroLinha & "/AF" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "AQ" Then
            ws.Range(LugarFormula).FormulaLocal = "=SEERRO(AP" & NumeroLinha & "/AF" & NumeroLinha & ";0)"
        ElseIf Colunas(IDColunas) = "AS" Then
            SomarAtePrimeiraLinhaResumo.SomarAtePrimeiraLinhaResumo
        Else
            ws.Range(LugarFormula).FormulaLocal = "=" & CPrimeira & "+" & CSegunda & "+" & CTerceira
        End If
    Next IDColunas
End Sub
