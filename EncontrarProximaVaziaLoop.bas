
Attribute VB_Name = "EncontrarProximaVaziaLoop"

Sub EncontrarProximaVaziaLoop()
    Dim ws As Worksheet
    Dim celula As Range
    Dim primeiraCelula As Range
    Dim ultimaLinha As Long
    
    Set ws = ThisWorkbook.Sheets("RESUMO") ' Defina sua planilha
    
    Dim Colunas(0 To 28) As String
    Dim DColunas As Integer
    
    Colunas(0)  = "B"
    Colunas(1)  = "C"
    Colunas(2)  = "E"
    Colunas(3)  = "F"
    Colunas(4)  = "H"
    Colunas(5)  = "I"
    Colunas(6)  = "K"
    Colunas(7)  = "L"
    Colunas(8)  = "N"
    Colunas(9)  = "O"
    Colunas(10) = "Q"
    Colunas(11) = "T"
    Colunas(12) = "U"
    Colunas(13) = "W"
    Colunas(14) = "X"
    Colunas(15) = "Z"
    Colunas(16) = "AA"
    Colunas(17) = "AC"
    Colunas(18) = "AB"
    Colunas(19) = "AF"
    Colunas(20) = "AG"
    Colunas(21) = "AI"
    Colunas(22) = "AJ"
    Colunas(23) = "Al"
    Colunas(24) = "AM"
    Colunas(25) = "AO"
    Colunas(26) = "AP"
    Colunas(27) = "AR"
    Colunas(28) = "AT"
    
    For DColunas = LBound(Colunas) To UBound(Colunas)
        
        Set primeiraCelula = ws.Range(Colunas(DColunas) &"3") ' Ponto de partida (ex: Coluna A)
        
        ' Encontra a última célula preenchida na coluna A
        ultimaLinha = ws.Cells(ws.Rows.Count, Colunas(DColunas)).End(xlUp).Row
        
        ' Começa a procurar a partir da linha seguinte à última preenchida
        For Each celula In ws.Range(primeiraCelula.Offset(1, 0), ws.Cells(ultimaLinha + 1, Colunas(DColunas)))
            If IsEmpty(celula) Or celula.Value = "" Then ' Verifica se está vazia
                If contagem = 0 Then
                    celula.Select
                    AutoSomaInteligente.AutoSomaInteligente
                    contagem = 1
                ElseIf contagem = 1 Then
                    contagem = contagem + 1
                Else
                    contagem = 0
                End If
            End If
        Next celula
        contagem = 0
    Next DColunas
End Sub
