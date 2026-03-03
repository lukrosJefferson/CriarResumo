
Attribute VB_Name = "EncontrarProximaVaziaLoop"

Sub EncontrarProximaVaziaLoop()
    Dim ws As Worksheet
    Dim celula As Range
    Dim primeiraCelula As Range
    Dim ultimaLinha As Long
    
    Set ws = ThisWorkbook.Sheets("RESUMO") ' Defina sua planilha
    
    Dim Colunas(0 To 38) As String
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
    Colunas(11) = "R"
    Colunas(12) = "T"
    Colunas(13) = "U"
    Colunas(14) = "W"
    Colunas(15) = "X"
    Colunas(16) = "Z"
    Colunas(17) = "AA"
    Colunas(18) = "AC"
    Colunas(19) = "AF"
    Colunas(20) = "AD"
    Colunas(21) = "AG"
    Colunas(22) = "AI"
    Colunas(23) = "AJ"
    Colunas(24) = "Al"
    Colunas(25) = "AM"
    Colunas(26) = "AO"
    Colunas(27) = "AP"
    Colunas(28) = "AR"
    Colunas(29) = "AT"
    Colunas(30) = "D"
    Colunas(31) = "J"
    Colunas(32) = "P"
    Colunas(33) = "V"
    Colunas(34) = "AB"
    Colunas(35) = "AH"
    Colunas(36) = "AQ"
    Colunas(37) = "AU"
    Colunas(38) = "AV"
    
    For DColunas = LBound(Colunas) To UBound(Colunas)
        
        Set primeiraCelula = ws.Range(Colunas(DColunas) &"3") ' Ponto de partida (ex: Coluna A)
        
        ' Encontra a última célula preenchida na coluna A
        ultimaLinha = ws.Cells(ws.Rows.Count, Colunas(DColunas)).End(xlUp).Row
        
        ' Começa a procurar a partir da linha seguinte à última preenchida
        For Each celula In ws.Range(primeiraCelula.Offset(1, 0), ws.Cells(ultimaLinha + 1, Colunas(DColunas)))
            If IsEmpty(celula) Or celula.Value = "" Then ' Verifica se está vazia
                If contagem = 0 Then
                    If Colunas(DColunas) = "D" Or Colunas(DColunas) = "J" Or _
                    Colunas(DColunas) = "P" Or Colunas(DColunas) = "V" Or _
                    Colunas(DColunas) = "AB" Or Colunas(DColunas) = "AH" Then
                        celula.Select
                        MediaSeInteligente.MediaSeInteligente
                        contagem = 1
                    ElseIf Colunas(DColunas) = "AU" Then
                        celula.Value = "-"
                    Else
                        celula.Select
                        AutoSomaInteligente.AutoSomaInteligente
                        contagem = 1
                    End If
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
