
Attribute VB_Name = "EncontrarProximaVaziaLoop"

Sub EncontrarProximaVaziaLoop()
    Dim ws As Worksheet
    Dim celula As Range
    Dim primeiraCelula As Range
    Dim ultimaLinha As Long
    
    Set ws = ThisWorkbook.Sheets("RESUMO") ' Defina sua planilha
    
    Dim Colunas(0 To 1) As String
    Dim DColunas As Integer
    
    Colunas(0) = "B"
    Colunas(1) = "C"
    
    For DColunas = LBound(Colunas) To UBound(Colunas)
        
        Set primeiraCelula = ws.Range(Colunas(DColunas) &"3") ' Ponto de partida (ex: Coluna A)
        
        ' Encontra a última célula preenchida na coluna A
        ultimaLinha = ws.Cells(ws.Rows.Count, Colunas(DColunas)).End(xlUp).Row
        
        ' Começa a procurar a partir da linha seguinte à última preenchida
        For Each celula In ws.Range(primeiraCelula.Offset(1, 0), ws.Cells(ultimaLinha + 1, Colunas(DColunas)))
            If IsEmpty(celula) Or celula.Value = "" Then ' Verifica se está vazia
                If contagem = 0 Then
                    celula.Select
                    AutoSomaInteligente
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

Sub AutoSomaInteligente()
    Dim primeiraLinha As Long
    
    ' Se a célula acima estiver vazia, não faz nada para evitar erro
    If IsEmpty(ActiveCell.Offset(-1, 0)) Then Exit Sub
    
    ' Localiza o início do bloco de dados acima da célula ativa
    primeiraLinha = ActiveCell.Offset(-1, 0).End(xlUp).Row
    
    ' Aplica a fórmula apenas para o intervalo com dados
    ActiveCell.Formula = "=SUM(" & Range(Cells(primeiraLinha, ActiveCell.Column), _
                         ActiveCell.Offset(-1, 0)).Address(False, False) & ")"
End Sub
