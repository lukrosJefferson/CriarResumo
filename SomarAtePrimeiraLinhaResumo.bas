Attribute VB_Name = "SomarAtePrimeiraLinhaResumo"

Sub SomarAtePrimeiraLinhaResumo()
    Dim linhaAtual As Long
    Dim colunaAtual As Integer
    
    ' Captura a linha e coluna da célula selecionada
    linhaAtual = ActiveCell.Row
    colunaAtual = ActiveCell.Column
    
    ' Verifica se a linha é maior que 1 para evitar erro de referência
    If linhaAtual > 1 Then
        ' Insere a fórmula SUM(Coluna1Linha1 : ColunaAtualLinhaAcima)
        ' Exemplo se estiver na A10: =SUM(A1:A9)
        ActiveCell.Formula = "=SUM(" & Cells(1, colunaAtual).Address & ":" & _
                             Cells(linhaAtual - 1, colunaAtual).Address & ")"
    Else
        MsgBox "Você está na primeira linha. Não há dados acima para somar."
    End If
End Sub
