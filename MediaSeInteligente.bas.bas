
Attribute VB_Name = "MediaSeInteligente"

Sub MediaSeInteligente()
    Dim primeiraLinha As Long
    
    ' Se a célula acima estiver vazia, não faz nada para evitar erro
    If IsEmpty(ActiveCell.Offset(-1, 0)) Then Exit Sub
    
    ' Localiza o início do bloco de dados acima da célula ativa
    primeiraLinha = ActiveCell.Offset(-1, 0).End(xlUp).Row
    
    ' Aplica a fórmula apenas para o intervalo com dados
    ActiveCell.Formula = "=IFERROR" & "AVERAGEIF(" & Range(Cells(primeiraLinha, ActiveCell.Column), _
                     ActiveCell.Offset(-1, 0)).Address(False, False) & ", "">0"", " & _
                     Range(Cells(primeiraLinha, ActiveCell.Column), _
                     ActiveCell.Offset(-1, 0)).Address(False, False) & ")" & ",0)"

End Sub
