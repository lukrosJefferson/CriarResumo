
Attribute VB_Name = "DeletarLinhasComCelulaVazia"

Sub DeletarLinhasComCelulaVazia()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ColunaParaVerificar As Long ' Defina o número da coluna (ex: 1 para A, 2 para B)

    ' --- CONFIGURAÇÕES ---
    Set ws = ThisWorkbook.ActiveSheet ' Ou Sheets("NomeDaSuaPlanilha")
    ColunaParaVerificar = 1 ' Verifica a coluna A (1)

    ' Encontra a última linha com dados na coluna definida
    lastRow = ws.Cells(ws.Rows.Count, ColunaParaVerificar).End(xlUp).Row

    ' Loop de baixo para cima para evitar pular linhas após a exclusão
    For i = lastRow To 1 Step -1 ' Começa da última linha até a primeira
        ' Verifica se a célula na coluna especificada está em branco
        If ws.Cells(i, ColunaParaVerificar).Value = "" Then
            ' Se estiver em branco, deleta a linha inteira
            ws.Range("A" & i & ":AU" & i).ClearContents ' Limpa de A at� Z (ajuste o Z)
        End If
    Next i

    MsgBox "Processo concluído!", vbInformation
End Sub
