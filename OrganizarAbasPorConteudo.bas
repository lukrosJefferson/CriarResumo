Sub OrganizarAbasPorConteudo()
    Dim ws As Worksheet
    Dim arrSheetNames As Variant
    Dim arrValues As Variant
    Dim i As Long, j As Long, tempName As Variant, tempValue As Variant
    Dim chaveCelula As Range ' CÃ©lula que contÃ©m o valor para ordenaÃ§Ã£o
    Dim nomeDaAba As String
    Dim valorDaAba As String
    
    ' Desabilita a atualização da tela para acelerar a macro
    Application.ScreenUpdating = False

    ' Define a cÃ©lula que serÃ¡ usada como critÃ©rio de ordenaÃ§Ã£o (ex: T2)
    Set chaveCelula = Range("T2") ' <<<<< AJUSTE AQUI: Mude "T2" para a cÃ©lula desejada

    ReDim arrSheetNames(1 To ThisWorkbook.Sheets.Count)
    ReDim arrValues(1 To ThisWorkbook.Sheets.Count)

    
    ' 1. Coletar nomes das abas e seus valores
    For Each ws In ThisWorkbook.Sheets
        nomeDaAba = ws.Name
        valorDaAba = ws.Range("T2").Value ' LÃª o valor da cÃ©lula T2 da aba
        arrSheetNames(ws.Index) = nomeDaAba
        arrValues(ws.Index) = valorDaAba
    Next ws
    ' Verifica se a planilha está visível
    
    ' 2. Ordenar os nomes das abas com base nos valores (Bubble Sort simples para exemplo)
    For i = 1 To UBound(arrSheetNames) - 1
        For j = i + 1 To UBound(arrSheetNames)
            ' Ordena em ordem crescente (xlAscending) - Altere para xlDescending se precisar
            If arrValues(i) > arrValues(j) Then
                ' Troca os valores
                tempValue = arrValues(i): arrValues(i) = arrValues(j): arrValues(j) = tempValue
                tempName = arrSheetNames(i): arrSheetNames(i) = arrSheetNames(j): arrSheetNames(j) = tempName
            End If
        Next j
    Next i

    ' 3. Reordenar as abas no Excel
    For i = 1 To UBound(arrSheetNames)
        ' Move a aba para a posiÃ§Ã£o correta (i-1 pois o Ã­ndice comeÃ§a em 0 para o .Move)
        ThisWorkbook.Sheets(arrSheetNames(i)).Move Before:=ThisWorkbook.Sheets(i)
    Next i
    
    MsgBox "Planilhas organizadas", vbInformation
    
End Sub