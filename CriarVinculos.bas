' Criado repositorio com git
' commit para marca data

Attribute VB_Name = "CriarVinculos"

Sub CriarVinculos()
    Dim ws As Worksheet
    Dim novaPlanilha As Worksheet
    Dim linha As Long
    Dim categoria As String
    Dim comparacao As String
    
    ' Adiciona uma nova planilha para consolidar os links
    Set novaPlanilha = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    novaPlanilha.Name = "RESUMO"
    
    ' Define a linha inicial para inserir os links
    linha = 2
    
    ' Loop por todas as planilhas na pasta de trabalho
    For Each ws In ThisWorkbook.Sheets
        ' Ignora as planilhas ocultas
        If ws.Visible = True Then
            ' Ignora a própria planilha "RESUMO"
            If ws.Name <> "RESUMO" Then
                ' Guarda o valor da celula na variável
                categoria = ws.Range("T2").Value
                
                ' Verifica se a nescessidade de pular mais lihas.
                If comparacao = "" Then
                    linha = linha + 1
                ElseIf comparacao = categoria Then
                    linha = linha + 1
                Else    
                    ' Move para a próxima linha
                    linha = linha + 4
                End If
                
                ' Guarda o valor da celula na variável
                comparacao = ws.Range("T2").Value
                ' Adiciona o nome da planilha na Coluna A
                'novaPlanilha.Range("A" & linha).Value = ws.Name
                novaPlanilha.Range("A" & linha).Formula = "='" & ws.Name & "'!B2"
                
                ' Cria o vínculos para as células D10, F10, G10, E10 e H10 da planilha atual nans Colunas B, C, D, E e F.
                ' A fórmula usa a sintaxe de referência externa do Excel: ='Nome da Planilha'!Célula
                novaPlanilha.Range("B" & linha).Formula = "='" & ws.Name & "'!D10"
                novaPlanilha.Range("C" & linha).Formula = "='" & ws.Name & "'!F10"
                novaPlanilha.Range("D" & linha).Formula = "='" & ws.Name & "'!G10"
                novaPlanilha.Range("E" & linha).Formula = "='" & ws.Name & "'!E10"
                novaPlanilha.Range("F" & linha).Formula = "='" & ws.Name & "'!H10"
                
                ' Cria o vínculos para as células D20, F20, G20, E20 e H20 da planilha atual nans Colunas H, I, J, K e L.
                ' A fórmula usa a sintaxe de referência externa do Excel: ='Nome da Planilha'!Célula
                novaPlanilha.Range("H" & linha).Formula = "='" & ws.Name & "'!D20"
                novaPlanilha.Range("I" & linha).Formula = "='" & ws.Name & "'!F20"
                novaPlanilha.Range("J" & linha).Formula = "='" & ws.Name & "'!G20"
                novaPlanilha.Range("K" & linha).Formula = "='" & ws.Name & "'!E20"
                novaPlanilha.Range("L" & linha).Formula = "='" & ws.Name & "'!H20"
                
                ' Cria o vínculos para as células D30, F30, G30, E30 e H30 da planilha atual nans Colunas N, O, P, Q e R.
                ' A fórmula usa a sintaxe de referência externa do Excel: ='Nome da Planilha'!Célula
                novaPlanilha.Range("N" & linha).Formula = "='" & ws.Name & "'!D30"
                novaPlanilha.Range("O" & linha).Formula = "='" & ws.Name & "'!F30"
                novaPlanilha.Range("P" & linha).Formula = "='" & ws.Name & "'!G30"
                novaPlanilha.Range("Q" & linha).Formula = "='" & ws.Name & "'!E30"
                novaPlanilha.Range("R" & linha).Formula = "='" & ws.Name & "'!H30"
                
                ' Cria o vínculos para as células D40, F40, G40, E40 e H40 da planilha atual nans Colunas T, U, V, W e X.
                ' A fórmula usa a sintaxe de referência externa do Excel: ='Nome da Planilha'!Célula
                novaPlanilha.Range("T" & linha).Formula = "='" & ws.Name & "'!D40"
                novaPlanilha.Range("U" & linha).Formula = "='" & ws.Name & "'!F40"
                novaPlanilha.Range("V" & linha).Formula = "='" & ws.Name & "'!G40"
                novaPlanilha.Range("W" & linha).Formula = "='" & ws.Name & "'!E40"
                novaPlanilha.Range("X" & linha).Formula = "='" & ws.Name & "'!H40"
                                
                ' Cria o vínculos para as células D50, F50, G50, E50 e H50 da planilha atual nans Colunas Z, AA, AB, AC e AD.
                ' A fórmula usa a sintaxe de referência externa do Excel: ='Nome da Planilha'!Célula
                novaPlanilha.Range("Z" & linha).Formula = "='" & ws.Name & "'!D50"
                novaPlanilha.Range("AA" & linha).Formula = "='" & ws.Name & "'!F50"
                novaPlanilha.Range("AB" & linha).Formula = "='" & ws.Name & "'!G50"
                novaPlanilha.Range("AC" & linha).Formula = "='" & ws.Name & "'!E50"
                novaPlanilha.Range("AD" & linha).Formula = "='" & ws.Name & "'!H50"
                
                ' Cria o vínculoS para aS célulaS B52, O52 e Z52 da planilha atual na Colunas AL, AR e Z52.
                ' A fórmula usa a sintaxe de referência externa do Excel: ='Nome da Planilha'!Célula
                novaPlanilha.Range("AL" & linha).Formula = "='" & ws.Name & "'!O52"
                novaPlanilha.Range("AP" & linha).Formula = "='" & ws.Name & "'!AA52"
                novaPlanilha.Range("AQ" & linha).Formula = "='" & ws.Name & "'!AB52"
                novaPlanilha.Range("AT" & linha).Formula = "='" & ws.Name & "'!B52"
                novaPlanilha.Range("AV" & linha).Formula = "='" & ws.Name & "'!Z52"
                
                
                
            End If
        End If
    Next ws
    
    ' Ajusta a largura das colunas para melhor visualização
    novaPlanilha.Columns("A:B").AutoFit
    
End Sub

