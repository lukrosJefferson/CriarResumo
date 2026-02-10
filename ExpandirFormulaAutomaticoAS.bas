
Attribute VB_Name = "ExpandirFormulaAutomatico"

Sub ExpandirFormulaAutomatico()
    
    ' Declarando variáveis
    Dim ws As Worksheet
    Dim UltimaLinha As Long
    Dim CelulaOriginal As Range
    Dim IntervaloDestino As Range
    Dim FimFor As Variant
    
    ' Atribuindo varáveis
    FimForDF = Range("B3").Value
    
    ' Declarando o Array
    Dim CeluasDaFormula(0 To 0) As String
    Dim CeluasDF As Integer
    
    ' Atribuindo o valor do Array
    CeluasDaFormula(0) = "AS3"
    
    ' Declarando o Array
    Dim FormulaDaColuna(0 To 0) As String
    Dim FormulaDC As Integer
    
    ' Atribuindo o valor do Array
    FormulaDaColuna(0) = "=SEERRO(AR3/22;0)"
    
    ' Declarando o Array
    Dim RangeDaFormula(0 To 0) As String
    Dim RangeDF As Integer
    
    ' Atribuindo o valor do Array
    RangeDaFormula(0) = "AS3:AS"
    
    ' Define a planilha em que você está trabalhando
    Set ws = ThisWorkbook.Sheets("RESUMO") ' Altere para o nome real da sua planilha
    
    ' Define a coluna que tem a fómula a ser expandida
    For CeluasDF = LBound(CeluasDaFormula) To UBound(CeluasDaFormula)
        
        ' Inserindo a formula
        ws.Range(CeluasDaFormula(CeluasDF)).FormulaLocal = FormulaDaColuna(FormulaDC)

        ' Encontra a última linha preenchida na coluna adjacente (por exemplo, Coluna A)    
        ' Considera que a coluna A tem dados que determinam até onde a fórmula deve ir
        UltimaLinha = ws.Range("B" & ws.Rows.Count).End(xlUp).Row

        ' Define o intervalo da célula original com a fórmula (ex: B2)
        Set CelulaOriginal = ws.Range(CeluasDaFormula(CeluasDF)) ' Altere para a célula onde sua fórmula começa

        ' Define o intervalo de destino para onde a fórmula será expandida
        Set IntervaloDestino = ws.Range(RangeDaFormula(RangeDF) & UltimaLinha)

        ' Aplica o AutoFill
        ' Copia a fórmula da célula original para todo o intervalo de destino
        CelulaOriginal.AutoFill Destination:=IntervaloDestino, Type:=xlFillDefault
        
        FormulaDC = FormulaDC + 1
        
        RangeDF = RangeDF + 1
    
    Next CeluasDF
    
End Sub
