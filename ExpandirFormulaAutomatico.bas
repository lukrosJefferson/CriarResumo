
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
    Dim CeluasDaFormula(0 To 15) As String
    Dim CeluasDF As Integer
    
    ' Atribuindo o valor do Array
    CeluasDaFormula(0)  = "G3"
    CeluasDaFormula(1)  = "M3"
    CeluasDaFormula(2)  = "S3"
    CeluasDaFormula(3)  = "Y3"
    CeluasDaFormula(4)  = "AE3"
    CeluasDaFormula(5)  = "AF3"
    CeluasDaFormula(6)  = "AG3"
    CeluasDaFormula(7)  = "AH3"
    CeluasDaFormula(8)  = "AI3"
    CeluasDaFormula(9)  = "AJ3"
    CeluasDaFormula(10) = "AK3"
    CeluasDaFormula(11) = "AM3"
    CeluasDaFormula(12) = "AN3"
    CeluasDaFormula(13) = "AR3"
    CeluasDaFormula(14) = "AS3"
    CeluasDaFormula(15) = "AW3"
    
    ' Declarando o Array
    Dim FormulaDaColuna(0 To 15) As String
    Dim FormulaDC As Integer
    
    ' Atribuindo o valor do Array
    FormulaDaColuna(0)  = "=SEERRO(F3/B3;0)"
    FormulaDaColuna(1)  = "=SEERRO(L3/H3;0)"
    FormulaDaColuna(2)  = "=SEERRO(R3/N3;0)"
    FormulaDaColuna(3)  = "=SEERRO(X3/T3;0)"
    FormulaDaColuna(4)  = "=SEERRO(AD3/Z3;0)"
    FormulaDaColuna(5)  = "=SEERRO(B3+H3+N3+T3+Z3;0)"
    FormulaDaColuna(6)  = "=SEERRO(C3+I3+O3+U3+AA3;0)"
    FormulaDaColuna(7)  = "=SEERRO((D3+J3+P3+V3)/4;0)"
    FormulaDaColuna(8)  = "=SEERRO(E3+K3+Q3+W3+AC3;0)"
    FormulaDaColuna(9)  = "=SEERRO(F3+L3+R3+X3+AD3;0)"
    FormulaDaColuna(10) = "=SEERRO(AJ3/AF3;0)"
    FormulaDaColuna(11) = "=SEERRO(AJ3+AL3;0)"
    FormulaDaColuna(12) = "=SEERRO(AM3/AF3;0)"
    FormulaDaColuna(13) = "=SEERRO(AM3+AO3+AP3+AQ3;0)"
    FormulaDaColuna(14) = "=SEERRO(AP3/AF3;0)"
    FormulaDaColuna(15) = "=SEERRO(AG3/AT3;0)"
    
    ' Declarando o Array
    Dim RangeDaFormula(0 To 15) As String
    Dim RangeDF As Integer
    
    ' Atribuindo o valor do Array
    RangeDaFormula(0)  = "G3:G"
    RangeDaFormula(1)  = "M3:M"
    RangeDaFormula(2)  = "S3:S"
    RangeDaFormula(3)  = "Y3:Y"
    RangeDaFormula(4)  = "AE3:AE"
    RangeDaFormula(5)  = "AF3:AF"
    RangeDaFormula(6)  = "AG3:AG"
    RangeDaFormula(7)  = "AH3:AH"
    RangeDaFormula(8)  = "AI3:AI"
    RangeDaFormula(9)  = "AJ3:AJ"
    RangeDaFormula(10) = "AK3:AK"
    RangeDaFormula(11) = "AM3:AM"
    RangeDaFormula(12) = "AN3:AN"
    RangeDaFormula(13) = "AR3:AR"
    RangeDaFormula(14) = "AS3:AS"
    RangeDaFormula(15) = "AW3:AW"
    
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
