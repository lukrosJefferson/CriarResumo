Attribute VB_Name = "CriarReseumo"

Sub CriarReseumo()

    ' Desativar o cálculo automático e outras funcionalidades para otimização
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False ' Desativa a atualização da tela para evitar piscar
    
    OrganizarAbasPorConteudo.OrganizarAbasPorConteudo
    
    CriarVinculos.CriarVinculos
    
    EncontrarProximaVaziaLoop.EncontrarProximaVaziaLoop
    
    ExpandirFormulaAutomatico.ExpandirFormulaAutomatico
    
    DeletarLinhasComCelulaVazia.DeletarLinhasComCelulaVazia
    
    ' Reativar o cálculo automático e a atualização da tela
    Application.ScreenUpdating = True ' Reativa a atualização da tela
    Application.Calculation = xlCalculationAutomatic
    
    ' RESUMO foi criado
    MsgBox "RESUMO foi criado"

End Sub