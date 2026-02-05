
Attribute VB_Name = "ImportarTodosModulos"

Sub ImportarTodosModulos()
    Dim strPath As String
    Dim strFile As String
    Dim vbp As Object
    
    ' --- AJUSTE O CAMINHO DA PASTA AQUI ---
    strPath = "C:\Users\Logistica 3\OneDrive - LUKRO'S DISTRIBUIDORA DE PRODUTOS ALIMENTICIOS LTD\Documents\arquivos.bas\CriarResumo"
    ' ---------------------------------------
    
    Set vbp = ActiveWorkbook.VBProject
    strFile = Dir(strPath & "*.bas")
    
    Do While strFile <> ""
        On Error Resume Next
        ' Importa o arquivo
        vbp.VBComponents.Import strPath & strFile
        On Error GoTo 0
        ' Próximo arquivo
        strFile = Dir
    Loop
    
    MsgBox "Todos os módulos foram importados com sucesso!", vbInformation
End Sub
