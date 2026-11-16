Attribute VB_Name = "LimparEspacosTexto"
Sub LimparEspacosTexto()
    Dim celula As Range
    Dim textoOriginal As String
    Dim textoLimpo As String
    Dim contador As Long
    
    ' Desativa a atualização de tela para executar mais rápido
    Application.ScreenUpdating = False
    
    contador = 0
    
    ' Percorre apenas as células selecionadas
    For Each celula In Selection
        ' Verifica se não está vazia, não é fórmula e se o conteúdo é estritamente Texto
        If Not IsEmpty(celula.Value) And Not celula.HasFormula Then
            If VarType(celula.Value) = vbString Then
                
                textoOriginal = celula.Value
                
                ' WorksheetFunction.Trim equivale à função ARRUMAR() do Excel:
                '    - Remove espaços no início e no final
                '    - Reduz espaços duplos/múltiplos no meio para apenas 1 espaço
                textoLimpo = Application.WorksheetFunction.Trim(textoOriginal)
                
                ' Atualiza a célula apenas se algo mudou
                If textoOriginal <> textoLimpo Then
                    celula.Value = textoLimpo
                    contador = contador + 1
                End If
                
            End If
        End If
    Next celula
    
    ' Reativa a atualização de tela
    Application.ScreenUpdating = True
    
    ' Exibe mensagem com o resultado
    MsgBox contador & " célula(s) de texto foram corrigidas!", vbInformation, "Limpeza Concluída"
End Sub
