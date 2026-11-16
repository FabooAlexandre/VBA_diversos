Attribute VB_Name = "ExcluirLinha"
Sub ExcluirLinha()
' apaga a linha que contém a célula selecionada

    If Not ActiveCell Is Nothing Then
        Selection.EntireRow.Delete
    Else
        MsgBox "Selecione uma célula para apagar sua linha.", vbExclamation
    End If
End Sub
