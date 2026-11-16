Attribute VB_Name = "Checks"
Sub Verde()
Attribute Verde.VB_Description = "Pinta de verde a célula selecionada"
Attribute Verde.VB_ProcData.VB_Invoke_Func = " \n14"
' PintarVerde Macro
' Pinta de verde a célula selecionada

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Vermelho()
Attribute Vermelho.VB_Description = "Pintar de vermelho a célula selecionada"
Attribute Vermelho.VB_ProcData.VB_Invoke_Func = " \n14"
' PintarVermelho Macro
' Pintar de vermelho a célula selecionada

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub CancelarRecepção()
Attribute CancelarRecepção.VB_Description = "Digita o texto ""Cancelar recepção"" na célula selecionada"
Attribute CancelarRecepção.VB_ProcData.VB_Invoke_Func = " \n14"
' CancelarRecepção Macro
' Digita o texto "Cancelar recepção" na célula selecionada

    ActiveCell.FormulaR1C1 = "Cancelar recepção"
    
End Sub
Sub RecepcionarAtendimento()
Attribute RecepcionarAtendimento.VB_Description = "Digita o texto ""Recepcionar atendimento"" na célula selecionada"
Attribute RecepcionarAtendimento.VB_ProcData.VB_Invoke_Func = " \n14"
' RecepcionarAtendimento Macro
' Digita o texto "Recepcionar atendimento" na célula selecionada

    ActiveCell.FormulaR1C1 = "Recepcionar atendimento"
    
End Sub
Sub RecepcionarTeleatendimento()
Attribute RecepcionarTeleatendimento.VB_Description = "Escreve ""Recepcionar Teleatendimento"" na célula selecionada"
Attribute RecepcionarTeleatendimento.VB_ProcData.VB_Invoke_Func = " \n14"
' RecepcionarTeleatendimento Macro
' Escreve "Recepcionar Teleatendimento" na célula selecionada

    ActiveCell.FormulaR1C1 = "Recepcionar teleatendimento"
    
End Sub
Private Sub AlterarDesistencia()
' AlterarDesistencia Macro
' Escreve "Alterar para desistência" na célula selecionada

    ActiveCell.FormulaR1C1 = "Alterar para desistência"
    
End Sub
