Attribute VB_Name = "FiltrosAIH"
Sub consulta_por_selecao()
Attribute consulta_por_selecao.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' Filtra pela célula selecionada em qualquer região da planilha
    ' Atalho: Ctrl + Q
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim valor As Variant
    Dim campo As Long
    
    Set ws = ActiveSheet
    
    ' Verifica se há célula ativa
    If ActiveCell Is Nothing Then Exit Sub
    
    ' Verifica se a célula tem valor
    If IsEmpty(ActiveCell.Value) Then
        MsgBox "Selecione uma célula para aplicar o filtro.", vbExclamation
        Exit Sub
    End If
    
    ' Define a região de dados com base na célula ativa
    Set rng = ActiveCell.CurrentRegion
    
    ' Calcula o campo relativo à região
    campo = ActiveCell.Column - rng.Columns(1).Column + 1
    
    ' Valor da célula
    valor = ActiveCell.Value
    
    ' Se já existe filtro, limpa antes
    If ws.AutoFilterMode Then
        rng.AutoFilter
    End If
    
    ' Aplica filtro
    rng.AutoFilter Field:=campo, Criteria1:=valor
    
End Sub

Sub Filtra_Prontuario_Selecionado()
' Macro que filtra por prontuario selecionado nas planilhas AIH

    If ActiveCell <> vbNullString Then
        Classifica_Tabela
        Sheets("Plan1").Select
        Range("B2") = ActiveCell
        Range("A2").Select
        Selection.ClearContents
        Range("C2:N2").Select
        Selection.ClearContents
        Sheets("AIH").Range("Consulta1[#All]").AdvancedFilter Action:= _
            xlFilterCopy, CriteriaRange:=Range("A1:N2"), CopyToRange:=Range("A5:N5"), _
            Unique:=False
        Sheets("Plan1").Select
        CalcularDiferencaDias
        Sheets("Plan1").Range("B2").Select
    Else
        MsgBox "Nenhum prontuário selecionado...", vbInformation
    End If
    
End Sub

Sub Gerar_Lista_Pacientes()
' Reabilitações do dia

    Sheets("Plan1").Select
    Range("H2") = "Alta por Decisão Médica"
    Range("I2") = "Reabilitação"
    Range("O5") = "Reabilitação: " & Range("A2")
    Range("O6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Filtro_Selecao
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("O6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    Range("A2:N2").Select
    Selection.ClearContents
    
End Sub

Sub Gerar_Lista_Pacientes_Cirurgicos()
' Cirurgias do dia

    Sheets("Plan1").Select
    Range("H2") = "Alta por Decisão Médica"
    Range("I2") = "Cirúrgico"
    Range("O5") = "Cirurgias: " & Range("A2")
    Range("O6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Filtro_Selecao
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("O6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    Range("A2:N2").Select
    Selection.ClearContents
    
End Sub
Sub Filtro_Selecao()
' Filtra campos em geral

    Classifica_Tabela
    Sheets("Plan1").Select
    Range("A2:N2").Select
    'Range("I2") = "Reabilitação"
    Sheets("AIH").Range("Consulta1[#All]").AdvancedFilter Action:= _
        xlFilterCopy, CriteriaRange:=Range("A1:N2"), CopyToRange:=Range("A5:N5"), _
        Unique:=False
    Sheets("Plan1").Select
    
End Sub

Sub CalcularDiferencaDias()
'faz as marcações onde fechar a AIH ao lado das datas

    Dim dataP, dataA As Date
    Dim diferencaDias As Long
    Dim x As Integer

    Sheets("Plan1").Select
    x = 6 'começa na linha 6 da "Plan1"
    Range("F" & x).Select
    Negritar
    dataP = Range("F" & x + 1).Value 'data posterior
    dataA = Range("G" & x).Value 'data anterior

    Do Until dataP = vbNullString 'percorre até não encontrar dia seguinte
        'Calcula a diferença em dias
        diferencaDias = DateDiff("d", dataP, dataA)
        diferencaDias = Abs(diferencaDias) 'diferença em números positivos absolutos

        If diferencaDias > 3 Then 'escreve onde fechar a AIH nesta condição específica
            Sheets("Plan1").Range("J" & x) = "<-Fecha aqui"
            Range("G" & x).Select
            Negritar
            Range("F" & x + 1).Select
            Negritar
        End If
        
        x = x + 1
        dataP = Range("F" & x + 1).Value
        dataA = Range("G" & x).Value
    Loop

    If Sheets("Plan1").Range("H" & x) = "Alta por Decisão Médica" Then 'escreve na última linha desde que seja alta por decisão médica
        Sheets("Plan1").Range("J" & x) = "<-Fecha aqui"
        Range("G" & x).Select
        Negritar
    End If

End Sub

Sub Negritar()
    Selection.Font.Bold = True
End Sub

Sub Classifica_Tabela()
' Macro para classificar a tabela

    Sheets("AIH").Select
    Range("Consulta1[[#Headers],[MÊS_ALTA]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("AIH").ListObjects("Consulta1").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("AIH").ListObjects("Consulta1").Sort. _
        SortFields.Add Key:=Range("Consulta1[DT_ALTA]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("AIH").ListObjects("Consulta1").Sort. _
        SortFields.Add Key:=Range("Consulta1[NR_REGISTRO]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("AIH").ListObjects("Consulta1"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Sub pacientes_Sem_AIH()
' Macro busca a lista de pacientes sem AIH no dia específico

    Sheets("Plan1").Select
    Range("B2:N2").Select
    Selection.ClearContents
    'Range("H2") = "Alta por Decisão Médica"
    'Range("I2") = "Reabilitação"
    Range("K2") = "="    'sem numeração de AIH (<> corresponde ao inverso)
    Range("O5") = "SEM AIH: " & Range("A2")
    Range("O6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Filtro_Selecao
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("O6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:N2").Select
    Selection.ClearContents
    
End Sub

