Attribute VB_Name = "AlternarDestaque"
Sub AlternarDestaque()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim colPessoa As Long
    Dim i As Long, corAtual As Long
    Dim nomeAnterior As String
    Dim corA As Long, corB As Long
    Dim rngAll As Range, rngSortKey As Range
    Dim temCor As Boolean
    
    'Set ws = ActiveWorkbook.Sheets("Data")
    Set ws = ActiveWorkbook.ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    colPessoa = 4 ' Coluna D = NM_PESSOA
    
    ' Verifica se já existe cor aplicada na primeira linha de dados
    temCor = (ws.Cells(2, 1).Interior.ColorIndex <> xlNone)
    
    If temCor Then
        ' LIMPAR preenchimento
        ws.Range(ws.Cells(2, 1), ws.Cells(ultimaLinha, 5)).Interior.ColorIndex = xlNone
    Else
        ' ORDENAR pela coluna NM_PESSOA
        Set rngAll = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, 5))
        Set rngSortKey = ws.Range(ws.Cells(2, colPessoa), ws.Cells(ultimaLinha, colPessoa))
        
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=rngSortKey, _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange rngAll
            .Header = xlYes
            .Apply
        End With
        
        ' DESTACAR com cores alternadas
        corA = RGB(220, 230, 241) ' azul claro
        corB = RGB(255, 242, 204) ' amarelo claro
        corAtual = corA
        nomeAnterior = Trim(ws.Cells(2, colPessoa).Value)
        
        For i = 2 To ultimaLinha
            If Trim(ws.Cells(i, colPessoa).Value) <> nomeAnterior Then
                corAtual = IIf(corAtual = corA, corB, corA)
                nomeAnterior = Trim(ws.Cells(i, colPessoa).Value)
            End If
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 5)).Interior.Color = corAtual
        Next i
    End If
End Sub

