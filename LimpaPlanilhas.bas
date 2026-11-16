Attribute VB_Name = "LimpaPlanilhas"
Sub limpa_ponto()
Attribute limpa_ponto.VB_ProcData.VB_Invoke_Func = " \n14"
' limpa os pontos dos profissionais na planilha registro de ponto

    Application.ScreenUpdating = False
    
    Sheets("registro de ponto").Select
    'Range("C2:AG518").Select
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub
Sub limpa_dados()
Attribute limpa_dados.VB_Description = "limpa os dados dos profissionais na planilha dados"
Attribute limpa_dados.VB_ProcData.VB_Invoke_Func = " \n14"
' limpa os dados dos profissionais na planilha dados

    Application.ScreenUpdating = False
    
    Sheets("dados").Select
    
    ' limpa os dados dos profissionais educacao artistica
    Range("C3:AG10").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais fonoaudiologia
    Range("C15:AG26").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais nutricao
    Range("C31:AG52").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais pedagogia
    Range("C56:AG76").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais psicologia
    Range("C80:AG104").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais educacao fisica
    Range("C109:AG130").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais servico social
    Range("C135:AG149").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais expressao corporal
    Range("C154:AG156").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais terapia ocupacional
    Range("C169:AG181").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais terapia funcional
    Range("C186:AG541").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais enfermagem
    Range("C547:AG799").Select
    Selection.ClearContents
    
    ' limpa os dados dos profissionais tecnicos de enfermagem
    Range("C804:AG1040").Select
    Selection.ClearContents
    
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub
