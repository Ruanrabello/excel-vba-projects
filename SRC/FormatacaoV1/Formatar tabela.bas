Attribute VB_Name = "Score"
Sub automacao()

    Dim planilha As Worksheet
    Dim ultimaColuna As Long
    Dim ultimaLinha As Long
    Dim i As Long
    Dim colFaixa As Long
    Dim colScore As Long
    
    Set planilha = ThisWorkbook.Sheets("Planilha1")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Identificar última coluna e linha
    ultimaColuna = planilha.Cells(1, planilha.Columns.Count).End(xlToLeft).Column
    ultimaLinha = planilha.Cells(planilha.Rows.Count, 1).End(xlUp).Row
    
    ' Encontrar colunas "Faixa" e "Score"
    For i = 1 To ultimaColuna
        If Trim(planilha.Cells(1, i).Value) = "Faixa" Then colFaixa = i
        If Trim(planilha.Cells(1, i).Value) = "Score" Then colScore = i
    Next i
    
    ' Preencher Faixa de acordo com Score
    For i = 2 To ultimaLinha
        If IsNumeric(planilha.Cells(i, colScore).Value) Then
            Select Case planilha.Cells(i, colScore).Value
                Case Is >= 80
                    planilha.Cells(i, colFaixa).Value = "Alta"
                Case 50 To 79
                    planilha.Cells(i, colFaixa).Value = "Media"
                Case Is < 50
                    planilha.Cells(i, colFaixa).Value = "Baixa"
                Case Else
                    planilha.Cells(i, colFaixa).Value = "Indefinido"
            End Select
        Else
            planilha.Cells(i, colFaixa).Value = "Indefinido"
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

            
 









