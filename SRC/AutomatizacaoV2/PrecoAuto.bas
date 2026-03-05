Attribute VB_Name = "Módulo1"
Sub preco()
    Dim planilha As Worksheet
    Dim ColPreco As Long, ColAtivo As Long, ultimaLinha As Long, ultimaColuna As Long, i As Long
    Dim tbl As ListObject
    Dim novaColuna As ListColumn
    
    Set planilha = ThisWorkbook.Sheets("Relatorio")
    Set tbl = planilha.ListObjects("Tabela1")
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ultimaLinha = planilha.Cells(planilha.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = planilha.Cells(1, planilha.Columns.Count).End(xlToLeft).Column
    
    ' Encontrar colunas "Preco" e "Ativo"
    For i = 1 To ultimaColuna
        If Trim(planilha.Cells(1, i).Value) = "Preco" Then ColPreco = i
        If Trim(planilha.Cells(1, i).Value) = "Ativo" Then ColAtivo = i
    Next i
    
    ' Adiciona nova coluna à tabela
    Set novaColuna = tbl.ListColumns.Add
    novaColuna.Name = "Preco5%"
    
    ' Preenche e arredonda apenas a nova coluna
    For i = 2 To ultimaLinha
        If Trim(planilha.Cells(i, ColPreco).Value) <> "" And Trim(planilha.Cells(i, ColAtivo).Value) = "Sim" Then
            planilha.Cells(i, novaColuna.Index).Value = WorksheetFunction.Round(planilha.Cells(i, ColPreco).Value * 1.05, 2)
        End If
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub




