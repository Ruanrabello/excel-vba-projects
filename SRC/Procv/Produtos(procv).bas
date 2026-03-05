Attribute VB_Name = "Produtos"
Sub Produtos()

    Dim planilha As Worksheet
    Dim planilha1 As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim resultado As Variant
    
    Set planilha = ThisWorkbook.Sheets("Planilha1")
    Set planilha1 = ThisWorkbook.Sheets("Planilha2")
    
    ultimaLinha = planilha1.Cells(planilha1.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        'Meio que um procv
        '=VLOOKUP(valor_procurado; tabela; numero_coluna; [procurar_exato_ou_aproximado])
        'planilha1.Cells(i, 2).Value = Application.VLookup(planilha1.Cells(i, 1).Value, planilha.Range("A:B"), 2, False) ( pode usar assim tambem, porem nao vai aparecer aqulea ajuda do excel, sobre oque vem/ oque vc precisa colocar)
        planilha1.Cells(i, 2).Value = Application.WorksheetFunction.VLookup(planilha1.Cells(i, 1).Value, planilha.Range("A:B"), 2, False)
        
        
    Next i
End Sub


