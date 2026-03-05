Attribute VB_Name = "Módulo7"
Sub alterar()
    Dim planilha As Worksheet
    Dim UltimaLinha As Long
    Dim UltimaColuna As Long
    Dim tbl As ListObject
    Dim rngTabela As Range
    
    Set planilha = ThisWorkbook.Sheets("Valores")
    
    Application.ScreenUpdating = False
    
    
    ' Encontrar a última linha usada na coluna A
    UltimaLinha = planilha.Cells(planilha.Rows.Count, 1).End(xlUp).Row
    ' Encontrar a ultima coluna usada na linha 1 (cabeçalho)
    UltimaColuna = planilha.Cells(1, planilha.Columns.Count).End(xlToLeft).Column
    
    
    
   With planilha
    ' remover aspas
    .Range("B2:B" & UltimaLinha).Replace What:=Chr(34), Replacement:="", LookAt:=xlPart
    .Range("C2:C" & UltimaLinha).Replace What:=Chr(34), Replacement:="", LookAt:=xlPart

    ' remover separador de milhar (ex: "1.234,50" -> "1234,50")
    .Range("B2:B" & UltimaLinha).Replace What:=".", Replacement:="", LookAt:=xlPart
    .Range("C2:C" & UltimaLinha).Replace What:=".", Replacement:="", LookAt:=xlPart

    ' converter texto em número
    .Range("B2:B" & UltimaLinha).TextToColumns Destination:=.Range("B2"), _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierNone, FieldInfo:=Array(1, 1)
    .Range("C2:C" & UltimaLinha).TextToColumns Destination:=.Range("C2"), _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierNone, FieldInfo:=Array(1, 1)

    ' aplicar formato
    .Range("B2:B" & UltimaLinha).NumberFormat = "0"
    .Range("C2:C" & UltimaLinha).NumberFormat = "$ #.##0,00"
End With




    
    ' Definir o intervalo que será a tabela (do cabeçalho até a última linha/coluna usada)
    Set rngTabela = planilha.Range(planilha.Cells(1, 1), planilha.Cells(UltimaLinha, UltimaColuna))
    
    ' Verificar se a tabela já existe
    On Error Resume Next
    Set tbl = planilha.ListObjects("Tabela1")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        ' Criar nova tabela
        Set tbl = planilha.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=rngTabela, _
            XlListObjectHasHeaders:=xlYes)
        On Error Resume Next
        tbl.Name = "Tabela1" ' tenta nomear; se já existir, ignora erro
        On Error GoTo 0
    Else
        ' Redimensionar tabela existente para cobrir o novo intervalo
        tbl.Resize rngTabela
    End If
    
    ' Aplicar estilo à tabela
    tbl.TableStyle = "TableStyleMedium2"
    
    Application.ScreenUpdating = True
End Sub


