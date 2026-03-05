Attribute VB_Name = "GraficoAuto"
Sub GraficoDistribuicaoGenero() ' Cria a funcao

    ' Declara as variáveis
    Dim ws As Worksheet
    Dim wr As Worksheet
    Dim grafico As ChartObject
    Dim masculino As Integer, feminino As Integer
    Dim ultimaLinha As Long
    Dim i As Long
    Dim rngGrafico As Range
    Dim pastaDestino As String
    Dim nomePDF As String

    Application.ScreenUpdating = False 'Congela a atualização da tela durante a execução da macro.'
    Application.DisplayAlerts = False 'Desativa as mensagens de alerta e caixas de diálogo automáticas do Excel.'

    ' Define as planilhas
    Set ws = ThisWorkbook.Sheets("Dados")
    Set wr = ThisWorkbook.Sheets("Relatório")
    pastaDestino = "C:\Users\rrsilva\Downloads\Excel\Modulos\Excel.testes\" ' Define a pasta aonde vai ser salvo

    ' Conta os gêneros
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Identifica a ultima linha prenchida
    For i = 2 To ultimaLinha
        If Trim(LCase(ws.Cells(i, 3).Value)) = "feminino" Then ' trim(tira espaço) LCase(converte em minusculo) = ao valor ali ele soma mais um e armazena na variavel
            feminino = feminino + 1
        ElseIf Trim(LCase(ws.Cells(i, 3).Value)) = "masculino" Then
            masculino = masculino + 1
        End If
    Next i
    
    ' Remove gráficos anteriores
    For Each grafico In wr.ChartObjects 'SE tiver outro grafico na planilha relatorios, entao delete (planilha relatorios esta atribuida ao objeto wr)
        grafico.Delete
    Next grafico


    ' Insere os dados de contagem na planilha "Relatório"
    ws.Range("E1").Value = "Gênero"
    ws.Range("F1").Value = "Quantidade"
    ws.Range("E2").Value = "Feminino"
    ws.Range("F2").Value = feminino
    ws.Range("E3").Value = "Masculino"
    ws.Range("F3").Value = masculino

    Set rngGrafico = ws.Range("E1:F3") 'atribui a planilha criada com os dados a uma variavel, vai ser usada para desenvolver o grafico

    
    ' Cria gráfico de pizza
    Set grafico = wr.ChartObjects.Add(Left:=100, Width:=500, Top:=50, Height:=400)
    With grafico.Chart
        .SetSourceData Source:=rngGrafico
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Distribuição por Gênero"

        ' Estilo visual
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.ForeColor.RGB = RGB(0, 0, 0)

        ' Rótulos
        .SeriesCollection(1).ApplyDataLabels
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.ShowPercentage = True
    End With

    ' Configura página para exportação/impressão
    With wr.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    ' Exporta como PDF
    nomePDF = pastaDestino & "Grafico_Distribuicao_Genero.pdf"
    wr.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nomePDF, Quality:=xlQualityStandard
    
    ' Remove os dados temporários
    ws.Range("E1:F3").ClearContents

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "PDF gerado com sucesso!", vbInformation  ' Mensagem

End Sub

