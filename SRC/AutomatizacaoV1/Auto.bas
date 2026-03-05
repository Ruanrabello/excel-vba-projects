Attribute VB_Name = "Auto"
Sub Auto()
    ' Declarações de variáveis
    Dim planilha As Worksheet
    Dim Destino As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaColuna As Long
    Dim dataRange As Range
    Dim i As Long
    Dim ii As Long
    Dim tbl As ListObject
    Dim rngTabela As Range
    Dim caminhoPDF As String
    Dim nomeArquivoPDF As String
    Dim ColCliente As Long          ' índice da coluna "Cliente"
    Dim OutlookApp As Object        ' instância da aplicação Outlook
    Dim OutlookMail As Object       ' item de e-mail (MailItem)
       
    ' Evita popups e acelera a execução
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Cria/obtém instância do Outlook (uma vez só)
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Define a planilha de origem
    Set planilha = ThisWorkbook.Sheets("Nomes")

    ' Tentar obter a folha destino; se não existir, cria
    On Error Resume Next
    Set Destino = ThisWorkbook.Sheets("Alunos_aprovados")
    On Error GoTo 0
    If Destino Is Nothing Then
        Set Destino = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        Destino.Name = "Alunos_aprovados"
    End If

    ' Encontrar última linha e última coluna da planilha de origem
    ultimaLinha = planilha.Cells(planilha.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = planilha.Cells(1, planilha.Columns.Count).End(xlToLeft).Column
    
    ' Copiar cabeçalho para a folha Destino (linha 1)
    planilha.Rows(1).Copy Destination:=Destino.Rows(1)

    ' Definir o intervalo de dados (inclui cabeçalho)
    Set dataRange = planilha.Range(planilha.Cells(1, 1), planilha.Cells(ultimaLinha, ultimaColuna))

    ' Aplica filtro para selecionar APROVADO na coluna 4
    dataRange.AutoFilter Field:=4, Criteria1:="=APROVADO"
    
    ' Percorre as linhas da origem de baixo para cima
    For i = ultimaLinha To 2 Step -1
        If Trim(planilha.Cells(i, 4).Value) = "APROVADO" Then

            ' Copia a linha aprovada para a linha 2 do Destino (mantendo cabeçalho na linha 1)
            planilha.Rows(i).Copy Destination:=Destino.Rows(2)

            ' Define o intervalo que contém o cabeçalho + a linha copiada (linhas 1 a 2)
            Set rngTabela = Destino.Range(Destino.Cells(1, 1), Destino.Cells(2, ultimaColuna))

            ' Verificar se a tabela já existe; se não, cria; se sim, redimensiona
            On Error Resume Next
            Set tbl = Destino.ListObjects("Tabela1")
            On Error GoTo 0

            If tbl Is Nothing Then
                Set tbl = Destino.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngTabela, XlListObjectHasHeaders:=xlYes)
                On Error Resume Next
                tbl.Name = "Tabela1"
                On Error GoTo 0
            Else
                ' Redimensiona a tabela para incluir somente cabeçalho + a linha atual
                tbl.Resize rngTabela
            End If

            ' Aplica estilo e ajusta colunas
            tbl.TableStyle = "TableStyleMedium2"
            Destino.Columns.AutoFit
            
            ' Encontrar a coluna "Cliente" no cabeçalho da planilha Destino
            ColCliente = 0
            For ii = 1 To ultimaColuna
                If Trim(Destino.Cells(1, ii).Value) = "Cliente" Then
                    ColCliente = ii
                    Exit For
                End If
            Next ii
            
            ' Se não encontrou a coluna "Cliente", usa um nome genérico para o arquivo
            If ColCliente = 0 Then
                nomeArquivoPDF = "C:\Users\rrsilva\Downloads\Excel\Modulos\testepdf\Relatorio_Sem_Cliente.pdf"
            Else
                ' Monta o caminho/nome do arquivo PDF (valor da coluna Cliente na linha 2)
                caminhoPDF = "C:\Users\rrsilva\Downloads\Excel\Modulos\testepdf\"
                nomeArquivoPDF = caminhoPDF & Trim(Destino.Cells(2, ColCliente).Value) & ".pdf"
            End If

            ' Exporta o cabeçalho + linha atual como PDF
            Destino.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nomeArquivoPDF, Quality:=xlQualityStandard

            ' Cria o e-mail e anexa o PDF
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = "destinatario@email.com"
                .CC = "destinatario1@email.com"
                .Subject = "Relatório Automático"
                .Body = "Segue em anexo o relatório."
                .Attachments.Add nomeArquivoPDF
                .Display   ' usar .Send para enviar direto
            End With

            ' Limpa somente a linha 2 (prepara para a próxima cópia)
            Destino.Rows(2).ClearContents

            ' Opcional: se quiser remover a tabela completamente, descomente:
            ' On Error Resume Next: tbl.Delete: On Error GoTo 0

        End If
    Next i

    ' Remove o filtro aplicado (boa prática)
    On Error Resume Next
    planilha.ShowAllData
    On Error GoTo 0

    ' Mensagem final
    MsgBox "E-mail(s) preparados/enviados com sucesso!"

    ' Restaura configurações da aplicação
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

