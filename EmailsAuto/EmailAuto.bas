Attribute VB_Name = "EmailAuto"
Sub EmailAuto()

    Dim planilha As Worksheet
    Dim i As Long
    Dim ultimaLinha As Long
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim email As String, nome As String, msg As String

    Set planilha = ThisWorkbook.Sheets("Planilha1")
    Set OutlookApp = CreateObject("Outlook.Application")

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ultimaLinha = planilha.Cells(planilha.Rows.Count, 1).End(xlUp).Row

    For i = 2 To ultimaLinha
        email = planilha.Cells(i, 1).Value
        nome = planilha.Cells(i, 2).Value
        msg = planilha.Cells(i, 3).Value

        'Criar um novo e-mail a cada linha
        Set OutlookMail = OutlookApp.CreateItem(0)

        With OutlookMail
            .To = email                       'Destinatário
            .CC = "destinatario1@email.com"   'Cópia fixa
            .Subject = nome                   'Assunto
            .Body = msg                       'Corpo do e-mail
            .Display                          'Exibir antes de enviar
            '.Send                            'Usar se quiser enviar direto
        End With

    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

