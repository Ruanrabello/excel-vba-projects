Sub Verificar_Condicoes()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.ActiveSheet 'Pode trocar para Worksheets("NomeDaAba") se quiser fixar
    ultimaLinha = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row 'Última linha com dados na coluna K
    
    For i = 3 To ultimaLinha 'Começa da linha 2 (caso linha 1 seja cabeçalho)
        
        '--- CASO 1: Coluna K <= 1 ---
        If ws.Cells(i, "K").Value <= 1 Then
            
            'Se P <= 1 ? próxima linha
            If ws.Cells(i, "P").Value <= 1 Then
                GoTo ProximaLinha
                
            'Se P > 1 ? compara P com G
            Else
                If ws.Cells(i, "P").Value = ws.Cells(i, "G").Value Then
                    GoTo ProximaLinha
                Else
                    'Pinta de azul escuro e letra branca da coluna B até P
                    With ws.Range("B" & i & ":P" & i)
                        .Interior.Color = RGB(0, 0, 139) 'Azul escuro
                        .Font.Color = RGB(255, 255, 255) 'Branco
                    End With
                End If
            End If
        
        '--- CASO 2: Coluna K > 1 ---
        Else
            'Compara K com G
            If ws.Cells(i, "K").Value = ws.Cells(i, "G").Value Then
                GoTo ProximaLinha
            Else
                'Pinta de vermelho e letra branca da coluna B até P
                With ws.Range("B" & i & ":P" & i)
                    .Interior.Color = RGB(255, 0, 0) 'Vermelho
                    .Font.Color = RGB(255, 255, 255) 'Branco
                End With
            End If
        End If
        
ProximaLinha:
    Next i

    MsgBox "Verificação concluída!", vbInformation

End Sub