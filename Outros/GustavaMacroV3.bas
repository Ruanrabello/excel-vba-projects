Attribute VB_Name = "GustavaMacroV3"
Option Explicit

Sub compararNOTAFISCAL()
    Dim ws As Worksheet
    Dim ultimaCol As Long
    Dim colNFCE As Long, colNOTA As Long
    Dim linha As Long
    Dim valorRaw As String
    Dim Barraposicao As Long
    Dim partBeforeSlash As String
    Dim tmp As String
    Dim tmpDigits As String
    Dim listaNFCE As Collection, listaNOTA As Collection
    Dim nfceDict As Object, notaDict As Object
    Dim nfceOrigDict As Object, notaOrigDict As Object
    Dim startOrigEndNFCE As Long, startOrigEndNOTA As Long
    Dim lastUsedRowNFCE As Long, lastUsedRowNOTA As Long
    Dim writeRowNFCE As Long, writeRowNOTA As Long
    Dim item As Variant
    Dim writtenDict As Object
    
    Set ws = ThisWorkbook.Sheets("Planilha1")
    Set nfceDict = CreateObject("Scripting.Dictionary")      ' Dicionario NFCE JA TRATADO (keys normalizados (ex: "2414"))
    Set notaDict = CreateObject("Scripting.Dictionary")      ' Dicionario NOTA JA TRATADO (keys normalizados (ex: "2414"))
    Set nfceOrigDict = CreateObject("Scripting.Dictionary")  ' Dicionario NFCE (key -> original full NFCE string)
    Set notaOrigDict = CreateObject("Scripting.Dictionary")  ' Dicionario NOTA(key -> original nota string with zeros)
    Set writtenDict = CreateObject("Scripting.Dictionary")   ' para evitar duplicatas ao escrever
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' localizar última coluna de cabeçalho (linha 2)
    ultimaCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    ' localizar colunas NFCE e NOTA_FISCAL
    colNFCE = 0: colNOTA = 0
    For linha = 1 To ultimaCol
        If UCase(Trim(ws.Cells(2, linha).value)) = "NFCE" Then colNFCE = linha
        If UCase(Trim(ws.Cells(2, linha).value)) = "NOTA_FISCAL" Then colNOTA = linha
    Next linha
    If colNFCE = 0 Or colNOTA = 0 Then
        MsgBox "Coluna NFCE ou NOTA_FISCAL não encontrada!", vbExclamation
        GoTo Finaliza
    End If
    
    'Encontrar a ultima linha preenchida do bloco orginal(com a fiunção que criamos)
    startOrigEndNFCE = FindOriginalBlockEnd(ws, colNFCE, 3)
    startOrigEndNOTA = FindOriginalBlockEnd(ws, colNOTA, 3)
    If startOrigEndNFCE < 3 Then startOrigEndNFCE = 2
    If startOrigEndNOTA < 3 Then startOrigEndNOTA = 2
    
    ' Apaga execuções anteriores(para caso vc execute o codigo 2 vezes nao ficar acumulando na tabela e dar conflito)
    lastUsedRowNFCE = ws.Cells(ws.Rows.Count, colNFCE).End(xlUp).Row    'Procura a ultima linha usada  na coluna NFCE
    If lastUsedRowNFCE > startOrigEndNFCE Then  'Se essa linha for maior q ultima linha orginal
        ws.Range(ws.Cells(startOrigEndNFCE + 1, colNFCE), ws.Cells(lastUsedRowNFCE, colNFCE)).Clear  'Ele limpa
    End If
    lastUsedRowNOTA = ws.Cells(ws.Rows.Count, colNOTA).End(xlUp).Row 'Mesmo processo porem com nota
    If lastUsedRowNOTA > startOrigEndNOTA Then
        ws.Range(ws.Cells(startOrigEndNOTA + 1, colNOTA), ws.Cells(lastUsedRowNOTA, colNOTA)).Clear
    End If
    
    
    ' prepara listas
    Set listaNFCE = New Collection
    Set listaNOTA = New Collection
    
   
    ' Mas guardar a string ORIGINAL completa para escrita posterior ---
    ' Tratar os dados de NFCE retirando a barra
    For linha = 3 To startOrigEndNFCE   'Loop de 3 ate a ultima linha preenchida da coluna original
        valorRaw = Trim(CStr(ws.Cells(linha, colNFCE).value))
            
        If valorRaw <> "" Then  ' Se o valor for diferente de nulp
            Barraposicao = InStr(valorRaw, "/") 'Procura a posição da / dentro da variavel
            
            If Barraposicao > 0 Then  ' Se for maior que 0(ou seja se tiver a /)
                partBeforeSlash = Left(valorRaw, Barraposicao - 1)  ' Ele pega o valor(da variavel) a esquerda da barra(exemplo:(4567/00 = 4567)por isso que ta -1
            Else
                partBeforeSlash = valorRaw  'ou seja se nao tiver barra ele igual
            End If
            
            If Len(partBeforeSlash) > 4 Then    'Se a quantidade de numeros(dos ja tratados sem a barra) for maior q 4
                tmp = Right(partBeforeSlash, 4) 'Ele pega só os 4 digitos a direita
            Else
                tmp = partBeforeSlash   'Se nao(else): só iguala o valor de uma varivael a outra variavel(tmp)
            End If
            tmp = OnlyDigitsString(tmp) 'Função do excel que ele pega apenas os digitos de uma string(exemplo: 24A1" --> "241")

            If tmp <> "" Then   'Se a variavel for diferente de zero(isso e caso nao tenha numero nela e fique(nada(pq a função pega só digito))
                ' chave normalizada (por exemplo "2414" ou "416")
                If Not nfceDict.Exists(tmp) Then    'Se esse valor de temp nao existe ainda no Dicionario
                    nfceDict.Add tmp, True  'Ele adiciona
                    nfceOrigDict.Add tmp, valorRaw ' Guarda o valor normal sem ser tratado(usando como chave o valor tratado)
                    listaNFCE.Add tmp
                End If
            End If
        End If
    Next linha
    
    
    ' Tratar os dados de NOTA Removendo os zeros
    For linha = 3 To startOrigEndNOTA   'Loop de 3 até a ultima linha preenchida do bloco de codigo original
        valorRaw = Trim(CStr(ws.Cells(linha, colNOTA).value))   'Armazena o valor na celula(forçando a virar uma string e tirando os espaços com trim)
        
        If valorRaw <> "" Then  'Se o valor for diferente de nada
            ' valorRaw vem sem apóstrofo de exibição (apóstrofo é apenas formatador no Excel),
            ' por exemplo "0002414" estará em valorRaw.
            tmpDigits = OnlyDigitsString(valorRaw)  'Pega só os digitos  e guarda na variavel
            
            If tmpDigits <> "" Then 'Se for diferente de nada
                ' chave normalizada (por exemplo "2414" ou "416")
                Dim keyNota As String
                keyNota = CStr(Val(tmpDigits))  'função Val() converte uma string numérica em número real (tipo Double). --> e ela acaba ignorando os zeros
                                                'E depois converte de volta para string, exemplo(Val("0002414") -> 2414)
                
                If keyNota = "" Then keyNota = tmpDigits 'Se a variavel for diferente de zero(isso e caso nao tenha numero nela e fique(nada(pq a função pega só digito))
                If Not notaDict.Exists(keyNota) Then 'caso o valor da variavel nao exista no dicionario
                    notaDict.Add keyNota, True  'Ela adiciona, em formato de chave
                    notaOrigDict.Add keyNota, tmpDigits ' Armazena tambem em outro dicionario o valor bruto, usando o valor trtado como chave
                    listaNOTA.Add keyNota
                End If
            End If
        End If
    Next linha
    
    
    ' --- Agora comparar e escrever de volta NA MESMA planilha,
    '     a partir da primeira linha vazia logo após o bloco original ---
    writeRowNFCE = startOrigEndNFCE + 1
    writeRowNOTA = startOrigEndNOTA + 1
    
    ' escrever itens em NFCE que não existem em NOTA -> escrever a STRING ORIGINAL completa
    For Each item In listaNFCE
        If Not notaDict.Exists(CStr(item)) Then
            If Not writtenDict.Exists("NFCE|" & CStr(item)) Then
                ws.Cells(writeRowNFCE, colNFCE).value = nfceOrigDict(CStr(item)) ' escrever string completa
                ' formatar: amarelo + negrito
                With ws.Cells(writeRowNFCE, colNFCE)
                    .Interior.Color = RGB(255, 255, 0)
                    .Font.Bold = True
                End With
                writtenDict.Add "NFCE|" & CStr(item), True
                writeRowNFCE = writeRowNFCE + 1
            End If
        End If
    Next item
    ' pular uma linha em branco se algo foi escrito
    If writeRowNFCE > startOrigEndNFCE + 1 Then writeRowNFCE = writeRowNFCE + 1
    
    ' escrever itens em NOTA que não existem em NFCE -> escrever com apóstrofo para preservar zeros
    For Each item In listaNOTA
        If Not nfceDict.Exists(CStr(item)) Then
            If Not writtenDict.Exists("NOTA|" & CStr(item)) Then
                ' escrever com apóstrofo à frente para forçar exibição com zeros (ex.: '0002414)
                ws.Cells(writeRowNOTA, colNOTA).value = "'" & notaOrigDict(CStr(item))
                With ws.Cells(writeRowNOTA, colNOTA)
                    .Interior.Color = RGB(255, 255, 0)
                    .Font.Bold = True
                End With
                writtenDict.Add "NOTA|" & CStr(item), True
                writeRowNOTA = writeRowNOTA + 1
            End If
        End If
    Next item
    If writeRowNOTA > startOrigEndNOTA + 1 Then writeRowNOTA = writeRowNOTA + 1
    
    MsgBox "Processo finalizado. Faltantes escritos em Planilha1 (formatados)."
    
Finaliza:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' --- função: encontra o fim do bloco original assumindo dados contíguos a partir de startRow até a primeira linha vazia ---
Private Function FindOriginalBlockEnd(ws As Worksheet, col As Long, startRow As Long) As Long
    Dim r As Long
    For r = startRow To ws.Rows.Count
        If Trim(CStr(ws.Cells(r, col).value)) = "" Then
            FindOriginalBlockEnd = r - 1
            Exit Function
        End If
    Next r
    FindOriginalBlockEnd = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' --- função auxiliar: deixa só dígitos (retorna string vazia se nada) ---
Private Function OnlyDigitsString(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    s = CStr(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    OnlyDigitsString = out
End Function

