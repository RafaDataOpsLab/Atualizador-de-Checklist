Attribute VB_Name = "Módulo2"
Sub GerarPDFporTST_Fornecedor()
    Dim wsPlan As Worksheet, wsDados As Worksheet, wsFornec As Worksheet
    Dim dictTST As Object, dictFornec As Object
    Dim i As Long, ultimaLinha As Long
    Dim TST As String, fornecedor As String, chaveTST As Variant, chaveFornec As Variant
    Dim lojas As String, siglas As String, cidades As String, estados As String, enderecos As String
    Dim tecnicos As String, funcoes As String, atividades As String
    Dim idx As Long, j As Long
    Dim caminhoModelo As String, caminhoPDF As String, pastaModelo As String, novoModelo As String
    Dim wdApp As Object, wdDoc As Object

    Set wsPlan = ThisWorkbook.Sheets("bd")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsFornec = ThisWorkbook.Sheets("Planilha3")
    Set dictTST = CreateObject("Scripting.Dictionary")
    ultimaLinha = wsPlan.Cells(wsPlan.Rows.Count, 7).End(xlUp).Row

    ' 1. Agrupar por TST
    For i = 2 To ultimaLinha
        TST = Trim(wsPlan.Cells(i, 7).Value)   ' G = TST Regional
        If TST <> "" Then
            If Not dictTST.Exists(TST) Then dictTST.Add TST, New Collection
            dictTST(TST).Add i
        End If
    Next i

    If dictTST.Count = 0 Then
        MsgBox "Nenhum TST encontrado na coluna G!", vbExclamation
        Exit Sub
    End If

    caminhoModelo = Application.GetOpenFilename("Arquivos do Word (*.docx), *.docx", , "Selecione o modelo Word")
    If caminhoModelo = "Falso" Then Exit Sub
    pastaModelo = Left(caminhoModelo, InStrRev(caminhoModelo, "\"))

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    For Each chaveTST In dictTST.Keys
        ' 2. Agrupar por fornecedor DENTRO desse TST
        Set dictFornec = CreateObject("Scripting.Dictionary")
        For idx = 1 To dictTST(chaveTST).Count
            j = dictTST(chaveTST).Item(idx)
            fornecedor = Trim(wsPlan.Cells(j, 8).Value) ' H = Fornecedor
            If fornecedor <> "" Then
                If Not dictFornec.Exists(fornecedor) Then dictFornec.Add fornecedor, New Collection
                dictFornec(fornecedor).Add j
            End If
        Next idx
        
        ' 3. Para cada fornecedor dentro do TST
        For Each chaveFornec In dictFornec.Keys
            ' Limpar variáveis
            lojas = "": siglas = "": cidades = "": estados = "": enderecos = ""
            tecnicos = "": funcoes = "": atividades = ""  ' <--- DECLARA COMO STRING

            ' a) Montar os campos de loja desse fornecedor naquele TST
            For idx = 1 To dictFornec(chaveFornec).Count
                j = dictFornec(chaveFornec).Item(idx)
                If wsPlan.Cells(j, 2).Value <> "" Then
                    siglas = siglas & wsPlan.Cells(j, 2).Value & vbCrLf
                    lojas = lojas & wsPlan.Cells(j, 3).Value & vbCrLf
                    cidades = cidades & wsPlan.Cells(j, 4).Value & vbCrLf
                    estados = estados & wsPlan.Cells(j, 5).Value & vbCrLf
                    enderecos = enderecos & wsPlan.Cells(j, 6).Value & vbCrLf
                End If
            Next idx
            siglas = RemoveUltimaLinha(siglas)
            lojas = RemoveUltimaLinha(lojas)
            cidades = RemoveUltimaLinha(cidades)
            estados = RemoveUltimaLinha(estados)
            enderecos = RemoveUltimaLinha(enderecos)

            ' b) Buscar técnicos do fornecedor na Planilha3 (Nome, Função, Atividade)
            Call BuscarTecnicos(wsFornec, CStr(chaveFornec), tecnicos, funcoes, atividades)

            ' c) Preencher a aba Dados com os campos normais + campos dos técnicos
            Call PreencherValor(wsDados, "<<Lojas>>", lojas)
            Call PreencherValor(wsDados, "<<Sigla>>", siglas)
            Call PreencherValor(wsDados, "<<Cidade>>", cidades)
            Call PreencherValor(wsDados, "<<Estado>>", estados)
            Call PreencherValor(wsDados, "<<Endereços>>", enderecos)
            Call PreencherValor(wsDados, "<<técnicos>>", tecnicos)
            Call PreencherValor(wsDados, "<<Função>>", funcoes)
            Call PreencherValor(wsDados, "<<Atividade>>", atividades)
            
            ' d) Cria e preenche o Word
            novoModelo = pastaModelo & "TEMP_" & LimparNomeArquivo(CStr(chaveTST) & "_" & CStr(chaveFornec)) & ".docx"
            FileCopy caminhoModelo, novoModelo
            Set wdDoc = wdApp.Documents.Open(novoModelo)
            Call PreencherChavesWordLong(wdDoc, wsDados)
            caminhoPDF = pastaModelo & LimparNomeArquivo(CStr(chaveTST) & "_" & CStr(chaveFornec)) & ".pdf"
            wdDoc.ExportAsFixedFormat OutputFileName:=caminhoPDF, ExportFormat:=17
            wdDoc.Close False
            Kill novoModelo
        Next chaveFornec
    Next chaveTST

    wdApp.Quit
    Set wdApp = Nothing
    MsgBox "Checklists gerados para cada TST e Fornecedor!"
End Sub

' Remove última quebra de linha
Function RemoveUltimaLinha(str As String) As String
    If Right(str, 2) = vbCrLf Then
        RemoveUltimaLinha = Left(str, Len(str) - 2)
    Else
        RemoveUltimaLinha = str
    End If
End Function

' Busca técnicos do fornecedor na Planilha3 (colunas certas!)
Sub BuscarTecnicos(wsFornec As Worksheet, fornecedor As String, ByRef tecnicos As String, ByRef funcoes As String, ByRef atividades As String)
    Dim ultimaLinhaFornec As Long, i As Long
    tecnicos = "": funcoes = "": atividades = ""
    ultimaLinhaFornec = wsFornec.Cells(wsFornec.Rows.Count, 5).End(xlUp).Row ' E = Fornecedor
    For i = 2 To ultimaLinhaFornec
        If Trim(wsFornec.Cells(i, 5).Value) = fornecedor Then
            tecnicos = tecnicos & wsFornec.Cells(i, 1).Value & vbCrLf      ' A = Nome
            funcoes = funcoes & wsFornec.Cells(i, 2).Value & vbCrLf        ' B = Função
            atividades = atividades & wsFornec.Cells(i, 3).Value & vbCrLf   ' C = Atividade
        End If
    Next i
    tecnicos = RemoveUltimaLinha(tecnicos)
    funcoes = RemoveUltimaLinha(funcoes)
    atividades = RemoveUltimaLinha(atividades)
End Sub

' Preenche valor na planilha Dados conforme chave
Sub PreencherValor(ws As Worksheet, chave As String, valor As String)
    Dim rng As Range
    Set rng = ws.Range("A:A").Find(chave, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then rng.Offset(0, 1).Value = valor
End Sub

' Preenche chaves no Word (mantendo lógica para campos longos)
Sub PreencherChavesWordLong(wdDoc As Object, wsDados As Worksheet)
    Dim i As Long, chave As String, valor As String
    i = 2
    Do While wsDados.Cells(i, 1).Value <> ""
        chave = Trim(wsDados.Cells(i, 1).Value)
        valor = Trim(wsDados.Cells(i, 2).Value)
        If Len(valor) > 250 Then
            Call SubstituirChavePorTextoLongo(wdDoc, chave, valor)
        Else
            With wdDoc.Content.Find
                .Text = chave
                .Replacement.Text = valor
                .Wrap = 1
                .Execute Replace:=2
            End With
        End If
        i = i + 1
    Loop
End Sub

Sub SubstituirChavePorTextoLongo(wdDoc As Object, chave As String, textoSubstituir As String)
    Dim rngBusca As Object
    Set rngBusca = wdDoc.Content
    With rngBusca.Find
        .Text = chave
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        Do While .Execute
            rngBusca.Text = textoSubstituir
            rngBusca.Collapse 0
        Loop
    End With
End Sub

Function LimparNomeArquivo(str As String) As String
    Dim invalidChars As Variant, ch As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    LimparNomeArquivo = str
    For Each ch In invalidChars
        LimparNomeArquivo = Replace(LimparNomeArquivo, ch, "_")
    Next ch
End Function

