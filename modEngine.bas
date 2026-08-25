Attribute VB_Name = "modEngine"
'================================================================================
' MÓDULO: modEngine
' DESCRIÇÃO: Motor de performance em arrays para leitura/escrita otimizada de OPs
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterDadosOPsEmArray
' PROPÓSITO: Ler toda a TabelaOPs para memória RAM (Variant Array) otimizando
'            a velocidade de preenchimento de ListBoxes e demais controles
' RETORNO: Variant Array (1-based) contendo todos os dados da tabela
'--------------------------------------------------------------------------------
Public Function ObterDadosOPsEmArray() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterDados
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    dados = tbl.Range.Value
    
    ObterDadosOPsEmArray = dados
    
Sair:
    Exit Function
    
ErroObterDados:
    ObterDadosOPsEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: SalvarNovaOP
' PROPÓSITO: Inserir uma nova Ordem de Produção na última linha livre da TabelaOPs
' PARÂMETROS: Dados da OP a ser persistida
'--------------------------------------------------------------------------------
Public Sub SalvarNovaOP(pID As String, _
                        pProduto As String, _
                        pEquipamento As String, _
                        pQtd As Long, _
                        pInicio As Date, _
                        pFim As Date, _
                        pDuracao As Double, _
                        pStatus As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroSalvar
    
    ' Validação de campos obrigatórios
    If Trim(pID) = "" Or Trim(pProduto) = "" Or Trim(pEquipamento) = "" Or pQtd <= 0 Then
        Err.Raise vbObjectError + 100, "SalvarNovaOP", _
            "Campos obrigatórios não foram preenchidos corretamente."
    End If
    
    If pFim < pInicio Then
        Err.Raise vbObjectError + 101, "SalvarNovaOP", _
            "A Data de Fim não pode ser anterior à Data de Início."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    Application.ScreenUpdating = False
    
    ' Localiza a última linha livre e insere nova linha
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_OP").Index).Value = pID
        .Cells(1, tbl.ListColumns("Produto").Index).Value = pProduto
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Quantidade").Index).Value = pQtd
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pFim
        .Cells(1, tbl.ListColumns("Duracao_Horas").Index).Value = pDuracao
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar nova OP: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: AtualizarOP
' PROPÓSITO: Atualizar uma Ordem de Produção existente na TabelaOPs
' PARÂMETROS: Dados da OP a ser atualizada
'--------------------------------------------------------------------------------
Public Sub AtualizarOP(pID As String, _
                       pProduto As String, _
                       pEquipamento As String, _
                       pQtd As Long, _
                       pInicio As Date, _
                       pFim As Date, _
                       pDuracao As Double, _
                       pStatus As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim totalLinhas As Long
    Dim dados As Variant
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroAtualizar
    
    If Trim(pID) = "" Then
        Err.Raise vbObjectError + 102, "AtualizarOP", "ID da OP não informado."
    End If
    
    If pFim < pInicio Then
        Err.Raise vbObjectError + 103, "AtualizarOP", "A Data de Fim não pode ser anterior à Data de Início."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 104, "AtualizarOP", "OP não encontrada para atualização."
    End If
    
    With tbl.ListRows(linhaEncontrada - 1).Range
        .Cells(1, tbl.ListColumns("ID_OP").Index).Value = pID
        .Cells(1, tbl.ListColumns("Produto").Index).Value = pProduto
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Quantidade").Index).Value = pQtd
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pFim
        .Cells(1, tbl.ListColumns("Duracao_Horas").Index).Value = pDuracao
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroAtualizar:
    MsgBox "Erro ao atualizar OP: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExcluirOP
' PROPÓSITO: Excluir uma Ordem de Produção existente da TabelaOPs
' PARÂMETROS: pID As String – ID da OP a ser excluída
'--------------------------------------------------------------------------------
Public Sub ExcluirOP(pID As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim totalLinhas As Long
    Dim dados As Variant
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroExcluir
    
    If Trim(pID) = "" Then
        Err.Raise vbObjectError + 105, "ExcluirOP", "ID da OP não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 106, "ExcluirOP", "OP não encontrada para exclusão."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Delete
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroExcluir:
    MsgBox "Erro ao excluir OP: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO GENÉRICA: ObterDadosTabelaEmArray
' PROPÓSITO: Ler qualquer tabela do banco de dados para memória
' PARÂMETROS: pNomePlanilha As String, pNomeTabela As String
' RETORNO: Variant Array (1-based) ou CVErr(xlErrRef) em caso de erro
'--------------------------------------------------------------------------------
Public Function ObterDadosTabelaEmArray(pNomePlanilha As String, pNomeTabela As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterDadosTabela
    
    Set ws = ThisWorkbook.Worksheets(pNomePlanilha)
    Set tbl = ws.ListObjects(pNomeTabela)
    
    dados = tbl.Range.Value
    ObterDadosTabelaEmArray = dados
    
Sair:
    Exit Function
    
ErroObterDadosTabela:
    ObterDadosTabelaEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA GENÉRICA: InserirLinhaTabela
' PROPÓSITO: Inserir nova linha em qualquer tabela do banco de dados
' PARÂMETROS: pNomePlanilha, pNomeTabela, pValores (array de valores)
'--------------------------------------------------------------------------------
Public Sub InserirLinhaTabela(pNomePlanilha As String, pNomeTabela As String, pValores As Variant)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    Dim i As Long
    
    On Error GoTo ErroInserir
    
    Set ws = ThisWorkbook.Worksheets(pNomePlanilha)
    Set tbl = ws.ListObjects(pNomeTabela)
    
    Application.ScreenUpdating = False
    
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        For i = LBound(pValores) To UBound(pValores)
            .Cells(1, i + 1).Value = pValores(i)
        Next i
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroInserir:
    MsgBox "Erro ao inserir registro: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA GENÉRICA: AtualizarLinhaTabela
' PROPÓSITO: Atualizar linha existente em qualquer tabela do banco de dados
' PARÂMETROS: pNomePlanilha, pNomeTabela, pColunaChave, pValorChave, pValores
'--------------------------------------------------------------------------------
Public Sub AtualizarLinhaTabela(pNomePlanilha As String, pNomeTabela As String, _
                                pColunaChave As String, pValorChave As String, pValores As Variant)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    Dim colunaChaveIndex As Long
    
    On Error GoTo ErroAtualizarTabela
    
    Set ws = ThisWorkbook.Worksheets(pNomePlanilha)
    Set tbl = ws.ListObjects(pNomeTabela)
    
    Application.ScreenUpdating = False
    
    colunaChaveIndex = tbl.ListColumns(pColunaChave).Index
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, colunaChaveIndex))) = Trim(pValorChave) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 200, "AtualizarLinhaTabela", "Registro não encontrado."
    End If
    
    With tbl.ListRows(linhaEncontrada - 1).Range
        For i = LBound(pValores) To UBound(pValores)
            .Cells(1, i + 1).Value = pValores(i)
        Next i
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroAtualizarTabela:
    MsgBox "Erro ao atualizar registro: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA GENÉRICA: ExcluirLinhaTabela
' PROPÓSITO: Excluir linha de qualquer tabela do banco de dados
' PARÂMETROS: pNomePlanilha, pNomeTabela, pColunaChave, pValorChave
'--------------------------------------------------------------------------------
Public Sub ExcluirLinhaTabela(pNomePlanilha As String, pNomeTabela As String, _
                              pColunaChave As String, pValorChave As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    Dim colunaChaveIndex As Long
    
    On Error GoTo ErroExcluirTabela
    
    Set ws = ThisWorkbook.Worksheets(pNomePlanilha)
    Set tbl = ws.ListObjects(pNomeTabela)
    
    Application.ScreenUpdating = False
    
    colunaChaveIndex = tbl.ListColumns(pColunaChave).Index
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, colunaChaveIndex))) = Trim(pValorChave) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 201, "ExcluirLinhaTabela", "Registro não encontrado."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Delete
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroExcluirTabela:
    MsgBox "Erro ao excluir registro: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub
