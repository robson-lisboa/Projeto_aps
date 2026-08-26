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

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterDadosSimulacaoEmArray
' PROPÓSITO: Ler todas as simulações para memória
' RETORNO: Variant Array (1-based)
'--------------------------------------------------------------------------------
Public Function ObterDadosSimulacaoEmArray() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterSimulacoes
    
    Set ws = ThisWorkbook.Worksheets("BD_Simulacoes")
    Set tbl = ws.ListObjects("TabelaSimulacoes")
    
    dados = tbl.Range.Value
    ObterDadosSimulacaoEmArray = dados
    
Sair:
    Exit Function
    
ErroObterSimulacoes:
    ObterDadosSimulacaoEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterOPsSimulacaoEmArray
' PROPÓSITO: Ler todas as OPs de uma simulação para memória
' PARÂMETROS: pID_Simulacao As String
' RETORNO: Variant Array (1-based)
'--------------------------------------------------------------------------------
Public Function ObterOPsSimulacaoEmArray(pID_Simulacao As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim resultados() As Variant
    Dim contResultado As Long
    
    On Error GoTo ErroObterOPsSimulacao
    
    Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tbl = ws.ListObjects("TabelaOPsSimulacao")
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    
    ReDim resultados(1 To totalLinhas, 1 To 12)
    contResultado = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) Then
            contResultado = contResultado + 1
            Dim j As Long
            For j = 1 To 12
                resultados(contResultado, j) = dados(i, j)
            Next j
        End If
    Next i
    
    If contResultado = 0 Then
        ObterOPsSimulacaoEmArray = CVErr(xlErrRef)
    Else
        ReDim Preserve resultados(1 To contResultado, 1 To 12)
        ObterOPsSimulacaoEmArray = resultados
    End If
    
Sair:
    Exit Function
    
ErroObterOPsSimulacao:
    ObterOPsSimulacaoEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: SalvarSimulacao
' PROPÓSITO: Inserir nova simulação
' PARÂMETROS: Dados da simulação
'--------------------------------------------------------------------------------
Public Sub SalvarSimulacao(pID As String, pNome As String, pDataCriacao As Date, _
                           pDataInicio As Date, pDataFim As Date, pStatus As String, pObs As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroSalvarSimulacao
    
    Set ws = ThisWorkbook.Worksheets("BD_Simulacoes")
    Set tbl = ws.ListObjects("TabelaSimulacoes")
    
    Application.ScreenUpdating = False
    
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_Simulacao").Index).Value = pID
        .Cells(1, tbl.ListColumns("Nome_Simulacao").Index).Value = pNome
        .Cells(1, tbl.ListColumns("Data_Criacao").Index).Value = pDataCriacao
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pDataInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pDataFim
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
        .Cells(1, tbl.ListColumns("Observacao").Index).Value = pObs
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroSalvarSimulacao:
    MsgBox "Erro ao salvar simulação: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: SalvarOPSimulacao
' PROPÓSITO: Inserir nova OP na simulação
'--------------------------------------------------------------------------------
Public Sub SalvarOPSimulacao(pID_Simulacao As String, pID_OP_Simulacao As String, pID_OP_Origem As String, _
                             pProduto As String, pEquipamento As String, pQtd As Long, _
                             pInicio As Date, pFim As Date, pDuracao As Double, _
                             pStatus As String, pPrioridade As String, pObs As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroSalvarOPSimulacao
    
    Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tbl = ws.ListObjects("TabelaOPsSimulacao")
    
    Application.ScreenUpdating = False
    
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_Simulacao").Index).Value = pID_Simulacao
        .Cells(1, tbl.ListColumns("ID_OP_Simulacao").Index).Value = pID_OP_Simulacao
        .Cells(1, tbl.ListColumns("ID_OP_Origem").Index).Value = pID_OP_Origem
        .Cells(1, tbl.ListColumns("Produto").Index).Value = pProduto
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Quantidade").Index).Value = pQtd
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pFim
        .Cells(1, tbl.ListColumns("Duracao_Horas").Index).Value = pDuracao
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
        .Cells(1, tbl.ListColumns("Prioridade").Index).Value = pPrioridade
        .Cells(1, tbl.ListColumns("Observacao").Index).Value = pObs
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroSalvarOPSimulacao:
    MsgBox "Erro ao salvar OP da simulação: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: AtualizarOPSimulacao
' PROPÓSITO: Atualizar OP existente na simulação
'--------------------------------------------------------------------------------
Public Sub AtualizarOPSimulacao(pID_Simulacao As String, pID_OP_Simulacao As String, _
                                pProduto As String, pEquipamento As String, pQtd As Long, _
                                pInicio As Date, pFim As Date, pDuracao As Double, _
                                pStatus As String, pPrioridade As String, pObs As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroAtualizarOPSimulacao
    
    Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tbl = ws.ListObjects("TabelaOPsSimulacao")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) And _
           Trim(CStr(dados(i, 2))) = Trim(pID_OP_Simulacao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 300, "AtualizarOPSimulacao", "OP da simulação não encontrada."
    End If
    
    With tbl.ListRows(linhaEncontrada - 1).Range
        .Cells(1, tbl.ListColumns("Produto").Index).Value = pProduto
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Quantidade").Index).Value = pQtd
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pFim
        .Cells(1, tbl.ListColumns("Duracao_Horas").Index).Value = pDuracao
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
        .Cells(1, tbl.ListColumns("Prioridade").Index).Value = pPrioridade
        .Cells(1, tbl.ListColumns("Observacao").Index).Value = pObs
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroAtualizarOPSimulacao:
    MsgBox "Erro ao atualizar OP da simulação: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExcluirOPSimulacao
' PROPÓSITO: Excluir OP da simulação
'--------------------------------------------------------------------------------
Public Sub ExcluirOPSimulacao(pID_Simulacao As String, pID_OP_Simulacao As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroExcluirOPSimulacao
    
    Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tbl = ws.ListObjects("TabelaOPsSimulacao")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) And _
           Trim(CStr(dados(i, 2))) = Trim(pID_OP_Simulacao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 301, "ExcluirOPSimulacao", "OP da simulação não encontrada."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Delete
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroExcluirOPSimulacao:
    MsgBox "Erro ao excluir OP da simulação: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExcluirSimulacao
' PROPÓSITO: Excluir simulação e todas as suas OPs
'--------------------------------------------------------------------------------
Public Sub ExcluirSimulacao(pID_Simulacao As String)
    Dim wsSim As Worksheet
    Dim wsOPs As Worksheet
    Dim tblSim As ListObject
    Dim tblOPs As ListObject
    Dim dadosSim As Variant
    Dim dadosOPs As Variant
    Dim i As Long, totalLinhasSim As Long
    Dim totalLinhasOPs As Long
    
    On Error GoTo ErroExcluirSimulacao
    
    Set wsSim = ThisWorkbook.Worksheets("BD_Simulacoes")
    Set wsOPs = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tblSim = wsSim.ListObjects("TabelaSimulacoes")
    Set tblOPs = wsOPs.ListObjects("TabelaOPsSimulacao")
    
    Application.ScreenUpdating = False
    
    dadosSim = tblSim.Range.Value
    totalLinhasSim = UBound(dadosSim, 1)
    
    ' Exclui a simulação (da última para a primeira para não desalinhar índices)
    For i = totalLinhasSim To 2 Step -1
        If Trim(CStr(dadosSim(i, 1))) = Trim(pID_Simulacao) Then
            tblSim.ListRows(i - 1).Delete
            Exit For
        End If
    Next i
    
    ' Exclui todas as OPs da simulação
    dadosOPs = tblOPs.Range.Value
    totalLinhasOPs = UBound(dadosOPs, 1)
    
    For i = totalLinhasOPs To 2 Step -1
        If Trim(CStr(dadosOPs(i, 1))) = Trim(pID_Simulacao) Then
            tblOPs.ListRows(i - 1).Delete
        End If
    Next i
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroExcluirSimulacao:
    MsgBox "Erro ao excluir simulação: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ImportarOPsParaSimulacao
' PROPÓSITO: Copiar OPs do planejamento para uma simulação
' PARÂMETROS: pID_Simulacao, pDataInicio (filtro opcional), pDataFim (filtro opcional)
'--------------------------------------------------------------------------------
Public Sub ImportarOPsParaSimulacao(pID_Simulacao As String, _
                                    Optional pDataInicio As Date = 0, _
                                    Optional pDataFim As Date = 0)
    Dim wsOPs As Worksheet
    Dim tblOPs As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    
    On Error GoTo ErroImportar
    
    Set wsOPs = ThisWorkbook.Worksheets("BD_OPs")
    Set tblOPs = wsOPs.ListObjects("TabelaOPs")
    
    Application.ScreenUpdating = False
    
    dados = tblOPs.Range.Value
    totalLinhas = UBound(dados, 1)
    
    Dim idSimulacao As String
    Dim idOPSimulacao As String
    Dim contador As Long
    contador = 0
    
    For i = 2 To totalLinhas
        Dim dataInicioOP As Date
        Dim dataFimOP As Date
        
        If IsDate(dados(i, 5)) Then dataInicioOP = CDate(dados(i, 5))
        If IsDate(dados(i, 6)) Then dataFimOP = CDate(dados(i, 6))
        
        ' Aplica filtros de data se fornecidos
        If pDataInicio <> 0 And dataInicioOP < pDataInicio Then GoTo ProximaOP
        If pDataFim <> 0 And dataFimOP > pDataFim Then GoTo ProximaOP
        
        contador = contador + 1
        idOPSimulacao = pID_Simulacao & "_OP" & Format(contador, "000")
        
        Call SalvarOPSimulacao( _
            pID_Simulacao, _
            idOPSimulacao, _
            CStr(dados(i, 1)), _
            CStr(dados(i, 2)), _
            CStr(dados(i, 3)), _
            CLng(dados(i, 4)), _
            dataInicioOP, _
            dataFimOP, _
            CDbl(dados(i, 7)), _
            CStr(dados(i, 8)), _
            "Média", _
            "Importado do planejamento")
ProximaOP:
    Next i
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroImportar:
    MsgBox "Erro ao importar OPs para simulação: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: CalcularCapacidadePorPosto
' PROPÓSITO: Calcular horas planejadas, disponíveis e ocupação por posto
' PARÂMETROS: pID_Simulacao As String
' RETORNO: Dictionary com capacidade por posto (chave = nome posto)
'--------------------------------------------------------------------------------
Public Function CalcularCapacidadePorPosto(pID_Simulacao As String) As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim capacidade As Object
    Dim nomePosto As String
    Dim horasPlanejadas As Double
    
    On Error GoTo ErroCapacidade
    
    Set capacidade = CreateObject("Scripting.Dictionary")
    
    ' Inicializa postos com capacidade padrão de 8h/dia (assumindo 22 dias úteis = 176h)
    ' Se houver tabela de equipamentos, usar capacidade real
    Dim capPadrao As Double
    capPadrao = 176 ' 8h * 22 dias
    
    ' Carrega OPs da simulação
    Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tbl = ws.ListObjects("TabelaOPsSimulacao")
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    
    ' Inicializa todos os postos com capacidade padrão
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) Then
            nomePosto = Trim(CStr(dados(i, 5)))
            If nomePosto <> "" Then
                If Not capacidade.Exists(nomePosto) Then
                    capacidade.Add nomePosto, Array(capPadrao, 0) ' (disponivel, planejado)
                End If
            End If
        End If
    Next i
    
    ' Calcula horas planejadas
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) Then
            nomePosto = Trim(CStr(dados(i, 5)))
            If nomePosto <> "" And capacidade.Exists(nomePosto) Then
                horasPlanejadas = capacidade(nomePosto)(1) + CDbl(dados(i, 9))
                capacidade(nomePosto) = Array(capacidade(nomePosto)(0), horasPlanejadas)
            End If
        End If
    Next i
    
    Set CalcularCapacidadePorPosto = capacidade
    
Sair:
    Exit Function
    
ErroCapacidade:
    Set CalcularCapacidadePorPosto = Nothing
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterIndicadoresSimulacao
' PROPÓSITO: Calcular indicadores gerais da simulação
' PARÂMETROS: pID_Simulacao As String
' RETORNO: Array com (totalOPs, horasPlanejadas, horasDisponiveis, ocupacao, conflitos, postosSobrecarga, atrasadas)
'--------------------------------------------------------------------------------
Public Function ObterIndicadoresSimulacao(pID_Simulacao As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim totalOPs As Long
    Dim horasPlanejadas As Double
    Dim horasDisponiveis As Double
    Dim ocupacao As Double
    Dim conflitos As Collection
    Dim postosSobrecarga As Long
    Dim atrasadas As Long
    
    On Error GoTo ErroIndicadores
    
    ' Carrega OPs da simulação
    Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
    Set tbl = ws.ListObjects("TabelaOPsSimulacao")
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    
    totalOPs = 0
    horasPlanejadas = 0
    atrasadas = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) Then
            totalOPs = totalOPs + 1
            horasPlanejadas = horasPlanejadas + CDbl(dados(i, 9))
            If Trim(CStr(dados(i, 10))) = "Atrasado" Then
                atrasadas = atrasadas + 1
            End If
        End If
    Next i
    
    ' Calcula capacidade
    Dim capacidade As Object
    Set capacidade = CalcularCapacidadePorPosto(pID_Simulacao)
    
    horasDisponiveis = 0
    postosSobrecarga = 0
    
    If Not capacidade Is Nothing Then
        Dim chave As Variant
        For Each chave In capacidade.Keys
            horasDisponiveis = horasDisponiveis + capacidade(chave)(0)
            If capacidade(chave)(1) > capacidade(chave)(0) Then
                postosSobrecarga = postosSobrecarga + 1
            End If
        Next chave
    End If
    
    If horasDisponiveis > 0 Then
        ocupacao = (horasPlanejadas / horasDisponiveis) * 100
    Else
        ocupacao = 0
    End If
    
    ' Detecta conflitos
    Dim colOPs As New Collection
    Dim op As clsCardProducao
    Dim j As Long
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Simulacao) Then
            Set op = New clsCardProducao
            With op
                .ID_OP = CStr(dados(i, 2))
                .Produto = CStr(dados(i, 4))
                .Equipamento = CStr(dados(i, 5))
                .Quantidade = CLng(dados(i, 6))
                If IsDate(dados(i, 7)) Then .DataInicio = CDate(dados(i, 7))
                If IsDate(dados(i, 8)) Then .DataFim = CDate(dados(i, 8))
                .Duracao = CDbl(dados(i, 9))
                .Status = CStr(dados(i, 10))
            End With
            colOPs.Add op
        End If
    Next i
    
    Dim colPostos As New Collection
    Dim nomePosto As String
    Dim existePosto As Boolean
    
    For i = 1 To colOPs.Count
        nomePosto = Trim(colOPs(i).Equipamento)
        If nomePosto <> "" Then
            existePosto = False
            For j = 1 To colPostos.Count
                If colPostos(j) = nomePosto Then
                    existePosto = True
                    Exit For
                End If
            Next j
            If Not existePosto Then
                colPostos.Add nomePosto
            End If
        End If
    Next i
    
    Set conflitos = DetectarConflitos(colOPs, colPostos)
    
    ObterIndicadoresSimulacao = Array(totalOPs, horasPlanejadas, horasDisponiveis, ocupacao, conflitos.Count, postosSobrecarga, atrasadas)
    
Sair:
    Exit Function
    
ErroIndicadores:
    ObterIndicadoresSimulacao = Array(0, 0, 0, 0, 0, 0, 0)
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: BuscarOPPorID
' PROPÓSITO: Buscar uma OP específica na TabelaOPs pelo ID
' RETORNO: Variant Array com os dados da OP (1=ID, 2=Produto, 3=Equipamento, 4=Quantidade, 5=Data_Inicio, 6=Data_Fim)
'          ou Empty se não encontrada
'--------------------------------------------------------------------------------
Public Function BuscarOPPorID(pID_OP As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    
    On Error GoTo ErroBuscar
    
    If Trim(pID_OP) = "" Then
        BuscarOPPorID = Empty
        Exit Function
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_OP) Then
            BuscarOPPorID = Array(dados(i, 1), dados(i, 2), dados(i, 3), dados(i, 4), dados(i, 5), dados(i, 6))
            Exit Function
        End If
    Next i
    
    BuscarOPPorID = Empty
    
Sair:
    Exit Function
    
ErroBuscar:
    BuscarOPPorID = Empty
    Resume Sair
End Function

'================================================================================
' MÓDULO DE PRODUÇÃO REAL
'================================================================================

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterProducoesEmArray
' PROPÓSITO: Ler todas as produções para memória
' RETORNO: Variant Array (1-based)
'--------------------------------------------------------------------------------
Public Function ObterProducoesEmArray() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterProducoes
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    dados = tbl.Range.Value
    ObterProducoesEmArray = dados
    
Sair:
    Exit Function
    
ErroObterProducoes:
    ObterProducoesEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterParadasEmArray
' PROPÓSITO: Ler todas as paradas para memória
' RETORNO: Variant Array (1-based)
'--------------------------------------------------------------------------------
Public Function ObterParadasEmArray() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterParadas
    
    Set ws = ThisWorkbook.Worksheets("BD_Paradas")
    Set tbl = ws.ListObjects("TabelaParadas")
    
    dados = tbl.Range.Value
    ObterParadasEmArray = dados
    
Sair:
    Exit Function
    
ErroObterParadas:
    ObterParadasEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: IniciarProducao
' PROPÓSITO: Iniciar uma nova produção
'--------------------------------------------------------------------------------
Public Sub IniciarProducao(pID_Producao As String, pID_OP As String, pID_Operador As String, _
                           pEquipamento As String, pQtdPlanejada As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroIniciar
    
    If Trim(pID_Producao) = "" Then
        Err.Raise vbObjectError + 400, "IniciarProducao", "ID da produção não informado."
    End If
    
    If Trim(pID_OP) = "" Then
        Err.Raise vbObjectError + 401, "IniciarProducao", "ID da OP não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    Application.ScreenUpdating = False
    
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_Producao").Index).Value = pID_Producao
        .Cells(1, tbl.ListColumns("ID_OP").Index).Value = pID_OP
        .Cells(1, tbl.ListColumns("ID_Operador").Index).Value = pID_Operador
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = Now
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = Empty
        .Cells(1, tbl.ListColumns("Quantidade_Planejada").Index).Value = pQtdPlanejada
        .Cells(1, tbl.ListColumns("Quantidade_Produzida").Index).Value = 0
        .Cells(1, tbl.ListColumns("Quantidade_Rejeitada").Index).Value = 0
        .Cells(1, tbl.ListColumns("Status").Index).Value = "EM PRODUÇÃO"
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroIniciar:
    MsgBox "Erro ao iniciar produção: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: PausarProducao
' PROPÓSITO: Pausar uma produção em andamento
'--------------------------------------------------------------------------------
Public Sub PausarProducao(pID_Producao As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroPausar
    
    If Trim(pID_Producao) = "" Then
        Err.Raise vbObjectError + 402, "PausarProducao", "ID da produção não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Producao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 403, "PausarProducao", "Produção não encontrada."
    End If
    
    If Trim(CStr(dados(linhaEncontrada, 10))) <> "EM PRODUÇÃO" Then
        Err.Raise vbObjectError + 404, "PausarProducao", "Produção não está em andamento."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Range.Cells(1, tbl.ListColumns("Status").Index).Value = "PAUSADA"
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroPausar:
    MsgBox "Erro ao pausar produção: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: RetomarProducao
' PROPÓSITO: Retomar uma produção pausada
'--------------------------------------------------------------------------------
Public Sub RetomarProducao(pID_Producao As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroRetomar
    
    If Trim(pID_Producao) = "" Then
        Err.Raise vbObjectError + 405, "RetomarProducao", "ID da produção não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Producao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 406, "RetomarProducao", "Produção não encontrada."
    End If
    
    If Trim(CStr(dados(linhaEncontrada, 10))) <> "PAUSADA" Then
        Err.Raise vbObjectError + 407, "RetomarProducao", "Produção não está pausada."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Range.Cells(1, tbl.ListColumns("Status").Index).Value = "EM PRODUÇÃO"
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroRetomar:
    MsgBox "Erro ao retomar produção: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: FinalizarProducao
' PROPÓSITO: Finalizar uma produção em andamento
'--------------------------------------------------------------------------------
Public Sub FinalizarProducao(pID_Producao As String, pQtdProduzida As Long, pQtdRejeitada As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroFinalizar
    
    If Trim(pID_Producao) = "" Then
        Err.Raise vbObjectError + 408, "FinalizarProducao", "ID da produção não informado."
    End If
    
    If pQtdProduzida < 0 Or pQtdRejeitada < 0 Then
        Err.Raise vbObjectError + 409, "FinalizarProducao", "Quantidades não podem ser negativas."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Producao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 410, "FinalizarProducao", "Produção não encontrada."
    End If
    
    With tbl.ListRows(linhaEncontrada - 1).Range
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = Now
        .Cells(1, tbl.ListColumns("Quantidade_Produzida").Index).Value = pQtdProduzida
        .Cells(1, tbl.ListColumns("Quantidade_Rejeitada").Index).Value = pQtdRejeitada
        .Cells(1, tbl.ListColumns("Status").Index).Value = "FINALIZADA"
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroFinalizar:
    MsgBox "Erro ao finalizar produção: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: RegistrarParada
' PROPÓSITO: Registrar início de parada
'--------------------------------------------------------------------------------
Public Sub RegistrarParada(pID_Parada As String, pID_OP As String, pEquipamento As String, _
                           pID_Operador As String, pID_Motivo As String, pObservacao As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroRegistrarParada
    
    If Trim(pID_Parada) = "" Then
        Err.Raise vbObjectError + 411, "RegistrarParada", "ID da parada não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Paradas")
    Set tbl = ws.ListObjects("TabelaParadas")
    
    Application.ScreenUpdating = False
    
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_Parada").Index).Value = pID_Parada
        .Cells(1, tbl.ListColumns("ID_OP").Index).Value = pID_OP
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("ID_Operador").Index).Value = pID_Operador
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = Now
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = Empty
        .Cells(1, tbl.ListColumns("ID_Motivo").Index).Value = pID_Motivo
        .Cells(1, tbl.ListColumns("Observacao").Index).Value = pObservacao
        .Cells(1, tbl.ListColumns("Duracao_Minutos").Index).Value = 0
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroRegistrarParada:
    MsgBox "Erro ao registrar parada: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: FinalizarParada
' PROPÓSITO: Finalizar uma parada em andamento
'--------------------------------------------------------------------------------
Public Sub FinalizarParada(pID_Parada As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    Dim inicioParada As Date
    Dim fimParada As Date
    Dim duracaoMinutos As Long
    
    On Error GoTo ErroFinalizarParada
    
    If Trim(pID_Parada) = "" Then
        Err.Raise vbObjectError + 412, "FinalizarParada", "ID da parada não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Paradas")
    Set tbl = ws.ListObjects("TabelaParadas")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Parada) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 413, "FinalizarParada", "Parada não encontrada."
    End If
    
    If IsDate(dados(linhaEncontrada, 6)) And Not IsEmpty(dados(linhaEncontrada, 6)) Then
        Err.Raise vbObjectError + 414, "FinalizarParada", "Parada já foi finalizada."
    End If
    
    fimParada = Now
    inicioParada = CDate(dados(linhaEncontrada, 5))
    duracaoMinutos = (fimParada - inicioParada) * 24 * 60
    
    With tbl.ListRows(linhaEncontrada - 1).Range
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = fimParada
        .Cells(1, tbl.ListColumns("Duracao_Minutos").Index).Value = duracaoMinutos
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroFinalizarParada:
    MsgBox "Erro ao finalizar parada: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: AtualizarProducao
' PROPÓSITO: Atualizar uma produção existente
'--------------------------------------------------------------------------------
Public Sub AtualizarProducao(pID_Producao As String, pID_OP As String, pID_Operador As String, _
                             pEquipamento As String, pDataInicio As Date, pDataFim As Date, _
                             pQtdPlanejada As Long, pQtdProduzida As Long, pQtdRejeitada As Long, pStatus As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroAtualizarProducao
    
    If Trim(pID_Producao) = "" Then
        Err.Raise vbObjectError + 415, "AtualizarProducao", "ID da produção não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Producao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 416, "AtualizarProducao", "Produção não encontrada."
    End If
    
    With tbl.ListRows(linhaEncontrada - 1).Range
        .Cells(1, tbl.ListColumns("ID_OP").Index).Value = pID_OP
        .Cells(1, tbl.ListColumns("ID_Operador").Index).Value = pID_Operador
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pDataInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pDataFim
        .Cells(1, tbl.ListColumns("Quantidade_Planejada").Index).Value = pQtdPlanejada
        .Cells(1, tbl.ListColumns("Quantidade_Produzida").Index).Value = pQtdProduzida
        .Cells(1, tbl.ListColumns("Quantidade_Rejeitada").Index).Value = pQtdRejeitada
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroAtualizarProducao:
    MsgBox "Erro ao atualizar produção: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExcluirProducao
' PROPÓSITO: Excluir uma produção
'--------------------------------------------------------------------------------
Public Sub ExcluirProducao(pID_Producao As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroExcluirProducao
    
    If Trim(pID_Producao) = "" Then
        Err.Raise vbObjectError + 417, "ExcluirProducao", "ID da produção não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Producoes")
    Set tbl = ws.ListObjects("TabelaProducoes")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Producao) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 418, "ExcluirProducao", "Produção não encontrada."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Delete
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroExcluirProducao:
    MsgBox "Erro ao excluir produção: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExcluirParada
' PROPÓSITO: Excluir uma parada
'--------------------------------------------------------------------------------
Public Sub ExcluirParada(pID_Parada As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim linhaEncontrada As Long
    
    On Error GoTo ErroExcluirParada
    
    If Trim(pID_Parada) = "" Then
        Err.Raise vbObjectError + 419, "ExcluirParada", "ID da parada não informado."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_Paradas")
    Set tbl = ws.ListObjects("TabelaParadas")
    
    Application.ScreenUpdating = False
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    linhaEncontrada = 0
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID_Parada) Then
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If linhaEncontrada = 0 Then
        Err.Raise vbObjectError + 420, "ExcluirParada", "Parada não encontrada."
    End If
    
    tbl.ListRows(linhaEncontrada - 1).Delete
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroExcluirParada:
    MsgBox "Erro ao excluir parada: " & Err.Description, vbCritical, "APS PURAN - Engine"
    Resume Sair
End Sub
