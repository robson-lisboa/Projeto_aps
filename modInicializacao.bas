Attribute VB_Name = "modInicializacao"
'================================================================================
' MÓDULO: modInicializacao
' DESCRIÇÃO: Inicialização da estrutura de planilhas e tabelas do APS PURAN
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' SUBROTINA: InicializarEstruturaAPS
' PROPÓSITO: Verificar e criar todas as planilhas e tabelas necessárias
'            para o funcionamento do sistema APS PURAN
'--------------------------------------------------------------------------------
Public Sub InicializarEstruturaAPS()
    On Error GoTo ErroInicializacao
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '--- BD_OPs ------------------------------------------------------------------
    If Not PlanilhaExiste("BD_OPs") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_OPs"
        CriarTabela ws, "TabelaOPs", Array("ID_OP", "Produto", "Equipamento", _
            "Quantidade", "Data_Inicio", "Data_Fim", "Duracao_Horas", "Status")
    Else
        Set ws = wb.Worksheets("BD_OPs")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaOPs") Then
            CriarTabela ws, "TabelaOPs", Array("ID_OP", "Produto", "Equipamento", _
                "Quantidade", "Data_Inicio", "Data_Fim", "Duracao_Horas", "Status")
        End If
    End If
    
    '--- BD_Equipamentos ---------------------------------------------------------
    If Not PlanilhaExiste("BD_Equipamentos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Equipamentos"
        CriarTabela ws, "TabelaEquipamentos", Array("ID_Equipamento", "Nome", _
            "Capacidade_Hora", "Status_Manutencao")
    Else
        Set ws = wb.Worksheets("BD_Equipamentos")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaEquipamentos") Then
            CriarTabela ws, "TabelaEquipamentos", Array("ID_Equipamento", "Nome", _
                "Capacidade_Hora", "Status_Manutencao")
        End If
    End If
    
    '--- BD_Eventos --------------------------------------------------------------
    If Not PlanilhaExiste("BD_Eventos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Eventos"
        CriarTabela ws, "TabelaEventos", Array("ID_Evento", "Tipo", "Equipamento", _
            "Inicio", "Fim", "Motivo")
    Else
        Set ws = wb.Worksheets("BD_Eventos")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaEventos") Then
            CriarTabela ws, "TabelaEventos", Array("ID_Evento", "Tipo", "Equipamento", _
                "Inicio", "Fim", "Motivo")
        End If
    End If
    
    '--- BD_Config ---------------------------------------------------------------
    If Not PlanilhaExiste("BD_Config") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Config"
        CriarTabela ws, "TabelaConfig", Array("Chave", "Valor")
    Else
        Set ws = wb.Worksheets("BD_Config")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaConfig") Then
            CriarTabela ws, "TabelaConfig", Array("Chave", "Valor")
        End If
    End If
    
    '--- BD_Produtos -------------------------------------------------------------
    If Not PlanilhaExiste("BD_Produtos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Produtos"
        CriarTabela ws, "TabelaProdutos", Array("ID_Produto", "Nome", "Descricao", "Tempo_Padrao_Horas", "Status")
    Else
        Set ws = wb.Worksheets("BD_Produtos")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaProdutos") Then
            CriarTabela ws, "TabelaProdutos", Array("ID_Produto", "Nome", "Descricao", _
                "Tempo_Padrao_Horas", "Status")
        End If
    End If
    
    '--- BD_Operadores -----------------------------------------------------------
    If Not PlanilhaExiste("BD_Operadores") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Operadores"
        CriarTabela ws, "TabelaOperadores", Array("ID_Operador", "Nome", "Status")
    Else
        Set ws = wb.Worksheets("BD_Operadores")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaOperadores") Then
            CriarTabela ws, "TabelaOperadores", Array("ID_Operador", "Nome", "Status")
        End If
    End If
    
    '--- BD_Turnos ---------------------------------------------------------------
    If Not PlanilhaExiste("BD_Turnos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Turnos"
        CriarTabela ws, "TabelaTurnos", Array("ID_Turno", "Nome", "Hora_Inicio", "Hora_Fim", "Dias_Funcionamento")
    Else
        Set ws = wb.Worksheets("BD_Turnos")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaTurnos") Then
            CriarTabela ws, "TabelaTurnos", Array("ID_Turno", "Nome", "Hora_Inicio", _
                "Hora_Fim", "Dias_Funcionamento")
        End If
    End If
    
    '--- BD_Processos ------------------------------------------------------------
    If Not PlanilhaExiste("BD_Processos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Processos"
        CriarTabela ws, "TabelaProcessos", Array("ID_Processo", "Produto", "Posto", "Sequencia", "Tempo_Padrao_Horas")
    Else
        Set ws = wb.Worksheets("BD_Processos")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaProcessos") Then
            CriarTabela ws, "TabelaProcessos", Array("ID_Processo", "Produto", _
                "Posto", "Sequencia", "Tempo_Padrao_Horas")
        End If
    End If
    
    '--- BD_MotivosParada --------------------------------------------------------
    If Not PlanilhaExiste("BD_MotivosParada") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_MotivosParada"
        CriarTabela ws, "TabelaMotivosParada", Array("ID_Motivo", "Descricao", "Categoria", "Status")
    Else
        Set ws = wb.Worksheets("BD_MotivosParada")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaMotivosParada") Then
            CriarTabela ws, "TabelaMotivosParada", Array("ID_Motivo", "Descricao", _
                "Categoria", "Status")
        End If
    End If
    
    '--- BD_Simulacoes -----------------------------------------------------------
    If Not PlanilhaExiste("BD_Simulacoes") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Simulacoes"
        CriarTabela ws, "TabelaSimulacoes", Array("ID_Simulacao", "Nome_Simulacao", "Data_Criacao", _
            "Data_Inicio", "Data_Fim", "Status", "Observacao")
    Else
        Set ws = wb.Worksheets("BD_Simulacoes")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaSimulacoes") Then
            CriarTabela ws, "TabelaSimulacoes", Array("ID_Simulacao", "Nome_Simulacao", _
                "Data_Criacao", "Data_Inicio", "Data_Fim", "Status", "Observacao")
        End If
    End If
    
    '--- BD_OPsSimulacao ---------------------------------------------------------
    If Not PlanilhaExiste("BD_OPsSimulacao") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_OPsSimulacao"
        CriarTabela ws, "TabelaOPsSimulacao", Array("ID_Simulacao", "ID_OP_Simulacao", "ID_OP_Origem", _
            "Produto", "Equipamento", "Quantidade", "Data_Inicio", "Data_Fim", "Duracao_Horas", _
            "Status", "Prioridade", "Observacao")
    Else
        Set ws = wb.Worksheets("BD_OPsSimulacao")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaOPsSimulacao") Then
            CriarTabela ws, "TabelaOPsSimulacao", Array("ID_Simulacao", "ID_OP_Simulacao", _
                "ID_OP_Origem", "Produto", "Equipamento", "Quantidade", "Data_Inicio", "Data_Fim", _
                "Duracao_Horas", "Status", "Prioridade", "Observacao")
        End If
    End If
    
    '--- BD_Producoes ------------------------------------------------------------
    If Not PlanilhaExiste("BD_Producoes") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Producoes"
        CriarTabela ws, "TabelaProducoes", Array("ID_Producao", "ID_OP", "ID_Operador", _
            "Equipamento", "Data_Inicio", "Data_Fim", "Quantidade_Planejada", _
            "Quantidade_Produzida", "Quantidade_Rejeitada", "Status")
    Else
        Set ws = wb.Worksheets("BD_Producoes")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaProducoes") Then
            CriarTabela ws, "TabelaProducoes", Array("ID_Producao", "ID_OP", "ID_Operador", _
                "Equipamento", "Data_Inicio", "Data_Fim", "Quantidade_Planejada", _
                "Quantidade_Produzida", "Quantidade_Rejeitada", "Status")
        End If
    End If
    
    '--- BD_Paradas --------------------------------------------------------------
    If Not PlanilhaExiste("BD_Paradas") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Paradas"
        CriarTabela ws, "TabelaParadas", Array("ID_Parada", "ID_OP", "Equipamento", _
            "ID_Operador", "Data_Inicio", "Data_Fim", "ID_Motivo", "Observacao", "Duracao_Minutos")
    Else
        Set ws = wb.Worksheets("BD_Paradas")
        If Not TabelaExisteEstaPlanilha(ws, "TabelaParadas") Then
            CriarTabela ws, "TabelaParadas", Array("ID_Parada", "ID_OP", "Equipamento", _
                "ID_Operador", "Data_Inicio", "Data_Fim", "ID_Motivo", "Observacao", "Duracao_Minutos")
        End If
    End If
    
    '--- Validação somente leitura das tabelas existentes -------------------------
    Call ValidarTabelasExistentes
    
Sair:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar estrutura do APS PURAN:" & vbCrLf & _
           "Descrição: " & Err.Description & vbCrLf & _
           "Linha: " & Erl, vbCritical + vbOKOnly, "APS PURAN - Inicialização"
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ValidarTabelasExistentes
' PROPÓSITO: Verificar se as tabelas existentes estão corretas SEM apagar dados.
'            Apenas informa inconsistências encontradas.
'--------------------------------------------------------------------------------
Private Sub ValidarTabelasExistentes()
    On Error Resume Next
    
    Dim inconsistencia As String
    inconsistencia = ""
    
    '--- BD_OPs / TabelaOPs ------------------------------------------------------
    If Not VerificarTabela("BD_OPs", "TabelaOPs", Array("ID_OP", "Produto", "Equipamento", _
        "Quantidade", "Data_Inicio", "Data_Fim", "Duracao_Horas", "Status")) Then
        inconsistencia = inconsistencia & "- BD_OPs/TabelaOPs: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Equipamentos / TabelaEquipamentos -------------------------------------
    If Not VerificarTabela("BD_Equipamentos", "TabelaEquipamentos", Array("ID_Equipamento", "Nome", _
        "Capacidade_Hora", "Status_Manutencao")) Then
        inconsistencia = inconsistencia & "- BD_Equipamentos/TabelaEquipamentos: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Eventos / TabelaEventos ----------------------------------------------
    If Not VerificarTabela("BD_Eventos", "TabelaEventos", Array("ID_Evento", "Tipo", "Equipamento", _
        "Inicio", "Fim", "Motivo")) Then
        inconsistencia = inconsistencia & "- BD_Eventos/TabelaEventos: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Config / TabelaConfig -------------------------------------------------
    If Not VerificarTabela("BD_Config", "TabelaConfig", Array("Chave", "Valor")) Then
        inconsistencia = inconsistencia & "- BD_Config/TabelaConfig: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Produtos / TabelaProdutos ---------------------------------------------
    If Not VerificarTabela("BD_Produtos", "TabelaProdutos", Array("ID_Produto", "Nome", "Descricao", _
        "Tempo_Padrao_Horas", "Status")) Then
        inconsistencia = inconsistencia & "- BD_Produtos/TabelaProdutos: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Operadores / TabelaOperadores -----------------------------------------
    If Not VerificarTabela("BD_Operadores", "TabelaOperadores", Array("ID_Operador", "Nome", "Status")) Then
        inconsistencia = inconsistencia & "- BD_Operadores/TabelaOperadores: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Turnos / TabelaTurnos -------------------------------------------------
    If Not VerificarTabela("BD_Turnos", "TabelaTurnos", Array("ID_Turno", "Nome", "Hora_Inicio", _
        "Hora_Fim", "Dias_Funcionamento")) Then
        inconsistencia = inconsistencia & "- BD_Turnos/TabelaTurnos: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Processos / TabelaProcessos -------------------------------------------
    If Not VerificarTabela("BD_Processos", "TabelaProcessos", Array("ID_Processo", "Produto", _
        "Posto", "Sequencia", "Tempo_Padrao_Horas")) Then
        inconsistencia = inconsistencia & "- BD_Processos/TabelaProcessos: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_MotivosParada / TabelaMotivosParada -----------------------------------
    If Not VerificarTabela("BD_MotivosParada", "TabelaMotivosParada", Array("ID_Motivo", "Descricao", _
        "Categoria", "Status")) Then
        inconsistencia = inconsistencia & "- BD_MotivosParada/TabelaMotivosParada: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Simulacoes / TabelaSimulacoes -----------------------------------------
    If Not VerificarTabela("BD_Simulacoes", "TabelaSimulacoes", Array("ID_Simulacao", "Nome_Simulacao", _
        "Data_Criacao", "Data_Inicio", "Data_Fim", "Status", "Observacao")) Then
        inconsistencia = inconsistencia & "- BD_Simulacoes/TabelaSimulacoes: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_OPsSimulacao / TabelaOPsSimulacao -------------------------------------
    If Not VerificarTabela("BD_OPsSimulacao", "TabelaOPsSimulacao", Array("ID_Simulacao", "ID_OP_Simulacao", _
        "ID_OP_Origem", "Produto", "Equipamento", "Quantidade", "Data_Inicio", "Data_Fim", _
        "Duracao_Horas", "Status", "Prioridade", "Observacao")) Then
        inconsistencia = inconsistencia & "- BD_OPsSimulacao/TabelaOPsSimulacao: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Producoes / TabelaProducoes -------------------------------------------
    If Not VerificarTabela("BD_Producoes", "TabelaProducoes", Array("ID_Producao", "ID_OP", "ID_Operador", _
        "Equipamento", "Data_Inicio", "Data_Fim", "Quantidade_Planejada", _
        "Quantidade_Produzida", "Quantidade_Rejeitada", "Status")) Then
        inconsistencia = inconsistencia & "- BD_Producoes/TabelaProducoes: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    '--- BD_Paradas / TabelaParadas -----------------------------------------------
    If Not VerificarTabela("BD_Paradas", "TabelaParadas", Array("ID_Parada", "ID_OP", "Equipamento", _
        "ID_Operador", "Data_Inicio", "Data_Fim", "ID_Motivo", "Observacao", "Duracao_Minutos")) Then
        inconsistencia = inconsistencia & "- BD_Paradas/TabelaParadas: estrutura de colunas não corresponde ao esperado." & vbCrLf
    End If
    
    On Error GoTo 0
    
    If inconsistencia <> "" Then
        MsgBox "A estrutura das tabelas existentes não corresponde ao esperado pelo sistema." & vbCrLf & vbCrLf & _
               "As tabelas não foram alteradas." & vbCrLf & vbCrLf & _
               "Inconsistências encontradas:" & vbCrLf & inconsistencia & vbCrLf & _
               "Correção manual: recrie as tabelas ou restaure o arquivo base.", _
               vbExclamation + vbOKOnly, "APS PURAN - Validação"
    End If
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: VerificarTabela
' PROPÓSITO: Verificar se uma tabela existe e tem as colunas esperadas
' RETORNO: True se a tabela está correta, False se há inconsistência
'--------------------------------------------------------------------------------
Private Function VerificarTabela(pNomePlanilha As String, _
                                 pNomeTabela As String, _
                                 pColunasEsperadas As Variant) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim nomeColuna As String
    Dim colunaEncontrada As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(pNomePlanilha)
    On Error GoTo 0
    
    If ws Is Nothing Then
        VerificarTabela = False
        Exit Function
    End If
    
    On Error Resume Next
    Set tbl = ws.ListObjects(pNomeTabela)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        VerificarTabela = False
        Exit Function
    End If
    
    ' Verifica se todas as colunas esperadas existem
    For i = LBound(pColunasEsperadas) To UBound(pColunasEsperadas)
        nomeColuna = CStr(pColunasEsperadas(i))
        colunaEncontrada = False
        
        Dim coluna As ListColumn
        For Each coluna In tbl.ListColumns
            If coluna.Name = nomeColuna Then
                colunaEncontrada = True
                Exit For
            End If
        Next coluna
        
        If Not colunaEncontrada Then
            VerificarTabela = False
            Exit Function
        End If
    Next i
    
    VerificarTabela = True
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: PlanilhaExiste
' PROPÓSITO: Verificar se uma planilha com o nome informado existe na pasta de trabalho
'--------------------------------------------------------------------------------
Private Function PlanilhaExiste(pNome As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(pNome)
    PlanilhaExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: TabelaExisteEstaPlanilha
' PROPÓSITO: Verificar se uma ListObject com o nome informado existe na planilha
'--------------------------------------------------------------------------------
Private Function TabelaExisteEstaPlanilha(pWorksheet As Worksheet, pNomeTabela As String) As Boolean
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = pWorksheet.ListObjects(pNomeTabela)
    TabelaExisteEstaPlanilha = Not tbl Is Nothing
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: CriarTabela
' PROPÓSITO: Criar uma ListObject (Tabela Estruturada) com os cabeçalhos definidos
'--------------------------------------------------------------------------------
Private Sub CriarTabela(pWorksheet As Worksheet, _
                        pNomeTabela As String, _
                        pColunas As Variant)
    Dim i As Long
    Dim tbl As ListObject
    
    With pWorksheet
        For i = LBound(pColunas) To UBound(pColunas)
            .Cells(1, i + 1).Value = pColunas(i)
            .Cells(1, i + 1).Font.Bold = True
        Next i
        
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A1").Resize(1, UBound(pColunas) + 1), , xlYes)
        tbl.Name = pNomeTabela
        tbl.TableStyle = "TableStyleMedium2"
    End With
End Sub
