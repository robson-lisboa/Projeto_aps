Attribute VB_Name = "modDatabase"
'================================================================================
' MÓDULO: modDatabase
' DESCRIÇÃO: Gerenciamento do banco de dados oculto (ListObjects/ListObjects)
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' SUBROTINA: CriarEstruturasDeDados
' PROPÓSITO: Garantir a existência das planilhas ocultas e suas tabelas estruturadas
'--------------------------------------------------------------------------------
Public Sub CriarEstruturasDeDados()
    On Error GoTo ErroCriarEstruturas
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    Application.ScreenUpdating = False
    
    '--- BD_OPs ------------------------------------------------------------------
    If Not PlanilhaExiste("BD_OPs") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_OPs"
        CriarTabela ws, "TabelaOPs", Array("ID_OP", "Produto", "Equipamento", _
            "Quantidade", "Data_Inicio", "Data_Fim", "Duracao_Horas", "Status")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Equipamentos ---------------------------------------------------------
    If Not PlanilhaExiste("BD_Equipamentos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Equipamentos"
        CriarTabela ws, "TabelaEquipamentos", Array("ID_Equipamento", "Nome", _
            "Capacidade_Hora", "Status_Manutencao")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Eventos --------------------------------------------------------------
    If Not PlanilhaExiste("BD_Eventos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Eventos"
        CriarTabela ws, "TabelaEventos", Array("ID_Evento", "Tipo", "Equipamento", _
            "Inicio", "Fim", "Motivo")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Config ---------------------------------------------------------------
    If Not PlanilhaExiste("BD_Config") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Config"
        CriarTabela ws, "TabelaConfig", Array("Chave", "Valor")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Produtos -------------------------------------------------------------
    If Not PlanilhaExiste("BD_Produtos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Produtos"
        CriarTabela ws, "TabelaProdutos", Array("ID_Produto", "Nome", "Descricao", "Tempo_Padrao_Horas", "Status")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Operadores -----------------------------------------------------------
    If Not PlanilhaExiste("BD_Operadores") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Operadores"
        CriarTabela ws, "TabelaOperadores", Array("ID_Operador", "Nome", "Status")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Turnos ---------------------------------------------------------------
    If Not PlanilhaExiste("BD_Turnos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Turnos"
        CriarTabela ws, "TabelaTurnos", Array("ID_Turno", "Nome", "Hora_Inicio", "Hora_Fim", "Dias_Funcionamento")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Processos ------------------------------------------------------------
    If Not PlanilhaExiste("BD_Processos") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Processos"
        CriarTabela ws, "TabelaProcessos", Array("ID_Processo", "Produto", "Posto", "Sequencia", "Tempo_Padrao_Horas")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_MotivosParada --------------------------------------------------------
    If Not PlanilhaExiste("BD_MotivosParada") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_MotivosParada"
        CriarTabela ws, "TabelaMotivosParada", Array("ID_Motivo", "Descricao", "Categoria", "Status")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Simulacoes -----------------------------------------------------------
    If Not PlanilhaExiste("BD_Simulacoes") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Simulacoes"
        CriarTabela ws, "TabelaSimulacoes", Array("ID_Simulacao", "Nome_Simulacao", "Data_Criacao", _
            "Data_Inicio", "Data_Fim", "Status", "Observacao")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_OPsSimulacao ---------------------------------------------------------
    If Not PlanilhaExiste("BD_OPsSimulacao") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_OPsSimulacao"
        CriarTabela ws, "TabelaOPsSimulacao", Array("ID_Simulacao", "ID_OP_Simulacao", "ID_OP_Origem", _
            "Produto", "Equipamento", "Quantidade", "Data_Inicio", "Data_Fim", "Duracao_Horas", _
            "Status", "Prioridade", "Observacao")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Producoes ------------------------------------------------------------
    If Not PlanilhaExiste("BD_Producoes") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Producoes"
        CriarTabela ws, "TabelaProducoes", Array("ID_Producao", "ID_OP", "ID_Operador", _
            "Equipamento", "Data_Inicio", "Data_Fim", "Quantidade_Planejada", _
            "Quantidade_Produzida", "Quantidade_Rejeitada", "Status")
        ws.Visible = xlSheetVeryHidden
    End If
    
    '--- BD_Paradas --------------------------------------------------------------
    If Not PlanilhaExiste("BD_Paradas") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "BD_Paradas"
        CriarTabela ws, "TabelaParadas", Array("ID_Parada", "ID_OP", "Equipamento", _
            "ID_Operador", "Data_Inicio", "Data_Fim", "ID_Motivo", "Observacao", "Duracao_Minutos")
        ws.Visible = xlSheetVeryHidden
    End If
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroCriarEstruturas:
    MsgBox "Erro ao criar estruturas de dados: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN - Banco de Dados"
    Resume Sair
End Sub

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
