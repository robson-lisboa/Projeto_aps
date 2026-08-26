VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Simulação de Produção"
  ClientHeight = 600
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 900
  Height = 638
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 912
  StartUpPosition = 1
  Attribute VB_Name = "frmSimulacao"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = True
  Attribute VB_Exposed = False
  BeginProperty Font
    Name = "Segoe UI"
    CharSet = 0
    Weight = 400
    Size = 9
  EndProperty
  KeyPreview = -1
END
Attribute VB_Name = "frmSimulacao"
'================================================================================
' USERFORM: frmSimulacao
' DESCRIÇÃO: Gerenciamento de simulações de produção mensal
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES PRINCIPAIS ======================================================
Private lstSimulacoes As MSForms.ListBox
Private fraAcoes As MSForms.Frame
Private fraVisualizacao As MSForms.Frame
Private hScrollSimulacao As MSForms.ScrollBar

'=== BOTÕES =====================================================================
Private btnNovaSimulacao As MSForms.CommandButton
Private btnAbrirSimulacao As MSForms.CommandButton
Private btnImportarOPs As MSForms.CommandButton
Private btnExcluirSimulacao As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

'=== CAMPOS DE FILTRO ===========================================================
Private lblPeriodo As MSForms.Label
Private txtDataInicial As MSForms.TextBox
Private txtDataFinal As MSForms.TextBox
Private cmdAplicarPeriodo As MSForms.CommandButton

'=== ESTADO =====================================================================
Private m_ID_SimulacaoAtiva As String
Private m_DataInicioPeriodo As Date
Private m_DataFimPeriodo As Date
Private m_Zoom As Double
Private m_DragCard As MSForms.Label
Private m_DragStartX As Single
Private m_DragStartY As Single
Private m_DragOriginalLeft As Single
Private m_DragOriginalTop As Single

'=== PROPRIEDADES VISUAIS ======================================================
Private Const COR_FUNDO As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_AZUL As Long = 15773696
Private Const COR_VERDE As Long = 5287936
Private Const COR_ALERTA As Long = 255

'================================================================================
' EVENTO: UserForm_Initialize
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Me.BackColor = COR_FUNDO
    Me.Caption = "APS PURAN – Simulação de Produção"
    
    m_ID_SimulacaoAtiva = ""
    m_DataInicioPeriodo = DateSerial(Year(Date), Month(Date), 1)
    m_DataFimPeriodo = DateSerial(Year(Date), Month(Date) + 1, 0)
    m_Zoom = 1.0
    
    Call CriarControles
    Call CarregarListaSimulacoes
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarControles
'================================================================================
Private Sub CriarControles()
    On Error Resume Next
    
    '--- Título ----------------------------------------------------------------
    Dim lblTitulo As MSForms.Label
    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    With lblTitulo
        .Caption = "SIMULAÇÃO DE PRODUÇÃO MENSAL"
        .Left = 12
        .Top = 12
        .Width = 860
        .Height = 32
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Indicadores gerais ----------------------------------------------------
    Dim lblIndicadores As MSForms.Label
    Set lblIndicadores = Me.Controls.Add("Forms.Label.1", "lblIndicadores", True)
    With lblIndicadores
        .Caption = "OPs: 0 | Horas: 0h | Ocupação: 0% | Conflitos: 0"
        .Left = 12
        .Top = 52
        .Width = 860
        .Height = 18
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignLeft
    End With
    
    '--- Lista de simulações ---------------------------------------------------
    Set lstSimulacoes = Me.Controls.Add("Forms.ListBox.1", "lstSimulacoes", True)
    With lstSimulacoes
        .Left = 12
        .Top = 78
        .Width = 400
        .Height = 260
        .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
    End With
    
    '--- Botões de ação --------------------------------------------------------
    Dim btnLeft As Single
    btnLeft = 12
    
    Set btnNovaSimulacao = Me.Controls.Add("Forms.CommandButton.1", "btnNovaSimulacao", True)
    With btnNovaSimulacao
        .Caption = "Nova Simulação"
        .Left = btnLeft
        .Top = 348
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 130
    
    Set btnAbrirSimulacao = Me.Controls.Add("Forms.CommandButton.1", "btnAbrirSimulacao", True)
    With btnAbrirSimulacao
        .Caption = "Abrir"
        .Left = btnLeft
        .Top = 348
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnImportarOPs = Me.Controls.Add("Forms.CommandButton.1", "btnImportarOPs", True)
    With btnImportarOPs
        .Caption = "Importar OPs"
        .Left = btnLeft
        .Top = 348
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 130
    
    Set btnExcluirSimulacao = Me.Controls.Add("Forms.CommandButton.1", "btnExcluirSimulacao", True)
    With btnExcluirSimulacao
        .Caption = "Excluir"
        .Left = btnLeft
        .Top = 348
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Filtro de período ------------------------------------------------------
    Set lblPeriodo = Me.Controls.Add("Forms.Label.1", "lblPeriodo", True)
    With lblPeriodo
        .Caption = "Período:"
        .Left = 12
        .Top = 398
        .Width = 50
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set txtDataInicial = Me.Controls.Add("Forms.TextBox.1", "txtDataInicial", True)
    With txtDataInicial
        .Text = Format(m_DataInicioPeriodo, "dd/mm/yyyy")
        .Left = 68
        .Top = 398
        .Width = 90
        .Height = 22
        .Font.Size = 9
    End With
    
    Set txtDataFinal = Me.Controls.Add("Forms.TextBox.1", "txtDataFinal", True)
    With txtDataFinal
        .Text = Format(m_DataFimPeriodo, "dd/mm/yyyy")
        .Left = 165
        .Top = 398
        .Width = 90
        .Height = 22
        .Font.Size = 9
    End With
    
    Set cmdAplicarPeriodo = Me.Controls.Add("Forms.CommandButton.1", "cmdAplicarPeriodo", True)
    With cmdAplicarPeriodo
        .Caption = "Aplicar"
        .Left = 262
        .Top = 398
        .Width = 70
        .Height = 22
        .Font.Size = 9
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Filtros ---------------------------------------------------------------
    Dim lblFiltroStatus As MSForms.Label
    Set lblFiltroStatus = Me.Controls.Add("Forms.Label.1", "lblFiltroStatus", True)
    With lblFiltroStatus
        .Caption = "Status:"
        .Left = 12
        .Top = 428
        .Width = 50
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Dim cboFiltroStatus As MSForms.ComboBox
    Set cboFiltroStatus = Me.Controls.Add("Forms.ComboBox.1", "cboFiltroStatus", True)
    With cboFiltroStatus
        .AddItem "Todos"
        .AddItem "Planejada"
        .AddItem "Em Andamento"
        .AddItem "Concluído"
        .AddItem "Atrasado"
        .Value = "Todos"
        .Left = 68
        .Top = 428
        .Width = 100
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    Dim lblBusca As MSForms.Label
    Set lblBusca = Me.Controls.Add("Forms.Label.1", "lblBusca", True)
    With lblBusca
        .Caption = "Buscar:"
        .Left = 180
        .Top = 428
        .Width = 50
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Dim txtBusca As MSForms.TextBox
    Set txtBusca = Me.Controls.Add("Forms.TextBox.1", "txtBusca", True)
    With txtBusca
        .Text = ""
        .Left = 235
        .Top = 428
        .Width = 100
        .Height = 22
        .Font.Size = 9
        .PlaceholderText = "ID ou Produto"
    End With
    
    '--- Área de visualização do planejamento -----------------------------------
    Set fraVisualizacao = Me.Controls.Add("Forms.Frame.1", "fraVisualizacao", True)
    With fraVisualizacao
        .Caption = ""
        .Left = 430
        .Top = 60
        .Width = 440
        .Height = 380
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
        .Font.Size = 10
    End With
    
    '--- Scroll horizontal da simulação -----------------------------------------
    Set hScrollSimulacao = Me.Controls.Add("Forms.ScrollBar.1", "hScrollSimulacao", True)
    With hScrollSimulacao
        .Left = 430
        .Top = Me.ClientHeight - 60 - 20
        .Width = 440
        .Height = 16
        .Min = 0
        .Max = 100
        .SmallChange = 10
        .LargeChange = 50
        .Value = 0
    End With
    
    '--- Botão fechar -----------------------------------------------------------
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = "Fechar"
        .Left = 760
        .Top = 550
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
End Sub

'================================================================================
' SUBROTINA PRIVADA: CarregarListaSimulacoes
'================================================================================
Private Sub CarregarListaSimulacoes()
    On Error GoTo ErroCarregarLista
    
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim texto As String
    
    dados = ObterDadosSimulacaoEmArray()
    
    If IsError(dados) Then
        lstSimulacoes.Clear
        Exit Sub
    End If
    
    lstSimulacoes.Clear
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        texto = CStr(dados(i, 1)) & " - " & CStr(dados(i, 2)) & _
                " (" & Format(CDate(dados(i, 4)), "dd/mm/yyyy") & " a " & Format(CDate(dados(i, 5)), "dd/mm/yyyy") & ")"
        lstSimulacoes.AddItem texto
    Next i
    
Sair:
    Exit Sub
    
ErroCarregarLista:
    MsgBox "Erro ao carregar simulações: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: AbrirSimulacao
'================================================================================
Private Sub AbrirSimulacao()
    On Error GoTo ErroAbrir
    
    If lstSimulacoes.ListIndex = -1 Then
        MsgBox "Selecione uma simulação para abrir.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim chave As String
    chave = Split(lstSimulacoes.List(lstSimulacoes.ListIndex), " - ")(0)
    m_ID_SimulacaoAtiva = chave
    
    Call CarregarPlanejamentoSimulacao
    
Sair:
    Exit Sub
    
ErroAbrir:
    MsgBox "Erro ao abrir simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarPlanejamentoSimulacao
' PROPÓSITO: Renderizar planejamento da simulação ativa
'================================================================================
Public Sub CarregarPlanejamentoSimulacao()
    On Error GoTo ErroCarregar
    
    If m_ID_SimulacaoAtiva = "" Then Exit Sub
    
    ' Limpa área de visualização
    Dim ctrl As MSForms.Control
    For Each ctrl In fraVisualizacao.Controls
        fraVisualizacao.Controls.Remove ctrl.Name
    Next ctrl
    
    ' Carrega OPs da simulação
    Dim dados As Variant
    dados = ObterOPsSimulacaoEmArray(m_ID_SimulacaoAtiva)
    
    If IsError(dados) Then
        Dim lblVazio As MSForms.Label
        Set lblVazio = fraVisualizacao.Controls.Add("Forms.Label.1", "lblVazio", True)
        With lblVazio
            .Caption = "Nenhuma OP cadastrada nesta simulação."
            .Left = 20
            .Top = 20
            .Width = 400
            .Height = 40
            .Font.Size = 10
        End With
        Exit Sub
    End If
    
    ' Cria coleção de OPs para o modGantt
    Dim colOPs As New Collection
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    For i = 1 To totalLinhas
        Dim op As clsCardProducao
        Set op = New clsCardProducao
        With op
            .ID_OP = CStr(dados(i, 2)) & " (" & CStr(dados(i, 1)) & ")"
            .Produto = CStr(dados(i, 4))
            .Equipamento = CStr(dados(i, 5))
            .Quantidade = CLng(dados(i, 6))
            If IsDate(dados(i, 7)) Then .DataInicio = CDate(dados(i, 7))
            If IsDate(dados(i, 8)) Then .DataFim = CDate(dados(i, 8))
            .Duracao = CDbl(dados(i, 9))
            .Status = CStr(dados(i, 10))
        End With
        colOPs.Add op
    Next i
    
    ' Extrai postos distintos
    Dim colPostos As New Collection
    Dim nomePosto As String
    Dim existePosto As Boolean
    Dim j As Long
    
    For i = 1 To totalLinhas
        nomePosto = Trim(CStr(dados(i, 5)))
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
    
    ' Renderiza usando lógica similar ao modGantt (simplificada para reutilização)
    Call RenderizarPlanejamentoSimulacao(fraVisualizacao, colOPs, colPostos, m_DataInicioPeriodo, m_DataFimPeriodo, m_Zoom)
    
Sair:
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar planejamento da simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: RenderizarPlanejamentoSimulacao
' PROPÓSITO: Renderizar cards da simulação (reutiliza lógica do modGantt)
'================================================================================
Private Sub RenderizarPlanejamentoSimulacao(pContainer As MSForms.Frame, _
                                            pColOPs As Collection, _
                                            pColPostos As Collection, _
                                            pDataInicio As Date, _
                                            pDataFim As Date, _
                                            pZoom As Double)
    On Error GoTo ErroRenderizar
    
    Dim escala As EscalaTemporal
    escala = CalcularEscalaTemporal(pDataInicio, pDataFim, pZoom)
    
    Dim topoTimeline As Single
    topoTimeline = 32
    
    Dim margemEsquerda As Single
    margemEsquerda = 120
    
    Dim larguraTimeline As Double
    larguraTimeline = escala.LarguraTotal
    
    Dim numPostos As Long
    numPostos = pColPostos.Count
    
    ' Fundo da timeline
    Dim fundoTimeline As MSForms.Label
    Set fundoTimeline = pContainer.Controls.Add("Forms.Label.1", "lblSimFundoTimeline", True)
    With fundoTimeline
        .Left = margemEsquerda
        .Top = topoTimeline
        .Width = larguraTimeline
        .Height = 28
        .BackColor = COR_HEADER
    End With
    
    ' Marcadores de tempo
    Dim intervaloMarcador As Double
    Dim unidade As String
    Dim totalDias As Long
    totalDias = DateDiff("d", pDataInicio, pDataFim) + 1
    
    If totalDias = 1 Then
        intervaloMarcador = 1 / 24
        unidade = "hora"
    ElseIf totalDias <= 7 Then
        intervaloMarcador = 1
        unidade = "dia"
    Else
        intervaloMarcador = 1
        unidade = "dia"
    End If
    
    Dim passo As Double
    passo = intervaloMarcador
    Dim contadorMarcador As Long
    contadorMarcador = 0
    
    Do While pDataInicio + passo <= pDataFim Or (contadorMarcador = 0)
        Dim posMarcador As Double
        Dim dataMarcador As Date
        dataMarcador = pDataInicio + passo
        
        posMarcador = ConverterDataParaX(dataMarcador, escala)
        
        Dim textoMarcador As String
        If unidade = "hora" Then
            textoMarcador = Format(dataMarcador, "HH:MM")
        Else
            textoMarcador = Format(dataMarcador, "dd/mm")
        End If
        
        Dim lblMarcador As MSForms.Label
        Set lblMarcador = pContainer.Controls.Add("Forms.Label.1", "lblSimMarc_" & contadorMarcador, True)
        With lblMarcador
            .Caption = textoMarcador
            .Left = posMarcador
            .Top = topoTimeline + 4
            .Width = 60
            .Height = 20
            .Font.Size = 8
            .ForeColor = COR_TEXTO_CLARO
            .BackColor = COR_HEADER
            .TextAlign = fmTextAlignCenter
        End With
        
        passo = passo + intervaloMarcador
        contadorMarcador = contadorMarcador + 1
    Loop
    
    ' Labels dos postos
    Dim topoPostos As Single
    topoPostos = topoTimeline + 28
    
    Dim i As Long
    Dim nomePostoAtual As String
    Dim topoCard As Single
    
    For i = 1 To numPostos
        nomePostoAtual = pColPostos(i)
        topoCard = topoPostos + 8 + ((i - 1) * 80)
        
        Dim lblPosto As MSForms.Label
        Set lblPosto = pContainer.Controls.Add("Forms.Label.1", "lblSimPosto_" & i, True)
        With lblPosto
            .Caption = nomePostoAtual
            .Left = 4
            .Top = topoCard
            .Width = margemEsquerda - 8
            .Height = 24
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
            .TextAlign = fmTextAlignRight
        End With
        
        ' Linha separadora
        Dim linhaH As MSForms.Label
        Set linhaH = pContainer.Controls.Add("Forms.Label.1", "lblSimLinhaH_" & i, True)
        With linhaH
            .Left = margemEsquerda
            .Top = topoCard + 26
            .Width = larguraTimeline
            .Height = 1
            .BackColor = COR_TEXTO_ESCURO
        End With
    Next i
    
    ' Detecta conflitos
    Dim conflitos As Collection
    Set conflitos = DetectarConflitos(pColOPs, pColPostos)
    
    Dim conflitoMarcado() As Boolean
    ReDim conflitoMarcado(1 To pColOPs.Count)
    
    Dim k As Long
    For i = 1 To conflitos.Count
        Dim par As Collection
        Set par = conflitos(i)
        If par.Count >= 2 Then
            Dim idxA As Long, idxB As Long
            idxA = 0
            idxB = 0
            For j = 1 To pColOPs.Count
                If pColOPs(j).ID_OP = par(1).ID_OP Then idxA = j
                If pColOPs(j).ID_OP = par(2).ID_OP Then idxB = j
            Next j
            If idxA > 0 Then conflitoMarcado(idxA) = True
            If idxB > 0 Then conflitoMarcado(idxB) = True
        End If
    Next i
    
    ' Renderiza cards
    Dim alturaCard As Single
    alturaCard = 80 - 8
    
    Dim uniqueId As Long
    uniqueId = 0
    
    Dim opAtual As clsCardProducao
    Dim indicePosto As Long
    Dim posX As Double
    Dim larguraCard As Double
    Dim corCard As Long
    Dim corBorda As Long
    
    For i = 1 To pColOPs.Count
        Set opAtual = pColOPs(i)
        indicePosto = ObterIndicePosto(opAtual.Equipamento, pColPostos)
        posX = ConverterDataParaX(opAtual.DataInicio, escala)
        larguraCard = CalcularLarguraCard(opAtual.DataInicio, opAtual.DataFim, escala)
        topoCard = topoPostos + 8 + (indicePosto * 80)
        
        uniqueId = uniqueId + 1
        
        ' Cores
        corBorda = COR_HEADER
        corCard = COR_FUNDO
        
        Select Case Trim(opAtual.Status)
            Case "Atrasado"
                corBorda = COR_ALERTA
                corCard = 10092543
            Case "Em Andamento", "Concluído"
                corBorda = COR_AZUL
                corCard = COR_FUNDO
            Case Else
                corBorda = COR_HEADER
                corCard = COR_FUNDO
        End Select
        
        If conflitoMarcado(i) Then
            corBorda = COR_ALERTA
        End If
        
        ' Card
        Dim card As MSForms.Label
        Set card = pContainer.Controls.Add("Forms.Label.1", "cardSim_" & uniqueId, True)
        With card
            .Caption = ""
            .Tag = opAtual.ID_OP
            .Left = posX
            .Top = topoCard
            .Width = larguraCard
            .Height = alturaCard
            .BackColor = corCard
            .BorderStyle = fmBorderStyleSingle
        End With
        
        ' Borda esquerda
        Dim bordaEsq As MSForms.Label
        Set bordaEsq = pContainer.Controls.Add("Forms.Label.1", "cardSimBorda_" & uniqueId, True)
        With bordaEsq
            .Left = posX
            .Top = topoCard
            .Width = 4
            .Height = alturaCard
            .BackColor = corBorda
        End With
        
        ' ID OP
        Dim lblID As MSForms.Label
        Set lblID = pContainer.Controls.Add("Forms.Label.1", "cardSimID_" & uniqueId, True)
        With lblID
            .Caption = opAtual.ID_OP
            .Left = posX + 8
            .Top = topoCard + 4
            .Width = larguraCard - 12
            .Height = 16
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Produto
        Dim lblProd As MSForms.Label
        Set lblProd = pContainer.Controls.Add("Forms.Label.1", "cardSimProd_" & uniqueId, True)
        With lblProd
            .Caption = Left(opAtual.Produto, 20)
            .Left = posX + 8
            .Top = topoCard + 22
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Período
        Dim lblPer As MSForms.Label
        Set lblPer = pContainer.Controls.Add("Forms.Label.1", "cardSimPer_" & uniqueId, True)
        With lblPer
            If IsDate(opAtual.DataInicio) And IsDate(opAtual.DataFim) Then
                .Caption = Format(opAtual.DataInicio, "HH:MM") & " - " & Format(opAtual.DataFim, "HH:MM")
            Else
                .Caption = "--:-- - --:--"
            End If
            .Left = posX + 8
            .Top = topoCard + 38
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Duração
        Dim lblDur As MSForms.Label
        Set lblDur = pContainer.Controls.Add("Forms.Label.1", "cardSimDur_" & uniqueId, True)
        With lblDur
            .Caption = Format(opAtual.Duracao, "0.0") & "h"
            .Left = posX + 8
            .Top = topoCard + 54
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Alerta de conflito
        If conflitoMarcado(i) Then
            Dim lblAlerta As MSForms.Label
            Set lblAlerta = pContainer.Controls.Add("Forms.Label.1", "cardSimAlerta_" & uniqueId, True)
            With lblAlerta
                .Caption = "!"
                .Left = posX + larguraCard - 20
                .Top = topoCard + 4
                .Width = 16
                .Height = 16
                .Font.Size = 10
                .Font.Bold = True
                .ForeColor = vbWhite
                .BackColor = COR_ALERTA
                .TextAlign = fmTextAlignCenter
            End With
        End If
    Next i
    
    ' Linha AGORA
    Dim agora As Date
    agora = Now
    
    If agora >= pDataInicio And agora <= pDataFim Then
        Dim posAgora As Double
        posAgora = ConverterDataParaX(agora, escala)
        
        Dim linhaAgora As MSForms.Label
        Set linhaAgora = pContainer.Controls.Add("Forms.Label.1", "lblSimAgora", True)
        With linhaAgora
            .Left = posAgora
            .Top = topoTimeline
            .Width = 2
            .Height = numPostos * 80 + 40
            .BackColor = COR_ALERTA
        End With
        
        Dim lblAgora As MSForms.Label
        Set lblAgora = pContainer.Controls.Add("Forms.Label.1", "lblSimAgoraTxt", True)
        With lblAgora
            .Caption = "AGORA"
            .Left = posAgora + 4
            .Top = topoTimeline
            .Width = 50
            .Height = 16
            .Font.Size = 8
            .Font.Bold = True
            .ForeColor = COR_ALERTA
            .BackColor = COR_FUNDO
            .TextAlign = fmTextAlignLeft
        End With
    End If
    
    ' Scroll horizontal
    hScrollSimulacao.Max = larguraTimeline + margemEsquerda
    If hScrollSimulacao.Max < pContainer.Width Then hScrollSimulacao.Max = 0
    hScrollSimulacao.LargeChange = pContainer.Width / 2
    hScrollSimulacao.SmallChange = 50
    hScrollSimulacao.Value = 0
    
Sair:
    Exit Sub
    
ErroRenderizar:
    MsgBox "Erro ao renderizar planejamento da simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' EVENTOS DOS BOTÕES ===========================================================
'================================================================================

Private Sub btnNovaSimulacao_Click()
    On Error GoTo ErroNova
    
    Dim nome As String
    nome = InputBox("Nome da simulação:", "Nova Simulação")
    If Trim(nome) = "" Then Exit Sub
    
    Dim idSim As String
    idSim = "SIM_" & Format(Now, "yyyymmdd_hhmmss")
    
    Call SalvarSimulacao(idSim, nome, Now, m_DataInicioPeriodo, m_DataFimPeriodo, "Ativa", "")
    MsgBox "Simulação criada com sucesso!", vbInformation, "APS PURAN"
    
    Call CarregarListaSimulacoes
    
Sair:
    Exit Sub
    
ErroNova:
    MsgBox "Erro ao criar simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnAbrirSimulacao_Click()
    Call AbrirSimulacao
End Sub

Private Sub btnImportarOPs_Click()
    On Error GoTo ErroImportar
    
    If m_ID_SimulacaoAtiva = "" Then
        MsgBox "Selecione uma simulação para importar OPs.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja importar todas as OPs do planejamento para esta simulação?" & vbCrLf & _
                      "As OPs serão copiadas como registros independentes.", _
                      vbQuestion + vbYesNo, "Importar OPs")
    If resposta = vbNo Then Exit Sub
    
    Call ImportarOPsParaSimulacao(m_ID_SimulacaoAtiva, m_DataInicioPeriodo, m_DataFimPeriodo)
    MsgBox "OPs importadas com sucesso!", vbInformation, "APS PURAN"
    
    Call CarregarPlanejamentoSimulacao
    
Sair:
    Exit Sub
    
ErroImportar:
    MsgBox "Erro ao importar OPs: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnExcluirSimulacao_Click()
    On Error GoTo ErroExcluir
    
    If lstSimulacoes.ListIndex = -1 Then
        MsgBox "Selecione uma simulação para excluir.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim chave As String
    chave = Split(lstSimulacoes.List(lstSimulacoes.ListIndex), " - ")(0)
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente excluir a simulação selecionada?" & vbCrLf & _
                      "Todas as OPs da simulação também serão excluídas.", _
                      vbQuestion + vbYesNo, "Confirmação")
    If resposta = vbNo Then Exit Sub
    
    Call ExcluirSimulacao(chave)
    MsgBox "Simulação excluída com sucesso!", vbInformation, "APS PURAN"
    
    m_ID_SimulacaoAtiva = ""
    Call CarregarListaSimulacoes
    Call CarregarPlanejamentoSimulacao
    
Sair:
    Exit Sub
    
ErroExcluir:
    MsgBox "Erro ao excluir simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdAplicarPeriodo_Click()
    On Error GoTo ErroPeriodo
    
    If Not IsDate(txtDataInicial.Text) Or Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe datas válidas.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    m_DataInicioPeriodo = CDate(txtDataInicial.Text)
    m_DataFimPeriodo = CDate(txtDataFinal.Text)
    
    If m_DataFimPeriodo < m_DataInicioPeriodo Then
        MsgBox "Data final não pode ser anterior à data inicial.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    If m_ID_SimulacaoAtiva <> "" Then
        Call CarregarPlanejamentoSimulacao
    End If
    
Sair:
    Exit Sub
    
ErroPeriodo:
    MsgBox "Erro ao aplicar período: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub hScrollSimulacao_Change()
    On Error Resume Next
    Dim offsetX As Single
    offsetX = -hScrollSimulacao.Value
    fraVisualizacao.Left = 430 + offsetX
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    
    If fraVisualizacao Is Nothing Then Exit Sub
    
    Dim ctrl As MSForms.Control
    Dim relX As Single
    Dim relY As Single
    
    relX = X - fraVisualizacao.Left
    relY = Y - fraVisualizacao.Top
    
    For Each ctrl In fraVisualizacao.Controls
        If TypeOf ctrl Is MSForms.Label Then
            If ctrl.Tag <> "" Then
                If relX >= ctrl.Left And relX <= ctrl.Left + ctrl.Width And _
                   relY >= ctrl.Top And relY <= ctrl.Top + ctrl.Height Then
                     If Button = 1 Then
                         Set m_DragCard = ctrl
                         m_DragStartX = X
                         m_DragStartY = Y
                         m_DragOriginalLeft = ctrl.Left
                         m_DragOriginalTop = ctrl.Top
                     ElseIf Button = 2 Then
                        frmDetalheOP.CarregarDetalhesSimulacao ctrl.Tag, m_ID_SimulacaoAtiva
                        frmDetalheOP.Show vbModal
                        Call CarregarPlanejamentoSimulacao
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next ctrl
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    
    If Not m_DragCard Is Nothing And Button = 1 Then
        Dim deltaX As Single
        Dim deltaY As Single
        deltaX = X - m_DragStartX
        deltaY = Y - m_DragStartY
        
        m_DragCard.Left = m_DragOriginalLeft + deltaX
        m_DragCard.Top = m_DragOriginalTop + deltaY
    End If
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErroMouseUp
    
    If m_DragCard Is Nothing Then Exit Sub
    
    Dim deltaX As Single
    Dim deltaY As Single
    deltaX = X - m_DragStartX
    deltaY = Y - m_DragStartY
    
    If Abs(deltaX) < 5 And Abs(deltaY) < 5 Then
        Set m_DragCard = Nothing
        Exit Sub
    End If
    
    Dim novoInicio As Date
    Dim novoFim As Date
    Dim novaDuracao As Double
    Dim indicePosto As Long
    Dim novoEquipamento As String
    
    ' Encontra o OP correspondente ao card
    Dim dados() As Variant
    dados = ObterOPsSimulacaoEmArray(m_ID_SimulacaoAtiva)
    
    If IsError(dados) Then
        Set m_DragCard = Nothing
        Exit Sub
    End If
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    Dim idOP As String
    idOP = m_DragCard.Tag
    
    For i = 1 To totalLinhas
        If CStr(dados(i, 2)) = idOP And CStr(dados(i, 1)) = m_ID_SimulacaoAtiva Then
            Dim dataInicioOP As Date
            Dim dataFimOP As Date
            Dim duracaoOP As Double
            Dim equipamentoAtual As String
            
            If IsDate(dados(i, 7)) Then dataInicioOP = CDate(dados(i, 7))
            If IsDate(dados(i, 8)) Then dataFimOP = CDate(dados(i, 8))
            duracaoOP = CDbl(dados(i, 9))
            equipamentoAtual = CStr(dados(i, 5))
            
            ' Calcula novo horário baseado no movimento horizontal
            novaDuracao = duracaoOP
            novoInicio = dataInicioOP + (deltaX / 100) / 24
            novoFim = dataFimOP + (deltaX / 100) / 24
            
            ' Calcula novo posto baseado no movimento vertical
            Dim colPostos As New Collection
            Dim j As Long
            Dim existe As Boolean
            
            For j = 2 To UBound(dados, 1)
                Dim eq As String
                eq = Trim(CStr(dados(j, 5)))
                If eq <> "" Then
                    existe = False
                    Dim k As Long
                    For k = 1 To colPostos.Count
                        If colPostos(k) = eq Then
                            existe = True
                            Exit For
                        End If
                    Next k
                    If Not existe Then colPostos.Add eq
                End If
            Next j
            
            indicePosto = ObterIndicePosto(equipamentoAtual, colPostos)
            Dim deltaPostos As Long
            deltaPostos = Round(deltaY / ALTURA_LINHA_POSTO)
            
            Dim novoIndicePosto As Long
            novoIndicePosto = indicePosto + deltaPostos
            
            If novoIndicePosto < 0 Then novoIndicePosto = 0
            If novoIndicePosto >= colPostos.Count Then novoIndicePosto = colPostos.Count - 1
            
            novoEquipamento = colPostos(novoIndicePosto + 1)
            
            ' Atualiza Data_Inicio, Data_Fim e Equipamento
            Call AtualizarOPSimulacao(m_ID_SimulacaoAtiva, idOP, CStr(dados(i, 4)), novoEquipamento, CLng(dados(i, 6)), novoInicio, novoFim, novaDuracao, CStr(dados(i, 10)), CStr(dados(i, 11)), CStr(dados(i, 12)))
            
            Exit For
        End If
    Next i
    
    Set m_DragCard = Nothing
    Call CarregarPlanejamentoSimulacao
    
Sair:
    Exit Sub
    
ErroMouseUp:
    Set m_DragCard = Nothing
    Resume Sair
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
