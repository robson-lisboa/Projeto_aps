VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 12632256
  BorderStyle = 3
  Caption = "APS PURAN – Sistema de Planejamento de Produção Industrial"
  ClientHeight = 756
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 1188
  Height = 800
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 1200
  StartUpPosition = 1
  Attribute VB_Name = "frmPrincipal"
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
Attribute VB_Name = "frmPrincipal"
'================================================================================
' USERFORM: frmPrincipal
' DESCRIÇÃO: Interface principal do APS PURAN – Sistema de Planejamento de Produção Industrial
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== API DO WINDOWS (COMPATÍVEL 32/64-BIT) =====================================
#If VBA7 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If

Private Const SW_MINIMIZE As Long = 6
Private Const SW_MAXIMIZE As Long = 3
Private Const SW_RESTORE As Long = 9

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_SIZEBOX As Long = WS_THICKFRAME

'================================================================================
' FUNÇÃO PRIVADA: AplicarEstiloRedimensionavel
' PROPÓSITO: Adicionar estilo WS_THICKFRAME à janela do UserForm para permitir
'            redimensionamento manual pelas bordas e cantos, como uma janela normal.
'================================================================================
Private Sub AplicarEstiloRedimensionavel()
    On Error Resume Next
    
    Dim hWnd As LongPtr
    Dim estilo As LongPtr
    
    #If VBA7 Then
        hWnd = Me.hWnd
    #Else
        hWnd = GetActiveWindow()
    #End If
    
    If hWnd <> 0 Then
        estilo = GetWindowLongPtr(hWnd, GWL_STYLE)
        If estilo <> 0 Then
            estilo = estilo Or WS_THICKFRAME
            SetWindowLongPtr hWnd, GWL_STYLE, estilo
        End If
    End If
End Sub

'=== CONTROLES DO FORMULÁRIO ===================================================
Private btnDashboard As MSForms.CommandButton
Private btnTimeline As MSForms.CommandButton
Private btnCards As MSForms.CommandButton
Private btnProducao As MSForms.CommandButton
Private btnDashboardProducao As MSForms.CommandButton
Private btnEventos As MSForms.CommandButton
Private btnCadastros As MSForms.CommandButton
Private btnConfig As MSForms.CommandButton
Private btnPlanejamento As MSForms.CommandButton
Private btnSimulacao As MSForms.CommandButton
Private btnComparar As MSForms.CommandButton
Private btnCapacidade As MSForms.CommandButton
Private btnRelatorios As MSForms.CommandButton

Private m_btnDashboardEvents As clsButtonEvents
Private m_btnTimelineEvents As clsButtonEvents
Private m_btnCardsEvents As clsButtonEvents
Private m_btnPlanejamentoEvents As clsButtonEvents
Private m_btnProducaoEvents As clsButtonEvents
Private m_btnDashboardProducaoEvents As clsButtonEvents
Private m_btnEventosEvents As clsButtonEvents
Private m_btnCadastrosEvents As clsButtonEvents
Private m_btnConfigEvents As clsButtonEvents
Private m_btnSimulacaoEvents As clsButtonEvents
Private m_btnCompararEvents As clsButtonEvents
Private m_btnCapacidadeEvents As clsButtonEvents
Private m_btnRelatoriosEvents As clsButtonEvents

Private fraMenu As MSForms.Frame
Private fraConteudo As MSForms.Frame
Private lblTitulo As MSForms.Label
Private btnMinimizar As MSForms.CommandButton
Private btnMaximizar As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

' Área de produção (planejamento temporal)
Private fraProducao As MSForms.Frame
Private hScrollProducao As MSForms.ScrollBar

' Controles do cabeçalho de planejamento
Private cmdHoje As MSForms.CommandButton
Private cmdDia As MSForms.CommandButton
Private cmdSemana As MSForms.CommandButton
Private cmdMes As MSForms.CommandButton
Private cmdPersonalizado As MSForms.CommandButton
Private txtDataInicial As MSForms.TextBox
Private txtDataFinal As MSForms.TextBox
Private cmdAplicarPeriodo As MSForms.CommandButton
Private cmdZoomMenos As MSForms.CommandButton
Private cmdZoomMais As MSForms.CommandButton
Private lblZoom As MSForms.Label
Private cmdAtualizar As MSForms.CommandButton
Private cmdAdicionarOP As MSForms.CommandButton
Private cmdSimularProducao As MSForms.CommandButton
Private cmdExportarPlanejamentoExcel As MSForms.CommandButton
Private cmdExportarPlanejamentoPDF As MSForms.CommandButton
Private cmdImprimirPlanejamento As MSForms.CommandButton
Private cboFiltroStatus As MSForms.ComboBox
Private txtBusca As MSForms.TextBox
Private cboOrdenar As MSForms.ComboBox

' Labels do Dashboard (nível de módulo para permitir atualização)
Private lblKPIPlanejadas As MSForms.Label
Private lblKPIConcluidas As MSForms.Label
Private lblKPIEmAndamento As MSForms.Label
Private lblKPIAtrasadas As MSForms.Label
Private lblKPIHorasPlanejadas As MSForms.Label
Private lblKPIHorasRealizadas As MSForms.Label

' Armazena o último painel não-modal para recarregamento após fechar modais
Private m_UltimoPainelNaoModal As String

' Estado de maximização
Private m_Maximizado As Boolean
Private m_TamanhoNormalLargura As Single
Private m_TamanhoNormalAltura As Single
Private m_TamanhoNormalLeft As Single
Private m_TamanhoNormalTop As Single

' Dados dos KPIs para reprocessamento durante resize
Private m_DadosKPIs As Variant
Private m_PainelAtivo As String

'=== PROPRIEDADES VISUAIS CORPORATIVAS =========================================
Private Const COR_FUNDO As Long = 4474559
Private Const COR_PAINEL As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_ALERTA As Long = 255
Private Const COR_AZUL As Long = 15773696
Private Const COR_VERDE As Long = 5287936
Private Const COR_AMARELO As Long = 65535

Private Const LARGURA_BOTAO As Single = 140
Private Const ALTURA_BOTAO As Single = 36
Private Const ESPACAMENTO As Single = 8

'=== ESTADO DO SISTEMA =========================================================
Private m_DataInicioPeriodo As Date
Private m_DataFimPeriodo As Date
Private m_Zoom As Double
Private m_FiltroStatus As String
Private m_TextoBusca As String
Private m_OrdenarPor As String

'================================================================================
' EVENTO: UserForm_Initialize
' PROPÓSITO: Inicializar os controles e exibir o painel inicial (Dashboard)
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Dim i As Long
    Dim posY As Single
    Dim botoes As Variant
    
    ' Configura propriedades visuais do formulário
    Me.BackColor = COR_FUNDO
    m_Maximizado = False
    m_PainelAtivo = ""
    m_DadosKPIs = Empty
    
    '--- Carrega preferências salvas -------------------------------------------
    Call CarregarPreferencias
    
    ' Inicializar período padrão (mês atual)
    m_DataInicioPeriodo = DateSerial(Year(Date), Month(Date), 1)
    m_DataFimPeriodo = DateSerial(Year(Date), Month(Date) + 1, 0)
    m_Zoom = 1.0
    
    ' Aplica estilo de janela redimensionável do Windows
    AplicarEstiloRedimensionavel
    
    '--- Cria Header -----------------------------------------------------------
    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    With lblTitulo
        .Caption = "APS PURAN – Sistema de Planejamento de Produção"
        .Left = 12
        .Top = 12
        .Width = Me.ClientWidth - 24
        .Height = 48
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Cria Botões de Controle da Janela --------------------------------------
    Set btnMinimizar = Me.Controls.Add("Forms.CommandButton.1", "btnMinimizar", True)
    With btnMinimizar
        .Caption = ChrW(8722)
        .Left = Me.ClientWidth - 90
        .Top = 18
        .Width = 24
        .Height = 24
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .BorderStyle = fmBorderStyleNone
    End With
    
    Set btnMaximizar = Me.Controls.Add("Forms.CommandButton.1", "btnMaximizar", True)
    With btnMaximizar
        .Caption = ChrW(9633)
        .Left = Me.ClientWidth - 60
        .Top = 18
        .Width = 24
        .Height = 24
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .BorderStyle = fmBorderStyleNone
    End With
    
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = ChrW(215)
        .Left = Me.ClientWidth - 30
        .Top = 18
        .Width = 24
        .Height = 24
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = vbWhite
        .BackColor = COR_ALERTA
        .BorderStyle = fmBorderStyleNone
    End With
    
    '--- Cria Frame do Menu Lateral --------------------------------------------
    Set fraMenu = Me.Controls.Add("Forms.Frame.1", "fraMenu", True)
    With fraMenu
        .Caption = ""
        .Left = 12
        .Top = 72
        .Width = 160
        .Height = Me.ClientHeight - 96
        .BackColor = COR_PAINEL
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 10
    End With
    
    '--- Cria Botões de Navegação ----------------------------------------------
    botoes = Array("btnDashboard", "btnTimeline", "btnCards", "btnPlanejamento", "btnSimulacao", "btnComparar", "btnCapacidade", "btnRelatorios", "btnProducao", "btnDashboardProducao", "btnEventos", "btnCadastros", "btnConfig")
    posY = 12
    
    For i = LBound(botoes) To UBound(botoes)
        Dim btn As MSForms.CommandButton
        Set btn = Me.Controls.Add("Forms.CommandButton.1", botoes(i), True)
        
        btn.Caption = _
            IIf(i = 0, "Dashboard", _
            IIf(i = 1, "Timeline", _
            IIf(i = 2, "Cards", _
            IIf(i = 3, "Planejamento", _
            IIf(i = 4, "Simulação", _
            IIf(i = 5, "Comparar", _
            IIf(i = 6, "Capacidade", _
            IIf(i = 7, "Relatórios", _
            IIf(i = 8, "Produção", _
            IIf(i = 9, "Dash. Prod.", _
            IIf(i = 10, "Eventos", _
            IIf(i = 11, "Cadastros", _
            IIf(i = 12, "Configurações", ""))))))))))))
        btn.Left = ESPACAMENTO
        btn.Top = posY
        btn.Width = LARGURA_BOTAO - (2 * ESPACAMENTO)
        btn.Height = ALTURA_BOTAO
        btn.Font.Size = 10
        btn.Font.Bold = True
        btn.BackColor = COR_HEADER
        btn.ForeColor = COR_TEXTO_CLARO
        
        posY = posY + ALTURA_BOTAO + ESPACAMENTO
        
        Select Case botoes(i)
            Case "btnDashboard": Set btnDashboard = btn
            Case "btnTimeline": Set btnTimeline = btn
            Case "btnCards": Set btnCards = btn
            Case "btnPlanejamento": Set btnPlanejamento = btn
            Case "btnSimulacao": Set btnSimulacao = btn
            Case "btnComparar": Set btnComparar = btn
            Case "btnCapacidade": Set btnCapacidade = btn
            Case "btnRelatorios": Set btnRelatorios = btn
            Case "btnProducao": Set btnProducao = btn
            Case "btnDashboardProducao": Set btnDashboardProducao = btn
            Case "btnEventos": Set btnEventos = btn
            Case "btnCadastros": Set btnCadastros = btn
            Case "btnConfig": Set btnConfig = btn
        End Select
    Next i
    
    Set m_btnDashboardEvents = New clsButtonEvents
    Set m_btnDashboardEvents.Button = btnDashboard
    
    Set m_btnTimelineEvents = New clsButtonEvents
    Set m_btnTimelineEvents.Button = btnTimeline
    
    Set m_btnCardsEvents = New clsButtonEvents
    Set m_btnCardsEvents.Button = btnCards
    
    Set m_btnPlanejamentoEvents = New clsButtonEvents
    Set m_btnPlanejamentoEvents.Button = btnPlanejamento
    
    Set m_btnProducaoEvents = New clsButtonEvents
    Set m_btnProducaoEvents.Button = btnProducao
    
    Set m_btnDashboardProducaoEvents = New clsButtonEvents
    Set m_btnDashboardProducaoEvents.Button = btnDashboardProducao
    
    Set m_btnEventosEvents = New clsButtonEvents
    Set m_btnEventosEvents.Button = btnEventos
    
    Set m_btnCadastrosEvents = New clsButtonEvents
    Set m_btnCadastrosEvents.Button = btnCadastros
    
    Set m_btnConfigEvents = New clsButtonEvents
    Set m_btnConfigEvents.Button = btnConfig
    
    Set m_btnSimulacaoEvents = New clsButtonEvents
    Set m_btnSimulacaoEvents.Button = btnSimulacao
    
    Set m_btnCompararEvents = New clsButtonEvents
    Set m_btnCompararEvents.Button = btnComparar
    
    Set m_btnCapacidadeEvents = New clsButtonEvents
    Set m_btnCapacidadeEvents.Button = btnCapacidade
    
    Set m_btnRelatoriosEvents = New clsButtonEvents
    Set m_btnRelatoriosEvents.Button = btnRelatorios
    
    '--- Cria Frame de Conteúdo Central -----------------------------------------
    Set fraConteudo = Me.Controls.Add("Forms.Frame.1", "fraConteudo", True)
    With fraConteudo
        .Caption = ""
        .Left = 184
        .Top = 72
        .Width = Me.ClientWidth - 196
        .Height = Me.ClientHeight - 96
        .BackColor = COR_PAINEL
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 2000
        .Font.Size = 10
    End With
    
    '--- Cria controles do Planejamento -----------------------------------------
    Set fraProducao = Me.Controls.Add("Forms.Frame.1", "fraProducao", True)
    With fraProducao
        .Caption = ""
        .Left = 184
        .Top = 72
        .Width = Me.ClientWidth - 196
        .Height = Me.ClientHeight - 96
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
        .Font.Size = 10
    End With
    
    ' Botões de período
    Dim periodoLeft As Single
    periodoLeft = 12
    
    Set cmdHoje = Me.Controls.Add("Forms.CommandButton.1", "cmdHoje", True)
    With cmdHoje
        .Caption = "Hoje"
        .Left = periodoLeft
        .Top = 12
        .Width = 60
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 66
    
    Set cmdDia = Me.Controls.Add("Forms.CommandButton.1", "cmdDia", True)
    With cmdDia
        .Caption = "Dia"
        .Left = periodoLeft
        .Top = 12
        .Width = 50
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 56
    
    Set cmdSemana = Me.Controls.Add("Forms.CommandButton.1", "cmdSemana", True)
    With cmdSemana
        .Caption = "Semana"
        .Left = periodoLeft
        .Top = 12
        .Width = 60
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 66
    
    Set cmdMes = Me.Controls.Add("Forms.CommandButton.1", "cmdMes", True)
    With cmdMes
        .Caption = "Mês"
        .Left = periodoLeft
        .Top = 12
        .Width = 50
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 56
    
    Set cmdPersonalizado = Me.Controls.Add("Forms.CommandButton.1", "cmdPersonalizado", True)
    With cmdPersonalizado
        .Caption = "Personalizado"
        .Left = periodoLeft
        .Top = 12
        .Width = 90
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 96
    
    ' Datas personalizadas
    Set txtDataInicial = Me.Controls.Add("Forms.TextBox.1", "txtDataInicial", True)
    With txtDataInicial
        .Text = Format(m_DataInicioPeriodo, "dd/mm/yyyy")
        .Left = periodoLeft
        .Top = 12
        .Width = 90
        .Height = 22
        .Font.Size = 9
    End With
    periodoLeft = periodoLeft + 96
    
    Set txtDataFinal = Me.Controls.Add("Forms.TextBox.1", "txtDataFinal", True)
    With txtDataFinal
        .Text = Format(m_DataFimPeriodo, "dd/mm/yyyy")
        .Left = periodoLeft
        .Top = 12
        .Width = 90
        .Height = 22
        .Font.Size = 9
    End With
    periodoLeft = periodoLeft + 96
    
    Set cmdAplicarPeriodo = Me.Controls.Add("Forms.CommandButton.1", "cmdAplicarPeriodo", True)
    With cmdAplicarPeriodo
        .Caption = "Aplicar"
        .Left = periodoLeft
        .Top = 12
        .Width = 70
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 76
    
    ' Zoom
    Set cmdZoomMenos = Me.Controls.Add("Forms.CommandButton.1", "cmdZoomMenos", True)
    With cmdZoomMenos
        .Caption = "-"
        .Left = periodoLeft
        .Top = 12
        .Width = 24
        .Height = 24
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 28
    
    Set lblZoom = Me.Controls.Add("Forms.Label.1", "lblZoom", True)
    With lblZoom
        .Caption = "100%"
        .Left = periodoLeft
        .Top = 12
        .Width = 40
        .Height = 24
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignCenter
    End With
    periodoLeft = periodoLeft + 44
    
    Set cmdZoomMais = Me.Controls.Add("Forms.CommandButton.1", "cmdZoomMais", True)
    With cmdZoomMais
        .Caption = "+"
        .Left = periodoLeft
        .Top = 12
        .Width = 24
        .Height = 24
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 28
    
    ' Filtro de status
    Set cboFiltroStatus = Me.Controls.Add("Forms.ComboBox.1", "cboFiltroStatus", True)
    With cboFiltroStatus
        .AddItem "Todos"
        .AddItem "Planejada"
        .AddItem "Em Andamento"
        .AddItem "Concluído"
        .AddItem "Atrasado"
        .Value = "Todos"
        .Left = periodoLeft
        .Top = 12
        .Width = 100
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    periodoLeft = periodoLeft + 106
    
    ' Busca
    Set txtBusca = Me.Controls.Add("Forms.TextBox.1", "txtBusca", True)
    With txtBusca
        .Text = ""
        .Left = periodoLeft
        .Top = 12
        .Width = 120
        .Height = 22
        .Font.Size = 9
        .PlaceholderText = "Buscar OP..."
    End With
    periodoLeft = periodoLeft + 126
    
    ' Ordenação
    Set cboOrdenar = Me.Controls.Add("Forms.ComboBox.1", "cboOrdenar", True)
    With cboOrdenar
        .AddItem "Ordenar por..."
        .AddItem "ID_OP"
        .AddItem "Data Início"
        .AddItem "Duração"
        .AddItem "Status"
        .Value = "Ordenar por..."
        .Left = periodoLeft
        .Top = 12
        .Width = 110
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    periodoLeft = periodoLeft + 116
    
    ' Ações
    Set cmdAtualizar = Me.Controls.Add("Forms.CommandButton.1", "cmdAtualizar", True)
    With cmdAtualizar
        .Caption = "Atualizar"
        .Left = periodoLeft
        .Top = 12
        .Width = 80
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 86
    
    Set cmdAdicionarOP = Me.Controls.Add("Forms.CommandButton.1", "cmdAdicionarOP", True)
    With cmdAdicionarOP
        .Caption = "Nova OP"
        .Left = periodoLeft
        .Top = 12
        .Width = 80
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 86
    
    Set cmdSimularProducao = Me.Controls.Add("Forms.CommandButton.1", "cmdSimularProducao", True)
    With cmdSimularProducao
        .Caption = "Simular"
        .Left = periodoLeft
        .Top = 12
        .Width = 80
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_AMARELO
        .ForeColor = COR_TEXTO_ESCURO
    End With
    periodoLeft = periodoLeft + 86
    
    Set cmdExportarPlanejamentoExcel = Me.Controls.Add("Forms.CommandButton.1", "cmdExportarPlanejamentoExcel", True)
    With cmdExportarPlanejamentoExcel
        .Caption = "Exportar Excel"
        .Left = periodoLeft
        .Top = 12
        .Width = 80
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    periodoLeft = periodoLeft + 86
    
    Set cmdExportarPlanejamentoPDF = Me.Controls.Add("Forms.CommandButton.1", "cmdExportarPlanejamentoPDF", True)
    With cmdExportarPlanejamentoPDF
        .Caption = "Exportar PDF"
        .Left = periodoLeft
        .Top = 12
        .Width = 80
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_AMARELO
        .ForeColor = COR_TEXTO_ESCURO
    End With
    periodoLeft = periodoLeft + 86
    
    Set cmdImprimirPlanejamento = Me.Controls.Add("Forms.CommandButton.1", "cmdImprimirPlanejamento", True)
    With cmdImprimirPlanejamento
        .Caption = "Imprimir"
        .Left = periodoLeft
        .Top = 12
        .Width = 80
        .Height = 24
        .Font.Size = 9
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    ' Scroll horizontal da área de produção
    Set hScrollProducao = Me.Controls.Add("Forms.ScrollBar.1", "hScrollProducao", True)
    With hScrollProducao
        .Left = 12
        .Top = Me.ClientHeight - 96 - 20
        .Width = Me.ClientWidth - 24
        .Height = 16
        .Min = 0
        .Max = 100
        .SmallChange = 10
        .LargeChange = 50
        .Value = 0
    End With
    
    ' Exibe o painel inicial
    m_UltimoPainelNaoModal = "Dashboard"
    ExibirPainel "Dashboard"
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar interface: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Interface"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA: ExibirPainel
' PROPÓSITO: Alterar dinamicamente o conteúdo do frame central e título do form
' PARÂMETROS: NomePainel As String – nome do painel a ser exibido
'================================================================================
Public Sub ExibirPainel(NomePainel As String)
    On Error GoTo ErroExibirPainel
    
    ' Limpa controles anteriores dentro do frame de conteúdo
    Dim ctrl As MSForms.Control
    For Each ctrl In fraConteudo.Controls
        fraConteudo.Controls.Remove ctrl.Name
    Next ctrl
    
    ' Atualiza título do formulário
    Me.Caption = "APS PURAN – Sistema de Planejamento de Produção Industrial" & _
                 " | Módulo: " & NomePainel
    
    ' Armazena o último painel não-modal para recuperação após modais
    If NomePainel <> "Produção" And NomePainel <> "Eventos" Then
        m_UltimoPainelNaoModal = NomePainel
    End If
    
    m_PainelAtivo = NomePainel
    
    ' Mostra/esconde área de planejamento conforme módulo
    If NomePainel = "Planejamento" Then
        If Not fraProducao Is Nothing Then fraProducao.Visible = True
        If Not hScrollProducao Is Nothing Then hScrollProducao.Visible = True
        If Not fraConteudo Is Nothing Then fraConteudo.Visible = False
    Else
        If Not fraProducao Is Nothing Then fraProducao.Visible = False
        If Not hScrollProducao Is Nothing Then hScrollProducao.Visible = False
        If Not fraConteudo Is Nothing Then fraConteudo.Visible = True
    End If
    
    ' Direciona para a rotina específica de cada módulo
    Select Case NomePainel
        Case "Dashboard"
            CarregarDashboard
        Case "Timeline"
            modTimeline.CarregarTimeline fraConteudo
        Case "Cards"
            modCards.CarregarCards fraConteudo
        Case "Planejamento"
            CarregarPlanejamento
        Case "Simulação"
            frmSimulacao.Show vbModeless
        Case "Produção"
            frmProducao.Show vbModeless
        Case "Eventos"
            frmEventos.Show vbModal
            ExibirPainel m_UltimoPainelNaoModal
        Case "Cadastros"
            frmCadastros.Show vbModeless
        Case Else
            ExibirPainelGenerico NomePainel
    End Select
    
Sair:
    Exit Sub
    
ErroExibirPainel:
    MsgBox "Erro ao exibir painel: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Navegação"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: ExibirPainelGenerico
' PROPÓSITO: Exibir painéis placeholder (Configurações)
'================================================================================
Private Sub ExibirPainelGenerico(NomePainel As String)
    On Error GoTo ErroGenerico
    
    Dim lblPainel As MSForms.Label
    
    Set lblPainel = fraConteudo.Controls.Add("Forms.Label.1", "lblPainel_" & NomePainel, True)
    With lblPainel
        .Caption = "Painel Ativo: " & NomePainel
        .Left = 20
        .Top = 20
        .Width = fraConteudo.Width - 40
        .Height = 400
        .Font.Size = 16
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_PAINEL
        .TextAlign = fmTextAlignCenter
    End With
    
Sair:
    Exit Sub
    
ErroGenerico:
    MsgBox "Erro ao exibir painel genérico: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Navegação"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarPlanejamento
' PROPÓSITO: Renderizar a visualização de planejamento (cards flutuantes) no frame
'================================================================================
Public Sub CarregarPlanejamento()
    On Error GoTo ErroCarregarPlanejamento
    
    fraProducao.Visible = True
    fraProducao.Left = 184
    fraProducao.Top = 72
    fraProducao.Width = Me.ClientWidth - 196
    fraProducao.Height = Me.ClientHeight - 96
    
    hScrollProducao.Visible = True
    hScrollProducao.Left = 184
    hScrollProducao.Top = Me.ClientHeight - 96 - 20
    hScrollProducao.Width = Me.ClientWidth - 196
    
    Call modGantt.CarregarGantt(fraProducao, hScrollProducao, m_DataInicioPeriodo, m_DataFimPeriodo, m_Zoom, m_FiltroStatus, m_TextoBusca, m_OrdenarPor)
    
Sair:
    Exit Sub
    
ErroCarregarPlanejamento:
    MsgBox "Erro ao carregar planejamento: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Planejamento"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarDashboard
' PROPÓSITO: Renderizar indicadores de desempenho (KPIs) no frame de conteúdo
'================================================================================
Public Sub CarregarDashboard()
    On Error GoTo ErroCarregarDashboard
    
    Dim dados() As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    Dim totalPlanejadas As Long
    Dim totalConcluidas As Long
    Dim totalEmAndamento As Long
    Dim totalAtrasadas As Long
    Dim totalHorasPlanejadas As Double
    Dim totalHorasRealizadas As Double
    
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then
        Err.Raise vbObjectError + 200, "CarregarDashboard", _
            "Não foi possível carregar dados da TabelaOPs."
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        Dim statusAtual As String
        statusAtual = CStr(dados(i, 8))
        
        Select Case Trim(statusAtual)
            Case "Concluído"
                totalConcluidas = totalConcluidas + 1
            Case "Em Andamento"
                totalEmAndamento = totalEmAndamento + 1
            Case "Atrasado"
                totalAtrasadas = totalAtrasadas + 1
            Case Else
                totalPlanejadas = totalPlanejadas + 1
        End Select
        
        totalHorasPlanejadas = totalHorasPlanejadas + CDbl(dados(i, 7))
        totalHorasRealizadas = totalHorasRealizadas + CDbl(dados(i, 7))
    Next i
    
    totalPlanejadas = totalPlanejadas + totalConcluidas + totalEmAndamento + totalAtrasadas
    
    ' Armazena dados para reprocessamento durante resize
    m_DadosKPIs = Array(totalPlanejadas, totalConcluidas, totalEmAndamento, totalAtrasadas, totalHorasPlanejadas, totalHorasRealizadas)
    m_PainelAtivo = "Dashboard"
    
    RenderizarKPIs
    
Sair:
    Exit Sub
    
ErroCarregarDashboard:
    MsgBox "Erro ao carregar Dashboard: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Dashboard"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: RenderizarKPIs
' PROPÓSITO: Recriar todos os controles do Dashboard com base nos dados armazenados
'================================================================================
Private Sub RenderizarKPIs()
    On Error GoTo ErroRenderizar
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraConteudo.Controls
        fraConteudo.Controls.Remove ctrl.Name
    Next ctrl
    
    If IsEmpty(m_DadosKPIs) Then Exit Sub
    
    Dim titulo As MSForms.Label
    Set titulo = fraConteudo.Controls.Add("Forms.Label.1", "lblDashTitulo", True)
    With titulo
        .Caption = "INDICADORES DE PRODUÇÃO – DASHBOARD"
        .Left = 20
        .Top = 20
        .Width = fraConteudo.Width - 40
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_PAINEL
        .TextAlign = fmTextAlignCenter
    End With
    
    AjustarLayoutKPIs
    
Sair:
    Exit Sub
    
ErroRenderizar:
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: AjustarLayoutKPIs
' PROPÓSITO: Recalcular posições e larguras das colunas de KPI conforme espaço disponível
'================================================================================
Private Sub AjustarLayoutKPIs()
    On Error GoTo ErroAjustar
    
    If IsEmpty(m_DadosKPIs) Then Exit Sub
    
    Dim larguraDisponivel As Single
    larguraDisponivel = fraConteudo.Width - 40
    
    Dim COLUNA1_X As Single
    Dim COLUNA2_X As Single
    Dim LARGURA_KPI As Single
    Dim LARGURA_KPI_MENOR As Single
    Dim LINHA_INICIO_Y As Single
    Dim ALTURA_KPI As Single
    Dim ESPACO_Y As Single
    Dim espacoEntreColunas As Single
    
    COLUNA1_X = 20
    espacoEntreColunas = 40
    LINHA_INICIO_Y = 80
    ALTURA_KPI = 100
    ESPACO_Y = 140
    
    ' Calcula largura da coluna maior: 65% do espaço disponível
    LARGURA_KPI = larguraDisponivel * 0.65
    If LARGURA_KPI < 300 Then LARGURA_KPI = 300
    If LARGURA_KPI > 700 Then LARGURA_KPI = 700
    
    ' Calcula largura da coluna menor: 30% do espaço disponível
    LARGURA_KPI_MENOR = larguraDisponivel * 0.30
    If LARGURA_KPI_MENOR < 200 Then LARGURA_KPI_MENOR = 200
    If LARGURA_KPI_MENOR > 400 Then LARGURA_KPI_MENOR = 400
    
    ' Garante que as duas colunas caibam
    If COLUNA1_X + LARGURA_KPI + espacoEntreColunas + LARGURA_KPI_MENOR > larguraDisponivel Then
        LARGURA_KPI = (larguraDisponivel - espacoEntreColunas) * 0.65
        LARGURA_KPI_MENOR = (larguraDisponivel - espacoEntreColunas) * 0.35
    End If
    
    COLUNA2_X = COLUNA1_X + LARGURA_KPI + espacoEntreColunas
    
    ' Ajusta espaçamento vertical se altura for pequena
    Dim alturaDisponivel As Single
    alturaDisponivel = fraConteudo.Height - LINHA_INICIO_Y - 40
    If alturaDisponivel < 500 Then
        ESPACO_Y = alturaDisponivel / 3
    End If
    
    Dim totalPlanejadas As Long
    Dim totalConcluidas As Long
    Dim totalEmAndamento As Long
    Dim totalAtrasadas As Long
    Dim totalHorasPlanejadas As Double
    Dim totalHorasRealizadas As Double
    
    totalPlanejadas = m_DadosKPIs(0)
    totalConcluidas = m_DadosKPIs(1)
    totalEmAndamento = m_DadosKPIs(2)
    totalAtrasadas = m_DadosKPIs(3)
    totalHorasPlanejadas = m_DadosKPIs(4)
    totalHorasRealizadas = m_DadosKPIs(5)
    
    Set lblKPIPlanejadas = CriaKPI(COLUNA1_X, LINHA_INICIO_Y, _
        "Total de Produções Planejadas", CStr(totalPlanejadas), COR_TEXTO_ESCURO, LARGURA_KPI)
    
    Set lblKPIConcluidas = CriaKPI(COLUNA2_X, LINHA_INICIO_Y, _
        "Produções Concluídas", CStr(totalConcluidas), vbGreen, LARGURA_KPI_MENOR)
    
    Set lblKPIEmAndamento = CriaKPI(COLUNA1_X, LINHA_INICIO_Y + ESPACO_Y, _
        "Produções em Andamento", CStr(totalEmAndamento), COR_HEADER, LARGURA_KPI)
    
    Dim corAtrasos As Long
    If totalAtrasadas > 0 Then
        corAtrasos = COR_ALERTA
    Else
        corAtrasos = COR_TEXTO_ESCURO
    End If
    
    Set lblKPIAtrasadas = CriaKPI(COLUNA2_X, LINHA_INICIO_Y + ESPACO_Y, _
        "Produções Atrasadas", CStr(totalAtrasadas), corAtrasos, LARGURA_KPI_MENOR)
    
    Set lblKPIHorasPlanejadas = CriaKPI(COLUNA1_X, LINHA_INICIO_Y + 2 * ESPACO_Y, _
        "Total de Horas Planejadas", Format(totalHorasPlanejadas, "0.00"), COR_TEXTO_ESCURO, LARGURA_KPI)
    
    Set lblKPIHorasRealizadas = CriaKPI(COLUNA2_X, LINHA_INICIO_Y + 2 * ESPACO_Y, _
        "Total de Horas Realizadas", Format(totalHorasRealizadas, "0.00"), COR_TEXTO_ESCURO, LARGURA_KPI_MENOR)
    
Sair:
    Exit Sub
    
ErroAjustar:
    Resume Sair
End Sub

'================================================================================
' FUNÇÃO PRIVADA: CriaKPI
' PROPÓSITO: Fábrica de controles Label estilizados para exibição de indicadores
'================================================================================
Private Function CriaKPI(pLeft As Single, pTop As Single, _
                         pTitulo As String, pValor As String, _
                         pCorValor As Long, _
                         Optional pLargura As Single = 0) As MSForms.Label
    Dim lbl As MSForms.Label
    Dim lblValor As MSForms.Label
    Dim lblTituloKPI As MSForms.Label
    
    Dim uniqueID As String
    uniqueID = Format(Now, "SSSSS") & "_" & CStr(Int(Rnd * 100000))
    
    Dim larguraUsada As Single
    larguraUsada = IIf(pLargura > 0, pLargura, 300)
    
    Set lbl = fraConteudo.Controls.Add("Forms.Label.1", "lblKPI_" & uniqueID, True)
    With lbl
        .Left = pLeft
        .Top = pTop
        .Width = larguraUsada
        .Height = 100
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set lblTituloKPI = fraConteudo.Controls.Add("Forms.Label.1", "lblKPI_" & uniqueID & "_T", True)
    With lblTituloKPI
        .Caption = pTitulo
        .Left = pLeft + 10
        .Top = pTop + 10
        .Width = larguraUsada - 20
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
        .TextAlign = fmTextAlignCenter
    End With
    
    Set lblValor = fraConteudo.Controls.Add("Forms.Label.1", "lblKPI_" & uniqueID & "_V", True)
    With lblValor
        .Caption = pValor
        .Left = pLeft + 10
        .Top = pTop + 45
        .Width = larguraUsada - 20
        .Height = 40
        .Font.Size = 20
        .Font.Bold = True
        .ForeColor = pCorValor
        .BackColor = vbWhite
        .TextAlign = fmTextAlignCenter
    End With
    
    Set CriaKPI = lbl
End Function

'================================================================================
' EVENTOS DOS BOTÕES DE NAVEGAÇÃO
'================================================================================

Private Sub btnDashboard_Click()
    ExibirPainel "Dashboard"
End Sub

Private Sub btnTimeline_Click()
    ExibirPainel "Timeline"
End Sub

Private Sub btnCards_Click()
    ExibirPainel "Cards"
End Sub

Private Sub btnProducao_Click()
    frmProducao.Show vbModeless
End Sub

Private Sub btnDashboardProducao_Click()
    frmDashboardProducao.Show vbModeless
End Sub

Private Sub btnEventos_Click()
    ExibirPainel "Eventos"
End Sub

Private Sub btnConfig_Click()
    frmConfiguracoes.Show vbModeless
End Sub

Private Sub btnCadastros_Click()
    ExibirPainel "Cadastros"
End Sub

Private Sub btnSimulacao_Click()
    ExibirPainel "Simulação"
End Sub

Private Sub btnComparar_Click()
    frmComparacaoSimulacao.Show vbModeless
End Sub

Private Sub btnCapacidade_Click()
    frmCapacidade.Show vbModeless
End Sub

Private Sub btnRelatorios_Click()
    frmRelatorios.Show vbModeless
End Sub

Private Sub m_btnSimulacaoEvents_Clicked()
    btnSimulacao_Click
End Sub

Private Sub m_btnCompararEvents_Clicked()
    btnComparar_Click
End Sub

Private Sub m_btnCapacidadeEvents_Clicked()
    btnCapacidade_Click
End Sub

Private Sub m_btnRelatoriosEvents_Clicked()
    btnRelatorios_Click
End Sub

Private Sub m_btnCadastrosEvents_Clicked()
    btnCadastros_Click
End Sub

Private Sub m_btnDashboardEvents_Clicked()
    btnDashboard_Click
End Sub

Private Sub m_btnTimelineEvents_Clicked()
    btnTimeline_Click
End Sub

Private Sub m_btnCardsEvents_Clicked()
    btnCards_Click
End Sub

Private Sub m_btnProducaoEvents_Clicked()
    btnProducao_Click
End Sub

Private Sub m_btnDashboardProducaoEvents_Clicked()
    btnDashboardProducao_Click
End Sub

Private Sub m_btnEventosEvents_Clicked()
    btnEventos_Click
End Sub

Private Sub m_btnConfigEvents_Clicked()
    btnConfig_Click
End Sub

Private Sub btnMinimizar_Click()
    On Error Resume Next
    Dim hWnd As LongPtr
    #If VBA7 Then
        hWnd = Me.hWnd
    #Else
        hWnd = GetActiveWindow()
    #End If
    If hWnd <> 0 Then
        ShowWindow hWnd, SW_MINIMIZE
    End If
End Sub

Private Sub btnMaximizar_Click()
    On Error Resume Next
    Dim hWnd As LongPtr
    #If VBA7 Then
        hWnd = Me.hWnd
    #Else
        hWnd = GetActiveWindow()
    #End If
    If hWnd <> 0 Then
        If m_Maximizado Then
            ShowWindow hWnd, SW_RESTORE
            Me.Width = m_TamanhoNormalLargura
            Me.Height = m_TamanhoNormalAltura
            Me.Left = m_TamanhoNormalLeft
            Me.Top = m_TamanhoNormalTop
            m_Maximizado = False
            btnMaximizar.Caption = ChrW(9633)
        Else
            m_TamanhoNormalLargura = Me.Width
            m_TamanhoNormalAltura = Me.Height
            m_TamanhoNormalLeft = Me.Left
            m_TamanhoNormalTop = Me.Top
            ShowWindow hWnd, SW_MAXIMIZE
            m_Maximizado = True
            btnMaximizar.Caption = ChrW(9634)
        End If
    End If
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    UserForm_QueryClose 0, vbFormControlMenu
End Sub

Private Sub btnPlanejamento_Click()
    ExibirPainel "Planejamento"
End Sub

Private Sub m_btnPlanejamentoEvents_Clicked()
    btnPlanejamento_Click
End Sub

Private Sub cmdHoje_Click()
    m_DataInicioPeriodo = DateSerial(Year(Now), Month(Now), Day(Now))
    m_DataFimPeriodo = DateSerial(Year(Now), Month(Now), Day(Now)) + 1
    Call CarregarPlanejamento
End Sub

Private Sub cmdDia_Click()
    m_DataInicioPeriodo = DateSerial(Year(Now), Month(Now), Day(Now))
    m_DataFimPeriodo = DateSerial(Year(Now), Month(Now), Day(Now)) + 1
    Call CarregarPlanejamento
End Sub

Private Sub cmdSemana_Click()
    Dim diaSemana As Long
    diaSemana = Weekday(Now, vbMonday)
    m_DataInicioPeriodo = DateAdd("d", -(diaSemana - 1), DateSerial(Year(Now), Month(Now), Day(Now)))
    m_DataFimPeriodo = DateAdd("d", 7, m_DataInicioPeriodo)
    Call CarregarPlanejamento
End Sub

Private Sub cmdMes_Click()
    m_DataInicioPeriodo = DateSerial(Year(Now), Month(Now), 1)
    m_DataFimPeriodo = DateSerial(Year(Now), Month(Now) + 1, 0)
    Call CarregarPlanejamento
End Sub

Private Sub cmdPersonalizado_Click()
    On Error GoTo ErroPersonalizado
    
    If Not IsDate(txtDataInicial.Text) Or Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe datas válidas para o período personalizado.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    m_DataInicioPeriodo = CDate(txtDataInicial.Text)
    m_DataFimPeriodo = CDate(txtDataFinal.Text)
    
    If m_DataFimPeriodo < m_DataInicioPeriodo Then
        MsgBox "A data final não pode ser anterior à data inicial.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Call CarregarPlanejamento
    
Sair:
    Exit Sub
    
ErroPersonalizado:
    MsgBox "Erro ao aplicar período personalizado: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub cmdAplicarPeriodo_Click()
    On Error GoTo ErroAplicarPeriodo
    
    If Not IsDate(txtDataInicial.Text) Or Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe datas válidas para o período personalizado.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    m_DataInicioPeriodo = CDate(txtDataInicial.Text)
    m_DataFimPeriodo = CDate(txtDataFinal.Text)
    
    If m_DataFimPeriodo < m_DataInicioPeriodo Then
        MsgBox "A data final não pode ser anterior à data inicial.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Call CarregarPlanejamento
    
Sair:
    Exit Sub
    
ErroAplicarPeriodo:
    MsgBox "Erro ao aplicar período personalizado: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub cmdZoomMenos_Click()
    m_Zoom = m_Zoom - 0.25
    If m_Zoom < 0.5 Then m_Zoom = 0.5
    lblZoom.Caption = Format(m_Zoom, "0%")
    Call CarregarPlanejamento
End Sub

Private Sub cmdZoomMais_Click()
    m_Zoom = m_Zoom + 0.25
    If m_Zoom > 3.0 Then m_Zoom = 3.0
    lblZoom.Caption = Format(m_Zoom, "0%")
    Call CarregarPlanejamento
End Sub

Private Sub cmdAtualizar_Click()
    Call CarregarPlanejamento
End Sub

Private Sub cmdAdicionarOP_Click()
    frmCadastroOP.Show vbModal
    If m_PainelAtivo = "Planejamento" Then
        Call CarregarPlanejamento
    End If
End Sub

Private Sub cmdSimularProducao_Click()
    Call CarregarPlanejamento
End Sub

Private Sub cmdExportarPlanejamentoExcel_Click()
    On Error Resume Next
    Call ExportarPlanejamentoExcel(m_DataInicioPeriodo, m_DataFimPeriodo, m_FiltroStatus, m_TextoBusca, m_OrdenarPor)
End Sub

Private Sub cmdExportarPlanejamentoPDF_Click()
    On Error Resume Next
    Call ExportarPlanejamentoPDF(m_DataInicioPeriodo, m_DataFimPeriodo, m_FiltroStatus, m_TextoBusca, m_OrdenarPor)
End Sub

Private Sub cmdImprimirPlanejamento_Click()
    On Error Resume Next
    Call ImprimirPlanejamento(m_DataInicioPeriodo, m_DataFimPeriodo, m_FiltroStatus, m_TextoBusca, m_OrdenarPor)
End Sub

Private Sub cboFiltroStatus_Change()
    m_FiltroStatus = cboFiltroStatus.Value
    Call CarregarPlanejamento
End Sub

Private Sub txtBusca_Change()
    m_TextoBusca = Trim(txtBusca.Text)
    Call CarregarPlanejamento
End Sub

Private Sub cboOrdenar_Change()
    If cboOrdenar.Value <> "Ordenar por..." Then
        m_OrdenarPor = cboOrdenar.Value
    Else
        m_OrdenarPor = ""
    End If
    Call CarregarPlanejamento
End Sub

Private Sub hScrollProducao_Change()
    On Error Resume Next
    Dim offsetX As Single
    offsetX = -hScrollProducao.Value
    fraProducao.Left = 184 + offsetX
End Sub

'================================================================================
' SUBROTINAS DE PERSISTÊNCIA DE PREFERÊNCIAS
'================================================================================

Private Sub CarregarPreferencias()
    On Error Resume Next
    
    Dim config As Object
    Set config = CarregarTodasConfiguracoes
    
    If config.Count = 0 Then Exit Sub
    
    ' Período
    If config.Exists("DataInicioPlanejamento") Then
        If IsDate(config("DataInicioPlanejamento")) Then
            m_DataInicioPeriodo = CDate(config("DataInicioPlanejamento"))
        End If
    End If
    
    If config.Exists("DataFimPlanejamento") Then
        If IsDate(config("DataFimPlanejamento")) Then
            m_DataFimPeriodo = CDate(config("DataFimPlanejamento"))
        End If
    End If
    
    ' Zoom
    If config.Exists("ZoomPlanejamento") Then
        Dim z As Double
        z = CDbl(config("ZoomPlanejamento"))
        If z >= 0.5 And z <= 3# Then m_Zoom = z
    End If
    
    ' Filtros
    If config.Exists("FiltroStatusPlanejamento") Then
        m_FiltroStatus = CStr(config("FiltroStatusPlanejamento"))
    End If
    
    If config.Exists("TextoBuscaPlanejamento") Then
        m_TextoBusca = CStr(config("TextoBuscaPlanejamento"))
    End If
    
    If config.Exists("OrdenarPorPlanejamento") Then
        m_OrdenarPor = CStr(config("OrdenarPorPlanejamento"))
    End If
    
    ' Painel ativo
    If config.Exists("UltimoPainelPrincipal") Then
        m_PainelAtivo = CStr(config("UltimoPainelPrincipal"))
    End If
    
    On Error GoTo 0
End Sub

Private Sub SalvarPreferencias()
    On Error Resume Next
    
    Call SalvarConfiguracao("DataInicioPlanejamento", Format(m_DataInicioPeriodo, "yyyy-mm-dd"))
    Call SalvarConfiguracao("DataFimPlanejamento", Format(m_DataFimPeriodo, "yyyy-mm-dd"))
    Call SalvarConfiguracao("ZoomPlanejamento", CStr(m_Zoom))
    Call SalvarConfiguracao("FiltroStatusPlanejamento", m_FiltroStatus)
    Call SalvarConfiguracao("TextoBuscaPlanejamento", m_TextoBusca)
    Call SalvarConfiguracao("OrdenarPorPlanejamento", m_OrdenarPor)
    Call SalvarConfiguracao("UltimoPainelPrincipal", m_PainelAtivo)
    
    On Error GoTo 0
End Sub

'================================================================================
' EVENTO: UserForm_QueryClose
' PROPÓSITO: Confirmação de saída segura e restauração do Excel
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Call SalvarPreferencias
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Deseja realmente sair do APS PURAN?", _
                          vbQuestion + vbYesNo, "Confirmação de Saída")
        If resposta = vbNo Then
            Cancel = True
        Else
            Application.Visible = True
        End If
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyF5 Then
        If m_PainelAtivo = "Planejamento" Then
            Call CarregarPlanejamento
        ElseIf m_PainelAtivo = "Dashboard" Then
            Call CarregarDashboard
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
        If m_PainelAtivo = "Planejamento" Then
            If Not txtBusca Is Nothing Then txtBusca.SetFocus
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN Then
        If m_PainelAtivo = "Planejamento" Then
            Call cmdAdicionarOP_Click
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If m_PainelAtivo = "Planejamento" Then
            Call cmdExportarPlanejamentoExcel_Click
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
        If m_PainelAtivo = "Planejamento" Then
            Call cmdImprimirPlanejamento_Click
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyEscape Then
        If m_PainelAtivo = "Planejamento" Then
            m_FiltroStatus = "Todos"
            m_TextoBusca = ""
            m_OrdenarPor = ""
            If Not cboFiltroStatus Is Nothing Then cboFiltroStatus.Value = "Todos"
            If Not txtBusca Is Nothing Then txtBusca.Text = ""
            If Not cboOrdenar Is Nothing Then cboOrdenar.Value = "Ordenar por..."
            Call CarregarPlanejamento
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyAdd Or KeyCode = vbKeyOemplus Then
        If m_PainelAtivo = "Planejamento" Then
            m_Zoom = m_Zoom + 0.1
            If m_Zoom > 3# Then m_Zoom = 3#
            lblZoom.Caption = Format(m_Zoom, "0%")
            Call CarregarPlanejamento
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeySubtract Or KeyCode = vbKeyOemMinus Then
        If m_PainelAtivo = "Planejamento" Then
            m_Zoom = m_Zoom - 0.1
            If m_Zoom < 0.5 Then m_Zoom = 0.5
            lblZoom.Caption = Format(m_Zoom, "0%")
            Call CarregarPlanejamento
        End If
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    
    Dim menuLargura As Single
    Dim margem As Single
    
    menuLargura = 160
    margem = 12
    
    ' Limita tamanho mínimo para não quebrar o layout
    If Me.ClientWidth < 600 Then Me.ClientWidth = 600
    If Me.ClientHeight < 500 Then Me.ClientHeight = 500
    
    fraMenu.Height = Me.ClientHeight - 72 - margem
    
    ' Garante que o painel de conteúdo nunca invada a área do menu lateral
    fraConteudo.Left = margem + menuLargura + margem
    fraConteudo.Width = Me.ClientWidth - fraConteudo.Left - margem
    If fraConteudo.Width < 400 Then fraConteudo.Width = 400
    fraConteudo.Height = fraMenu.Height
    lblTitulo.Width = Me.ClientWidth - margem * 2
    
    ' Reposiciona botões de controle da janela
    If Not btnMinimizar Is Nothing Then
        btnMinimizar.Left = Me.ClientWidth - 90
        btnMaximizar.Left = Me.ClientWidth - 60
        btnFechar.Left = Me.ClientWidth - 30
    End If
    
    ' Ajusta layout dos KPIs se houver dados carregados
    If m_PainelAtivo = "Dashboard" And Not IsEmpty(m_DadosKPIs) Then
        RenderizarKPIs
    End If
    
    ' Ajusta área de planejamento se visível
    If m_PainelAtivo = "Planejamento" Then
        If Not fraProducao Is Nothing Then
            fraProducao.Width = Me.ClientWidth - 196
            fraProducao.Height = Me.ClientHeight - 96
        End If
        If Not hScrollProducao Is Nothing Then
            hScrollProducao.Width = Me.ClientWidth - 196
            hScrollProducao.Top = Me.ClientHeight - 96 - 20
        End If
        Call CarregarPlanejamento
    End If
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    
    If m_PainelAtivo <> "Planejamento" Then Exit Sub
    If fraProducao Is Nothing Then Exit Sub
    
    Dim ctrl As MSForms.Control
    Dim relX As Single
    Dim relY As Single
    
    relX = X - fraProducao.Left
    relY = Y - fraProducao.Top
    
    For Each ctrl In fraProducao.Controls
        If TypeOf ctrl Is MSForms.Label Then
            If ctrl.Tag <> "" Then
                If relX >= ctrl.Left And relX <= ctrl.Left + ctrl.Width And _
                   relY >= ctrl.Top And relY <= ctrl.Top + ctrl.Height Then
                    frmDetalheOP.CarregarDetalhes ctrl.Tag
                    frmDetalheOP.Show vbModal
                    Call CarregarPlanejamento
                    Exit Sub
                End If
            End If
        End If
    Next ctrl
End Sub
