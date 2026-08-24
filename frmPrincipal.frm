VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 12632256
  BorderStyle = 1
  Caption = "APS PURAN – Sistema de Planejamento de Produção Industrial"
  ClientHeight = 620
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 980
  Height = 664
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 992
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

'=== CONTROLES DO FORMULÁRIO ===================================================
Private WithEvents btnDashboard As MSForms.CommandButton
Private WithEvents btnTimeline As MSForms.CommandButton
Private WithEvents btnCards As MSForms.CommandButton
Private WithEvents btnProducao As MSForms.CommandButton
Private WithEvents btnEventos As MSForms.CommandButton
Private WithEvents btnConfig As MSForms.CommandButton

Private fraMenu As MSForms.Frame
Private fraConteudo As MSForms.Frame
Private lblTitulo As MSForms.Label

' Labels do Dashboard (nível de módulo para permitir atualização)
Private lblKPIPlanejadas As MSForms.Label
Private lblKPIConcluidas As MSForms.Label
Private lblKPIEmAndamento As MSForms.Label
Private lblKPIAtrasadas As MSForms.Label
Private lblKPIHorasPlanejadas As MSForms.Label
Private lblKPIHorasRealizadas As MSForms.Label

'=== PROPRIEDADES VISUAIS CORPORATIVAS =========================================
Private Const COR_FUNDO As Long = 4474559
Private Const COR_PAINEL As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_ALERTA As Long = 255

Private Const LARGURA_BOTAO As Single = 140
Private Const ALTURA_BOTAO As Single = 36
Private Const ESPACAMENTO As Single = 8

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
    
    '--- Cria Header -----------------------------------------------------------
    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    With lblTitulo
        .Caption = "APS PURAN – Sistema de Planejamento de Produção"
        .Left = 12
        .Top = 12
        .Width = 960
        .Height = 48
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Cria Frame do Menu Lateral --------------------------------------------
    Set fraMenu = Me.Controls.Add("Forms.Frame.1", "fraMenu", True)
    With fraMenu
        .Caption = ""
        .Left = 12
        .Top = 72
        .Width = 160
        .Height = 540
        .BackColor = COR_PAINEL
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 10
    End With
    
    '--- Cria Botões de Navegação ----------------------------------------------
    botoes = Array("btnDashboard", "btnTimeline", "btnCards", "btnProducao", "btnEventos", "btnConfig")
    posY = 12
    
    For i = LBound(botoes) To UBound(botoes)
        Dim btn As MSForms.CommandButton
        Set btn = Me.Controls.Add("Forms.CommandButton.1", botoes(i), True)
        
        btn.Caption = _
            IIf(i = 0, "Dashboard", _
            IIf(i = 1, "Timeline", _
            IIf(i = 2, "Cards", _
            IIf(i = 3, "Produção", _
            IIf(i = 4, "Eventos", _
            IIf(i = 5, "Configurações", ""))))))
        btn.Left = ESPACAMENTO
        btn.Top = posY
        btn.Width = LARGURA_BOTAO - (2 * ESPACAMENTO)
        btn.Height = ALTURA_BOTAO
        btn.Font.Size = 10
        btn.Font.Bold = True
        btn.BackColor = COR_HEADER
        btn.ForeColor = COR_TEXTO_CLARO
        
        posY = posY + ALTURA_BOTAO + ESPACAMENTO
    Next i
    
    '--- Cria Frame de Conteúdo Central -----------------------------------------
    Set fraConteudo = Me.Controls.Add("Forms.Frame.1", "fraConteudo", True)
    With fraConteudo
        .Caption = ""
        .Left = 184
        .Top = 72
        .Width = 788
        .Height = 540
        .BackColor = COR_PAINEL
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 10
    End With
    
    ' Exibe o painel inicial
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
    
    ' Direciona para a rotina específica de cada módulo
    Select Case NomePainel
        Case "Dashboard"
            CarregarDashboard
        Case "Timeline"
            modTimeline.CarregarTimeline fraConteudo
        Case "Cards"
            modCards.CarregarCards fraConteudo
        Case "Produção"
            frmCadastroOP.Show vbModal
        Case "Eventos"
            frmEventos.Show vbModal
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
        statusAtual = CStr(dados(i, 7))
        
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
    
    ' Renderiza interface do Dashboard
    Dim ctrl As MSForms.Control
    For Each ctrl In fraConteudo.Controls
        fraConteudo.Controls.Remove ctrl.Name
    Next ctrl
    
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
    
    Const COLUNA1_X As Single = 40
    Const COLUNA2_X As Single = 420
    Const LINHA_INICIO_Y As Single = 80
    Const LARGURA_KPI As Single = 320
    Const ALTURA_KPI As Single = 100
    Const ESPACO_Y As Single = 120
    
    Set lblKPIPlanejadas = CriaKPI(COLUNA1_X, LINHA_INICIO_Y, _
        "Total de Produções Planejadas", CStr(totalPlanejadas), COR_TEXTO_ESCURO)
    
    Set lblKPIConcluidas = CriaKPI(COLUNA2_X, LINHA_INICIO_Y, _
        "Produções Concluídas", CStr(totalConcluidas), vbGreen)
    
    Set lblKPIEmAndamento = CriaKPI(COLUNA1_X, LINHA_INICIO_Y + ESPACO_Y, _
        "Produções em Andamento", CStr(totalEmAndamento), COR_HEADER)
    
    Dim corAtrasos As Long
    If totalAtrasadas > 0 Then
        corAtrasos = COR_ALERTA
    Else
        corAtrasos = COR_TEXTO_ESCURO
    End If
    
    Set lblKPIAtrasadas = CriaKPI(COLUNA2_X, LINHA_INICIO_Y + ESPACO_Y, _
        "Produções Atrasadas", CStr(totalAtrasadas), corAtrasos)
    
    Set lblKPIHorasPlanejadas = CriaKPI(COLUNA1_X, LINHA_INICIO_Y + 2 * ESPACO_Y, _
        "Total de Horas Planejadas", Format(totalHorasPlanejadas, "0.00"), COR_TEXTO_ESCURO)
    
    Set lblKPIHorasRealizadas = CriaKPI(COLUNA2_X, LINHA_INICIO_Y + 2 * ESPACO_Y, _
        "Total de Horas Realizadas", Format(totalHorasRealizadas, "0.00"), COR_TEXTO_ESCURO)
    
Sair:
    Exit Sub
    
ErroCarregarDashboard:
    MsgBox "Erro ao carregar Dashboard: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Dashboard"
    Resume Sair
End Sub

'================================================================================
' FUNÇÃO PRIVADA: CriaKPI
' PROPÓSITO: Fábrica de controles Label estilizados para exibição de indicadores
'================================================================================
Private Function CriaKPI(pLeft As Single, pTop As Single, _
                         pTitulo As String, pValor As String, _
                         pCorValor As Long) As MSForms.Label
    Dim lbl As MSForms.Label
    Dim lblValor As MSForms.Label
    Dim lblTituloKPI As MSForms.Label
    
    Dim uniqueID As String
    uniqueID = Format(Now, "SSSSS") & "_" & CStr(Int(Rnd * 100000))
    
    Set lbl = fraConteudo.Controls.Add("Forms.Label.1", "lblKPI_" & uniqueID, True)
    With lbl
        .Left = pLeft
        .Top = pTop
        .Width = 320
        .Height = 100
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set lblTituloKPI = fraConteudo.Controls.Add("Forms.Label.1", "lblKPI_" & uniqueID & "_T", True)
    With lblTituloKPI
        .Caption = pTitulo
        .Left = pLeft + 10
        .Top = pTop + 10
        .Width = 300
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
        .Width = 300
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
    ExibirPainel "Produção"
End Sub

Private Sub btnEventos_Click()
    ExibirPainel "Eventos"
End Sub

Private Sub btnConfig_Click()
    ExibirPainel "Configurações"
End Sub

'================================================================================
' EVENTO: UserForm_QueryClose
' PROPÓSITO: Confirmação de saída segura e restauração do Excel
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
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
