VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Produção Real"
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
  Attribute VB_Name = "frmProducao"
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
Attribute VB_Name = "frmProducao"
'================================================================================
' USERFORM: frmProducao
' DESCRIÇÃO: Apontamento de produção real
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES PRINCIPAIS ======================================================
Private lstProducoes As MSForms.ListBox
Private lstParadas As MSForms.ListBox
Private fraAcoes As MSForms.Frame
Private fraDetalhe As MSForms.Frame

'=== BOTÕES =====================================================================
Private btnIniciar As MSForms.CommandButton
Private btnPausar As MSForms.CommandButton
Private btnRetomar As MSForms.CommandButton
Private btnFinalizar As MSForms.CommandButton
Private btnRegistrarParada As MSForms.CommandButton
Private btnFinalizarParada As MSForms.CommandButton
Private btnExcluirProducao As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

'=== CAMPOS DE EDIÇÃO ===========================================================
Private txtID_Producao As MSForms.TextBox
Private txtID_OP As MSForms.TextBox
Private txtProduto As MSForms.TextBox
Private txtEquipamento As MSForms.TextBox
Private txtOperador As MSForms.TextBox
Private txtQtdPlanejada As MSForms.TextBox
Private txtQtdProduzida As MSForms.TextBox
Private txtQtdRejeitada As MSForms.TextBox
Private txtDataInicio As MSForms.TextBox
Private txtDataFim As MSForms.TextBox
Private cboStatus As MSForms.ComboBox
Private lblTempoDecorrido As MSForms.Label
Private lblStatus As MSForms.Label

'=== CAMPOS DE PARADA ===========================================================
Private txtID_Parada As MSForms.TextBox
Private txtMotivoParada As MSForms.TextBox
Private txtObsParada As MSForms.TextBox
Private txtInicioParada As MSForms.TextBox
Private txtFimParada As MSForms.TextBox

'=== ESTADO =====================================================================
Private m_ID_ProducaoAtiva As String
Private m_ID_ParadaAtiva As String
Private m_TimerAtivo As Boolean

'=== PROPRIEDADES VISUAIS ======================================================
Private Const COR_FUNDO As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_AZUL As Long = 15773696
Private Const COR_VERDE As Long = 5287936
Private Const COR_ALERTA As Long = 255
Private Const COR_AMARELO As Long = 65535

'================================================================================
' EVENTO: UserForm_Initialize
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Me.BackColor = COR_FUNDO
    Me.Caption = "APS PURAN – Produção Real"
    
    m_ID_ProducaoAtiva = ""
    m_ID_ParadaAtiva = ""
    m_TimerAtivo = False
    
    Call CriarControles
    Call CarregarProducoes
    Call CarregarParadas
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar produção: " & Err.Description, vbCritical, "APS PURAN"
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
        .Caption = "APONTAMENTO DE PRODUÇÃO REAL"
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
    
    '--- Lista de produções ---------------------------------------------------
    Set lstProducoes = Me.Controls.Add("Forms.ListBox.1", "lstProducoes", True)
    With lstProducoes
        .Left = 12
        .Top = 60
        .Width = 400
        .Height = 200
        .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
    End With
    
    '--- Lista de paradas ------------------------------------------------------
    Set lstParadas = Me.Controls.Add("Forms.ListBox.1", "lstParadas", True)
    With lstParadas
        .Left = 12
        .Top = 280
        .Width = 400
        .Height = 150
        .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
    End With
    
    '--- Botões de ação --------------------------------------------------------
    Dim btnLeft As Single
    btnLeft = 12
    
    Set btnIniciar = Me.Controls.Add("Forms.CommandButton.1", "btnIniciar", True)
    With btnIniciar
        .Caption = "Iniciar"
        .Left = btnLeft
        .Top = 450
        .Width = 80
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 90
    
    Set btnPausar = Me.Controls.Add("Forms.CommandButton.1", "btnPausar", True)
    With btnPausar
        .Caption = "Pausar"
        .Left = btnLeft
        .Top = 450
        .Width = 80
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AMARELO
        .ForeColor = COR_TEXTO_ESCURO
    End With
    btnLeft = btnLeft + 90
    
    Set btnRetomar = Me.Controls.Add("Forms.CommandButton.1", "btnRetomar", True)
    With btnRetomar
        .Caption = "Retomar"
        .Left = btnLeft
        .Top = 450
        .Width = 80
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 90
    
    Set btnFinalizar = Me.Controls.Add("Forms.CommandButton.1", "btnFinalizar", True)
    With btnFinalizar
        .Caption = "Finalizar"
        .Left = btnLeft
        .Top = 450
        .Width = 80
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 255
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 90
    
    Set btnRegistrarParada = Me.Controls.Add("Forms.CommandButton.1", "btnRegistrarParada", True)
    With btnRegistrarParada
        .Caption = "Registrar Parada"
        .Left = btnLeft
        .Top = 450
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 130
    
    Set btnFinalizarParada = Me.Controls.Add("Forms.CommandButton.1", "btnFinalizarParada", True)
    With btnFinalizarParada
        .Caption = "Finalizar Parada"
        .Left = btnLeft
        .Top = 450
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 130
    
    Set btnExcluirProducao = Me.Controls.Add("Forms.CommandButton.1", "btnExcluirProducao", True)
    With btnExcluirProducao
        .Caption = "Excluir"
        .Left = btnLeft
        .Top = 450
        .Width = 90
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 8421504
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Campos de detalhe da produção ------------------------------------------
    Dim leftCampo As Single
    Dim topCampo As Single
    leftCampo = 430
    topCampo = 60
    
    Dim lblID As MSForms.Label
    Set lblID = Me.Controls.Add("Forms.Label.1", "lblID", True)
    With lblID
        .Caption = "ID Produção:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblOP As MSForms.Label
    Set lblOP = Me.Controls.Add("Forms.Label.1", "lblOP", True)
    With lblOP
        .Caption = "OP:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblProd As MSForms.Label
    Set lblProd = Me.Controls.Add("Forms.Label.1", "lblProd", True)
    With lblProd
        .Caption = "Produto:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblEquip As MSForms.Label
    Set lblEquip = Me.Controls.Add("Forms.Label.1", "lblEquip", True)
    With lblEquip
        .Caption = "Equipamento:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblOperador As MSForms.Label
    Set lblOperador = Me.Controls.Add("Forms.Label.1", "lblOperador", True)
    With lblOperador
        .Caption = "Operador:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblQtd As MSForms.Label
    Set lblQtd = Me.Controls.Add("Forms.Label.1", "lblQtd", True)
    With lblQtd
        .Caption = "Qtd Planejada:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblProduzida As MSForms.Label
    Set lblProduzida = Me.Controls.Add("Forms.Label.1", "lblProduzida", True)
    With lblProduzida
        .Caption = "Produzida:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblRejeitada As MSForms.Label
    Set lblRejeitada = Me.Controls.Add("Forms.Label.1", "lblRejeitada", True)
    With lblRejeitada
        .Caption = "Rejeitada:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblDataInicio As MSForms.Label
    Set lblDataInicio = Me.Controls.Add("Forms.Label.1", "lblDataInicio", True)
    With lblDataInicio
        .Caption = "Início:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    Dim lblDataFim As MSForms.Label
    Set lblDataFim = Me.Controls.Add("Forms.Label.1", "lblDataFim", True)
    With lblDataFim
        .Caption = "Fim:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    ' Campos de texto
    leftCampo = 520
    topCampo = 60
    
    Set txtID_Producao = Me.Controls.Add("Forms.TextBox.1", "txtID_Producao", True)
    With txtID_Producao
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set txtID_OP = Me.Controls.Add("Forms.TextBox.1", "txtID_OP", True)
    With txtID_OP
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtProduto = Me.Controls.Add("Forms.TextBox.1", "txtProduto", True)
    With txtProduto
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set txtEquipamento = Me.Controls.Add("Forms.TextBox.1", "txtEquipamento", True)
    With txtEquipamento
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set txtOperador = Me.Controls.Add("Forms.TextBox.1", "txtOperador", True)
    With txtOperador
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtQtdPlanejada = Me.Controls.Add("Forms.TextBox.1", "txtQtdPlanejada", True)
    With txtQtdPlanejada
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set txtQtdProduzida = Me.Controls.Add("Forms.TextBox.1", "txtQtdProduzida", True)
    With txtQtdProduzida
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtQtdRejeitada = Me.Controls.Add("Forms.TextBox.1", "txtQtdRejeitada", True)
    With txtQtdRejeitada
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtDataInicio = Me.Controls.Add("Forms.TextBox.1", "txtDataInicio", True)
    With txtDataInicio
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set txtDataFim = Me.Controls.Add("Forms.TextBox.1", "txtDataFim", True)
    With txtDataFim
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set cboStatus = Me.Controls.Add("Forms.ComboBox.1", "cboStatus", True)
    With cboStatus
        .AddItem "NÃO INICIADA"
        .AddItem "EM PRODUÇÃO"
        .AddItem "PAUSADA"
        .AddItem "FINALIZADA"
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    ' Status label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Caption = ""
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
    End With
    topCampo = topCampo + 26
    
    ' Tempo decorrido
    Set lblTempoDecorrido = Me.Controls.Add("Forms.Label.1", "lblTempoDecorrido", True)
    With lblTempoDecorrido
        .Caption = "Tempo: 00:00:00"
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
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
' SUBROTINAS PÚBLICAS
'================================================================================

Public Sub CarregarProducoes()
    On Error GoTo ErroCarregar
    
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim texto As String
    
    dados = ObterProducoesEmArray()
    
    If IsError(dados) Then
        lstProducoes.Clear
        Exit Sub
    End If
    
    lstProducoes.Clear
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        texto = CStr(dados(i, 1)) & " | " & CStr(dados(i, 2)) & " | " & CStr(dados(i, 4)) & _
                " | " & Format(CDate(dados(i, 5)), "dd/mm/yyyy HH:MM") & " | " & CStr(dados(i, 10))
        lstProducoes.AddItem texto
    Next i
    
Sair:
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar produções: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Public Sub CarregarParadas()
    On Error GoTo ErroCarregarParadas
    
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim texto As String
    
    dados = ObterParadasEmArray()
    
    If IsError(dados) Then
        lstParadas.Clear
        Exit Sub
    End If
    
    lstParadas.Clear
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        texto = CStr(dados(i, 1)) & " | " & CStr(dados(i, 2)) & " | " & CStr(dados(i, 4)) & _
                " | " & Format(CDate(dados(i, 5)), "dd/mm/yyyy HH:MM")
        lstParadas.AddItem texto
    Next i
    
Sair:
    Exit Sub
    
ErroCarregarParadas:
    MsgBox "Erro ao carregar paradas: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub LimparCampos()
    On Error Resume Next
    
    txtID_Producao.Text = ""
    txtID_OP.Text = ""
    txtProduto.Text = ""
    txtEquipamento.Text = ""
    txtOperador.Text = ""
    txtQtdPlanejada.Text = ""
    txtQtdProduzida.Text = ""
    txtQtdRejeitada.Text = ""
    txtDataInicio.Text = ""
    txtDataFim.Text = ""
    cboStatus.Value = ""
    lblStatus.Caption = ""
    lblTempoDecorrido.Caption = "Tempo: 00:00:00"
    m_ID_ProducaoAtiva = ""
    m_ID_ParadaAtiva = ""
End Sub

Private Sub PreencherCamposProducao(pID_Producao As String)
    On Error Resume Next
    
    Dim dados As Variant
    dados = ObterProducoesEmArray()
    
    If IsError(dados) Then Exit Sub
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If CStr(dados(i, 1)) = pID_Producao Then
            txtID_Producao.Text = CStr(dados(i, 1))
            txtID_OP.Text = CStr(dados(i, 2))
            txtProduto.Text = ""
            txtEquipamento.Text = CStr(dados(i, 4))
            txtOperador.Text = CStr(dados(i, 3))
            txtQtdPlanejada.Text = CStr(dados(i, 7))
            txtQtdProduzida.Text = CStr(dados(i, 8))
            txtQtdRejeitada.Text = CStr(dados(i, 9))
            If IsDate(dados(i, 5)) Then txtDataInicio.Text = Format(CDate(dados(i, 5)), "dd/mm/yyyy HH:MM")
            If IsDate(dados(i, 6)) Then txtDataFim.Text = Format(CDate(dados(i, 6)), "dd/mm/yyyy HH:MM")
            cboStatus.Value = CStr(dados(i, 10))
            lblStatus.Caption = CStr(dados(i, 10))
            m_ID_ProducaoAtiva = pID_Producao
            Exit For
        End If
    Next i
End Sub

Private Sub txtID_OP_Change()
    On Error Resume Next
    
    Dim idOP As String
    idOP = Trim(txtID_OP.Text)
    
    If idOP = "" Then
        txtProduto.Text = ""
        txtEquipamento.Text = ""
        txtQtdPlanejada.Text = ""
        txtDataInicio.Text = ""
        txtDataFim.Text = ""
        Exit Sub
    End If
    
    Dim dadosOP As Variant
    dadosOP = BuscarOPPorID(idOP)
    
    If IsEmpty(dadosOP) Then
        txtProduto.Text = "OP não encontrada"
        txtEquipamento.Text = ""
        txtQtdPlanejada.Text = ""
        txtDataInicio.Text = ""
        txtDataFim.Text = ""
    Else
        txtProduto.Text = CStr(dadosOP(1))
        txtEquipamento.Text = CStr(dadosOP(2))
        txtQtdPlanejada.Text = CStr(dadosOP(3))
        If IsDate(dadosOP(4)) Then txtDataInicio.Text = Format(CDate(dadosOP(4)), "dd/mm/yyyy")
        If IsDate(dadosOP(5)) Then txtDataFim.Text = Format(CDate(dadosOP(5)), "dd/mm/yyyy")
    End If
End Sub

'================================================================================
' EVENTOS DOS BOTÕES ===========================================================
'================================================================================

Private Sub btnIniciar_Click()
    On Error GoTo ErroIniciar
    
    Dim idOP As String
    Dim equipamento As String
    Dim operador As String
    Dim qtdPlanejada As Long
    
    idOP = Trim(txtID_OP.Text)
    If idOP = "" Then
        MsgBox "Informe a ID da OP.", vbExclamation, "Validação"
        txtID_OP.SetFocus
        Exit Sub
    End If
    
    If txtProduto.Text = "OP não encontrada" Or txtProduto.Text = "" Then
        MsgBox "OP não encontrada. Verifique o ID informado.", vbCritical, "Validação"
        txtID_OP.SetFocus
        Exit Sub
    End If
    
    equipamento = Trim(txtEquipamento.Text)
    If equipamento = "" Then
        MsgBox "Equipamento não carregado da OP.", vbCritical, "Validação"
        Exit Sub
    End If
    
    operador = Trim(txtOperador.Text)
    If operador = "" Then operador = "NÃO INFORMADO"
    
    qtdPlanejada = 0
    On Error Resume Next
    qtdPlanejada = CLng(txtQtdPlanejada.Text)
    If qtdPlanejada <= 0 Then qtdPlanejada = 0
    On Error GoTo ErroIniciar
    
    Dim idProducao As String
    idProducao = "PRD_" & Format(Now, "yyyymmdd_hhmmss")
    
    Call IniciarProducao(idProducao, idOP, operador, equipamento, qtdPlanejada)
    MsgBox "Produção iniciada com sucesso!", vbInformation, "APS PURAN"
    
    Call CarregarProducoes
    
Sair:
    Exit Sub
    
ErroIniciar:
    MsgBox "Erro ao iniciar produção: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnPausar_Click()
    On Error GoTo ErroPausar
    
    If m_ID_ProducaoAtiva = "" Then
        MsgBox "Selecione uma produção para pausar.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Call PausarProducao(m_ID_ProducaoAtiva)
    MsgBox "Produção pausada!", vbInformation, "APS PURAN"
    
    Call CarregarProducoes
    Call PreencherCamposProducao(m_ID_ProducaoAtiva)
    
Sair:
    Exit Sub
    
ErroPausar:
    MsgBox "Erro ao pausar produção: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnRetomar_Click()
    On Error GoTo ErroRetomar
    
    If m_ID_ProducaoAtiva = "" Then
        MsgBox "Selecione uma produção para retomar.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Call RetomarProducao(m_ID_ProducaoAtiva)
    MsgBox "Produção retomada!", vbInformation, "APS PURAN"
    
    Call CarregarProducoes
    Call PreencherCamposProducao(m_ID_ProducaoAtiva)
    
Sair:
    Exit Sub
    
ErroRetomar:
    MsgBox "Erro ao retomar produção: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnFinalizar_Click()
    On Error GoTo ErroFinalizar
    
    If m_ID_ProducaoAtiva = "" Then
        MsgBox "Selecione uma produção para finalizar.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim qtdProduzida As Long
    Dim qtdRejeitada As Long
    
    qtdProduzida = 0
    qtdRejeitada = 0
    
    On Error Resume Next
    qtdProduzida = CLng(txtQtdProduzida.Text)
    If qtdProduzida < 0 Then qtdProduzida = 0
    On Error GoTo ErroFinalizar
    
    On Error Resume Next
    qtdRejeitada = CLng(txtQtdRejeitada.Text)
    If qtdRejeitada < 0 Then qtdRejeitada = 0
    On Error GoTo ErroFinalizar
    
    Call FinalizarProducao(m_ID_ProducaoAtiva, qtdProduzida, qtdRejeitada)
    MsgBox "Produção finalizada!", vbInformation, "APS PURAN"
    
    Call CarregarProducoes
    Call LimparCampos
    
Sair:
    Exit Sub
    
ErroFinalizar:
    MsgBox "Erro ao finalizar produção: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnRegistrarParada_Click()
    On Error GoTo ErroRegParada
    
    If m_ID_ProducaoAtiva = "" Then
        MsgBox "Selecione uma produção para registrar parada.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim motivo As String
    Dim observacao As String
    Dim idParada As String
    
    motivo = InputBox("Motivo da parada:", "Registrar Parada")
    If Trim(motivo) = "" Then
        MsgBox "Informe o motivo da parada.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    observacao = InputBox("Observação (opcional):", "Registrar Parada")
    idParada = "PAR_" & Format(Now, "yyyymmdd_hhmmss")
    
    Call RegistrarParada(idParada, Trim(txtID_OP.Text), Trim(txtEquipamento.Text), Trim(txtOperador.Text), motivo, observacao)
    MsgBox "Parada registrada!", vbInformation, "APS PURAN"
    
    Call CarregarParadas
    
Sair:
    Exit Sub
    
ErroRegParada:
    MsgBox "Erro ao registrar parada: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnFinalizarParada_Click()
    On Error GoTo ErroFinParada
    
    If lstParadas.ListIndex = -1 Then
        MsgBox "Selecione uma parada para finalizar.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim chaveParada As String
    chaveParada = Split(lstParadas.List(lstParadas.ListIndex), " | ")(0)
    
    Call FinalizarParada(chaveParada)
    MsgBox "Parada finalizada!", vbInformation, "APS PURAN"
    
    Call CarregarParadas
    
Sair:
    Exit Sub
    
ErroFinParada:
    MsgBox "Erro ao finalizar parada: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnExcluirProducao_Click()
    On Error GoTo ErroExcluir
    
    If m_ID_ProducaoAtiva = "" Then
        MsgBox "Selecione uma produção para excluir.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente excluir a produção selecionada?", vbQuestion + vbYesNo, "Confirmação")
    If resposta = vbNo Then Exit Sub
    
    Call ExcluirProducao(m_ID_ProducaoAtiva)
    MsgBox "Produção excluída!", vbInformation, "APS PURAN"
    
    Call CarregarProducoes
    Call LimparCampos
    
Sair:
    Exit Sub
    
ErroExcluir:
    MsgBox "Erro ao excluir produção: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub lstProducoes_Click()
    On Error Resume Next
    
    If lstProducoes.ListIndex = -1 Then Exit Sub
    
    Dim chave As String
    chave = Split(lstProducoes.List(lstProducoes.ListIndex), " | ")(0)
    
    Call PreencherCamposProducao(chave)
End Sub

Private Sub lstParadas_Click()
    On Error Resume Next
    
    If lstParadas.ListIndex = -1 Then Exit Sub
    
    Dim chave As String
    chave = Split(lstParadas.List(lstParadas.ListIndex), " | ")(0)
    m_ID_ParadaAtiva = chave
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
