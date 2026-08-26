VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Relatórios de Produção"
  ClientHeight = 650
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 900
  Height = 688
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 912
  StartUpPosition = 1
  Attribute VB_Name = "frmRelatorios"
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
Attribute VB_Name = "frmRelatorios"
'================================================================================
' USERFORM: frmRelatorios
' DESCRIÇÃO: Relatórios de produção com indicadores agregados
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES ================================================================
Private cboPeriodo As MSForms.ComboBox
Private dtpInicio As MSForms.TextBox
Private dtpFim As MSForms.TextBox
Private cboPosto As MSForms.ComboBox
Private cboProduto As MSForms.ComboBox
Private cboStatus As MSForms.ComboBox
Private btnGerar As MSForms.CommandButton
Private btnAtualizar As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton
Private fraResultado As MSForms.Frame
Private lblResumo As MSForms.Label

'=== ESTADO =====================================================================
Private m_PeriodoInicio As Date
Private m_PeriodoFim As Date

'=== PROPRIEDADES VISUAIS ======================================================
Private Const COR_FUNDO As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_AZUL As Long = 15773696
Private Const COR_VERDE As Long = 5287936
Private Const COR_AMARELO As Long = 65535
Private Const COR_ALERTA As Long = 255

'================================================================================
' EVENTO: UserForm_Initialize
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Me.BackColor = COR_FUNDO
    Me.Caption = "APS PURAN – Relatórios de Produção"
    
    m_PeriodoInicio = DateSerial(Year(Date), Month(Date), 1)
    m_PeriodoFim = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    Call CriarControles
    Call CarregarFiltros
    Call GerarRelatorio
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar relatórios: " & Err.Description, vbCritical, "APS PURAN"
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
        .Caption = "RELATÓRIOS DE PRODUÇÃO"
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
    
    '--- Período ---------------------------------------------------------------
    Dim lblPeriodo As MSForms.Label
    Set lblPeriodo = Me.Controls.Add("Forms.Label.1", "lblPeriodo", True)
    With lblPeriodo
        .Caption = "Período:"
        .Left = 12
        .Top = 56
        .Width = 60
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboPeriodo = Me.Controls.Add("Forms.ComboBox.1", "cboPeriodo", True)
    With cboPeriodo
        .Left = 80
        .Top = 56
        .Width = 120
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    Set dtpInicio = Me.Controls.Add("Forms.TextBox.1", "dtpInicio", True)
    With dtpInicio
        .Text = Format(m_PeriodoInicio, "dd/mm/yyyy")
        .Left = 210
        .Top = 56
        .Width = 100
        .Height = 22
        .Font.Size = 9
        .Locked = True
        .BackColor = vbWhite
    End With
    
    Dim lblAte As MSForms.Label
    Set lblAte = Me.Controls.Add("Forms.Label.1", "lblAte", True)
    With lblAte
        .Caption = "até"
        .Left = 320
        .Top = 56
        .Width = 30
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignCenter
    End With
    
    Set dtpFim = Me.Controls.Add("Forms.TextBox.1", "dtpFim", True)
    With dtpFim
        .Text = Format(m_PeriodoFim, "dd/mm/yyyy")
        .Left = 360
        .Top = 56
        .Width = 100
        .Height = 22
        .Font.Size = 9
        .Locked = True
        .BackColor = vbWhite
    End With
    
    '--- Filtros ---------------------------------------------------------------
    Dim lblPosto As MSForms.Label
    Set lblPosto = Me.Controls.Add("Forms.Label.1", "lblPosto", True)
    With lblPosto
        .Caption = "Posto:"
        .Left = 12
        .Top = 88
        .Width = 60
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboPosto = Me.Controls.Add("Forms.ComboBox.1", "cboPosto", True)
    With cboPosto
        .Left = 80
        .Top = 88
        .Width = 200
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    Dim lblProduto As MSForms.Label
    Set lblProduto = Me.Controls.Add("Forms.Label.1", "lblProduto", True)
    With lblProduto
        .Caption = "Produto:"
        .Left = 300
        .Top = 88
        .Width = 60
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboProduto = Me.Controls.Add("Forms.ComboBox.1", "cboProduto", True)
    With cboProduto
        .Left = 370
        .Top = 88
        .Width = 200
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    Dim lblStatus As MSForms.Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Caption = "Status:"
        .Left = 12
        .Top = 120
        .Width = 60
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboStatus = Me.Controls.Add("Forms.ComboBox.1", "cboStatus", True)
    With cboStatus
        .Left = 80
        .Top = 120
        .Width = 200
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    '--- Botões -----------------------------------------------------------------
    Set btnGerar = Me.Controls.Add("Forms.CommandButton.1", "btnGerar", True)
    With btnGerar
        .Caption = "Gerar"
        .Left = 12
        .Top = 152
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnAtualizar = Me.Controls.Add("Forms.CommandButton.1", "btnAtualizar", True)
    With btnAtualizar
        .Caption = "Atualizar"
        .Left = 122
        .Top = 152
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = "Fechar"
        .Left = 232
        .Top = 152
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Container de resultado -------------------------------------------------
    Set fraResultado = Me.Controls.Add("Forms.Frame.1", "fraResultado", True)
    With fraResultado
        .Caption = ""
        .Left = 12
        .Top = 190
        .Width = 860
        .Height = 400
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 2000
    End With
    
    '--- Resumo -----------------------------------------------------------------
    Set lblResumo = Me.Controls.Add("Forms.Label.1", "lblResumo", True)
    With lblResumo
        .Caption = ""
        .Left = 12
        .Top = 600
        .Width = 860
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
    End With
End Sub

'================================================================================
' EVENTOS: btnGerar_Click / btnAtualizar_Click / btnFechar_Click
'================================================================================
Private Sub btnGerar_Click()
    On Error Resume Next
    Call GerarRelatorio
End Sub

Private Sub btnAtualizar_Click()
    On Error Resume Next
    Call CarregarFiltros
    Call GerarRelatorio
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

'================================================================================
' EVENTO: cboPeriodo_Change
'================================================================================
Private Sub cboPeriodo_Change()
    On Error Resume Next
    Call AtualizarPeriodo
End Sub

'================================================================================
' SUBROTINA PRIVADA: CarregarFiltros
'================================================================================
Private Sub CarregarFiltros()
    On Error Resume Next
    
    cboPeriodo.Clear
    cboPosto.Clear
    cboProduto.Clear
    cboStatus.Clear
    
    cboPeriodo.AddItem "Mês"
    cboPeriodo.AddItem "Semana"
    cboPeriodo.AddItem "Dia"
    cboPeriodo.AddItem "Personalizado"
    cboPeriodo.ListIndex = 0
    
    Dim postos As Collection
    Set postos = ObterListaPostos()
    Dim p As Variant
    cboPosto.AddItem ""
    For Each p In postos
        cboPosto.AddItem p
    Next p
    cboPosto.ListIndex = 0
    
    Dim produtos As Collection
    Set produtos = ObterListaProdutos()
    Dim prod As Variant
    cboProduto.AddItem ""
    For Each prod In produtos
        cboProduto.AddItem prod
    Next prod
    cboProduto.ListIndex = 0
    
    Dim statusList As Collection
    Set statusList = ObterListaStatus()
    Dim st As Variant
    cboStatus.AddItem ""
    For Each st In statusList
        cboStatus.AddItem st
    Next st
    cboStatus.ListIndex = 0
End Sub

'================================================================================
' SUBROTINA PRIVADA: AtualizarPeriodo
'================================================================================
Private Sub AtualizarPeriodo()
    On Error Resume Next
    
    Dim periodo As String
    periodo = cboPeriodo.Text
    
    Select Case periodo
        Case "Dia"
            m_PeriodoInicio = DateSerial(Year(Now), Month(Now), Day(Now))
            m_PeriodoFim = DateSerial(Year(Now), Month(Now), Day(Now)) + 1
        Case "Semana"
            Dim diaSemana As Long
            diaSemana = Weekday(Now, vbMonday)
            m_PeriodoInicio = DateAdd("d", -(diaSemana - 1), DateSerial(Year(Now), Month(Now), Day(Now)))
            m_PeriodoFim = DateAdd("d", 7, m_PeriodoInicio)
        Case "Mês"
            m_PeriodoInicio = DateSerial(Year(Now), Month(Now), 1)
            m_PeriodoFim = DateSerial(Year(Now), Month(Now) + 1, 0)
        Case Else
            Exit Sub
    End Select
    
    dtpInicio.Text = Format(m_PeriodoInicio, "dd/mm/yyyy")
    dtpFim.Text = Format(m_PeriodoFim, "dd/mm/yyyy")
End Sub

'================================================================================
' SUBROTINA PRIVADA: GerarRelatorio
'================================================================================
Private Sub GerarRelatorio()
    On Error GoTo ErroGerar
    
    If Not IsDate(dtpInicio.Text) Or Not IsDate(dtpFim.Text) Then
        MsgBox "Informe datas válidas.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    m_PeriodoInicio = CDate(dtpInicio.Text)
    m_PeriodoFim = CDate(dtpFim.Text)
    
    If m_PeriodoFim < m_PeriodoInicio Then
        MsgBox "A data final não pode ser anterior à inicial.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim postoFiltro As String
    Dim produtoFiltro As String
    Dim statusFiltro As String
    
    postoFiltro = IIf(cboPosto.ListIndex = -1, "", cboPosto.Text)
    produtoFiltro = IIf(cboProduto.ListIndex = -1, "", cboProduto.Text)
    statusFiltro = IIf(cboStatus.ListIndex = -1, "", cboStatus.Text)
    
    Dim relatorio As clsRelatorioProducao
    Set relatorio = GerarRelatorioProducao(m_PeriodoInicio, m_PeriodoFim, postoFiltro, produtoFiltro, statusFiltro)
    
    Call RenderizarRelatorio(relatorio)
    
Sair:
    Exit Sub
    
ErroGerar:
    MsgBox "Erro ao gerar relatório: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: RenderizarRelatorio
'================================================================================
Private Sub RenderizarRelatorio(pRelatorio As clsRelatorioProducao)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraResultado.Controls
        fraResultado.Controls.Remove ctrl.Name
    Next ctrl
    
    Dim topPos As Single
    topPos = 10
    
    ' Card 1: Total de OPs
    Call CriarCardIndicador(fraResultado, "Total de OPs", CStr(pRelatorio.TotalOPs), COR_TEXTO_ESCURO, topPos)
    topPos = topPos + 60
    
    ' Card 2: Quantidade Planejada
    Call CriarCardIndicador(fraResultado, "Quantidade Planejada", CStr(pRelatorio.QuantidadePlanejada), COR_TEXTO_ESCURO, topPos)
    topPos = topPos + 60
    
    ' Card 3: Horas Programadas
    Call CriarCardIndicador(fraResultado, "Horas Programadas", Format(pRelatorio.HorasProgramadas, "0.00") & "h", COR_TEXTO_ESCURO, topPos)
    topPos = topPos + 60
    
    ' Card 4: OPs por Status
    Call CriarCardIndicador(fraResultado, "Planejadas / Em Andamento / Concluídas / Atrasadas", _
                           CStr(pRelatorio.OPsPlanejadas) & " / " & CStr(pRelatorio.OPsEmAndamento) & " / " & CStr(pRelatorio.OPsConcluidas) & " / " & CStr(pRelatorio.OPsAtrasadas), _
                           COR_TEXTO_ESCURO, topPos)
    topPos = topPos + 60
    
    ' Card 5: Postos Utilizados
    Call CriarCardIndicador(fraResultado, "Postos Utilizados", CStr(pRelatorio.QuantidadePostos), COR_TEXTO_ESCURO, topPos)
    topPos = topPos + 60
    
    ' Card 6: Percentual de Ocupação
    Dim corOcupacao As Long
    If pRelatorio.PercentualOcupacao <= 80 Then
        corOcupacao = COR_VERDE
    ElseIf pRelatorio.PercentualOcupacao <= 100 Then
        corOcupacao = COR_AMARELO
    Else
        corOcupacao = COR_ALERTA
    End If
    
    Call CriarCardIndicador(fraResultado, "Percentual de Ocupação", Format(pRelatorio.PercentualOcupacao, "0.0") & "%", corOcupacao, topPos)
    
    lblResumo.Caption = "Período: " & Format(m_PeriodoInicio, "dd/mm/yyyy") & " a " & Format(m_PeriodoFim, "dd/mm/yyyy") & _
                        " | Posto: " & IIf(postoFiltro = "", "Todos", postoFiltro) & _
                        " | Produto: " & IIf(produtoFiltro = "", "Todos", produtoFiltro) & _
                        " | Status: " & IIf(statusFiltro = "", "Todos", statusFiltro)
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarCardIndicador
'================================================================================
Private Sub CriarCardIndicador(pContainer As MSForms.Frame, pTitulo As String, pValor As String, pCorValor As Long, pTop As Single)
    On Error Resume Next
    
    Dim uniqueID As String
    uniqueID = Format(Now, "SSSSS") & "_" & CStr(Int(Rnd * 100000))
    
    Dim lbl As MSForms.Label
    Set lbl = pContainer.Controls.Add("Forms.Label.1", "lblRel_" & uniqueID, True)
    With lbl
        .Left = 10
        .Top = pTop
        .Width = pContainer.Width - 20
        .Height = 50
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Dim lblTitulo As MSForms.Label
    Set lblTitulo = pContainer.Controls.Add("Forms.Label.1", "lblRel_" & uniqueID & "_T", True)
    With lblTitulo
        .Caption = pTitulo
        .Left = 20
        .Top = pTop + 5
        .Width = pContainer.Width - 40
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
    End With
    
    Dim lblValor As MSForms.Label
    Set lblValor = pContainer.Controls.Add("Forms.Label.1", "lblRel_" & uniqueID & "_V", True)
    With lblValor
        .Caption = pValor
        .Left = 20
        .Top = pTop + 25
        .Width = pContainer.Width - 40
        .Height = 20
        .Font.Size = 16
        .Font.Bold = True
        .ForeColor = pCorValor
        .BackColor = vbWhite
    End With
End Sub

'================================================================================
' EVENTO: UserForm_QueryClose
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
