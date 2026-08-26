VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Dashboard Produção Real"
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
  Attribute VB_Name = "frmDashboardProducao"
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
Attribute VB_Name = "frmDashboardProducao"
'================================================================================
' USERFORM: frmDashboardProducao
' DESCRIÇÃO: Dashboard de Produção Real com indicadores e filtros
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES DE FILTRO ========================================================
Private cboPeriodo As MSForms.ComboBox
Private dtpInicio As MSForms.TextBox
Private dtpFim As MSForms.TextBox
Private cboEquipamento As MSForms.ComboBox
Private cboOperador As MSForms.ComboBox
Private cboStatus As MSForms.ComboBox

'=== BOTÕES =====================================================================
Private btnAtualizar As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

'=== CONTAINERS =================================================================
Private fraKPIs As MSForms.Frame
Private fraProducaoEquip As MSForms.Frame
Private fraProducaoOperador As MSForms.Frame
Private fraParadas As MSForms.Frame

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
    Me.Caption = "APS PURAN – Dashboard Produção Real"
    
    m_PeriodoInicio = DateSerial(Year(Date), Month(Date), 1)
    m_PeriodoFim = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    Call CriarControles
    Call CarregarFiltros
    Call AtualizarDashboard
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar dashboard: " & Err.Description, vbCritical, "APS PURAN"
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
        .Caption = "DASHBOARD DE PRODUÇÃO REAL"
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
    Dim lblEquip As MSForms.Label
    Set lblEquip = Me.Controls.Add("Forms.Label.1", "lblEquip", True)
    With lblEquip
        .Caption = "Equipamento:"
        .Left = 12
        .Top = 88
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboEquipamento = Me.Controls.Add("Forms.ComboBox.1", "cboEquipamento", True)
    With cboEquipamento
        .Left = 100
        .Top = 88
        .Width = 200
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    Dim lblOperador As MSForms.Label
    Set lblOperador = Me.Controls.Add("Forms.Label.1", "lblOperador", True)
    With lblOperador
        .Caption = "Operador:"
        .Left = 310
        .Top = 88
        .Width = 70
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboOperador = Me.Controls.Add("Forms.ComboBox.1", "cboOperador", True)
    With cboOperador
        .Left = 390
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
        .Left = 600
        .Top = 88
        .Width = 60
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboStatus = Me.Controls.Add("Forms.ComboBox.1", "cboStatus", True)
    With cboStatus
        .Left = 670
        .Top = 88
        .Width = 120
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    '--- Botões -----------------------------------------------------------------
    Set btnAtualizar = Me.Controls.Add("Forms.CommandButton.1", "btnAtualizar", True)
    With btnAtualizar
        .Caption = "Atualizar"
        .Left = 800
        .Top = 56
        .Width = 90
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = "Fechar"
        .Left = 800
        .Top = 88
        .Width = 90
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Container KPIs ---------------------------------------------------------
    Set fraKPIs = Me.Controls.Add("Forms.Frame.1", "fraKPIs", True)
    With fraKPIs
        .Caption = "INDICADORES PRINCIPAIS"
        .Left = 12
        .Top = 120
        .Width = 860
        .Height = 140
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    '--- Container Produção por Equipamento -------------------------------------
    Set fraProducaoEquip = Me.Controls.Add("Forms.Frame.1", "fraProducaoEquip", True)
    With fraProducaoEquip
        .Caption = "PRODUÇÃO POR EQUIPAMENTO"
        .Left = 12
        .Top = 270
        .Width = 420
        .Height = 160
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    '--- Container Produção por Operador -----------------------------------------
    Set fraProducaoOperador = Me.Controls.Add("Forms.Frame.1", "fraProducaoOperador", True)
    With fraProducaoOperador
        .Caption = "PRODUÇÃO POR OPERADOR"
        .Left = 450
        .Top = 270
        .Width = 420
        .Height = 160
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    '--- Container Paradas ------------------------------------------------------
    Set fraParadas = Me.Controls.Add("Forms.Frame.1", "fraParadas", True)
    With fraParadas
        .Caption = "PARADAS"
        .Left = 12
        .Top = 440
        .Width = 860
        .Height = 160
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Bold = True
    End With
End Sub

'================================================================================
' SUBROTINAS PÚBLICAS
'================================================================================

Public Sub CarregarFiltros()
    On Error Resume Next
    
    '--- Período ---------------------------------------------------------------
    cboPeriodo.Clear
    cboPeriodo.AddItem "Dia"
    cboPeriodo.AddItem "Semana"
    cboPeriodo.AddItem "Mês"
    cboPeriodo.AddItem "Personalizado"
    cboPeriodo.Value = "Mês"
    
    '--- Status ---------------------------------------------------------------
    cboStatus.Clear
    cboStatus.AddItem ""
    cboStatus.AddItem "EM PRODUÇÃO"
    cboStatus.AddItem "PAUSADA"
    cboStatus.AddItem "FINALIZADA"
    
    '--- Equipamentos ----------------------------------------------------------
    Dim dadosProd As Variant
    Dim dadosPar As Variant
    Dim i As Long
    Dim equipamentos As Object
    Set equipamentos = CreateObject("Scripting.Dictionary")
    
    dadosProd = ObterProducoesEmArray()
    If Not IsError(dadosProd) Then
        For i = 2 To UBound(dadosProd, 1)
            Dim eq As String
            eq = Trim(CStr(dadosProd(i, 4)))
            If eq <> "" Then equipamentos(eq) = 1
        Next i
    End If
    
    dadosPar = ObterParadasEmArray()
    If Not IsError(dadosPar) Then
        For i = 2 To UBound(dadosPar, 1)
            Dim eqPar As String
            eqPar = Trim(CStr(dadosPar(i, 3)))
            If eqPar <> "" Then equipamentos(eqPar) = 1
        Next i
    End If
    
    cboEquipamento.Clear
    cboEquipamento.AddItem ""
    Dim chave As Variant
    For Each chave In equipamentos.Keys
        cboEquipamento.AddItem CStr(chave)
    Next chave
    
    '--- Operadores ------------------------------------------------------------
    Dim operadores As Object
    Set operadores = CreateObject("Scripting.Dictionary")
    
    dadosProd = ObterProducoesEmArray()
    If Not IsError(dadosProd) Then
        For i = 2 To UBound(dadosProd, 1)
            Dim op As String
            op = Trim(CStr(dadosProd(i, 3)))
            If op <> "" Then operadores(op) = 1
        Next i
    End If
    
    dadosPar = ObterParadasEmArray()
    If Not IsError(dadosPar) Then
        For i = 2 To UBound(dadosPar, 1)
            Dim opPar As String
            opPar = Trim(CStr(dadosPar(i, 4)))
            If opPar <> "" Then operadores(opPar) = 1
        Next i
    End If
    
    cboOperador.Clear
    cboOperador.AddItem ""
    Dim chaveOp As Variant
    For Each chaveOp In operadores.Keys
        cboOperador.AddItem CStr(chaveOp)
    Next chaveOp
End Sub

Public Sub AtualizarDashboard()
    On Error GoTo ErroAtualizar
    
    Dim inicioPeriodo As Date
    Dim fimPeriodo As Date
    Dim equipamentoFiltro As String
    Dim operadorFiltro As String
    Dim statusFiltro As String
    
    ' Calcula período conforme seleção
    Call CalcularPeriodo(inicioPeriodo, fimPeriodo)
    
    m_PeriodoInicio = inicioPeriodo
    m_PeriodoFim = fimPeriodo
    
    dtpInicio.Text = Format(inicioPeriodo, "dd/mm/yyyy")
    dtpFim.Text = Format(fimPeriodo, "dd/mm/yyyy")
    
    equipamentoFiltro = Trim(cboEquipamento.Value)
    operadorFiltro = Trim(cboOperador.Value)
    statusFiltro = Trim(cboStatus.Value)
    
    Dim dashboard As clsDashboardProducao
    Set dashboard = CalcularDashboardProducao(inicioPeriodo, fimPeriodo, equipamentoFiltro, operadorFiltro, statusFiltro)
    
    Call RenderizarKPIs(dashboard)
    Call RenderizarProducaoEquipamento(dashboard)
    Call RenderizarProducaoOperador(dashboard)
    Call RenderizarParadas(dashboard)
    
Sair:
    Exit Sub
    
ErroAtualizar:
    MsgBox "Erro ao atualizar dashboard: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINAS PRIVADAS
'================================================================================

Private Sub CalcularPeriodo(ByRef pInicio As Date, ByRef pFim As Date)
    On Error Resume Next
    
    Dim periodo As String
    periodo = Trim(cboPeriodo.Value)
    
    If periodo = "Dia" Then
        pInicio = Date
        pFim = Date + 1
    ElseIf periodo = "Semana" Then
        pInicio = Date - Weekday(Date, vbMonday) + 1
        pFim = pInicio + 7
    ElseIf periodo = "Mês" Then
        pInicio = DateSerial(Year(Date), Month(Date), 1)
        pFim = DateSerial(Year(Date), Month(Date) + 1, 0)
    Else
        ' Personalizado: usa o período atual ou padrão
        pInicio = m_PeriodoInicio
        pFim = m_PeriodoFim
    End If
End Sub

Private Sub RenderizarKPIs(pDashboard As clsDashboardProducao)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraKPIs.Controls
        fraKPIs.Controls.Remove ctrl.Name
    Next ctrl
    
    Dim leftPos As Single
    Dim topPos As Single
    leftPos = 10
    topPos = 20
    
    Dim kpis As Variant
    Dim valores As Variant
    Dim rotulos As Variant
    
    kpis = Array("Total Produções", "Em Produção", "Pausadas", "Finalizadas", _
                 "Planejada", "Produzida", "Rejeitada", "% Produzido", _
                 "Horas Prod.", "Horas Parada", "Paradas", "Eficiência")
    
    valores = Array(CStr(pDashboard.TotalProducoes), _
                    CStr(pDashboard.ProducoesEmProducao), _
                    CStr(pDashboard.ProducoesPausadas), _
                    CStr(pDashboard.ProducoesFinalizadas), _
                    CStr(pDashboard.QuantidadePlanejada), _
                    CStr(pDashboard.QuantidadeProduzida), _
                    CStr(pDashboard.QuantidadeRejeitada), _
                    Format(pDashboard.PercentualProduzido, "0.0") & "%", _
                    Format(pDashboard.HorasProducao, "0.0") & "h", _
                    Format(pDashboard.HorasParada, "0.0") & "h", _
                    CStr(pDashboard.NumeroParadas), _
                    Format(pDashboard.EficienciaProducao, "0.0") & "%")
    
    Dim i As Long
    For i = 0 To UBound(kpis)
        Dim lblKPI As MSForms.Label
        Set lblKPI = fraKPIs.Controls.Add("Forms.Label.1", "lblKPI_" & i, True)
        With lblKPI
            .Caption = kpis(i)
            .Left = leftPos
            .Top = topPos
            .Width = 100
            .Height = 18
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        Dim lblValor As MSForms.Label
        Set lblValor = fraKPIs.Controls.Add("Forms.Label.1", "lblValor_" & i, True)
        With lblValor
            .Caption = valores(i)
            .Left = leftPos + 110
            .Top = topPos
            .Width = 60
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        leftPos = leftPos + 180
        If leftPos > 700 Then
            leftPos = 10
            topPos = topPos + 22
        End If
    Next i
End Sub

Private Sub RenderizarProducaoEquipamento(pDashboard As clsDashboardProducao)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraProducaoEquip.Controls
        fraProducaoEquip.Controls.Remove ctrl.Name
    Next ctrl
    
    If pDashboard.ProducaoPorEquipamento.Count = 0 Then
        Dim lblVazioEq As MSForms.Label
        Set lblVazioEq = fraProducaoEquip.Controls.Add("Forms.Label.1", "lblVazioEq", True)
        With lblVazioEq
            .Caption = "Nenhum dado disponível"
            .Left = 20
            .Top = 30
            .Width = 380
            .Height = 20
            .Font.Size = 10
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        Exit Sub
    End If
    
    Dim topPos As Single
    topPos = 20
    
    Dim chave As Variant
    For Each chave In pDashboard.ProducaoPorEquipamento.Keys
        Dim lblEq As MSForms.Label
        Set lblEq = fraProducaoEquip.Controls.Add("Forms.Label.1", "lblEq_" & CStr(chave), True)
        With lblEq
            .Caption = CStr(chave)
            .Left = 20
            .Top = topPos
            .Width = 200
            .Height = 18
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        Dim lblValorEq As MSForms.Label
        Set lblValorEq = fraProducaoEquip.Controls.Add("Forms.Label.1", "lblValorEq_" & CStr(chave), True)
        With lblValorEq
            .Caption = CStr(pDashboard.ProducaoPorEquipamento(chave)) & " produções"
            .Left = 220
            .Top = topPos
            .Width = 180
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        topPos = topPos + 20
    Next chave
End Sub

Private Sub RenderizarProducaoOperador(pDashboard As clsDashboardProducao)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraProducaoOperador.Controls
        fraProducaoOperador.Controls.Remove ctrl.Name
    Next ctrl
    
    If pDashboard.ProducaoPorOperador.Count = 0 Then
        Dim lblVazioOp As MSForms.Label
        Set lblVazioOp = fraProducaoOperador.Controls.Add("Forms.Label.1", "lblVazioOp", True)
        With lblVazioOp
            .Caption = "Nenhum dado disponível"
            .Left = 20
            .Top = 30
            .Width = 380
            .Height = 20
            .Font.Size = 10
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        Exit Sub
    End If
    
    Dim topPos As Single
    topPos = 20
    
    Dim chave As Variant
    For Each chave In pDashboard.ProducaoPorOperador.Keys
        Dim lblOp As MSForms.Label
        Set lblOp = fraProducaoOperador.Controls.Add("Forms.Label.1", "lblOp_" & CStr(chave), True)
        With lblOp
            .Caption = CStr(chave)
            .Left = 20
            .Top = topPos
            .Width = 200
            .Height = 18
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        Dim lblValorOp As MSForms.Label
        Set lblValorOp = fraProducaoOperador.Controls.Add("Forms.Label.1", "lblValorOp_" & CStr(chave), True)
        With lblValorOp
            .Caption = CStr(pDashboard.ProducaoPorOperador(chave)) & " produções"
            .Left = 220
            .Top = topPos
            .Width = 180
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        topPos = topPos + 20
    Next chave
End Sub

Private Sub RenderizarParadas(pDashboard As clsDashboardProducao)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraParadas.Controls
        fraParadas.Controls.Remove ctrl.Name
    Next ctrl
    
    If pDashboard.NumeroParadas = 0 Then
        Dim lblVazioPar As MSForms.Label
        Set lblVazioPar = fraParadas.Controls.Add("Forms.Label.1", "lblVazioPar", True)
        With lblVazioPar
            .Caption = "Nenhuma parada registrada no período"
            .Left = 20
            .Top = 30
            .Width = 820
            .Height = 20
            .Font.Size = 10
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        Exit Sub
    End If
    
    Dim topPos As Single
    topPos = 20
    
    Dim lblTotalPar As MSForms.Label
    Set lblTotalPar = fraParadas.Controls.Add("Forms.Label.1", "lblTotalPar", True)
    With lblTotalPar
        .Caption = "Total de paradas: " & CStr(pDashboard.NumeroParadas) & _
                   " | Horas paradas: " & Format(pDashboard.HorasParada, "0.0") & "h"
        .Left = 20
        .Top = topPos
        .Width = 820
        .Height = 18
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
    End With
    
    topPos = topPos + 25
    
    Dim chave As Variant
    For Each chave In pDashboard.MotivosParada.Keys
        Dim lblMotivo As MSForms.Label
        Set lblMotivo = fraParadas.Controls.Add("Forms.Label.1", "lblMotivo_" & CStr(chave), True)
        With lblMotivo
            .Caption = CStr(chave)
            .Left = 20
            .Top = topPos
            .Width = 400
            .Height = 18
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        Dim lblQtdMotivo As MSForms.Label
        Set lblQtdMotivo = fraParadas.Controls.Add("Forms.Label.1", "lblQtdMotivo_" & CStr(chave), True)
        With lblQtdMotivo
            .Caption = CStr(pDashboard.MotivosParada(chave)) & " ocorrências"
            .Left = 430
            .Top = topPos
            .Width = 200
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO
        End With
        
        topPos = topPos + 20
    Next chave
End Sub

'================================================================================
' EVENTOS DOS BOTÕES ===========================================================
'================================================================================

Private Sub btnAtualizar_Click()
    On Error Resume Next
    Call AtualizarDashboard
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cboPeriodo_Change()
    On Error Resume Next
    Call AtualizarDashboard
End Sub

Private Sub cboEquipamento_Change()
    On Error Resume Next
    Call AtualizarDashboard
End Sub

Private Sub cboOperador_Change()
    On Error Resume Next
    Call AtualizarDashboard
End Sub

Private Sub cboStatus_Change()
    On Error Resume Next
    Call AtualizarDashboard
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
