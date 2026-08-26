VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Comparação de Simulações"
  ClientHeight = 600
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 1000
  Height = 638
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 1012
  StartUpPosition = 1
  Attribute VB_Name = "frmComparacaoSimulacao"
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
Attribute VB_Name = "frmComparacaoSimulacao"
'================================================================================
' USERFORM: frmComparacaoSimulacao
' DESCRIÇÃO: Comparação visual entre duas simulações de produção
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES PRINCIPAIS ======================================================
Private cboSimulacaoA As MSForms.ComboBox
Private cboSimulacaoB As MSForms.ComboBox
Private btnComparar As MSForms.CommandButton
Private btnAtualizar As MSForms.CommandButton
Private btnExportarComparacaoExcel As MSForms.CommandButton
Private btnExportarComparacaoPDF As MSForms.CommandButton
Private btnImprimirComparacao As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

Private fraA As MSForms.Frame
Private fraB As MSForms.Frame
Private hScrollA As MSForms.ScrollBar
Private hScrollB As MSForms.ScrollBar
Private lblTituloA As MSForms.Label
Private lblTituloB As MSForms.Label
Private lblResumo As MSForms.Label

'=== ESTADO =====================================================================
Private m_ID_SimulacaoA As String
Private m_ID_SimulacaoB As String
Private m_ScrollPos As Single
Private m_SincronizandoScroll As Boolean
Private m_PosicoesSalvas As Boolean

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
    Me.Caption = "APS PURAN – Comparação de Simulações"
    
    m_ID_SimulacaoA = ""
    m_ID_SimulacaoB = ""
    
    Call CarregarPreferencias
    Call CriarControles
    Call CarregarListaSimulacoes
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar comparação: " & Err.Description, vbCritical, "APS PURAN"
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
        .Caption = "COMPARAÇÃO DE SIMULAÇÕES"
        .Left = 12
        .Top = 12
        .Width = 960
        .Height = 32
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Seleção Simulação A ---------------------------------------------------
    Dim lblA As MSForms.Label
    Set lblA = Me.Controls.Add("Forms.Label.1", "lblA", True)
    With lblA
        .Caption = "Simulação A:"
        .Left = 12
        .Top = 56
        .Width = 90
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboSimulacaoA = Me.Controls.Add("Forms.ComboBox.1", "cboSimulacaoA", True)
    With cboSimulacaoA
        .Left = 110
        .Top = 56
        .Width = 350
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    '--- Seleção Simulação B ---------------------------------------------------
    Dim lblB As MSForms.Label
    Set lblB = Me.Controls.Add("Forms.Label.1", "lblB", True)
    With lblB
        .Caption = "Simulação B:"
        .Left = 480
        .Top = 56
        .Width = 90
        .Height = 22
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    
    Set cboSimulacaoB = Me.Controls.Add("Forms.ComboBox.1", "cboSimulacaoB", True)
    With cboSimulacaoB
        .Left = 580
        .Top = 56
        .Width = 350
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    
    '--- Botões -----------------------------------------------------------------
    Set btnComparar = Me.Controls.Add("Forms.CommandButton.1", "btnComparar", True)
    With btnComparar
        .Caption = "Comparar"
        .Left = 12
        .Top = 88
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnAtualizar = Me.Controls.Add("Forms.CommandButton.1", "btnAtualizar", True)
    With btnAtualizar
        .Caption = "Atualizar"
        .Left = 142
        .Top = 88
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnExportarComparacaoExcel = Me.Controls.Add("Forms.CommandButton.1", "btnExportarComparacaoExcel", True)
    With btnExportarComparacaoExcel
        .Caption = "Exportar Excel"
        .Left = 272
        .Top = 88
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnExportarComparacaoPDF = Me.Controls.Add("Forms.CommandButton.1", "btnExportarComparacaoPDF", True)
    With btnExportarComparacaoPDF
        .Caption = "Exportar PDF"
        .Left = 402
        .Top = 88
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AMARELO
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set btnImprimirComparacao = Me.Controls.Add("Forms.CommandButton.1", "btnImprimirComparacao", True)
    With btnImprimirComparacao
        .Caption = "Imprimir"
        .Left = 532
        .Top = 88
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = "Fechar"
        .Left = 662
        .Top = 88
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Labels dos títulos das áreas -------------------------------------------
    Set lblTituloA = Me.Controls.Add("Forms.Label.1", "lblTimelineA", True)
    With lblTituloA
        .Caption = "08:00 | 09:00 | 10:00 | 11:00 | 12:00 | 13:00 | 14:00 | 15:00 | 16:00 | 17:00"
        .Left = 12
        .Top = 128
        .Width = 470
        .Height = 24
        .Font.Size = 8
        .Font.Bold = False
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
        .TextAlign = fmTextAlignCenter
    End With
    
    Set lblTituloB = Me.Controls.Add("Forms.Label.1", "lblTimelineB", True)
    With lblTituloB
        .Caption = "08:00 | 09:00 | 10:00 | 11:00 | 12:00 | 13:00 | 14:00 | 15:00 | 16:00 | 17:00"
        .Left = 500
        .Top = 128
        .Width = 470
        .Height = 24
        .Font.Size = 8
        .Font.Bold = False
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Frames das áreas ------------------------------------------------------
    Set fraA = Me.Controls.Add("Forms.Frame.1", "fraA", True)
    With fraA
        .Caption = ""
        .Left = 12
        .Top = 156
        .Width = 470
        .Height = 380
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set fraB = Me.Controls.Add("Forms.Frame.1", "fraB", True)
    With fraB
        .Caption = ""
        .Left = 500
        .Top = 156
        .Width = 470
        .Height = 380
        .BackColor = COR_FUNDO
        .BorderStyle = fmBorderStyleSingle
    End With
    
    '--- Scrolls horizontais ----------------------------------------------------
    Set hScrollA = Me.Controls.Add("Forms.ScrollBar.1", "hScrollA", True)
    With hScrollA
        .Left = 12
        .Top = 540
        .Width = 470
        .Height = 16
        .Min = 0
        .Max = 1000
        .Value = 0
    End With
    
    Set hScrollB = Me.Controls.Add("Forms.ScrollBar.1", "hScrollB", True)
    With hScrollB
        .Left = 500
        .Top = 540
        .Width = 470
        .Height = 16
        .Min = 0
        .Max = 1000
        .Value = 0
    End With
    
    '--- Resumo -----------------------------------------------------------------
    Set lblResumo = Me.Controls.Add("Forms.Label.1", "lblResumo", True)
    With lblResumo
        .Caption = ""
        .Left = 12
        .Top = 580
        .Width = 960
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
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
        cboSimulacaoA.Clear
        cboSimulacaoB.Clear
        Exit Sub
    End If
    
    cboSimulacaoA.Clear
    cboSimulacaoB.Clear
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        texto = CStr(dados(i, 1)) & " - " & CStr(dados(i, 2))
        cboSimulacaoA.AddItem texto
        cboSimulacaoB.AddItem texto
    Next i
    
Sair:
    Exit Sub
    
ErroCarregarLista:
    MsgBox "Erro ao carregar simulações: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' EVENTO: btnComparar_Click
'================================================================================
Private Sub btnComparar_Click()
    On Error GoTo ErroComparar
    
    If cboSimulacaoA.ListIndex = -1 Or cboSimulacaoB.ListIndex = -1 Then
        MsgBox "Selecione as duas simulações para comparar.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    m_ID_SimulacaoA = Split(cboSimulacaoA.List(cboSimulacaoA.ListIndex), " - ")(0)
    m_ID_SimulacaoB = Split(cboSimulacaoB.List(cboSimulacaoB.ListIndex), " - ")(0)
    
    If m_ID_SimulacaoA = m_ID_SimulacaoB Then
        MsgBox "Selecione simulações diferentes para comparar.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Call ExecutarComparacao
    
Sair:
    Exit Sub
    
ErroComparar:
    MsgBox "Erro ao comparar: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' EVENTO: btnAtualizar_Click
'================================================================================
Private Sub btnAtualizar_Click()
    On Error Resume Next
    Call CarregarListaSimulacoes
    If m_ID_SimulacaoA <> "" And m_ID_SimulacaoB <> "" Then
        Call ExecutarComparacao
    End If
End Sub

'================================================================================
' EVENTO: btnFechar_Click
'================================================================================
Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub btnExportarComparacaoExcel_Click()
    On Error Resume Next
    Call ExportarComparacaoExcel(m_ID_SimulacaoA, m_ID_SimulacaoB)
End Sub

Private Sub btnExportarComparacaoPDF_Click()
    On Error Resume Next
    Call ExportarComparacaoPDF(m_ID_SimulacaoA, m_ID_SimulacaoB)
End Sub

Private Sub btnImprimirComparacao_Click()
    On Error Resume Next
    Call ImprimirComparacao(m_ID_SimulacaoA, m_ID_SimulacaoB)
End Sub

'================================================================================
' SUBROTINA PRIVADA: ExecutarComparacao
'================================================================================
Private Sub ExecutarComparacao()
    On Error GoTo ErroExecutar
    
    Dim resultados As Collection
    Set resultados = CompararSimulacoes(m_ID_SimulacaoA, m_ID_SimulacaoB)
    
    Call LimparAreas
    Call RenderizarResultado(resultados)
    Call AtualizarResumo(resultados)
    
Sair:
    Exit Sub
    
ErroExecutar:
    MsgBox "Erro ao executar comparação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: LimparAreas
'================================================================================
Private Sub LimparAreas()
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    
    Call ResetarPosicoesScroll
    
    For Each ctrl In fraA.Controls
        fraA.Controls.Remove ctrl.Name
    Next ctrl
    
    For Each ctrl In fraB.Controls
        fraB.Controls.Remove ctrl.Name
    Next ctrl
    
    If Not hScrollA Is Nothing Then hScrollA.Value = 0
    If Not hScrollB Is Nothing Then hScrollB.Value = 0
    m_ScrollPos = 0
    m_PosicoesSalvas = False
End Sub

'================================================================================
' SUBROTINA PRIVADA: RenderizarResultado
'================================================================================
Private Sub RenderizarResultado(pResultados As Collection)
    On Error GoTo ErroRenderizar
    
    Dim i As Long
    Dim topA As Single
    Dim topB As Single
    topA = 10
    topB = 10
    
    Dim corFundoA As Long
    Dim corFundoB As Long
    Dim corBordaA As Long
    Dim corBordaB As Long
    
    Dim totalA As Long
    Dim totalB As Long
    Dim totalIguais As Long
    totalA = 0
    totalB = 0
    totalIguais = 0
    
    For i = 1 To pResultados.Count
        Dim item As clsComparacaoOP
        Set item = pResultados(i)
        
        ' Contadores
        If item.ExisteA Then totalA = totalA + 1
        If item.ExisteB Then totalB = totalB + 1
        If item.TipoDiferenca = "IGUAL" Then totalIguais = totalIguais + 1
        
        ' Cores por tipo
        If item.TipoDiferenca = "IGUAL" Then
            corFundoA = vbWhite
            corFundoB = vbWhite
            corBordaA = COR_VERDE
            corBordaB = COR_VERDE
        ElseIf item.TipoDiferenca = "ALTERADA" Then
            corFundoA = vbWhite
            corFundoB = 10092543
            corBordaA = COR_ALERTA
            corBordaB = COR_ALERTA
        ElseIf item.TipoDiferenca = "SOMENTE_A" Then
            corFundoA = vbWhite
            corFundoB = COR_FUNDO
            corBordaA = COR_AZUL
            corBordaB = COR_FUNDO
        ElseIf item.TipoDiferenca = "SOMENTE_B" Then
            corFundoA = COR_FUNDO
            corFundoB = vbWhite
            corBordaA = COR_FUNDO
            corBordaB = COR_AZUL
        End If
        
        ' Renderiza card A
        If item.ExisteA Then
            Call CriarCardComparacao(fraA, item.ID_OP, item, True, topA, corFundoA, corBordaA)
            topA = topA + 52
        End If
        
        ' Renderiza card B
        If item.ExisteB Then
            Call CriarCardComparacao(fraB, item.ID_OP, item, False, topB, corFundoB, corBordaB)
            topB = topB + 52
        End If
    Next i
    
    ' Ajusta scroll vertical se necessário
    If topA > fraA.Height Then
        fraA.ScrollBars = fmScrollBarsVertical
    Else
        fraA.ScrollBars = fmScrollBarsNone
    End If
    
    If topB > fraB.Height Then
        fraB.ScrollBars = fmScrollBarsVertical
    Else
        fraB.ScrollBars = fmScrollBarsNone
    End If
    
    Call ResetarPosicoesScroll
    If Not hScrollA Is Nothing Then hScrollA.Value = 0
    If Not hScrollB Is Nothing Then hScrollB.Value = 0
    m_ScrollPos = 0
    m_PosicoesSalvas = False
    
Sair:
    Exit Sub
    
ErroRenderizar:
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarCardComparacao
'================================================================================
Private Sub CriarCardComparacao(pContainer As MSForms.Frame, pID_OP As String, _
                                pItem As clsComparacaoOP, pIsA As Boolean, _
                                pTop As Single, pCorFundo As Long, pCorBorda As Long)
    On Error Resume Next
    
    Dim lbl As MSForms.Label
    Dim uniqueID As String
    uniqueID = Format(Now, "SSSSS") & "_" & CStr(Int(Rnd * 100000))
    
    Set lbl = pContainer.Controls.Add("Forms.Label.1", "lblComp_" & uniqueID, True)
    With lbl
        .Left = 10
        .Top = pTop
        .Width = pContainer.Width - 20
        .Height = 42
        .BackColor = pCorFundo
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Dim lblTexto As MSForms.Label
    Set lblTexto = pContainer.Controls.Add("Forms.Label.1", "lblComp_" & uniqueID & "_T", True)
    With lblTexto
        .Caption = pID_OP & " | "
        If pIsA Then
            .Caption = .Caption & "Posto: " & pItem.EquipamentoA & " | "
            .Caption = .Caption & "Início: " & Format(pItem.DataInicioA, "dd/mm/yyyy HH:MM") & " | "
            .Caption = .Caption & "Fim: " & Format(pItem.DataFimA, "dd/mm/yyyy HH:MM") & " | "
            .Caption = .Caption & "Duração: " & Format(pItem.DuracaoA, "0.0") & "h | "
            .Caption = .Caption & "Status: " & pItem.StatusA
        Else
            .Caption = .Caption & "Posto: " & pItem.EquipamentoB & " | "
            .Caption = .Caption & "Início: " & Format(pItem.DataInicioB, "dd/mm/yyyy HH:MM") & " | "
            .Caption = .Caption & "Fim: " & Format(pItem.DataFimB, "dd/mm/yyyy HH:MM") & " | "
            .Caption = .Caption & "Duração: " & Format(pItem.DuracaoB, "0.0") & "h | "
            .Caption = .Caption & "Status: " & pItem.StatusB
        End If
        .Left = 16
        .Top = pTop + 8
        .Width = pContainer.Width - 32
        .Height = 26
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = pCorFundo
    End With
    
    If pItem.TipoDiferenca = "ALTERADA" Or pItem.TipoDiferenca = "SOMENTE_A" Or pItem.TipoDiferenca = "SOMENTE_B" Then
        Dim lblDiff As MSForms.Label
        Set lblDiff = pContainer.Controls.Add("Forms.Label.1", "lblComp_" & uniqueID & "_D", True)
        With lblDiff
            .Caption = pItem.Observacao
            .Left = 16
            .Top = pTop + 28
            .Width = pContainer.Width - 32
            .Height = 12
            .Font.Size = 8
            .Font.Bold = True
            .ForeColor = COR_ALERTA
            .BackColor = pCorFundo
        End With
    End If
End Sub

'================================================================================
' SUBROTINA PRIVADA: AtualizarResumo
'================================================================================
Private Sub AtualizarResumo(pResultados As Collection)
    On Error Resume Next
    
    Dim iguais As Long
    Dim alteradas As Long
    Dim soA As Long
    Dim soB As Long
    Dim i As Long
    
    iguais = 0
    alteradas = 0
    soA = 0
    soB = 0
    
    For i = 1 To pResultados.Count
        Select Case pResultados(i).TipoDiferenca
            Case "IGUAL": iguais = iguais + 1
            Case "ALTERADA": alteradas = alteradas + 1
            Case "SOMENTE_A": soA = soA + 1
            Case "SOMENTE_B": soB = soB + 1
        End Select
    Next i
    
    lblResumo.Caption = "Total: " & pResultados.Count & " | Iguais: " & iguais & _
                        " | Alteradas: " & alteradas & _
                        " | Somente A: " & soA & _
                        " | Somente B: " & soB
End Sub

'================================================================================
' EVENTOS DE SCROLL HORIZONTAL SINCRONIZADO
'================================================================================
Private Sub hScrollA_Change()
    On Error GoTo ErroScrollA
    
    If m_SincronizandoScroll Then Exit Sub
    If fraA Is Nothing Or fraB Is Nothing Then Exit Sub
    If hScrollA Is Nothing Or hScrollB Is Nothing Then Exit Sub
    
    m_SincronizandoScroll = True
    
    Dim valor As Single
    valor = hScrollA.Value
    m_ScrollPos = valor
    
    Call AplicarScrollHorizontal(valor)
    hScrollB.Value = valor
    
    m_SincronizandoScroll = False
    
Sair:
    Exit Sub
    
ErroScrollA:
    m_SincronizandoScroll = False
    Resume Sair
End Sub

Private Sub hScrollB_Change()
    On Error GoTo ErroScrollB
    
    If m_SincronizandoScroll Then Exit Sub
    If fraA Is Nothing Or fraB Is Nothing Then Exit Sub
    If hScrollA Is Nothing Or hScrollB Is Nothing Then Exit Sub
    
    m_SincronizandoScroll = True
    
    Dim valor As Single
    valor = hScrollB.Value
    m_ScrollPos = valor
    
    Call AplicarScrollHorizontal(valor)
    hScrollA.Value = valor
    
    m_SincronizandoScroll = False
    
Sair:
    Exit Sub
    
ErroScrollB:
    m_SincronizandoScroll = False
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: AplicarScrollHorizontal
' PROPÓSITO: Mover todos os controles internos de fraA e fraB conforme offset
'================================================================================
Private Sub AplicarScrollHorizontal(pOffset As Single)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    
    For Each ctrl In fraA.Controls
        If m_PosicoesSalvas Then
            ctrl.Left = ctrl.Tag - pOffset
        Else
            ctrl.Tag = ctrl.Left
            ctrl.Left = ctrl.Left - pOffset
        End If
    Next ctrl
    
    For Each ctrl In fraB.Controls
        If m_PosicoesSalvas Then
            ctrl.Left = ctrl.Tag - pOffset
        Else
            ctrl.Tag = ctrl.Left
            ctrl.Left = ctrl.Left - pOffset
        End If
    Next ctrl
    
    ' Move timelines sincronizadas
    If Not lblTituloA Is Nothing Then
        If m_PosicoesSalvas Then
            lblTituloA.Left = lblTituloA.Tag - pOffset
        Else
            lblTituloA.Tag = lblTituloA.Left
            lblTituloA.Left = lblTituloA.Left - pOffset
        End If
    End If
    
    If Not lblTituloB Is Nothing Then
        If m_PosicoesSalvas Then
            lblTituloB.Left = lblTituloB.Tag - pOffset
        Else
            lblTituloB.Tag = lblTituloB.Left
            lblTituloB.Left = lblTituloB.Left - pOffset
        End If
    End If
    
    m_PosicoesSalvas = True
End Sub

'================================================================================
' SUBROTINA PRIVADA: ResetarPosicoesScroll
' PROPÓSITO: Restaurar posições originais dos controles
'================================================================================
Private Sub ResetarPosicoesScroll()
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    
    For Each ctrl In fraA.Controls
        If IsNumeric(ctrl.Tag) Then
            ctrl.Left = ctrl.Tag
        End If
    Next ctrl
    
    For Each ctrl In fraB.Controls
        If IsNumeric(ctrl.Tag) Then
            ctrl.Left = ctrl.Tag
        End If
    Next ctrl
    
    If Not lblTituloA Is Nothing Then
        If IsNumeric(lblTituloA.Tag) Then lblTituloA.Left = lblTituloA.Tag
    End If
    
    If Not lblTituloB Is Nothing Then
        If IsNumeric(lblTituloB.Tag) Then lblTituloB.Left = lblTituloB.Tag
    End If
    
    m_ScrollPos = 0
    m_PosicoesSalvas = False
End Sub

'================================================================================
' EVENTO: UserForm_QueryClose
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Call SalvarPreferencias
        Unload Me
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyF5 Then
        If m_ID_SimulacaoA <> "" And m_ID_SimulacaoB <> "" Then
            Call ExecutarComparacao
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
        Call btnExportarComparacaoExcel_Click
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
        Call btnImprimirComparacao_Click
        KeyCode = 0
    End If
End Sub

'================================================================================
' SUBROTINAS DE PERSISTÊNCIA DE PREFERÊNCIAS
'================================================================================

Private Sub CarregarPreferencias()
    On Error Resume Next
    
    Dim config As Object
    Set config = CarregarTodasConfiguracoes
    
    If config.Count = 0 Then Exit Sub
    
    If config.Exists("ComparacaoSimA") Then
        Dim idA As String
        idA = CStr(config("ComparacaoSimA"))
        Dim dados As Variant
        dados = ObterDadosSimulacaoEmArray()
        If Not IsError(dados) Then
            Dim i As Long
            Dim encontradaA As Boolean
            encontradaA = False
            For i = 2 To UBound(dados, 1)
                If CStr(dados(i, 1)) = idA Then
                    encontradaA = True
                    Exit For
                End If
            Next i
            If encontradaA Then m_ID_SimulacaoA = idA
        End If
    End If
    
    If config.Exists("ComparacaoSimB") Then
        Dim idB As String
        idB = CStr(config("ComparacaoSimB"))
        Dim dadosB As Variant
        dadosB = ObterDadosSimulacaoEmArray()
        If Not IsError(dadosB) Then
            Dim j As Long
            Dim encontradaB As Boolean
            encontradaB = False
            For j = 2 To UBound(dadosB, 1)
                If CStr(dadosB(j, 1)) = idB Then
                    encontradaB = True
                    Exit For
                End If
            Next j
            If encontradaB Then m_ID_SimulacaoB = idB
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub SalvarPreferencias()
    On Error Resume Next
    
    If m_ID_SimulacaoA <> "" Then
        Call SalvarConfiguracao("ComparacaoSimA", m_ID_SimulacaoA)
    End If
    If m_ID_SimulacaoB <> "" Then
        Call SalvarConfiguracao("ComparacaoSimB", m_ID_SimulacaoB)
    End If
    
    On Error GoTo 0
End Sub
