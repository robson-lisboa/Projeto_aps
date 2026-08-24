VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 12632256
  BorderStyle = 3
  Caption = "APS PURAN – Registro de Eventos e Paradas"
  ClientHeight = 420
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 520
  Height = 458
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 532
  StartUpPosition = 1
  Attribute VB_Name = "frmEventos"
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
Attribute VB_Name = "frmEventos"
'================================================================================
' USERFORM: frmEventos
' DESCRIÇÃO: Registro de ocorrências industriais (manutenções, paradas, etc.)
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES DO FORMULÁRIO ===================================================
Private txtID_Evento As MSForms.TextBox
Private cboTipo As MSForms.ComboBox
Private cboEquipamento As MSForms.ComboBox
Private txtInicio As MSForms.TextBox
Private txtFim As MSForms.TextBox
Private txtMotivo As MSForms.TextBox

Private btnSalvar As MSForms.CommandButton
Private btnCancelar As MSForms.CommandButton

'=== PROPRIEDADES VISUAIS CORPORATIVAS =========================================
Private Const COR_FUNDO As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_ALERTA As Long = 255
Private Const COR_AZUL As Long = 15773696

Private Const ESPACAMENTO As Single = 10
Private Const LARGURA_ROTULO As Single = 100
Private Const ALTURA_CAMPO As Single = 22
Private Const LARGURA_CAMPO As Single = 220

'================================================================================
' EVENTO: UserForm_Initialize
' PROPÓSITO: Criar controles dinamicamente e inicializar valores padrão
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Dim i As Long
    Dim posY As Single
    
    Me.BackColor = COR_FUNDO
    Me.Caption = "APS PURAN – Registro de Eventos e Paradas"
    
    '--- Cabeçalho do formulário ------------------------------------------------
    Dim hdr As MSForms.Label
    Set hdr = Me.Controls.Add("Forms.Label.1", "lblEventosHdr", True)
    With hdr
        .Caption = "REGISTRO DE EVENTOS E PARADAS"
        .Left = 12
        .Top = 12
        .Width = 490
        .Height = 32
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Campo: ID_Evento -------------------------------------------------------
    posY = 60
    CriaRotulo "lblID_Evento", "ID Evento:", 12, posY
    Set txtID_Evento = CriaCampoTexto("txtID_Evento", 120, posY)
    
    '--- Campo: Tipo ------------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblTipo", "Tipo:", 12, posY
    Set cboTipo = CriaCampoCombo("cboTipo", 120, posY)
    
    '--- Campo: Equipamento -----------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblEquipamento", "Equipamento:", 12, posY
    Set cboEquipamento = CriaCampoCombo("cboEquipamento", 120, posY)
    
    '--- Campo: Início ----------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblInicio", "Início:", 12, posY
    Set txtInicio = CriaCampoTexto("txtInicio", 120, posY)
    
    '--- Campo: Fim -------------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblFim", "Fim:", 12, posY
    Set txtFim = CriaCampoTexto("txtFim", 120, posY)
    
    '--- Campo: Motivo ----------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblMotivo", "Motivo:", 12, posY
    Set txtMotivo = CriaCampoTexto("txtMotivo", 120, posY)
    txtMotivo.Width = 340
    
    '--- Botões de ação --------------------------------------------------------
    posY = posY + 50
    
    Set btnSalvar = Me.Controls.Add("Forms.CommandButton.1", "btnSalvar", True)
    With btnSalvar
        .Caption = "Salvar"
        .Left = 120
        .Top = posY
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar", True)
    With btnCancelar
        .Caption = "Cancelar"
        .Left = 260
        .Top = posY
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Preenche combos -------------------------------------------------------
    With cboTipo
        .AddItem "Manutenção"
        .AddItem "Parada"
        .AddItem "Refeição"
        .AddItem "Outros"
        .Value = "Manutenção"
    End With
    
    ' Carrega equipamentos da TabelaEquipamentos
    CarregarEquipamentos
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar formulário de eventos: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Eventos"
    Resume Sair
End Sub

'================================================================================
' EVENTO: btnSalvar_Click
' PROPÓSITO: Validar campos e persistir evento na TabelaEventos
'================================================================================
Private Sub btnSalvar_Click()
    On Error GoTo ErroSalvar
    
    Dim idEvento As String
    Dim tipo As String
    Dim equipamento As String
    Dim inicio As Date
    Dim fim As Date
    Dim motivo As String
    
    '--- Validação de campos obrigatórios --------------------------------------
    idEvento = Trim(Me.txtID_Evento.Value)
    tipo = Trim(Me.cboTipo.Value)
    equipamento = Trim(Me.cboEquipamento.Value)
    motivo = Trim(Me.txtMotivo.Value)
    
    If idEvento = "" Then
        MsgBox "Informe o ID do Evento.", vbExclamation, "Validação"
        Me.txtID_Evento.SetFocus
        Exit Sub
    End If
    
    If tipo = "" Then
        MsgBox "Selecione o Tipo do evento.", vbExclamation, "Validação"
        Me.cboTipo.SetFocus
        Exit Sub
    End If
    
    If equipamento = "" Then
        MsgBox "Selecione o Equipamento.", vbExclamation, "Validação"
        Me.cboEquipamento.SetFocus
        Exit Sub
    End If
    
    If motivo = "" Then
        MsgBox "Informe o Motivo do evento.", vbExclamation, "Validação"
        Me.txtMotivo.SetFocus
        Exit Sub
    End If
    
    '--- Converte datas --------------------------------------------------------
    If Not IsDate(Me.txtInicio.Value) Then
        MsgBox "Informe uma data/hora de Início válida.", vbExclamation, "Validação"
        Me.txtInicio.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.txtFim.Value) Then
        MsgBox "Informe uma data/hora de Fim válida.", vbExclamation, "Validação"
        Me.txtFim.SetFocus
        Exit Sub
    End If
    
    inicio = CDate(Me.txtInicio.Value)
    fim = CDate(Me.txtFim.Value)
    
    If fim < inicio Then
        MsgBox "A Data/Hora de Fim não pode ser anterior à de Início.", vbExclamation, "Validação"
        Me.txtFim.SetFocus
        Exit Sub
    End If
    
    '--- Persiste na TabelaEventos ---------------------------------------------
    If ID_EventoExiste(idEvento) Then
        MsgBox "Já existe um evento cadastrado com o ID informado.", vbExclamation, "Validação"
        Me.txtID_Evento.SetFocus
        Exit Sub
    End If
    
    SalvarEvento idEvento, tipo, equipamento, inicio, fim, motivo
    
    '--- Feedback e fechamento -------------------------------------------------
    MsgBox "Evento registrado com sucesso!", vbInformation, "APS PURAN – Eventos"
    Unload Me
    
Sair:
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar evento: " & Err.Description, vbCritical, "APS PURAN – Eventos"
    Resume Sair
End Sub

'================================================================================
' EVENTO: btnCancelar_Click
' PROPÓSITO: Fechar formulário sem salvar alterações
'================================================================================
Private Sub btnCancelar_Click()
    On Error Resume Next
    Unload Me
End Sub

'================================================================================
' FUNÇÃO PRIVADA: ID_EventoExiste
' PROPÓSITO: Verificar se um ID_Evento já está cadastrado na TabelaEventos
' RETORNO: True se existir, False caso contrário
'================================================================================
Private Function ID_EventoExiste(pID As String) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroVerificar
    
    Set ws = ThisWorkbook.Worksheets("BD_Eventos")
    Set tbl = ws.ListObjects("TabelaEventos")
    
    dados = tbl.Range.Value
    
    If IsError(dados) Then
        ID_EventoExiste = False
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID) Then
            ID_EventoExiste = True
            Exit Function
        End If
    Next i
    
    ID_EventoExiste = False
    Exit Function
    
ErroVerificar:
    ID_EventoExiste = False
End Function

'================================================================================
' SUBROTINA PRIVADA: CarregarEquipamentos
' PROPÓSITO: Popular ComboBox de equipamentos a partir da TabelaEquipamentos
'================================================================================
Private Sub CarregarEquipamentos()
    On Error GoTo ErroCarregar
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    Set ws = Nothing
    Set tbl = Nothing
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_Equipamentos")
    On Error GoTo ErroCarregar
    
    If ws Is Nothing Then
        MsgBox "Planilha BD_Equipamentos não encontrada.", vbExclamation, "APS PURAN – Eventos"
        Exit Sub
    End If
    
    On Error Resume Next
    Set tbl = ws.ListObjects("TabelaEquipamentos")
    On Error GoTo ErroCarregar
    
    If tbl Is Nothing Then
        MsgBox "Tabela TabelaEquipamentos não encontrada.", vbExclamation, "APS PURAN – Eventos"
        Exit Sub
    End If
    
    dados = tbl.Range.Value
    
    If IsError(dados) Then Exit Sub
    
    totalLinhas = UBound(dados, 1)
    
    Me.cboEquipamento.Clear
    
    For i = 2 To totalLinhas
        Dim nomeEq As String
        nomeEq = CStr(dados(i, 2))
        If Trim(nomeEq) <> "" Then
            Me.cboEquipamento.AddItem nomeEq
        End If
    Next i
    
Sair:
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar lista de equipamentos: " & Err.Description, vbExclamation, "APS PURAN – Eventos"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: SalvarEvento
' PROPÓSITO: Inserir novo registro na TabelaEventos da planilha oculta BD_Eventos
'================================================================================
Private Sub SalvarEvento(pID As String, _
                         pTipo As String, _
                         pEquipamento As String, _
                         pInicio As Date, _
                         pFim As Date, _
                         pMotivo As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroSalvarEvento
    
    Set ws = ThisWorkbook.Worksheets("BD_Eventos")
    Set tbl = ws.ListObjects("TabelaEventos")
    
    Application.ScreenUpdating = False
    
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_Evento").Index).Value = pID
        .Cells(1, tbl.ListColumns("Tipo").Index).Value = pTipo
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Inicio").Index).Value = pInicio
        .Cells(1, tbl.ListColumns("Fim").Index).Value = pFim
        .Cells(1, tbl.ListColumns("Motivo").Index).Value = pMotivo
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroSalvarEvento:
    Application.ScreenUpdating = True
    Err.Raise vbObjectError + 600, "SalvarEvento", _
        "Erro ao gravar evento na TabelaEventos. " & Err.Description
End Sub

'================================================================================
' FUNÇÕES AUXILIARES DE CRIAÇÃO DE CONTROLES
'================================================================================

Private Function CriaRotulo(pNome As String, pCaption As String, _
                            pLeft As Single, pTop As Single) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", pNome, True)
    With lbl
        .Caption = pCaption
        .Left = pLeft
        .Top = pTop
        .Width = LARGURA_ROTULO
        .Height = ALTURA_CAMPO
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
        .TextAlign = fmTextAlignRight
    End With
    Set CriaRotulo = lbl
End Function

Private Function CriaCampoTexto(pNome As String, pLeft As Single, _
                                pTop As Single) As MSForms.TextBox
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", pNome, True)
    With txt
        .Left = pLeft
        .Top = pTop
        .Width = LARGURA_CAMPO
        .Height = ALTURA_CAMPO
        .Font.Size = 9
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
    Set CriaCampoTexto = txt
End Function

Private Function CriaCampoCombo(pNome As String, pLeft As Single, _
                                pTop As Single) As MSForms.ComboBox
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls.Add("Forms.ComboBox.1", pNome, True)
    With cbo
        .Left = pLeft
        .Top = pTop
        .Width = LARGURA_CAMPO
        .Height = ALTURA_CAMPO
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
    Set CriaCampoCombo = cbo
End Function
