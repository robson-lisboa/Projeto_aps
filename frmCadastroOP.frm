VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 12632256
  BorderStyle = 3
  Caption = "APS PURAN – Cadastro de Ordem de Produção"
  ClientHeight = 420
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 500
  Height = 458
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 512
  StartUpPosition = 1
  Attribute VB_Name = "frmCadastroOP"
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
Attribute VB_Name = "frmCadastroOP"
'================================================================================
' USERFORM: frmCadastroOP
' DESCRIÇÃO: Cadastro e edição de Ordens de Produção (OP) do APS PURAN
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES DO FORMULÁRIO ===================================================
Private txtID_OP As MSForms.TextBox
Private txtProduto As MSForms.TextBox
Private cboEquipamento As MSForms.ComboBox
Private txtQuantidade As MSForms.TextBox
Private txtData_Inicio As MSForms.TextBox
Private txtData_Fim As MSForms.TextBox
Private cboStatus As MSForms.ComboBox

Private btnSalvar As MSForms.CommandButton
Private btnCancelar As MSForms.CommandButton

Private m_btnSalvarEvents As clsButtonEvents
Private m_btnCancelarEvents As clsButtonEvents

Private m_EditandoID As String

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
Private Const LARGURA_CAMPO As Single = 200

'================================================================================
' EVENTO: UserForm_Initialize
' PROPÓSITO: Criar controles dinamicamente e inicializar valores padrão
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Dim i As Long
    Dim posY As Single
    
    Me.BackColor = COR_FUNDO
    Me.Caption = "APS PURAN – Cadastro de Ordem de Produção"
    
    '--- Cabeçalho do formulário ------------------------------------------------
    Dim hdr As MSForms.Label
    Set hdr = Me.Controls.Add("Forms.Label.1", "lblCadastroHdr", True)
    With hdr
        .Caption = "CADASTRO DE ORDEM DE PRODUÇÃO"
        .Left = 12
        .Top = 12
        .Width = 470
        .Height = 32
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Campo: ID_OP ----------------------------------------------------------
    posY = 60
    CriaRotulo "lblID_OP", "ID OP:", 12, posY
    Set txtID_OP = CriaCampoTexto("txtID_OP", 120, posY)
    
    '--- Campo: Produto --------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblProduto", "Produto:", 12, posY
    Set txtProduto = CriaCampoTexto("txtProduto", 120, posY)
    
    '--- Campo: Equipamento ----------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblEquipamento", "Equipamento:", 12, posY
    Set cboEquipamento = CriaCampoCombo("cboEquipamento", 120, posY)
    
    '--- Campo: Quantidade -----------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblQuantidade", "Quantidade:", 12, posY
    Set txtQuantidade = CriaCampoTexto("txtQuantidade", 120, posY)
    
    '--- Campo: Data Início ----------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblData_Inicio", "Data Início:", 12, posY
    Set txtData_Inicio = CriaCampoTexto("txtData_Inicio", 120, posY)
    
    '--- Campo: Data Fim -------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblData_Fim", "Data Fim:", 12, posY
    Set txtData_Fim = CriaCampoTexto("txtData_Fim", 120, posY)
    
    '--- Campo: Status ---------------------------------------------------------
    posY = posY + 36
    CriaRotulo "lblStatus", "Status:", 12, posY
    Set cboStatus = CriaCampoCombo("cboStatus", 120, posY)
    
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
    
    Set m_btnSalvarEvents = New clsButtonEvents
    Set m_btnSalvarEvents.Button = btnSalvar
    
    Set m_btnCancelarEvents = New clsButtonEvents
    Set m_btnCancelarEvents.Button = btnCancelar
    
    '--- Preenche combos -------------------------------------------------------
    With Me.Controls("cboStatus")
        .AddItem "Planejada"
        .AddItem "Em Andamento"
        .AddItem "Concluído"
        .AddItem "Atrasado"
        .Value = "Planejada"
    End With
    
    CarregarEquipamentos
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar formulário de cadastro: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Cadastro"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarParaEdicao
' PROPÓSITO: Carregar dados de uma OP existente para edição
' PARÂMETROS: pID_OP As String – ID da OP a ser editada
'================================================================================
Public Sub CarregarParaEdicao(pID_OP As String)
    On Error GoTo ErroCarregarEdicao
    
    m_EditandoID = pID_OP
    Me.Caption = "APS PURAN – Editar Ordem de Produção: " & pID_OP
    
    Me.Controls("txtID_OP").Enabled = False
    
    Dim dados() As Variant
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then Exit Sub
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If CStr(dados(i, 1)) = pID_OP Then
            Me.Controls("txtID_OP").Value = CStr(dados(i, 1))
            Me.Controls("txtProduto").Value = CStr(dados(i, 2))
            Me.Controls("cboEquipamento").Value = CStr(dados(i, 3))
            Me.Controls("txtQuantidade").Value = CStr(dados(i, 4))
            Me.Controls("txtData_Inicio").Value = Format(CDate(dados(i, 5)), "dd/mm/yyyy")
            Me.Controls("txtData_Fim").Value = Format(CDate(dados(i, 6)), "dd/mm/yyyy")
            Me.Controls("cboStatus").Value = CStr(dados(i, 8))
            Exit For
        End If
    Next i
    
Sair:
    Exit Sub
    
ErroCarregarEdicao:
    MsgBox "Erro ao carregar OP para edição: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' EVENTO: btnSalvar_Click
' PROPÓSITO: Validar campos e persistir nova OP via modEngine
'================================================================================
Private Sub btnSalvar_Click()
    On Error GoTo ErroSalvar
    
    Dim idOP As String
    Dim produto As String
    Dim equipamento As String
    Dim quantidade As Long
    Dim dataInicio As Date
    Dim dataFim As Date
    Dim duracao As Double
    Dim status As String
    
    '--- Validação de campos obrigatórios --------------------------------------
    idOP = Trim(Me.Controls("txtID_OP").Value)
    produto = Trim(Me.Controls("txtProduto").Value)
    equipamento = Trim(Me.Controls("cboEquipamento").Value)
    status = Trim(Me.Controls("cboStatus").Value)
    
    If idOP = "" Then
        MsgBox "Informe o ID da OP.", vbExclamation, "Validação"
        Me.Controls("txtID_OP").SetFocus
        Exit Sub
    End If
    
    If produto = "" Then
        MsgBox "Informe o Produto.", vbExclamation, "Validação"
        Me.Controls("txtProduto").SetFocus
        Exit Sub
    End If
    
    If equipamento = "" Then
        MsgBox "Selecione o Equipamento.", vbExclamation, "Validação"
        Me.Controls("cboEquipamento").SetFocus
        Exit Sub
    End If
    
    '--- Converte valores numéricos e datas -------------------------------------
    If Not IsNumeric(Me.Controls("txtQuantidade").Value) Or CLng(Me.Controls("txtQuantidade").Value) <= 0 Then
        MsgBox "Informe uma Quantidade válida maior que zero.", vbExclamation, "Validação"
        Me.Controls("txtQuantidade").SetFocus
        Exit Sub
    End If
    
    quantidade = CLng(Me.Controls("txtQuantidade").Value)
    
    If Not IsDate(Me.Controls("txtData_Inicio").Value) Then
        MsgBox "Informe uma Data de Início válida.", vbExclamation, "Validação"
        Me.Controls("txtData_Inicio").SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.Controls("txtData_Fim").Value) Then
        MsgBox "Informe uma Data de Fim válida.", vbExclamation, "Validação"
        Me.Controls("txtData_Fim").SetFocus
        Exit Sub
    End If
    
    dataInicio = CDate(Me.Controls("txtData_Inicio").Value)
    dataFim = CDate(Me.Controls("txtData_Fim").Value)
    
    If dataFim < dataInicio Then
        MsgBox "A Data de Fim não pode ser anterior à Data de Início.", vbExclamation, "Validação"
        Me.Controls("txtData_Fim").SetFocus
        Exit Sub
    End If
    
    '--- Cálculo da duração em horas -------------------------------------------
    duracao = (dataFim - dataInicio) * 24
    
    '--- Valida duplicidade de ID_OP apenas para nova OP ------------------------
    If m_EditandoID = "" Then
        If ID_OPExiste(idOP) Then
            MsgBox "Já existe uma OP cadastrada com o ID informado.", vbExclamation, "Validação"
            Me.Controls("txtID_OP").SetFocus
            Exit Sub
        End If
    End If
    
    '--- Persiste dados via modEngine ------------------------------------------
    If m_EditandoID = "" Then
        Call SalvarNovaOP(idOP, produto, equipamento, quantidade, dataInicio, dataFim, duracao, status)
        MsgBox "Ordem de Produção cadastrada com sucesso!", vbInformation, "APS PURAN – Cadastro"
    Else
        Call AtualizarOP(m_EditandoID, produto, equipamento, quantidade, dataInicio, dataFim, duracao, status)
        MsgBox "Ordem de Produção atualizada com sucesso!", vbInformation, "APS PURAN – Cadastro"
    End If
    
    Unload Me
    
Sair:
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar OP: " & Err.Description, vbCritical, "APS PURAN – Cadastro"
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

Private Sub m_btnSalvarEvents_Clicked()
    btnSalvar_Click
End Sub

Private Sub m_btnCancelarEvents_Clicked()
    btnCancelar_Click
End Sub

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
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_Equipamentos")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set tbl = ws.ListObjects("TabelaEquipamentos")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub
    
    dados = tbl.Range.Value
    
    If IsError(dados) Then Exit Sub
    
    totalLinhas = UBound(dados, 1)
    
    Me.Controls("cboEquipamento").Clear
    
    For i = 2 To totalLinhas
        Dim nomeEq As String
        nomeEq = CStr(dados(i, 2))
        If Trim(nomeEq) <> "" Then
            Me.Controls("cboEquipamento").AddItem nomeEq
        End If
    Next i
    
Sair:
    Exit Sub
    
ErroCarregar:
    Resume Sair
End Sub

'================================================================================
' FUNÇÃO PRIVADA: ID_OPExiste
' PROPÓSITO: Verificar se um ID_OP já está cadastrado na TabelaOPs
' RETORNO: True se existir, False caso contrário
'================================================================================
Private Function ID_OPExiste(pID As String) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroVerificar
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    dados = tbl.Range.Value
    
    If IsError(dados) Then
        ID_OPExiste = False
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pID) Then
            ID_OPExiste = True
            Exit Function
        End If
    Next i
    
    ID_OPExiste = False
    Exit Function
    
ErroVerificar:
    ID_OPExiste = False
End Function

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
