VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Detalhes da Ordem de Produção"
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
  Attribute VB_Name = "frmDetalheOP"
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
Attribute VB_Name = "frmDetalheOP"
'================================================================================
' USERFORM: frmDetalheOP
' DESCRIÇÃO: Detalhes da Ordem de Produção (visualização e edição)
' VERSÃO: 1.0
'================================================================================
Option Explicit

Private m_ID_OP As String
Private btnEditar As MSForms.CommandButton
Private btnExcluir As MSForms.CommandButton
Private btnDuplicar As MSForms.CommandButton
Private btnSalvar As MSForms.CommandButton
Private btnCancelar As MSForms.CommandButton

' Campos de edição
Private txtID_OP As MSForms.TextBox
Private txtProduto As MSForms.TextBox
Private txtEquipamento As MSForms.TextBox
Private txtQuantidade As MSForms.TextBox
Private txtDataInicio As MSForms.TextBox
Private txtDataFim As MSForms.TextBox
Private txtDuracao As MSForms.TextBox
Private cboStatus As MSForms.ComboBox
Private txtPrioridade As MSForms.TextBox
Private txtObservacao As MSForms.TextBox

' Estado de edição
Private m_Editando As Boolean
Private m_SimulacaoAtiva As String
Private m_LabelDetalhe As MSForms.Label

Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Me.BackColor = 14211288
    Me.Caption = "Detalhes da OP"
    
    Call CriarControlesEdicao
    Call ModoVisualizacao
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar detalhes: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub CriarControlesEdicao()
    On Error Resume Next
    
    Dim leftCampo As Single
    Dim topCampo As Single
    leftCampo = 20
    topCampo = 20
    
    ' ID OP
    Dim lblID As MSForms.Label
    Set lblID = Me.Controls.Add("Forms.Label.1", "lblID", True)
    With lblID
        .Caption = "ID OP:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Produto
    Dim lblProduto As MSForms.Label
    Set lblProduto = Me.Controls.Add("Forms.Label.1", "lblProduto", True)
    With lblProduto
        .Caption = "Produto:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Equipamento
    Dim lblEquip As MSForms.Label
    Set lblEquip = Me.Controls.Add("Forms.Label.1", "lblEquip", True)
    With lblEquip
        .Caption = "Posto:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Quantidade
    Dim lblQtd As MSForms.Label
    Set lblQtd = Me.Controls.Add("Forms.Label.1", "lblQtd", True)
    With lblQtd
        .Caption = "Quantidade:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Data Início
    Dim lblInicio As MSForms.Label
    Set lblInicio = Me.Controls.Add("Forms.Label.1", "lblInicio", True)
    With lblInicio
        .Caption = "Início:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Data Fim
    Dim lblFim As MSForms.Label
    Set lblFim = Me.Controls.Add("Forms.Label.1", "lblFim", True)
    With lblFim
        .Caption = "Fim:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Duração
    Dim lblDur As MSForms.Label
    Set lblDur = Me.Controls.Add("Forms.Label.1", "lblDur", True)
    With lblDur
        .Caption = "Duração:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Status
    Dim lblStatus As MSForms.Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Caption = "Status:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Prioridade
    Dim lblPrioridade As MSForms.Label
    Set lblPrioridade = Me.Controls.Add("Forms.Label.1", "lblPrioridade", True)
    With lblPrioridade
        .Caption = "Prioridade:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    topCampo = topCampo + 26
    
    ' Observação
    Dim lblObs As MSForms.Label
    Set lblObs = Me.Controls.Add("Forms.Label.1", "lblObs", True)
    With lblObs
        .Caption = "Observação:"
        .Left = leftCampo
        .Top = topCampo
        .Width = 80
        .Height = 22
        .Font.Size = 9
        .ForeColor = 0
        .BackColor = 14211288
        .TextAlign = fmTextAlignRight
    End With
    
    ' Campos de edição
    leftCampo = 110
    topCampo = 20
    
    Set txtID_OP = Me.Controls.Add("Forms.TextBox.1", "txtID_OP", True)
    With txtID_OP
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Enabled = False
    End With
    topCampo = topCampo + 26
    
    Set txtProduto = Me.Controls.Add("Forms.TextBox.1", "txtProduto", True)
    With txtProduto
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtEquipamento = Me.Controls.Add("Forms.TextBox.1", "txtEquipamento", True)
    With txtEquipamento
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtQuantidade = Me.Controls.Add("Forms.TextBox.1", "txtQuantidade", True)
    With txtQuantidade
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
    End With
    topCampo = topCampo + 26
    
    Set txtDataFim = Me.Controls.Add("Forms.TextBox.1", "txtDataFim", True)
    With txtDataFim
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtDuracao = Me.Controls.Add("Forms.TextBox.1", "txtDuracao", True)
    With txtDuracao
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set cboStatus = Me.Controls.Add("Forms.ComboBox.1", "cboStatus", True)
    With cboStatus
        .AddItem "Planejada"
        .AddItem "Em Andamento"
        .AddItem "Concluído"
        .AddItem "Atrasado"
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    topCampo = topCampo + 26
    
    Set txtPrioridade = Me.Controls.Add("Forms.TextBox.1", "txtPrioridade", True)
    With txtPrioridade
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 22
        .Font.Size = 9
    End With
    topCampo = topCampo + 26
    
    Set txtObservacao = Me.Controls.Add("Forms.TextBox.1", "txtObservacao", True)
    With txtObservacao
        .Left = leftCampo
        .Top = topCampo
        .Width = 360
        .Height = 60
        .Font.Size = 9
        .MultiLine = True
    End With
    
    ' Botões de ação
    Set btnEditar = Me.Controls.Add("Forms.CommandButton.1", "btnEditar", True)
    With btnEditar
        .Caption = "Editar"
        .Left = 20
        .Top = 340
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 15773696
        .ForeColor = 16777215
    End With
    
    Set btnSalvar = Me.Controls.Add("Forms.CommandButton.1", "btnSalvar", True)
    With btnSalvar
        .Caption = "Salvar"
        .Left = 130
        .Top = 340
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 5287936
        .ForeColor = 16777215
        .Visible = False
    End With
    
    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar", True)
    With btnCancelar
        .Caption = "Cancelar"
        .Left = 240
        .Top = 340
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 10092543
        .ForeColor = 16777215
        .Visible = False
    End With
    
    Set btnExcluir = Me.Controls.Add("Forms.CommandButton.1", "btnExcluir", True)
    With btnExcluir
        .Caption = "Excluir"
        .Left = 350
        .Top = 340
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 255
        .ForeColor = 16777215
    End With
End Sub

Public Sub CarregarDetalhes(pID_OP As String)
    On Error GoTo ErroCarregar
    
    m_ID_OP = pID_OP
    Me.Caption = "Detalhes da OP: " & pID_OP
    
    Dim dados() As Variant
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then Exit Sub
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If CStr(dados(i, 1)) = pID_OP Then
            Dim lbl As MSForms.Label
            Set lbl = Me.Controls.Add("Forms.Label.1", "lblDetalhe", True)
            With lbl
                .Caption = "ID OP: " & CStr(dados(i, 1)) & vbCrLf & _
                           "Produto: " & CStr(dados(i, 2)) & vbCrLf & _
                           "Equipamento: " & CStr(dados(i, 3)) & vbCrLf & _
                           "Quantidade: " & CStr(dados(i, 4)) & vbCrLf & _
                           "Início: " & Format(CDate(dados(i, 5)), "dd/mm/yyyy HH:MM") & vbCrLf & _
                           "Fim: " & Format(CDate(dados(i, 6)), "dd/mm/yyyy HH:MM") & vbCrLf & _
                           "Duração: " & Format(CDbl(dados(i, 7)), "0.0") & "h" & vbCrLf & _
                           "Status: " & CStr(dados(i, 8))
                .Left = 20
                .Top = 20
                .Width = 440
                .Height = 300
                .Font.Size = 10
            End With
            Exit For
        End If
    Next i
    
Sair:
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar detalhes: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnEditar_Click()
    On Error Resume Next
    frmCadastroOP.CarregarParaEdicao m_ID_OP
    frmCadastroOP.Show vbModal
    Unload Me
End Sub

Private Sub btnExcluir_Click()
    On Error GoTo ErroExcluir
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente excluir a OP " & m_ID_OP & "?", _
                      vbQuestion + vbYesNo, "Confirmação de Exclusão")
    If resposta = vbNo Then Exit Sub
    
    Call ExcluirOP(m_ID_OP)
    MsgBox "OP excluída com sucesso!", vbInformation, "APS PURAN"
    Unload Me
    
Sair:
    Exit Sub
    
ErroExcluir:
    MsgBox "Erro ao excluir OP: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnDuplicar_Click()
    On Error GoTo ErroDuplicar
    
    Dim dados() As Variant
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then Exit Sub
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    Dim idOP As String
    Dim produto As String
    Dim equipamento As String
    Dim quantidade As Long
    Dim dataInicio As Date
    Dim dataFim As Date
    Dim duracao As Double
    Dim status As String
    
    For i = 2 To totalLinhas
        If CStr(dados(i, 1)) = m_ID_OP Then
            idOP = "Copia_" & CStr(dados(i, 1))
            produto = CStr(dados(i, 2))
            equipamento = CStr(dados(i, 3))
            quantidade = CLng(dados(i, 4))
            dataInicio = CDate(dados(i, 5))
            dataFim = CDate(dados(i, 6))
            duracao = CDbl(dados(i, 7))
            status = CStr(dados(i, 8))
            Exit For
        End If
    Next i
    
    If idOP = "" Then Exit Sub
    
    frmCadastroOP.CarregarParaCopia idOP, produto, equipamento, quantidade, dataInicio, dataFim, status
    frmCadastroOP.Show vbModal
    Unload Me
    
Sair:
    Exit Sub
    
ErroDuplicar:
    MsgBox "Erro ao duplicar OP: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarDetalhes
' PROPÓSITO: Carregar dados de uma OP real
'================================================================================
Public Sub CarregarDetalhes(pID_OP As String)
    On Error GoTo ErroCarregar
    
    m_ID_OP = pID_OP
    m_SimulacaoAtiva = ""
    Me.Caption = "Detalhes da OP: " & pID_OP
    
    Dim dados() As Variant
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then Exit Sub
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If CStr(dados(i, 1)) = pID_OP Then
            txtID_OP.Text = CStr(dados(i, 1))
            txtProduto.Text = CStr(dados(i, 2))
            txtEquipamento.Text = CStr(dados(i, 3))
            txtQuantidade.Text = CStr(dados(i, 4))
            txtDataInicio.Text = Format(CDate(dados(i, 5)), "dd/mm/yyyy")
            txtDataFim.Text = Format(CDate(dados(i, 6)), "dd/mm/yyyy")
            txtDuracao.Text = Format(CDbl(dados(i, 7)), "0.0")
            cboStatus.Value = CStr(dados(i, 8))
            txtPrioridade.Text = ""
            txtObservacao.Text = ""
            Exit For
        End If
    Next i
    
    Call ModoVisualizacao
    
Sair:
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar detalhes: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarDetalhesSimulacao
' PROPÓSITO: Carregar dados de uma OP da simulação
'================================================================================
Public Sub CarregarDetalhesSimulacao(pID_OP_Simulacao As String, pID_Simulacao As String)
    On Error GoTo ErroCarregarSim
    
    m_ID_OP = pID_OP_Simulacao
    m_SimulacaoAtiva = pID_Simulacao
    Me.Caption = "Detalhes OP (Simulação): " & pID_OP_Simulacao
    
    Dim dados() As Variant
    dados = ObterOPsSimulacaoEmArray(pID_Simulacao)
    
    If IsError(dados) Then Exit Sub
    
    Dim i As Long
    Dim totalLinhas As Long
    totalLinhas = UBound(dados, 1)
    
    For i = 1 To totalLinhas
        If CStr(dados(i, 2)) = pID_OP_Simulacao And CStr(dados(i, 1)) = pID_Simulacao Then
            txtID_OP.Text = CStr(dados(i, 2))
            txtProduto.Text = CStr(dados(i, 4))
            txtEquipamento.Text = CStr(dados(i, 5))
            txtQuantidade.Text = CStr(dados(i, 6))
            txtDataInicio.Text = Format(CDate(dados(i, 7)), "dd/mm/yyyy")
            txtDataFim.Text = Format(CDate(dados(i, 8)), "dd/mm/yyyy")
            txtDuracao.Text = Format(CDbl(dados(i, 9)), "0.0")
            cboStatus.Value = CStr(dados(i, 10))
            txtPrioridade.Text = CStr(dados(i, 11))
            txtObservacao.Text = CStr(dados(i, 12))
            Exit For
        End If
    Next i
    
    Call ModoVisualizacao
    
Sair:
    Exit Sub
    
ErroCarregarSim:
    MsgBox "Erro ao carregar detalhes da simulação: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub ModoVisualizacao()
    On Error Resume Next
    
    m_Editando = False
    
    txtID_OP.Enabled = False
    txtProduto.Enabled = False
    txtEquipamento.Enabled = False
    txtQuantidade.Enabled = False
    txtDataInicio.Enabled = False
    txtDataFim.Enabled = False
    txtDuracao.Enabled = False
    cboStatus.Enabled = False
    txtPrioridade.Enabled = False
    txtObservacao.Enabled = False
    
    btnEditar.Visible = True
    btnSalvar.Visible = False
    btnCancelar.Visible = False
    btnExcluir.Visible = True
End Sub

Private Sub ModoEdicao()
    On Error Resume Next
    
    m_Editando = True
    
    txtID_OP.Enabled = False
    txtProduto.Enabled = True
    txtEquipamento.Enabled = True
    txtQuantidade.Enabled = True
    txtDataInicio.Enabled = True
    txtDataFim.Enabled = True
    txtDuracao.Enabled = True
    cboStatus.Enabled = True
    txtPrioridade.Enabled = True
    txtObservacao.Enabled = True
    
    btnEditar.Visible = False
    btnSalvar.Visible = True
    btnCancelar.Visible = True
    btnExcluir.Visible = False
End Sub

Private Sub btnEditar_Click()
    Call ModoEdicao
End Sub

Private Sub btnSalvar_Click()
    On Error GoTo ErroSalvar
    
    If m_SimulacaoAtiva = "" Then
        MsgBox "Edição de simulação não disponível.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim idOP As String
    Dim produto As String
    Dim equipamento As String
    Dim quantidade As Long
    Dim dataInicio As Date
    Dim dataFim As Date
    Dim duracao As Double
    Dim status As String
    Dim prioridade As String
    Dim observacao As String
    
    idOP = Trim(txtID_OP.Text)
    produto = Trim(txtProduto.Text)
    equipamento = Trim(txtEquipamento.Text)
    prioridade = Trim(txtPrioridade.Text)
    observacao = Trim(txtObservacao.Text)
    
    If produto = "" Then
        MsgBox "Informe o Produto.", vbExclamation, "Validação"
        txtProduto.SetFocus
        Exit Sub
    End If
    
    If equipamento = "" Then
        MsgBox "Informe o Posto/Equipamento.", vbExclamation, "Validação"
        txtEquipamento.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtQuantidade.Text) Or CLng(txtQuantidade.Text) <= 0 Then
        MsgBox "Informe uma Quantidade válida.", vbExclamation, "Validação"
        txtQuantidade.SetFocus
        Exit Sub
    End If
    quantidade = CLng(txtQuantidade.Text)
    
    If Not IsDate(txtDataInicio.Text) Then
        MsgBox "Informe uma Data de Início válida.", vbExclamation, "Validação"
        txtDataInicio.SetFocus
        Exit Sub
    End If
    dataInicio = CDate(txtDataInicio.Text)
    
    If Not IsDate(txtDataFim.Text) Then
        MsgBox "Informe uma Data de Fim válida.", vbExclamation, "Validação"
        txtDataFim.SetFocus
        Exit Sub
    End If
    dataFim = CDate(txtDataFim.Text)
    
    If dataFim < dataInicio Then
        MsgBox "A Data de Fim não pode ser anterior à Data de Início.", vbExclamation, "Validação"
        txtDataFim.SetFocus
        Exit Sub
    End If
    
    duracao = (dataFim - dataInicio) * 24
    txtDuracao.Text = Format(duracao, "0.0")
    
    status = Trim(cboStatus.Value)
    If status = "" Then status = "Planejada"
    
    Call AtualizarOPSimulacao(m_SimulacaoAtiva, m_ID_OP, produto, equipamento, quantidade, dataInicio, dataFim, duracao, status, prioridade, observacao)
    MsgBox "OP atualizada com sucesso!", vbInformation, "APS PURAN"
    
    Call ModoVisualizacao
    
Sair:
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnCancelar_Click()
    Call ModoVisualizacao
    If m_SimulacaoAtiva <> "" Then
        Call CarregarDetalhesSimulacao m_ID_OP, m_SimulacaoAtiva
    Else
        Call CarregarDetalhes m_ID_OP
    End If
End Sub

Private Sub btnExcluir_Click()
    On Error GoTo ErroExcluir
    
    If m_SimulacaoAtiva = "" Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Deseja realmente excluir a OP " & m_ID_OP & "?", vbQuestion + vbYesNo, "Confirmação")
        If resposta = vbNo Then Exit Sub
        
        Call ExcluirOP(m_ID_OP)
        MsgBox "OP excluída com sucesso!", vbInformation, "APS PURAN"
        Unload Me
    Else
        resposta = MsgBox("Deseja realmente excluir a OP da simulação?", vbQuestion + vbYesNo, "Confirmação")
        If resposta = vbNo Then Exit Sub
        
        Call ExcluirOPSimulacao(m_SimulacaoAtiva, m_ID_OP)
        MsgBox "OP excluída da simulação!", vbInformation, "APS PURAN"
        Unload Me
    End If
    
Sair:
    Exit Sub
    
ErroExcluir:
    MsgBox "Erro ao excluir: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
