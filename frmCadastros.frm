VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Cadastros"
  ClientHeight = 520
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 820
  Height = 558
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 832
  StartUpPosition = 1
  Attribute VB_Name = "frmCadastros"
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
Attribute VB_Name = "frmCadastros"
'================================================================================
' USERFORM: frmCadastros
' DESCRIÇÃO: Módulo de cadastros gerais (MultiPage)
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES PRINCIPAIS ======================================================
Private mpCadastros As MSForms.MultiPage
Private lblTitulo As MSForms.Label

'=== BOTÕES DE NAVEGAÇÃO DO MENU LATERAL ======================================
Private btnProdutos As MSForms.CommandButton
Private btnPostos As MSForms.CommandButton
Private btnOperadores As MSForms.CommandButton
Private btnTurnos As MSForms.CommandButton
Private btnProcessos As MSForms.CommandButton
Private btnMotivosParada As MSForms.CommandButton

'=== CONTROLES COMUNS ==========================================================
Private lstRegistros As MSForms.ListBox
Private btnNovo As MSForms.CommandButton
Private btnSalvar As MSForms.CommandButton
Private btnExcluir As MSForms.CommandButton
Private btnCancelar As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

'=== CONTROLES DE EDICAO (Dinâmicos por página) ================================
Private txtCampo1 As MSForms.TextBox
Private txtCampo2 As MSForms.TextBox
Private txtCampo3 As MSForms.TextBox
Private txtCampo4 As MSForms.TextBox
Private txtCampo5 As MSForms.TextBox
Private cboCampo1 As MSForms.ComboBox
Private cboCampo2 As MSForms.ComboBox
Private lblCampo1 As MSForms.Label
Private lblCampo2 As MSForms.Label
Private lblCampo3 As MSForms.Label
Private lblCampo4 As MSForms.Label
Private lblCampo5 As MSForms.Label

'=== ESTADO ====================================================================
Private m_PaginaAtiva As String
Private m_EditandoID As String

'=== PROPRIEDADES VISUAIS ======================================================
Private Const COR_FUNDO As Long = 14211288
Private Const COR_HEADER As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_AZUL As Long = 15773696
Private Const COR_VERDE As Long = 5287936
Private Const COR_ALERTA As Long = 255

Private Const ESPACAMENTO As Single = 10
Private Const LARGURA_ROTULO As Single = 120
Private Const ALTURA_CAMPO As Single = 22
Private Const LARGURA_CAMPO As Single = 240

'================================================================================
' EVENTO: UserForm_Initialize
'================================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Me.BackColor = COR_FUNDO
    Me.Caption = "APS PURAN – Cadastros"
    
    Call CriarControles
    Call ExibirPagina("Produtos")
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar cadastros: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarControles
' PROPÓSITO: Criar todos os controles dinamicamente
'================================================================================
Private Sub CriarControles()
    On Error Resume Next
    
    Dim i As Long
    Dim posY As Single
    
    '--- Título do formulário ---------------------------------------------------
    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    With lblTitulo
        .Caption = "CADASTROS DO SISTEMA"
        .Left = 12
        .Top = 12
        .Width = 780
        .Height = 32
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Menu lateral -----------------------------------------------------------
    Dim menuLeft As Single
    menuLeft = 12
    posY = 60
    
    Dim botoes As Variant
    botoes = Array("btnProdutos", "btnPostos", "btnOperadores", "btnTurnos", "btnProcessos", "btnMotivosParada")
    
    For i = LBound(botoes) To UBound(botoes)
        Dim btn As MSForms.CommandButton
        Set btn = Me.Controls.Add("Forms.CommandButton.1", botoes(i), True)
        
        btn.Caption = _
            IIf(i = 0, "Produtos", _
            IIf(i = 1, "Postos", _
            IIf(i = 2, "Operadores", _
            IIf(i = 3, "Turnos", _
            IIf(i = 4, "Processos", _
            IIf(i = 5, "Motivos Parada", ""))))))
        btn.Left = menuLeft
        btn.Top = posY
        btn.Width = 140
        btn.Height = 36
        btn.Font.Size = 10
        btn.Font.Bold = True
        btn.BackColor = COR_HEADER
        btn.ForeColor = COR_TEXTO_CLARO
        
        posY = posY + 44
        
        Select Case botoes(i)
            Case "btnProdutos": Set btnProdutos = btn
            Case "btnPostos": Set btnPostos = btn
            Case "btnOperadores": Set btnOperadores = btn
            Case "btnTurnos": Set btnTurnos = btn
            Case "btnProcessos": Set btnProcessos = btn
            Case "btnMotivosParada": Set btnMotivosParada = btn
        End Select
    Next i
    
    '--- MultiPage --------------------------------------------------------------
    Set mpCadastros = Me.Controls.Add("Forms.MultiPage.1", "mpCadastros", True)
    With mpCadastros
        .Left = 168
        .Top = 60
        .Width = 620
        .Height = 380
        .Font.Size = 10
    End With
    
    ' Cria páginas
    Dim paginas As Variant
    paginas = Array("Produtos", "Postos", "Operadores", "Turnos", "Processos", "MotivosParada")
    
    For i = LBound(paginas) To UBound(paginas)
        Dim pg As MSForms.Page
        Set pg = mpCadastros.Pages.Add()
        pg.Name = "page_" & paginas(i)
        pg.Caption = paginas(i)
    Next i
    
    '--- Lista de registros -----------------------------------------------------
    Set lstRegistros = Me.Controls.Add("Forms.ListBox.1", "lstRegistros", True)
    With lstRegistros
        .Left = 168
        .Top = 60
        .Width = 620
        .Height = 380
        .Font.Size = 9
        .BorderStyle = fmBorderStyleSingle
    End With
    
    '--- Botões de ação ---------------------------------------------------------
    Dim btnLeft As Single
    btnLeft = 168
    
    Set btnNovo = Me.Controls.Add("Forms.CommandButton.1", "btnNovo", True)
    With btnNovo
        .Caption = "Novo"
        .Left = btnLeft
        .Top = 460
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnSalvar = Me.Controls.Add("Forms.CommandButton.1", "btnSalvar", True)
    With btnSalvar
        .Caption = "Salvar"
        .Left = btnLeft
        .Top = 460
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnExcluir = Me.Controls.Add("Forms.CommandButton.1", "btnExcluir", True)
    With btnExcluir
        .Caption = "Excluir"
        .Left = btnLeft
        .Top = 460
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar", True)
    With btnCancelar
        .Caption = "Cancelar"
        .Left = btnLeft
        .Top = 460
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 10092543
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Botão fechar -----------------------------------------------------------
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = "Fechar"
        .Left = 660
        .Top = 460
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
End Sub

'================================================================================
' SUBROTINA PRIVADA: ExibirPagina
' PROPÓSITO: Configurar a página ativa do MultiPage
' PARÂMETROS: pNomePagina As String
'================================================================================
Private Sub ExibirPagina(pNomePagina As String)
    On Error Resume Next
    
    m_PaginaAtiva = pNomePagina
    m_EditandoID = ""
    
    Dim i As Long
    For i = 0 To mpCadastros.Pages.Count - 1
        If mpCadastros.Pages(i).Name = "page_" & pNomePagina Then
            mpCadastros.Value = i
            Exit For
        End If
    Next i
    
    Call CarregarLista
    Call LimparCampos
End Sub

'================================================================================
' SUBROTINA PRIVADA: CarregarLista
' PROPÓSITO: Carregar registros da tabela correspondente à página ativa
'================================================================================
Private Sub CarregarLista()
    On Error GoTo ErroCarregarLista
    
    Dim nomePlanilha As String
    Dim nomeTabela As String
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim texto As String
    
    nomePlanilha = ObterNomePlanilha(m_PaginaAtiva)
    nomeTabela = ObterNomeTabela(m_PaginaAtiva)
    
    If nomePlanilha = "" Or nomeTabela = "" Then Exit Sub
    
    dados = ObterDadosTabelaEmArray(nomePlanilha, nomeTabela)
    
    If IsError(dados) Then Exit Sub
    
    lstRegistros.Clear
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        texto = CStr(dados(i, 1))
        If m_PaginaAtiva = "Produtos" Then
            texto = texto & " - " & CStr(dados(i, 2))
        ElseIf m_PaginaAtiva = "Postos" Then
            texto = texto & " - " & CStr(dados(i, 2))
        ElseIf m_PaginaAtiva = "Operadores" Then
            texto = texto & " - " & CStr(dados(i, 2))
        ElseIf m_PaginaAtiva = "Turnos" Then
            texto = texto & " - " & CStr(dados(i, 2))
        ElseIf m_PaginaAtiva = "Processos" Then
            texto = texto & " - " & CStr(dados(i, 2))
        ElseIf m_PaginaAtiva = "MotivosParada" Then
            texto = texto & " - " & CStr(dados(i, 2))
        End If
        lstRegistros.AddItem texto
    Next i
    
Sair:
    Exit Sub
    
ErroCarregarLista:
    MsgBox "Erro ao carregar lista: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: LimparCampos
' PROPÓSITO: Limpar campos de edição
'================================================================================
Private Sub LimparCampos()
    On Error Resume Next
    
    m_EditandoID = ""
    
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is MSForms.ComboBox Then ctrl.Value = ""
    Next ctrl
End Sub

'================================================================================
' FUNÇÃO PRIVADA: ObterNomePlanilha
' PROPÓSITO: Retornar nome da planilha conforme página ativa
'================================================================================
Private Function ObterNomePlanilha(pPagina As String) As String
    Select Case pPagina
        Case "Produtos": ObterNomePlanilha = "BD_Produtos"
        Case "Postos": ObterNomePlanilha = "BD_Equipamentos"
        Case "Operadores": ObterNomePlanilha = "BD_Operadores"
        Case "Turnos": ObterNomePlanilha = "BD_Turnos"
        Case "Processos": ObterNomePlanilha = "BD_Processos"
        Case "MotivosParada": ObterNomePlanilha = "BD_MotivosParada"
        Case Else: ObterNomePlanilha = ""
    End Select
End Function

'================================================================================
' FUNÇÃO PRIVADA: ObterNomeTabela
' PROPÓSITO: Retornar nome da tabela conforme página ativa
'================================================================================
Private Function ObterNomeTabela(pPagina As String) As String
    Select Case pPagina
        Case "Produtos": ObterNomeTabela = "TabelaProdutos"
        Case "Postos": ObterNomeTabela = "TabelaEquipamentos"
        Case "Operadores": ObterNomeTabela = "TabelaOperadores"
        Case "Turnos": ObterNomeTabela = "TabelaTurnos"
        Case "Processos": ObterNomeTabela = "TabelaProcessos"
        Case "MotivosParada": ObterNomeTabela = "TabelaMotivosParada"
        Case Else: ObterNomeTabela = ""
    End Select
End Function

'================================================================================
' EVENTOS DOS BOTÕES DE NAVEGAÇÃO =============================================
'================================================================================

Private Sub btnProdutos_Click()
    Call ExibirPagina("Produtos")
End Sub

Private Sub btnPostos_Click()
    Call ExibirPagina("Postos")
End Sub

Private Sub btnOperadores_Click()
    Call ExibirPagina("Operadores")
End Sub

Private Sub btnTurnos_Click()
    Call ExibirPagina("Turnos")
End Sub

Private Sub btnProcessos_Click()
    Call ExibirPagina("Processos")
End Sub

Private Sub btnMotivosParada_Click()
    Call ExibirPagina("MotivosParada")
End Sub

Private Sub btnNovo_Click()
    Call LimparCampos
End Sub

Private Sub btnSalvar_Click()
    On Error GoTo ErroSalvar
    
    Dim nomePlanilha As String
    Dim nomeTabela As String
    Dim valores As Variant
    
    nomePlanilha = ObterNomePlanilha(m_PaginaAtiva)
    nomeTabela = ObterNomeTabela(m_PaginaAtiva)
    
    If nomePlanilha = "" Or nomeTabela = "" Then Exit Sub
    
    ' Coleta valores conforme página
    valores = ColetarValores()
    If IsError(valores) Then Exit Sub
    
    If m_EditandoID = "" Then
        Call InserirLinhaTabela(nomePlanilha, nomeTabela, valores)
        MsgBox "Registro cadastrado com sucesso!", vbInformation, "APS PURAN"
    Else
        Call AtualizarLinhaTabela(nomePlanilha, nomeTabela, ObterColunaChave(m_PaginaAtiva), m_EditandoID, valores)
        MsgBox "Registro atualizado com sucesso!", vbInformation, "APS PURAN"
    End If
    
    Call CarregarLista
    Call LimparCampos
    
Sair:
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnExcluir_Click()
    On Error GoTo ErroExcluir
    
    If m_EditandoID = "" Then
        MsgBox "Selecione um registro para excluir.", vbExclamation, "Validação"
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente excluir o registro selecionado?", vbQuestion + vbYesNo, "Confirmação")
    If resposta = vbNo Then Exit Sub
    
    Dim nomePlanilha As String
    Dim nomeTabela As String
    
    nomePlanilha = ObterNomePlanilha(m_PaginaAtiva)
    nomeTabela = ObterNomeTabela(m_PaginaAtiva)
    
    Call ExcluirLinhaTabela(nomePlanilha, nomeTabela, ObterColunaChave(m_PaginaAtiva), m_EditandoID)
    MsgBox "Registro excluído com sucesso!", vbInformation, "APS PURAN"
    
    Call CarregarLista
    Call LimparCampos
    
Sair:
    Exit Sub
    
ErroExcluir:
    MsgBox "Erro ao excluir: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnCancelar_Click()
    Call LimparCampos
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub lstRegistros_Click()
    On Error Resume Next
    
    If lstRegistros.ListIndex = -1 Then Exit Sub
    
    Dim nomePlanilha As String
    Dim nomeTabela As String
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim chave As String
    
    nomePlanilha = ObterNomePlanilha(m_PaginaAtiva)
    nomeTabela = ObterNomeTabela(m_PaginaAtiva)
    
    dados = ObterDadosTabelaEmArray(nomePlanilha, nomeTabela)
    If IsError(dados) Then Exit Sub
    
    totalLinhas = UBound(dados, 1)
    chave = Split(lstRegistros.List(lstRegistros.ListIndex), " - ")(0)
    
    For i = 2 To totalLinhas
        If CStr(dados(i, 1)) = chave Then
            m_EditandoID = chave
            Call PreencherCampos(dados, i)
            Exit For
        End If
    Next i
End Sub

'================================================================================
' FUNÇÕES AUXILIARES ===========================================================
'================================================================================

Private Function ColetarValores() As Variant
    On Error GoTo ErroColetar
    
    Dim valores(1 To 5) As Variant
    
    Select Case m_PaginaAtiva
        Case "Produtos"
            valores(1) = txtCampo1.Text ' ID_Produto
            valores(2) = txtCampo2.Text ' Nome
            valores(3) = txtCampo3.Text ' Descricao
            valores(4) = CDbl(txtCampo4.Text) ' Tempo_Padrao_Horas
            valores(5) = cboCampo1.Value ' Status
        Case "Postos"
            valores(1) = txtCampo1.Text ' ID_Equipamento
            valores(2) = txtCampo2.Text ' Nome
            valores(3) = CDbl(txtCampo3.Text) ' Capacidade_Hora
            valores(4) = cboCampo1.Value ' Status_Manutencao
            valores(5) = "" ' Campo extra não usado
        Case "Operadores"
            valores(1) = txtCampo1.Text ' ID_Operador
            valores(2) = txtCampo2.Text ' Nome
            valores(3) = cboCampo1.Value ' Status
            valores(4) = ""
            valores(5) = ""
        Case "Turnos"
            valores(1) = txtCampo1.Text ' ID_Turno
            valores(2) = txtCampo2.Text ' Nome
            valores(3) = txtCampo3.Text ' Hora_Inicio
            valores(4) = txtCampo4.Text ' Hora_Fim
            valores(5) = txtCampo5.Text ' Dias_Funcionamento
        Case "Processos"
            valores(1) = txtCampo1.Text ' ID_Processo
            valores(2) = cboCampo1.Value ' Produto
            valores(3) = cboCampo2.Value ' Posto
            valores(4) = CLng(txtCampo3.Text) ' Sequencia
            valores(5) = CDbl(txtCampo4.Text) ' Tempo_Padrao_Horas
        Case "MotivosParada"
            valores(1) = txtCampo1.Text ' ID_Motivo
            valores(2) = txtCampo2.Text ' Descricao
            valores(3) = cboCampo1.Value ' Categoria
            valores(4) = cboCampo2.Value ' Status
            valores(5) = ""
    End Select
    
    ColetarValores = valores
    Exit Function
    
ErroColetar:
    ColetarValores = CVErr(xlErrRef)
End Function

Private Sub PreencherCampos(pDados As Variant, pLinha As Long)
    On Error Resume Next
    
    Select Case m_PaginaAtiva
        Case "Produtos"
            txtCampo1.Text = CStr(pDados(pLinha, 1))
            txtCampo2.Text = CStr(pDados(pLinha, 2))
            txtCampo3.Text = CStr(pDados(pLinha, 3))
            txtCampo4.Text = CStr(pDados(pLinha, 4))
            cboCampo1.Value = CStr(pDados(pLinha, 5))
        Case "Postos"
            txtCampo1.Text = CStr(pDados(pLinha, 1))
            txtCampo2.Text = CStr(pDados(pLinha, 2))
            txtCampo3.Text = CStr(pDados(pLinha, 3))
            cboCampo1.Value = CStr(pDados(pLinha, 4))
        Case "Operadores"
            txtCampo1.Text = CStr(pDados(pLinha, 1))
            txtCampo2.Text = CStr(pDados(pLinha, 2))
            cboCampo1.Value = CStr(pDados(pLinha, 3))
        Case "Turnos"
            txtCampo1.Text = CStr(pDados(pLinha, 1))
            txtCampo2.Text = CStr(pDados(pLinha, 2))
            txtCampo3.Text = CStr(pDados(pLinha, 3))
            txtCampo4.Text = CStr(pDados(pLinha, 4))
            txtCampo5.Text = CStr(pDados(pLinha, 5))
        Case "Processos"
            txtCampo1.Text = CStr(pDados(pLinha, 1))
            cboCampo1.Value = CStr(pDados(pLinha, 2))
            cboCampo2.Value = CStr(pDados(pLinha, 3))
            txtCampo3.Text = CStr(pDados(pLinha, 4))
            txtCampo4.Text = CStr(pDados(pLinha, 5))
        Case "MotivosParada"
            txtCampo1.Text = CStr(pDados(pLinha, 1))
            txtCampo2.Text = CStr(pDados(pLinha, 2))
            cboCampo1.Value = CStr(pDados(pLinha, 3))
            cboCampo2.Value = CStr(pDados(pLinha, 4))
    End Select
End Sub

Private Function ObterColunaChave(pPagina As String) As String
    Select Case pPagina
        Case "Produtos": ObterColunaChave = "ID_Produto"
        Case "Postos": ObterColunaChave = "ID_Equipamento"
        Case "Operadores": ObterColunaChave = "ID_Operador"
        Case "Turnos": ObterColunaChave = "ID_Turno"
        Case "Processos": ObterColunaChave = "ID_Processo"
        Case "MotivosParada": ObterColunaChave = "ID_Motivo"
        Case Else: ObterColunaChave = ""
    End Select
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
