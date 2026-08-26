VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Configurações"
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
  Attribute VB_Name = "frmConfiguracoes"
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
Attribute VB_Name = "frmConfiguracoes"
'================================================================================
' USERFORM: frmConfiguracoes
' DESCRIÇÃO: Configurações gerais do sistema
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES PRINCIPAIS ======================================================
Private mpConfiguracoes As MSForms.MultiPage

'=== BOTÕES =====================================================================
Private btnSalvar As MSForms.CommandButton
Private btnRestaurarPadrao As MSForms.CommandButton
Private btnExportarConfig As MSForms.CommandButton
Private btnImportarConfig As MSForms.CommandButton
Private btnLimparPreferencias As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton

'=== CAMPOS - GERAIS ============================================================
Private txtNomeSistema As MSForms.TextBox
Private txtEmpresa As MSForms.TextBox
Private cboFormatoData As MSForms.ComboBox
Private txtHorarioPadrao As MSForms.TextBox

'=== CAMPOS - PLANEJAMENTO ======================================================
Private txtZoomPadrao As MSForms.TextBox
Private cboPeriodoPadrao As MSForms.ComboBox
Private chkMostrarLinhaAgora As MSForms.CheckBox
Private chkMostrarConflitos As MSForms.CheckBox
Private txtEscalaTimeline As MSForms.TextBox
Private txtAlturaLinhaPosto As MSForms.TextBox

'=== CAMPOS - PRODUÇÃO ==========================================================
Private cboTurnoPadrao As MSForms.ComboBox
Private txtHorarioInicio As MSForms.TextBox
Private txtHorarioFim As MSForms.TextBox
Private txtToleranciaAtraso As MSForms.TextBox

'=== CAMPOS - INTERFACE =========================================================
Private cboEstiloCards As MSForms.ComboBox
Private txtAlturaCards As MSForms.TextBox
Private chkMostrarScrollHorizontal As MSForms.CheckBox

'=== ESTADO =====================================================================
Private m_Modificado As Boolean

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
    Me.Caption = "APS PURAN – Configurações"
    
    m_Modificado = False
    
    Call CriarControles
    Call CarregarConfiguracoes
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar configurações: " & Err.Description, vbCritical, "APS PURAN"
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
        .Caption = "CONFIGURAÇÕES DO SISTEMA"
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
    
    '--- MultiPage -------------------------------------------------------------
    Set mpConfiguracoes = Me.Controls.Add("Forms.TabStrip.1", "mpConfiguracoes", True)
    With mpConfiguracoes
        .Left = 12
        .Top = 56
        .Width = 780
        .Height = 380
        .Font.Size = 9
    End With
    
    ' Adiciona páginas
    Dim idxPage As Long
    For idxPage = 1 To 5
        mpConfiguracoes.Pages.Add
    Next idxPage
    
    '--- Botões de ação -------------------------------------------------------
    Dim btnLeft As Single
    btnLeft = 12
    
    Set btnSalvar = Me.Controls.Add("Forms.CommandButton.1", "btnSalvar", True)
    With btnSalvar
        .Caption = "Salvar"
        .Left = btnLeft
        .Top = 450
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_VERDE
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnRestaurarPadrao = Me.Controls.Add("Forms.CommandButton.1", "btnRestaurarPadrao", True)
    With btnRestaurarPadrao
        .Caption = "Restaurar Padrão"
        .Left = btnLeft
        .Top = 450
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AMARELO
        .ForeColor = COR_TEXTO_ESCURO
    End With
    btnLeft = btnLeft + 130
    
    Set btnExportarConfig = Me.Controls.Add("Forms.CommandButton.1", "btnExportarConfig", True)
    With btnExportarConfig
        .Caption = "Exportar"
        .Left = btnLeft
        .Top = 450
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_AZUL
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnImportarConfig = Me.Controls.Add("Forms.CommandButton.1", "btnImportarConfig", True)
    With btnImportarConfig
        .Caption = "Importar"
        .Left = btnLeft
        .Top = 450
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    btnLeft = btnLeft + 110
    
    Set btnLimparPreferencias = Me.Controls.Add("Forms.CommandButton.1", "btnLimparPreferencias", True)
    With btnLimparPreferencias
        .Caption = "Limpar Prefs"
        .Left = btnLeft
        .Top = 450
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Caption = "Fechar"
        .Left = 680
        .Top = 450
        .Width = 100
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_HEADER
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Cria páginas do MultiPage ---------------------------------------------
    Call CriarPaginaGerais
    Call CriarPaginaPlanejamento
    Call CriarPaginaProducao
    Call CriarPaginaInterface
    Call CriarPaginaBackup
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarPaginaGerais
'================================================================================
Private Sub CriarPaginaGerais()
    On Error Resume Next
    
    Dim lbl As MSForms.Label
    Dim yPos As Single
    yPos = 10
    
    Set lbl = mpConfiguracoes.Pages(1).Controls.Add("Forms.Label.1", "lblSecaoGerais", True)
    With lbl
        .Caption = "Configurações Gerais"
        .Left = 10
        .Top = yPos
        .Width = 400
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
    End With
    yPos = yPos + 28
    
    Set lbl = mpConfiguracoes.Pages(1).Controls.Add("Forms.Label.1", "lblNomeSistema", True)
    With lbl
        .Caption = "Nome do Sistema:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtNomeSistema = mpConfiguracoes.Pages(1).Controls.Add("Forms.TextBox.1", "txtNomeSistema", True)
    With txtNomeSistema
        .Text = "APS PURAN"
        .Left = 140
        .Top = yPos
        .Width = 300
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(1).Controls.Add("Forms.Label.1", "lblEmpresa", True)
    With lbl
        .Caption = "Empresa/Unidade:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtEmpresa = mpConfiguracoes.Pages(1).Controls.Add("Forms.TextBox.1", "txtEmpresa", True)
    With txtEmpresa
        .Text = ""
        .Left = 140
        .Top = yPos
        .Width = 300
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(1).Controls.Add("Forms.Label.1", "lblFormatoData", True)
    With lbl
        .Caption = "Formato de Data:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set cboFormatoData = mpConfiguracoes.Pages(1).Controls.Add("Forms.ComboBox.1", "cboFormatoData", True)
    With cboFormatoData
        .AddItem "dd/mm/yyyy"
        .AddItem "mm/dd/yyyy"
        .AddItem "yyyy-mm-dd"
        .Value = "dd/mm/yyyy"
        .Left = 140
        .Top = yPos
        .Width = 150
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(1).Controls.Add("Forms.Label.1", "lblHorarioPadrao", True)
    With lbl
        .Caption = "Horário Padrão:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtHorarioPadrao = mpConfiguracoes.Pages(1).Controls.Add("Forms.TextBox.1", "txtHorarioPadrao", True)
    With txtHorarioPadrao
        .Text = "08:00 - 17:00"
        .Left = 140
        .Top = yPos
        .Width = 120
        .Height = 22
        .Font.Size = 9
    End With
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarPaginaPlanejamento
'================================================================================
Private Sub CriarPaginaPlanejamento()
    On Error Resume Next
    
    Dim lbl As MSForms.Label
    Dim yPos As Single
    yPos = 10
    
    Set lbl = mpConfiguracoes.Pages(2).Controls.Add("Forms.Label.1", "lblSecaoPlanejamento", True)
    With lbl
        .Caption = "Configurações do Planejamento"
        .Left = 10
        .Top = yPos
        .Width = 400
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
    End With
    yPos = yPos + 28
    
    Set lbl = mpConfiguracoes.Pages(2).Controls.Add("Forms.Label.1", "lblZoomPadrao", True)
    With lbl
        .Caption = "Zoom Padrão (0.5 - 3.0):"
        .Left = 10
        .Top = yPos
        .Width = 150
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtZoomPadrao = mpConfiguracoes.Pages(2).Controls.Add("Forms.TextBox.1", "txtZoomPadrao", True)
    With txtZoomPadrao
        .Text = "1.0"
        .Left = 170
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(2).Controls.Add("Forms.Label.1", "lblPeriodoPadrao", True)
    With lbl
        .Caption = "Período Padrão:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set cboPeriodoPadrao = mpConfiguracoes.Pages(2).Controls.Add("Forms.ComboBox.1", "cboPeriodoPadrao", True)
    With cboPeriodoPadrao
        .AddItem "Dia"
        .AddItem "Semana"
        .AddItem "Mês"
        .Value = "Mês"
        .Left = 140
        .Top = yPos
        .Width = 120
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    yPos = yPos + 30
    
    Set chkMostrarLinhaAgora = mpConfiguracoes.Pages(2).Controls.Add("Forms.CheckBox.1", "chkMostrarLinhaAgora", True)
    With chkMostrarLinhaAgora
        .Caption = "Mostrar Linha AGORA"
        .Left = 10
        .Top = yPos
        .Width = 200
        .Height = 20
        .Font.Size = 9
        .Value = True
    End With
    yPos = yPos + 26
    
    Set chkMostrarConflitos = mpConfiguracoes.Pages(2).Controls.Add("Forms.CheckBox.1", "chkMostrarConflitos", True)
    With chkMostrarConflitos
        .Caption = "Mostrar Conflitos"
        .Left = 10
        .Top = yPos
        .Width = 200
        .Height = 20
        .Font.Size = 9
        .Value = True
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(2).Controls.Add("Forms.Label.1", "lblEscalaTimeline", True)
    With lbl
        .Caption = "Escala Timeline (px/hora):"
        .Left = 10
        .Top = yPos
        .Width = 160
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtEscalaTimeline = mpConfiguracoes.Pages(2).Controls.Add("Forms.TextBox.1", "txtEscalaTimeline", True)
    With txtEscalaTimeline
        .Text = "100"
        .Left = 180
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(2).Controls.Add("Forms.Label.1", "lblAlturaLinhaPosto", True)
    With lbl
        .Caption = "Altura Linha/Posto (px):"
        .Left = 10
        .Top = yPos
        .Width = 160
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtAlturaLinhaPosto = mpConfiguracoes.Pages(2).Controls.Add("Forms.TextBox.1", "txtAlturaLinhaPosto", True)
    With txtAlturaLinhaPosto
        .Text = "80"
        .Left = 180
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarPaginaProducao
'================================================================================
Private Sub CriarPaginaProducao()
    On Error Resume Next
    
    Dim lbl As MSForms.Label
    Dim yPos As Single
    yPos = 10
    
    Set lbl = mpConfiguracoes.Pages(3).Controls.Add("Forms.Label.1", "lblSecaoProducao", True)
    With lbl
        .Caption = "Configurações de Produção"
        .Left = 10
        .Top = yPos
        .Width = 400
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
    End With
    yPos = yPos + 28
    
    Set lbl = mpConfiguracoes.Pages(3).Controls.Add("Forms.Label.1", "lblTurnoPadrao", True)
    With lbl
        .Caption = "Turno Padrão:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set cboTurnoPadrao = mpConfiguracoes.Pages(3).Controls.Add("Forms.ComboBox.1", "cboTurnoPadrao", True)
    With cboTurnoPadrao
        .AddItem "Manhã"
        .AddItem "Tarde"
        .AddItem "Noite"
        .AddItem "Integral"
        .Value = "Manhã"
        .Left = 140
        .Top = yPos
        .Width = 150
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(3).Controls.Add("Forms.Label.1", "lblHorarioInicio", True)
    With lbl
        .Caption = "Horário Início:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtHorarioInicio = mpConfiguracoes.Pages(3).Controls.Add("Forms.TextBox.1", "txtHorarioInicio", True)
    With txtHorarioInicio
        .Text = "08:00"
        .Left = 140
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(3).Controls.Add("Forms.Label.1", "lblHorarioFim", True)
    With lbl
        .Caption = "Horário Fim:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtHorarioFim = mpConfiguracoes.Pages(3).Controls.Add("Forms.TextBox.1", "txtHorarioFim", True)
    With txtHorarioFim
        .Text = "17:00"
        .Left = 140
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(3).Controls.Add("Forms.Label.1", "lblToleranciaAtraso", True)
    With lbl
        .Caption = "Tolerância Atraso (min):"
        .Left = 10
        .Top = yPos
        .Width = 150
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtToleranciaAtraso = mpConfiguracoes.Pages(3).Controls.Add("Forms.TextBox.1", "txtToleranciaAtraso", True)
    With txtToleranciaAtraso
        .Text = "30"
        .Left = 170
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarPaginaInterface
'================================================================================
Private Sub CriarPaginaInterface()
    On Error Resume Next
    
    Dim lbl As MSForms.Label
    Dim yPos As Single
    yPos = 10
    
    Set lbl = mpConfiguracoes.Pages(4).Controls.Add("Forms.Label.1", "lblSecaoInterface", True)
    With lbl
        .Caption = "Configurações de Interface"
        .Left = 10
        .Top = yPos
        .Width = 400
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
    End With
    yPos = yPos + 28
    
    Set lbl = mpConfiguracoes.Pages(4).Controls.Add("Forms.Label.1", "lblEstiloCards", True)
    With lbl
        .Caption = "Estilo dos Cards:"
        .Left = 10
        .Top = yPos
        .Width = 120
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set cboEstiloCards = mpConfiguracoes.Pages(4).Controls.Add("Forms.ComboBox.1", "cboEstiloCards", True)
    With cboEstiloCards
        .AddItem "Padrão"
        .AddItem "Compacto"
        .AddItem "Detalhado"
        .Value = "Padrão"
        .Left = 140
        .Top = yPos
        .Width = 150
        .Height = 22
        .Font.Size = 9
        .Style = fmStyleDropDownList
    End With
    yPos = yPos + 30
    
    Set lbl = mpConfiguracoes.Pages(4).Controls.Add("Forms.Label.1", "lblAlturaCards", True)
    With lbl
        .Caption = "Altura dos Cards (px):"
        .Left = 10
        .Top = yPos
        .Width = 140
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    
    Set txtAlturaCards = mpConfiguracoes.Pages(4).Controls.Add("Forms.TextBox.1", "txtAlturaCards", True)
    With txtAlturaCards
        .Text = "80"
        .Left = 160
        .Top = yPos
        .Width = 80
        .Height = 22
        .Font.Size = 9
    End With
    yPos = yPos + 30
    
    Set chkMostrarScrollHorizontal = mpConfiguracoes.Pages(4).Controls.Add("Forms.CheckBox.1", "chkMostrarScrollHorizontal", True)
    With chkMostrarScrollHorizontal
        .Caption = "Mostrar Scroll Horizontal"
        .Left = 10
        .Top = yPos
        .Width = 220
        .Height = 20
        .Font.Size = 9
        .Value = True
    End With
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarPaginaBackup
'================================================================================
Private Sub CriarPaginaBackup()
    On Error Resume Next
    
    Dim lbl As MSForms.Label
    Dim yPos As Single
    yPos = 10
    
    Set lbl = mpConfiguracoes.Pages(5).Controls.Add("Forms.Label.1", "lblSecaoBackup", True)
    With lbl
        .Caption = "Backup e Manutenção"
        .Left = 10
        .Top = yPos
        .Width = 400
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
    End With
    yPos = yPos + 28
    
    Set lbl = mpConfiguracoes.Pages(5).Controls.Add("Forms.Label.1", "lblInfoBackup", True)
    With lbl
        .Caption = "Exportar/Importar configurações e preferências do sistema."
        .Left = 10
        .Top = yPos
        .Width = 600
        .Height = 40
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
    End With
    yPos = yPos + 50
    
    Set lbl = mpConfiguracoes.Pages(5).Controls.Add("Forms.Label.1", "lblAvisoBackup", True)
    With lbl
        .Caption = "Atenção: A restauração de valores padrão ou limpeza de preferências não afeta dados operacionais."
        .Left = 10
        .Top = yPos
        .Width = 600
        .Height = 40
        .Font.Size = 9
        .ForeColor = COR_ALERTA
    End With
End Sub

'================================================================================
' EVENTOS DOS BOTÕES
'================================================================================
Private Sub btnSalvar_Click()
    On Error GoTo ErroSalvar
    
    Call SalvarConfiguracoes
    m_Modificado = False
    MsgBox "Configurações salvas com sucesso!", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar configurações: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnRestaurarPadrao_Click()
    On Error GoTo ErroRestaurar
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente restaurar todas as configurações para os valores padrão?" & vbCrLf & _
                      "Esta ação não afeta dados operacionais.", _
                      vbQuestion + vbYesNo, "Restaurar Padrão")
    If resposta = vbNo Then Exit Sub
    
    Call RestaurarPadrao
    Call CarregarConfiguracoes
    m_Modificado = False
    MsgBox "Configurações restauradas para os valores padrão.", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroRestaurar:
    MsgBox "Erro ao restaurar configurações: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnExportarConfig_Click()
    On Error GoTo ErroExportar
    
    Dim nomeArquivo As String
    nomeArquivo = ThisWorkbook.Path & "\APS_PURAN_Config_" & Format(Now, "yyyy-mm-dd") & ".json"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(nomeArquivo, True)
    
    Dim config As Object
    Set config = CarregarTodasConfiguracoes
    
    Dim chave As Variant
    Dim json As String
    json = "{"
    
    For Each chave In config.Keys
        json = json & """" & chave & """: """ & config(chave) & """, "
    Next chave
    
    If Len(json) > 1 Then
        json = Left(json, Len(json) - 2)
    End If
    json = json & "}"
    
    ts.Write json
    ts.Close
    
    MsgBox "Configurações exportadas com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroExportar:
    MsgBox "Erro ao exportar configurações: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnImportarConfig_Click()
    On Error GoTo ErroImportar
    
    Dim dialogo As FileDialog
    Set dialogo = Application.FileDialog(msoFileDialogFilePicker)
    
    With dialogo
        .Title = "Importar Configurações"
        .Filters.Clear
        .Filters.Add "Arquivos JSON", "*.json"
        .AllowMultiSelect = False
    End With
    
    If dialogo.Show <> -1 Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(dialogo.SelectedItems(1), 1)
    
    Dim json As String
    json = ts.ReadAll
    ts.Close
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja substituir todas as configurações atuais?" & vbCrLf & _
                      "As configurações atuais serão perdidas.", _
                      vbQuestion + vbYesNo, "Importar Configurações")
    If resposta = vbNo Then Exit Sub
    
    Call LimparTodasConfiguracoes
    Call ImportarConfiguracoesJSON(json)
    
    Call CarregarConfiguracoes
    m_Modificado = False
    MsgBox "Configurações importadas com sucesso!", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroImportar:
    MsgBox "Erro ao importar configurações: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnLimparPreferencias_Click()
    On Error GoTo ErroLimpar
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente limpar todas as preferências do usuário?" & vbCrLf & _
                      "As configurações do sistema serão mantidas.", _
                      vbQuestion + vbYesNo, "Limpar Preferências")
    If resposta = vbNo Then Exit Sub
    
    Call LimparPreferenciasUsuario
    Call CarregarConfiguracoes
    m_Modificado = False
    MsgBox "Preferências do usuário limpas com sucesso!", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroLimpar:
    MsgBox "Erro ao limpar preferências: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    If m_Modificado Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Existem alterações não salvas. Deseja salvar antes de fechar?", _
                          vbQuestion + vbYesNoCancel, "Alterações Pendentes")
        If resposta = vbYes Then
            Call SalvarConfiguracoes
        ElseIf resposta = vbCancel Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

'================================================================================
' SUBROTINAS DE PERSISTÊNCIA
'================================================================================
Private Sub CarregarConfiguracoes()
    On Error Resume Next
    
    Dim config As Object
    Set config = CarregarTodasConfiguracoes
    
    ' Gerais
    If config.Exists("NomeSistema") Then txtNomeSistema.Text = CStr(config("NomeSistema"))
    If config.Exists("Empresa") Then txtEmpresa.Text = CStr(config("Empresa"))
    If config.Exists("FormatoData") Then cboFormatoData.Value = CStr(config("FormatoData"))
    If config.Exists("HorarioPadrao") Then txtHorarioPadrao.Text = CStr(config("HorarioPadrao"))
    
    ' Planejamento
    If config.Exists("ZoomPadrao") Then
        Dim z As Double
        z = CDbl(config("ZoomPadrao"))
        If z >= 0.5 And z <= 3# Then txtZoomPadrao.Text = CStr(z)
    End If
    If config.Exists("PeriodoPadrao") Then cboPeriodoPadrao.Value = CStr(config("PeriodoPadrao"))
    If config.Exists("MostrarLinhaAgora") Then chkMostrarLinhaAgora.Value = CBool(config("MostrarLinhaAgora"))
    If config.Exists("MostrarConflitos") Then chkMostrarConflitos.Value = CBool(config("MostrarConflitos"))
    If config.Exists("EscalaTimeline") Then txtEscalaTimeline.Text = CStr(config("EscalaTimeline"))
    If config.Exists("AlturaLinhaPosto") Then txtAlturaLinhaPosto.Text = CStr(config("AlturaLinhaPosto"))
    
    ' Produção
    If config.Exists("TurnoPadrao") Then cboTurnoPadrao.Value = CStr(config("TurnoPadrao"))
    If config.Exists("HorarioInicio") Then txtHorarioInicio.Text = CStr(config("HorarioInicio"))
    If config.Exists("HorarioFim") Then txtHorarioFim.Text = CStr(config("HorarioFim"))
    If config.Exists("ToleranciaAtraso") Then txtToleranciaAtraso.Text = CStr(config("ToleranciaAtraso"))
    
    ' Interface
    If config.Exists("EstiloCards") Then cboEstiloCards.Value = CStr(config("EstiloCards"))
    If config.Exists("AlturaCards") Then txtAlturaCards.Text = CStr(config("AlturaCards"))
    If config.Exists("MostrarScrollHorizontal") Then chkMostrarScrollHorizontal.Value = CBool(config("MostrarScrollHorizontal"))
    
    On Error GoTo 0
End Sub

Private Sub SalvarConfiguracoes()
    On Error Resume Next
    
    ' Validações básicas
    Dim zoom As Double
    If Not IsNumeric(txtZoomPadrao.Text) Then
        txtZoomPadrao.Text = "1.0"
    Else
        zoom = CDbl(txtZoomPadrao.Text)
        If zoom < 0.5 Or zoom > 3# Then txtZoomPadrao.Text = "1.0"
    End If
    
    Dim escala As Double
    If Not IsNumeric(txtEscalaTimeline.Text) Then
        txtEscalaTimeline.Text = "100"
    Else
        escala = CDbl(txtEscalaTimeline.Text)
        If escala <= 0 Then txtEscalaTimeline.Text = "100"
    End If
    
    Dim altura As Double
    If Not IsNumeric(txtAlturaLinhaPosto.Text) Then
        txtAlturaLinhaPosto.Text = "80"
    Else
        altura = CDbl(txtAlturaLinhaPosto.Text)
        If altura < 40 Or altura > 200 Then txtAlturaLinhaPosto.Text = "80"
    End If
    
    Dim tolerancia As Double
    If Not IsNumeric(txtToleranciaAtraso.Text) Then
        txtToleranciaAtraso.Text = "30"
    Else
        tolerancia = CDbl(txtToleranciaAtraso.Text)
        If tolerancia < 0 Or tolerancia > 1440 Then txtToleranciaAtraso.Text = "30"
    End If
    
    If Not IsNumeric(txtAlturaCards.Text) Then
        txtAlturaCards.Text = "80"
    Else
        altura = CDbl(txtAlturaCards.Text)
        If altura < 40 Or altura > 200 Then txtAlturaCards.Text = "80"
    End If
    
    ' Gerais
    Call SalvarConfiguracao("NomeSistema", txtNomeSistema.Text)
    Call SalvarConfiguracao("Empresa", txtEmpresa.Text)
    Call SalvarConfiguracao("FormatoData", cboFormatoData.Value)
    Call SalvarConfiguracao("HorarioPadrao", txtHorarioPadrao.Text)
    
    ' Planejamento
    Call SalvarConfiguracao("ZoomPadrao", txtZoomPadrao.Text)
    Call SalvarConfiguracao("PeriodoPadrao", cboPeriodoPadrao.Value)
    Call SalvarConfiguracao("MostrarLinhaAgora", CStr(chkMostrarLinhaAgora.Value))
    Call SalvarConfiguracao("MostrarConflitos", CStr(chkMostrarConflitos.Value))
    Call SalvarConfiguracao("EscalaTimeline", txtEscalaTimeline.Text)
    Call SalvarConfiguracao("AlturaLinhaPosto", txtAlturaLinhaPosto.Text)
    
    ' Produção
    Call SalvarConfiguracao("TurnoPadrao", cboTurnoPadrao.Value)
    Call SalvarConfiguracao("HorarioInicio", txtHorarioInicio.Text)
    Call SalvarConfiguracao("HorarioFim", txtHorarioFim.Text)
    Call SalvarConfiguracao("ToleranciaAtraso", txtToleranciaAtraso.Text)
    
    ' Interface
    Call SalvarConfiguracao("EstiloCards", cboEstiloCards.Value)
    Call SalvarConfiguracao("AlturaCards", txtAlturaCards.Text)
    Call SalvarConfiguracao("MostrarScrollHorizontal", CStr(chkMostrarScrollHorizontal.Value))
    
    On Error GoTo 0
End Sub

Private Sub RestaurarPadrao()
    On Error Resume Next
    
    txtNomeSistema.Text = "APS PURAN"
    txtEmpresa.Text = ""
    cboFormatoData.Value = "dd/mm/yyyy"
    txtHorarioPadrao.Text = "08:00 - 17:00"
    
    txtZoomPadrao.Text = "1.0"
    cboPeriodoPadrao.Value = "Mês"
    chkMostrarLinhaAgora.Value = True
    chkMostrarConflitos.Value = True
    txtEscalaTimeline.Text = "100"
    txtAlturaLinhaPosto.Text = "80"
    
    cboTurnoPadrao.Value = "Manhã"
    txtHorarioInicio.Text = "08:00"
    txtHorarioFim.Text = "17:00"
    txtToleranciaAtraso.Text = "30"
    
    cboEstiloCards.Value = "Padrão"
    txtAlturaCards.Text = "80"
    chkMostrarScrollHorizontal.Value = True
    
    Call SalvarConfiguracoes
    On Error GoTo 0
End Sub

Private Sub LimparTodasConfiguracoes()
    On Error Resume Next
    
    Dim dados As Variant
    dados = ObterDadosTabelaEmArray("BD_Config", "TabelaConfig")
    
    If Not IsError(dados) Then
        Dim i As Long
        Dim totalLinhas As Long
        totalLinhas = UBound(dados, 1)
        
        For i = 2 To totalLinhas
            Call ExcluirConfiguracao(CStr(dados(i, 1)))
        Next i
    End If
    
    On Error GoTo 0
End Sub

Private Sub LimparPreferenciasUsuario()
    On Error Resume Next
    
    Dim chaves As Variant
    chaves = Array("DataInicioPlanejamento", "DataFimPlanejamento", "ZoomPlanejamento", _
                   "FiltroStatusPlanejamento", "TextoBuscaPlanejamento", "OrdenarPorPlanejamento", _
                   "UltimoPainelPrincipal", "SimulacaoDataInicio", "SimulacaoDataFim", _
                   "SimulacaoZoom", "SimulacaoAtivaID", "ComparacaoSimA", "ComparacaoSimB")
    
    Dim i As Long
    For i = LBound(chaves) To UBound(chaves)
        Call ExcluirConfiguracao(chaves(i))
    Next i
    
    On Error GoTo 0
End Sub

Private Sub ImportarConfiguracoesJSON(pJSON As String)
    On Error Resume Next
    
    Dim json As Object
    Set json = JsonParse(pJSON)
    
    If json Is Nothing Then Exit Sub
    
    Dim chave As Variant
    For Each chave In json.Keys
        Call SalvarConfiguracao(CStr(chave), CStr(json(chave)))
    Next chave
    
    On Error GoTo 0
End Sub

Private Function JsonParse(pJSON As String) As Object
    On Error Resume Next
    
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
    Dim conteudo As String
    conteudo = Trim(pJSON)
    
    If Left(conteudo, 1) = "{" And Right(conteudo, 1) = "}" Then
        conteudo = Mid(conteudo, 2, Len(conteudo) - 2)
    End If
    
    Dim pares() As String
    pares = Split(conteudo, ",")
    
    Dim i As Long
    For i = 0 To UBound(pares)
        Dim par As String
        par = Trim(pares(i))
        
        Dim doisPontos As Long
        doisPontos = InStr(par, ":")
        
        If doisPontos > 0 Then
            Dim chave As String
            Dim valor As String
            
            chave = Trim(Mid(par, 1, doisPontos - 1))
            If Left(chave, 1) = """" And Right(chave, 1) = """" Then
                chave = Mid(chave, 2, Len(chave) - 2)
            End If
            
            valor = Trim(Mid(par, doisPontos + 1))
            If Left(valor, 1) = """" And Right(valor, 1) = """" Then
                valor = Mid(valor, 2, Len(valor) - 2)
            End If
            
            If chave <> "" Then
                json(chave) = valor
            End If
        End If
    Next i
    
    Set JsonParse = json
    
    On Error GoTo 0
End Function

'================================================================================
' EVENTO: UserForm_QueryClose
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If m_Modificado Then
            Dim resposta As VbMsgBoxResult
            resposta = MsgBox("Existem alterações não salvas. Deseja salvar antes de fechar?", _
                              vbQuestion + vbYesNoCancel, "Alterações Pendentes")
            If resposta = vbYes Then
                Call SalvarConfiguracoes
            ElseIf resposta = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        Unload Me
    End If
End Sub
