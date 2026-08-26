VERSION 1.0 FORM
BEGIN
  MultiUse = -1
  BackColor = 14211288
  BorderStyle = 3
  Caption = "APS PURAN – Capacidade dos Postos"
  ClientHeight = 600
  ClientLeft = 2268
  ClientTop = 1128
  ClientWidth = 800
  Height = 638
  Left = 2268
  ScaleMode = 3
  Top = 1128
  Width = 812
  StartUpPosition = 1
  Attribute VB_Name = "frmCapacidade"
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
Attribute VB_Name = "frmCapacidade"
'================================================================================
' USERFORM: frmCapacidade
' DESCRIÇÃO: Indicadores de capacidade ocupada por posto
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== CONTROLES ================================================================
Private btnAtualizar As MSForms.CommandButton
Private btnFechar As MSForms.CommandButton
Private fraContainer As MSForms.Frame
Private lblResumo As MSForms.Label

'=== ESTADO =====================================================================
Private m_DataInicioPeriodo As Date
Private m_DataFimPeriodo As Date

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
    Me.Caption = "APS PURAN – Capacidade dos Postos"
    
    m_DataInicioPeriodo = 0
    m_DataFimPeriodo = 0
    
    Call CriarControles
    Call CarregarCapacidade
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar capacidade: " & Err.Description, vbCritical, "APS PURAN"
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
        .Caption = "CAPACIDADE DOS POSTOS"
        .Left = 12
        .Top = 12
        .Width = 760
        .Height = 32
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_HEADER
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- Botões -----------------------------------------------------------------
    Set btnAtualizar = Me.Controls.Add("Forms.CommandButton.1", "btnAtualizar", True)
    With btnAtualizar
        .Caption = "Atualizar"
        .Left = 12
        .Top = 56
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
        .Left = 142
        .Top = 56
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COR_ALERTA
        .ForeColor = COR_TEXTO_CLARO
    End With
    
    '--- Container -------------------------------------------------------------
    Set fraContainer = Me.Controls.Add("Forms.Frame.1", "fraContainer", True)
    With fraContainer
        .Caption = ""
        .Left = 12
        .Top = 96
        .Width = 760
        .Height = 460
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
        .Top = 560
        .Width = 760
        .Height = 20
        .Font.Size = 9
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = COR_FUNDO
    End With
End Sub

'================================================================================
' EVENTOS: btnAtualizar_Click / btnFechar_Click
'================================================================================
Private Sub btnAtualizar_Click()
    On Error Resume Next
    Call CarregarCapacidade
End Sub

Private Sub btnFechar_Click()
    On Error Resume Next
    Unload Me
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CarregarCapacidade
' PROPÓSITO: Carregar indicadores de capacidade usando o período atual do frmPrincipal
'================================================================================
Public Sub CarregarCapacidade()
    On Error GoTo ErroCarregar
    
    Dim inicioPeriodo As Date
    Dim fimPeriodo As Date
    
    ' Tenta obter período do frmPrincipal
    On Error Resume Next
    inicioPeriodo = frmPrincipal.m_DataInicioPeriodo
    fimPeriodo = frmPrincipal.m_DataFimPeriodo
    On Error GoTo ErroCarregar
    
    If inicioPeriodo = 0 Or fimPeriodo = 0 Then
        inicioPeriodo = DateSerial(Year(Date), Month(Date), 1)
        fimPeriodo = DateSerial(Year(Date), Month(Date) + 1, 0)
    End If
    
    m_DataInicioPeriodo = inicioPeriodo
    m_DataFimPeriodo = fimPeriodo
    
    Dim resultados As Collection
    Set resultados = CalcularCapacidadePostos(inicioPeriodo, fimPeriodo, "TabelaOPs")
    
    Call RenderizarResultados(resultados)
    
Sair:
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar capacidade: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
End Sub

'================================================================================
' SUBROTINA PRIVADA: RenderizarResultados
'================================================================================
Private Sub RenderizarResultados(pResultados As Collection)
    On Error Resume Next
    
    Dim ctrl As MSForms.Control
    For Each ctrl In fraContainer.Controls
        fraContainer.Controls.Remove ctrl.Name
    Next ctrl
    
    Dim i As Long
    Dim topPos As Single
    topPos = 10
    
    Dim totalPostos As Long
    Dim totalNormal As Long
    Dim totalAlta As Long
    Dim totalSobrecarga As Long
    Dim totalSemProg As Long
    Dim totalManutencao As Long
    
    totalPostos = 0
    totalNormal = 0
    totalAlta = 0
    totalSobrecarga = 0
    totalSemProg = 0
    totalManutencao = 0
    
    For i = 1 To pResultados.Count
        Dim item As clsCapacidadePosto
        Set item = pResultados(i)
        
        totalPostos = totalPostos + 1
        
        Select Case item.Classificacao
            Case "NORMAL": totalNormal = totalNormal + 1
            Case "ALTA UTILIZAÇÃO": totalAlta = totalAlta + 1
            Case "SOBRECARGA": totalSobrecarga = totalSobrecarga + 1
            Case "SEM PROGRAMAÇÃO": totalSemProg = totalSemProg + 1
            Case "MANUTENÇÃO": totalManutencao = totalManutencao + 1
        End Select
        
        Call CriarCardCapacidade(fraContainer, item, topPos)
        topPos = topPos + 70
    Next i
    
    lblResumo.Caption = "Postos: " & totalPostos & " | Normal: " & totalNormal & _
                        " | Alta Utilização: " & totalAlta & " | Sobrecarga: " & totalSobrecarga & _
                        " | Sem Programação: " & totalSemProg & " | Manutenção: " & totalManutencao
End Sub

'================================================================================
' SUBROTINA PRIVADA: CriarCardCapacidade
'================================================================================
Private Sub CriarCardCapacidade(pContainer As MSForms.Frame, pItem As clsCapacidadePosto, pTop As Single)
    On Error Resume Next
    
    Dim uniqueID As String
    uniqueID = Format(Now, "SSSSS") & "_" & CStr(Int(Rnd * 100000))
    
    ' Fundo do card
    Dim lblFundo As MSForms.Label
    Set lblFundo = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID, True)
    With lblFundo
        .Left = 10
        .Top = pTop
        .Width = pContainer.Width - 20
        .Height = 55
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
    
    ' Nome do posto
    Dim lblNome As MSForms.Label
    Set lblNome = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID & "_N", True)
    With lblNome
        .Caption = pItem.NomePosto
        .Left = 20
        .Top = pTop + 5
        .Width = 200
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
    End With
    
    ' Horas
    Dim lblHoras As MSForms.Label
    Set lblHoras = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID & "_H", True)
    With lblHoras
        .Caption = Format(pItem.HorasProgramadas, "0.0") & "h / " & Format(pItem.HorasDisponiveis, "0.0") & "h"
        .Left = 20
        .Top = pTop + 25
        .Width = 200
        .Height = 15
        .Font.Size = 8
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
    End With
    
    ' Percentual
    Dim lblPct As MSForms.Label
    Set lblPct = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID & "_P", True)
    With lblPct
        .Caption = Format(pItem.PercentualOcupacao, "0.0") & "%"
        .Left = 230
        .Top = pTop + 5
        .Width = 80
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = COR_TEXTO_ESCURO
        .BackColor = vbWhite
        .TextAlign = fmTextAlignRight
    End With
    
    ' Classificação
    Dim lblClass As MSForms.Label
    Set lblClass = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID & "_C", True)
    With lblClass
        .Caption = pItem.Classificacao
        .Left = 320
        .Top = pTop + 5
        .Width = 140
        .Height = 20
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        Select Case pItem.Classificacao
            Case "NORMAL": .BackColor = COR_VERDE
            Case "ALTA UTILIZAÇÃO": .BackColor = COR_AMARELO
            Case "SOBRECARGA": .BackColor = COR_ALERTA
            Case "SEM PROGRAMAÇÃO": .BackColor = COR_FUNDO
            Case "MANUTENÇÃO": .BackColor = COR_HEADER
            Case Else: .BackColor = COR_FUNDO
        End Select
    End With
    
    ' Barra de progresso
    Dim lblBarra As MSForms.Label
    Set lblBarra = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID & "_B", True)
    With lblBarra
        .Left = 20
        .Top = pTop + 40
        .Width = pContainer.Width - 40
        .Height = 8
        .BackColor = RGB(220, 220, 220)
        .BorderStyle = fmBorderStyleSingle
    End With
    
    ' Preenchimento da barra
    Dim larguraBarra As Single
    If pItem.HorasDisponiveis > 0 Then
        larguraBarra = (pItem.Width - 40) * (Application.WorksheetFunction.Min(pItem.PercentualOcupacao, 100) / 100)
    Else
        larguraBarra = 0
    End If
    
    If larguraBarra > 0 Then
        Dim lblPreenchimento As MSForms.Label
        Set lblPreenchimento = pContainer.Controls.Add("Forms.Label.1", "lblCap_" & uniqueID & "_BP", True)
        With lblPreenchimento
            .Left = 20
            .Top = pTop + 40
            .Width = larguraBarra
            .Height = 8
            Select Case pItem.Classificacao
                Case "NORMAL": .BackColor = COR_VERDE
                Case "ALTA UTILIZAÇÃO": .BackColor = COR_AMARELO
                Case "SOBRECARGA": .BackColor = COR_ALERTA
                Case "SEM PROGRAMAÇÃO": .BackColor = COR_FUNDO
                Case "MANUTENÇÃO": .BackColor = COR_HEADER
                Case Else: .BackColor = COR_FUNDO
            End Select
        End With
    End If
End Sub

'================================================================================
' EVENTO: UserForm_QueryClose
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
