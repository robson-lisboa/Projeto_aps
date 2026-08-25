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

Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    Me.BackColor = 14211288
    Me.Caption = "Detalhes da OP"
    
    Set btnEditar = Me.Controls.Add("Forms.CommandButton.1", "btnEditar", True)
    With btnEditar
        .Caption = "Editar"
        .Left = 180
        .Top = 340
        .Width = 120
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = 15773696
        .ForeColor = 16777215
    End With
    
Sair:
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro ao inicializar detalhes: " & Err.Description, vbCritical, "APS PURAN"
    Resume Sair
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
