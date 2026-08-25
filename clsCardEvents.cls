VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsCardEvents"
'================================================================================
' CLASSE: clsCardEvents
' DESCRIÇÃO: Encaminha eventos de clique de labels (cards) para o formulário
' VERSÃO: 1.0
'================================================================================
Option Explicit

Public WithEvents Card As MSForms.Label
Public Event Clicked(ID_OP As String)

Private Sub Card_Click()
    RaiseEvent Clicked(Card.Tag)
End Sub
