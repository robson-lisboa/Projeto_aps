VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsButtonEvents"
'================================================================================
' CLASSE: clsButtonEvents
' DESCRIÇÃO: Encaminha eventos de clique de botões dinâmicos para o formulário
' VERSÃO: 1.0
'================================================================================
Option Explicit

Public WithEvents Button As MSForms.CommandButton
Public Event Clicked()

Private Sub Button_Click()
    RaiseEvent Clicked
End Sub
