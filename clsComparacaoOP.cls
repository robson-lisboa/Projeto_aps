VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsComparacaoOP"
'================================================================================
' CLASSE: clsComparacaoOP
' DESCRIÇÃO: Resultado da comparação entre duas simulações para uma OP
' VERSÃO: 1.0
'================================================================================
Option Explicit

Public ID_OP As String
Public ExisteA As Boolean
Public ExisteB As Boolean
Public EquipamentoA As String
Public EquipamentoB As String
Public DataInicioA As Date
Public DataInicioB As Date
Public DataFimA As Date
Public DataFimB As Date
Public DuracaoA As Double
Public DuracaoB As Double
Public StatusA As String
Public StatusB As String
Public TipoDiferenca As String
Public Observacao As String
