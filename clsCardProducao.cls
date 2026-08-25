VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsCardProducao"
'================================================================================
' CLASSE: clsCardProducao
' DESCRIÇÃO: Representa um card de produção com posicionamento temporal
' VERSÃO: 1.0
'================================================================================
Option Explicit

Public ID_OP As String
Public Produto As String
Public Equipamento As String
Public Quantidade As Long
Public DataInicio As Date
Public DataFim As Date
Public Duracao As Double
Public Status As String

Public Left As Double
Public Top As Single
Public Width As Double
Public Height As Single
Public IndicePosto As Long

Private m_TemConflito As Boolean

Public Property Get TemConflito() As Boolean
    TemConflito = m_TemConflito
End Property

Public Property Let TemConflito(pValor As Boolean)
    m_TemConflito = pValor
End Property
