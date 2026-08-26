VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsCapacidadePosto"
'================================================================================
' CLASSE: clsCapacidadePosto
' DESCRIÇÃO: Representa o indicador de capacidade ocupada de um posto
' VERSÃO: 1.0
'================================================================================
Option Explicit

Public ID_Equipamento As String
Public NomePosto As String
Public HorasDisponiveis As Double
Public HorasProgramadas As Double
Public PercentualOcupacao As Double
Public Classificacao As String
Public QuantidadeOPs As Long
Public StatusManutencao As String
