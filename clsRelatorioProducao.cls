VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsRelatorioProducao"
'================================================================================
' CLASSE: clsRelatorioProducao
' DESCRIÇÃO: Representa um relatório de produção com indicadores agregados
' VERSÃO: 1.0
'================================================================================
Option Explicit

Public TotalOPs As Long
Public QuantidadePlanejada As Long
Public HorasProgramadas As Double
Public OPsPlanejadas As Long
Public OPsEmAndamento As Long
Public OPsConcluidas As Long
Public OPsAtrasadas As Long
Public QuantidadePostos As Long
Public PercentualOcupacao As Double

' Filtros aplicados
Public PeriodoInicio As Date
Public PeriodoFim As Date
Public FiltroPosto As String
Public FiltroProduto As String
Public FiltroStatus As String
