VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsDashboardProducao"
'================================================================================
' CLASSE: clsDashboardProducao
' DESCRIÇÃO: Representa os indicadores do Dashboard de Produção Real
' VERSÃO: 1.0
'================================================================================
Option Explicit

'=== INDICADORES PRINCIPAIS =====================================================
Public TotalProducoes As Long
Public ProducoesEmProducao As Long
Public ProducoesPausadas As Long
Public ProducoesFinalizadas As Long

Public QuantidadePlanejada As Long
Public QuantidadeProduzida As Long
Public QuantidadeRejeitada As Long

Public PercentualProduzido As Double
Public HorasProducao As Double
Public HorasParada As Double
Public NumeroParadas As Long
Public EficienciaProducao As Double

'=== PRODUÇÃO POR EQUIPAMENTO ===================================================
Public ProducaoPorEquipamento As Object

'=== PRODUÇÃO POR OPERADOR ======================================================
Public ProducaoPorOperador As Object

'=== MOTIVOS DE PARADA ==========================================================
Public MotivosParada As Object

'=== FILTROS APLICADOS ==========================================================
Public FiltroPeriodoInicio As Date
Public FiltroPeriodoFim As Date
Public FiltroEquipamento As String
Public FiltroOperador As String
Public FiltroStatus As String
