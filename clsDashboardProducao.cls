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
Public QuantidadeBoa As Long

Public PercentualProduzido As Double
Public HorasProducao As Double
Public HorasParada As Double
Public NumeroParadas As Long
Public EficienciaProducao As Double
Public QuantidadePorHora As Double

'=== META x REALIZADO ==========================================================
Public SaldoProduzir As Long
Public PercentualAtingido As Double

'=== OEE ========================================================================
' Disponibilidade = Tempo de produção planejado / Tempo disponível
' Performance = Produção real / Produção teórica esperada
'   ATENÇÃO: Performance/OEE completo depende de taxa padrão confiável.
'   Sem essa informação, Performance é calculada como 0 e OEE parcial.
' Qualidade = Quantidade boa / Quantidade produzida
' OEE = Disponibilidade × Performance × Qualidade
Public Disponibilidade As Double
Public Performance As Double
Public Qualidade As Double
Public OEE As Double
Public PerformanceDisponivel As Boolean

'=== PARADAS ====================================================================
Public MediaDuracaoParada As Double
Public MaiorParada As String
Public MaiorParadaDuracao As Double
Public RankingMotivosParada As Object
Public RankingEquipamentosParada As Object

'=== PRODUÇÃO POR EQUIPAMENTO ===================================================
Public ProducaoPorEquipamento As Object

'=== PRODUÇÃO POR OPERADOR ======================================================
Public ProducaoPorOperador As Object

'=== MOTIVOS DE PARADA ==========================================================
Public MotivosParada As Object

'=== ALERTAS OPERACIONAIS =======================================================
Public Alertas As Object

'=== FILTROS APLICADOS ==========================================================
Public FiltroPeriodoInicio As Date
Public FiltroPeriodoFim As Date
Public FiltroEquipamento As String
Public FiltroOperador As String
Public FiltroStatus As String
