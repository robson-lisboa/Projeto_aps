Attribute VB_Name = "modDashboardProducao"
'================================================================================
' MÓDULO: modDashboardProducao
' DESCRIÇÃO: Cálculo de indicadores do Dashboard de Produção Real
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: CalcularDashboardProducao
' PROPÓSITO: Calcular todos os indicadores do dashboard de produção
' PARÂMETROS: pDataInicio, pDataFim, pEquipamento, pOperador, pStatus
' RETORNO: clsDashboardProducao
'--------------------------------------------------------------------------------
Public Function CalcularDashboardProducao(pDataInicio As Date, pDataFim As Date, _
                                          Optional pEquipamento As String = "", _
                                          Optional pOperador As String = "", _
                                          Optional pStatus As String = "") As clsDashboardProducao
    Dim dashboard As New clsDashboardProducao
    Dim dadosProducoes As Variant
    Dim dadosParadas As Variant
    Dim i As Long, totalLinhas As Long
    
    On Error GoTo ErroCalcular
    
    dashboard.FiltroPeriodoInicio = pDataInicio
    dashboard.FiltroPeriodoFim = pDataFim
    dashboard.FiltroEquipamento = pEquipamento
    dashboard.FiltroOperador = pOperador
    dashboard.FiltroStatus = pStatus
    
    ' Inicializa dicionários
    Set dashboard.ProducaoPorEquipamento = CreateObject("Scripting.Dictionary")
    Set dashboard.ProducaoPorOperador = CreateObject("Scripting.Dictionary")
    Set dashboard.MotivosParada = CreateObject("Scripting.Dictionary")
    
    ' Carrega dados de produção
    dadosProducoes = ObterProducoesEmArray()
    If IsError(dadosProducoes) Then dadosProducoes = Array()
    
    ' Carrega dados de paradas
    dadosParadas = ObterParadasEmArray()
    If IsError(dadosParadas) Then dadosParadas = Array()
    
    ' Processa produções
    If Not IsEmpty(dadosProducoes) Then
        totalLinhas = UBound(dadosProducoes, 1)
        
        For i = 2 To totalLinhas
            Dim idProducao As String
            Dim idOP As String
            Dim idOperador As String
            Dim equipamento As String
            Dim inicioProd As Date
            Dim fimProd As Date
            Dim qtdPlanejada As Long
            Dim qtdProduzida As Long
            Dim qtdRejeitada As Long
            Dim status As String
            
            idProducao = Trim(CStr(dadosProducoes(i, 1)))
            idOP = Trim(CStr(dadosProducoes(i, 2)))
            idOperador = Trim(CStr(dadosProducoes(i, 3)))
            equipamento = Trim(CStr(dadosProducoes(i, 4)))
            
            If IsDate(dadosProducoes(i, 5)) Then inicioProd = CDate(dadosProducoes(i, 5))
            If IsDate(dadosProducoes(i, 6)) Then fimProd = CDate(dadosProducoes(i, 6))
            
            If IsNumeric(dadosProducoes(i, 7)) Then qtdPlanejada = CLng(dadosProducoes(i, 7))
            If IsNumeric(dadosProducoes(i, 8)) Then qtdProduzida = CLng(dadosProducoes(i, 8))
            If IsNumeric(dadosProducoes(i, 9)) Then qtdRejeitada = CLng(dadosProducoes(i, 9))
            status = Trim(CStr(dadosProducoes(i, 10)))
            
            ' Ignora registros inválidos
            If idProducao = "" Or idOP = "" Then GoTo ProximaProducao
            
            ' Verifica sobreposição com o período
            Dim inicioEfetivo As Date
            Dim fimEfetivo As Date
            Dim sobrepoePeriodo As Boolean
            
            ' Para produções em andamento, considera Now() como fim efetivo
            If status = "EM PRODUÇÃO" Or status = "PAUSADA" Then
                fimEfetivo = Now()
            Else
                fimEfetivo = fimProd
            End If
            
            ' Calcula sobreposição com o período
            inicioEfetivo = IIf(inicioProd > pDataInicio, inicioProd, pDataInicio)
            fimEfetivo = IIf(fimEfetivo < pDataFim, fimEfetivo, pDataFim)
            
            sobrepoePeriodo = (fimEfetivo > inicioEfetivo)
            
            If Not sobrepoePeriodo Then GoTo ProximaProducao
            
            ' Aplica filtros
            If pEquipamento <> "" And StrComp(equipamento, pEquipamento, vbTextCompare) <> 0 Then
                GoTo ProximaProducao
            End If
            
            If pOperador <> "" And StrComp(idOperador, pOperador, vbTextCompare) <> 0 Then
                GoTo ProximaProducao
            End If
            
            If pStatus <> "" And StrComp(status, pStatus, vbTextCompare) <> 0 Then
                GoTo ProximaProducao
            End If
            
            ' Contabiliza indicadores
            dashboard.TotalProducoes = dashboard.TotalProducoes + 1
            
            Select Case status
                Case "EM PRODUÇÃO": dashboard.ProducoesEmProducao = dashboard.ProducoesEmProducao + 1
                Case "PAUSADA": dashboard.ProducoesPausadas = dashboard.ProducoesPausadas + 1
                Case "FINALIZADA": dashboard.ProducoesFinalizadas = dashboard.ProducoesFinalizadas + 1
            End Select
            
            ' Quantidades (apenas se não negativas)
            If qtdPlanejada >= 0 Then
                dashboard.QuantidadePlanejada = dashboard.QuantidadePlanejada + qtdPlanejada
            End If
            
            If qtdProduzida >= 0 Then
                dashboard.QuantidadeProduzida = dashboard.QuantidadeProduzida + qtdProduzida
            End If
            
            If qtdRejeitada >= 0 Then
                dashboard.QuantidadeRejeitada = dashboard.QuantidadeRejeitada + qtdRejeitada
            End If
            
            ' Horas de produção (considera sobreposição com período)
            Dim horasProducao As Double
            horasProducao = (fimEfetivo - inicioEfetivo) * 24
            If horasProducao > 0 Then
                dashboard.HorasProducao = dashboard.HorasProducao + horasProducao
            End If
            
            ' Produção por equipamento
            If equipamento <> "" Then
                If Not dashboard.ProducaoPorEquipamento.Exists(equipamento) Then
                    dashboard.ProducaoPorEquipamento.Add equipamento, 0
                End If
                dashboard.ProducaoPorEquipamento(equipamento) = dashboard.ProducaoPorEquipamento(equipamento) + 1
            End If
            
            ' Produção por operador
            If idOperador <> "" Then
                If Not dashboard.ProducaoPorOperador.Exists(idOperador) Then
                    dashboard.ProducaoPorOperador.Add idOperador, 0
                End If
                dashboard.ProducaoPorOperador(idOperador) = dashboard.ProducaoPorOperador(idOperador) + 1
            End If
            
ProximaProducao:
        Next i
    End If
    
    ' Processa paradas
    If Not IsEmpty(dadosParadas) Then
        totalLinhas = UBound(dadosParadas, 1)
        
        For i = 2 To totalLinhas
            Dim idParada As String
            Dim equipamentoParada As String
            Dim operadorParada As String
            Dim inicioParada As Date
            Dim fimParada As Date
            Dim idMotivo As String
            
            idParada = Trim(CStr(dadosParadas(i, 1)))
            Dim idOPParada As String
            idOPParada = Trim(CStr(dadosParadas(i, 2)))
            equipamentoParada = Trim(CStr(dadosParadas(i, 3)))
            operadorParada = Trim(CStr(dadosParadas(i, 4)))
            
            If IsDate(dadosParadas(i, 5)) Then inicioParada = CDate(dadosParadas(i, 5))
            If IsDate(dadosParadas(i, 6)) Then fimParada = CDate(dadosParadas(i, 6))
            
            idMotivo = Trim(CStr(dadosParadas(i, 7)))
            
            ' Ignora registros inválidos
            If idParada = "" Then GoTo ProximaParada
            
            ' Aplica filtros de equipamento e operador
            If pEquipamento <> "" And StrComp(equipamentoParada, pEquipamento, vbTextCompare) <> 0 Then
                GoTo ProximaParada
            End If
            
            If pOperador <> "" And StrComp(operadorParada, pOperador, vbTextCompare) <> 0 Then
                GoTo ProximaParada
            End If
            
            ' Para paradas em andamento, considera Now() como fim efetivo
            Dim fimEfetivoParada As Date
            Dim inicioEfetivoParada As Date
            Dim sobrepoePeriodoParada As Boolean
            
            If Not IsDate(fimParada) Or IsEmpty(fimParada) Then
                fimEfetivoParada = Now()
            Else
                fimEfetivoParada = fimParada
            End If
            
            ' Calcula sobreposição com o período
            inicioEfetivoParada = IIf(inicioParada > pDataInicio, inicioParada, pDataInicio)
            fimEfetivoParada = IIf(fimEfetivoParada < pDataFim, fimEfetivoParada, pDataFim)
            
            sobrepoePeriodoParada = (fimEfetivoParada > inicioEfetivoParada)
            
            If Not sobrepoePeriodoParada Then GoTo ProximaParada
            
            ' Contabiliza parada
            dashboard.NumeroParadas = dashboard.NumeroParadas + 1
            
            ' Horas de parada
            Dim horasParada As Double
            horasParada = (fimEfetivoParada - inicioEfetivoParada) * 24
            If horasParada > 0 Then
                dashboard.HorasParada = dashboard.HorasParada + horasParada
            End If
            
            ' Motivos de parada
            If idMotivo <> "" Then
                If Not dashboard.MotivosParada.Exists(idMotivo) Then
                    dashboard.MotivosParada.Add idMotivo, 0
                End If
                dashboard.MotivosParada(idMotivo) = dashboard.MotivosParada(idMotivo) + 1
            End If
            
ProximaParada:
        Next i
    End If
    
    ' Calcula percentuais
    If dashboard.QuantidadePlanejada > 0 Then
        dashboard.PercentualProduzido = (dashboard.QuantidadeProduzida / dashboard.QuantidadePlanejada) * 100
    Else
        dashboard.PercentualProduzido = 0
    End If
    
    ' Eficiência = Quantidade Produzida / Quantidade Planejada × 100
    ' Regra consistente: mesma fórmula do percentual produzido
    If dashboard.QuantidadePlanejada > 0 Then
        dashboard.EficienciaProducao = (dashboard.QuantidadeProduzida / dashboard.QuantidadePlanejada) * 100
    Else
        dashboard.EficienciaProducao = 0
    End If
    
    Set CalcularDashboardProducao = dashboard
    
Sair:
    Exit Function
    
ErroCalcular:
    Set CalcularDashboardProducao = dashboard
    Resume Sair
End Function
