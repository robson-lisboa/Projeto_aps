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
    Set dashboard.RankingMotivosParada = CreateObject("Scripting.Dictionary")
    Set dashboard.RankingEquipamentosParada = CreateObject("Scripting.Dictionary")
    Set dashboard.Alertas = CreateObject("Scripting.Dictionary")
    
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
            
            ' Quantidade boa = produzida - rejeitada
            Dim qtdBoa As Long
            qtdBoa = qtdProduzida - qtdRejeitada
            If qtdBoa >= 0 Then
                dashboard.QuantidadeBoa = dashboard.QuantidadeBoa + qtdBoa
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
            
            ' Ranking de motivos de parada (mesmo cálculo, separado para clareza)
            If idMotivo <> "" Then
                If Not dashboard.RankingMotivosParada.Exists(idMotivo) Then
                    dashboard.RankingMotivosParada.Add idMotivo, 0
                End If
                dashboard.RankingMotivosParada(idMotivo) = dashboard.RankingMotivosParada(idMotivo) + 1
            End If
            
            ' Ranking de equipamentos com mais paradas
            If equipamentoParada <> "" Then
                If Not dashboard.RankingEquipamentosParada.Exists(equipamentoParada) Then
                    dashboard.RankingEquipamentosParada.Add equipamentoParada, 0
                End If
                dashboard.RankingEquipamentosParada(equipamentoParada) = dashboard.RankingEquipamentosParada(equipamentoParada) + horasParada
            End If
            
            ' Maior parada
            If horasParada > dashboard.MaiorParadaDuracao Then
                dashboard.MaiorParadaDuracao = horasParada
                If idMotivo <> "" Then
                    dashboard.MaiorParada = idMotivo
                Else
                    dashboard.MaiorParada = "Parada sem motivo"
                End If
            End If
            
ProximaParada:
        Next i
    End If
    
    ' Calcula média de duração das paradas
    If dashboard.NumeroParadas > 0 Then
        dashboard.MediaDuracaoParada = dashboard.HorasParada / dashboard.NumeroParadas
    Else
        dashboard.MediaDuracaoParada = 0
    End If
    
    ' Calcula percentuais
    If dashboard.QuantidadePlanejada > 0 Then
        dashboard.PercentualProduzido = (dashboard.QuantidadeProduzida / dashboard.QuantidadePlanejada) * 100
    Else
        dashboard.PercentualProduzido = 0
    End If
    
    ' Meta x Realizado
    dashboard.SaldoProduzir = dashboard.QuantidadePlanejada - dashboard.QuantidadeProduzida
    If dashboard.QuantidadePlanejada > 0 Then
        dashboard.PercentualAtingido = (dashboard.QuantidadeProduzida / dashboard.QuantidadePlanejada) * 100
    Else
        dashboard.PercentualAtingido = 0
    End If
    
    ' Eficiência = Quantidade Produzida / Quantidade Planejada × 100
    ' Regra consistente: mesma fórmula do percentual produzido
    If dashboard.QuantidadePlanejada > 0 Then
        dashboard.EficienciaProducao = (dashboard.QuantidadeProduzida / dashboard.QuantidadePlanejada) * 100
    Else
        dashboard.EficienciaProducao = 0
    End If
    
    ' Quantidade produzida por hora
    If dashboard.HorasProducao > 0 Then
        dashboard.QuantidadePorHora = dashboard.QuantidadeProduzida / dashboard.HorasProducao
    Else
        dashboard.QuantidadePorHora = 0
    End If
    
    ' OEE
    ' Disponibilidade = Tempo de produção / Tempo disponível (período)
    Dim tempoDisponivel As Double
    tempoDisponivel = (pDataFim - pDataInicio) * 24
    
    If tempoDisponivel > 0 Then
        dashboard.Disponibilidade = (dashboard.HorasProducao / tempoDisponivel) * 100
    Else
        dashboard.Disponibilidade = 0
    End If
    
    ' Performance = Produção real / Produção teórica esperada
    ' ATENÇÃO: Performance/OEE completo depende de uma taxa padrão confiável.
    ' Sem essa informação, não é possível calcular Performance corretamente.
    ' NÃO inventar velocidade. NÃO usar Capacidade_Hora arbitrariamente.
    dashboard.PerformanceDisponivel = False
    dashboard.Performance = 0
    
    ' Qualidade = Quantidade boa / Quantidade produzida
    If dashboard.QuantidadeProduzida > 0 Then
        dashboard.Qualidade = (dashboard.QuantidadeBoa / dashboard.QuantidadeProduzida) * 100
    Else
        dashboard.Qualidade = 0
    End If
    
    ' OEE = Disponibilidade × Performance × Qualidade
    ' Como Performance não está disponível, OEE é calculado como parcial
    dashboard.OEE = (dashboard.Disponibilidade / 100) * (dashboard.Qualidade / 100) * 100
    
    ' Alertas operacionais
    Call GerarAlertas(dashboard)
    
    Set CalcularDashboardProducao = dashboard
    
Sair:
    Exit Function
    
ErroCalcular:
    Set CalcularDashboardProducao = dashboard
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA PRIVADA: GerarAlertas
' PROPÓSITO: Gerar alertas operacionais baseados nos indicadores
'--------------------------------------------------------------------------------
Private Sub GerarAlertas(pDashboard As clsDashboardProducao)
    On Error Resume Next
    
    ' OP em produção
    If pDashboard.ProducoesEmProducao > 0 Then
        pDashboard.Alertas.Add "OP em produção", pDashboard.ProducoesEmProducao & " produção(ões) em andamento"
    End If
    
    ' OP pausada
    If pDashboard.ProducoesPausadas > 0 Then
        pDashboard.Alertas.Add "OP pausada", pDashboard.ProducoesPausadas & " produção(ões) pausada(s)"
    End If
    
    ' Produção atrasada - considera percentual baixo para produções finalizadas
    If pDashboard.ProducoesFinalizadas > 0 And pDashboard.PercentualAtingido < 50 Then
        pDashboard.Alertas.Add "Produção atrasada", "Percentual atingido abaixo de 50%"
    End If
    
    ' Rejeição elevada - mais de 10% de rejeição
    If pDashboard.QuantidadeProduzida > 0 Then
        Dim percentualRejeicao As Double
        percentualRejeicao = (pDashboard.QuantidadeRejeitada / pDashboard.QuantidadeProduzida) * 100
        If percentualRejeicao > 10 Then
            pDashboard.Alertas.Add "Rejeição elevada", Format(percentualRejeicao, "0.0") & "% de rejeição"
        End If
    End If
    
    ' Equipamento com muitas paradas (mais de 3 paradas)
    Dim chaveEq As Variant
    For Each chaveEq In pDashboard.RankingEquipamentosParada.Keys
        If pDashboard.RankingEquipamentosParada(chaveEq) > 3 Then
            pDashboard.Alertas.Add "Paradas no equipamento", CStr(chaveEq) & " com " & CStr(pDashboard.RankingEquipamentosParada(chaveEq)) & " paradas"
        End If
    Next chaveEq
    
    ' Horas de parada altas (mais de 20% do tempo disponível)
    Dim tempoDisponivel As Double
    tempoDisponivel = (pDashboard.FiltroPeriodoFim - pDashboard.FiltroPeriodoInicio) * 24
    If tempoDisponivel > 0 Then
        Dim percentualParada As Double
        percentualParada = (pDashboard.HorasParada / tempoDisponivel) * 100
        If percentualParada > 20 Then
            pDashboard.Alertas.Add "Alto tempo de parada", Format(percentualParada, "0.0") & "% do período parado"
        End If
    End If
End Sub
