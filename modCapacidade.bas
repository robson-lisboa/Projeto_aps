Attribute VB_Name = "modCapacidade"
'================================================================================
' MÓDULO: modCapacidade
' DESCRIÇÃO: Indicadores de capacidade ocupada por posto
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: CalcularCapacidadePostos
' PROPÓSITO: Calcular capacidade ocupada por posto para um período
' PARÂMETROS: pDataInicio, pDataFim, pTabelaOPs (nome da tabela)
' RETORNO: Collection de clsCapacidadePosto
'--------------------------------------------------------------------------------
Public Function CalcularCapacidadePostos(pDataInicio As Date, pDataFim As Date, _
                                         Optional pTabelaOPs As String = "TabelaOPs") As Collection
    Dim resultados As New Collection
    Dim dadosOPs As Variant
    Dim dadosEquipamentos As Variant
    Dim mapaPostos As Object
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroCalcular
    
    ' Carrega dados
    dadosOPs = ObterDadosParaCapacidade(pTabelaOPs, pDataInicio, pDataFim)
    dadosEquipamentos = ObterDadosEquipamentos()
    
    If IsError(dadosOPs) Then dadosOPs = Array()
    If IsError(dadosEquipamentos) Then dadosEquipamentos = Array()
    
    ' Inicializa mapa de postos
    Set mapaPostos = BuildMapaPostos(dadosEquipamentos)
    
    ' Calcula horas disponíveis do período
    Dim horasDisponiveis As Double
    horasDisponiveis = (pDataFim - pDataInicio) * 24
    
    ' Inicializa contadores
    Dim chave As Variant
    For Each chave In mapaPostos.Keys
        Dim item As clsCapacidadePosto
        Set item = mapaPostos(chave)
        item.HorasDisponiveis = horasDisponiveis
        item.HorasProgramadas = 0
        item.PercentualOcupacao = 0
        item.QuantidadeOPs = 0
        item.Classificacao = "SEM PROGRAMAÇÃO"
    Next chave
    
    ' Processa OPs
    If Not IsEmpty(dadosOPs) Then
        totalLinhas = UBound(dadosOPs, 1)
        
        For i = 2 To totalLinhas
            Dim equipamento As String
            equipamento = Trim(CStr(dadosOPs(i, 3)))
            
            If equipamento <> "" Then
                Dim inicioOP As Date
                Dim fimOP As Date
                Dim duracaoOP As Double
                
                If IsDate(dadosOPs(i, 5)) Then inicioOP = CDate(dadosOPs(i, 5))
                If IsDate(dadosOPs(i, 6)) Then fimOP = CDate(dadosOPs(i, 6))
                duracaoOP = CDbl(dadosOPs(i, 7))
                
                ' Calcula sobreposição com o período
                Dim inicioEfetivo As Date
                Dim fimEfetivo As Date
                Dim horasSobrepostas As Double
                
                inicioEfetivo = IIf(inicioOP > pDataInicio, inicioOP, pDataInicio)
                fimEfetivo = IIf(fimOP < pDataFim, fimOP, pDataFim)
                
                If fimEfetivo > inicioEfetivo Then
                    horasSobrepostas = (fimEfetivo - inicioEfetivo) * 24
                    
                    If mapaPostos.Exists(equipamento) Then
                        Dim itemOP As clsCapacidadePosto
                        Set itemOP = mapaPostos(equipamento)
                        itemOP.HorasProgramadas = itemOP.HorasProgramadas + horasSobrepostas
                        itemOP.QuantidadeOPs = itemOP.QuantidadeOPs + 1
                    Else
                        Dim novoItem As New clsCapacidadePosto
                        novoItem.ID_Equipamento = equipamento
                        novoItem.NomePosto = equipamento
                        novoItem.HorasDisponiveis = horasDisponiveis
                        novoItem.HorasProgramadas = horasSobrepostas
                        novoItem.QuantidadeOPs = 1
                        novoItem.StatusManutencao = ""
                        mapaPostos.Add equipamento, novoItem
                    End If
                End If
            End If
        Next i
    End If
    
    ' Calcula percentuais e classificação
    Dim key As Variant
    For Each key In mapaPostos.Keys
        Dim itemFinal As clsCapacidadePosto
        Set itemFinal = mapaPostos(key)
        
        If itemFinal.HorasDisponiveis > 0 Then
            itemFinal.PercentualOcupacao = (itemFinal.HorasProgramadas / itemFinal.HorasDisponiveis) * 100
        Else
            itemFinal.PercentualOcupacao = 0
        End If
        
        ' Verifica manutenção
        If InStr(1, UCase(itemFinal.StatusManutencao), "MANUTEN") > 0 Then
            itemFinal.Classificacao = "MANUTENÇÃO"
        ElseIf itemFinal.PercentualOcupacao = 0 Then
            itemFinal.Classificacao = "SEM PROGRAMAÇÃO"
        ElseIf itemFinal.PercentualOcupacao <= 80 Then
            itemFinal.Classificacao = "NORMAL"
        ElseIf itemFinal.PercentualOcupacao <= 100 Then
            itemFinal.Classificacao = "ALTA UTILIZAÇÃO"
        Else
            itemFinal.Classificacao = "SOBRECARGA"
        End If
        
        resultados.Add itemFinal
    Next key
    
    Set CalcularCapacidadePostos = resultados
    
Sair:
    Exit Function
    
ErroCalcular:
    MsgBox "Erro ao calcular capacidade: " & Err.Description, vbCritical, "APS PURAN"
    Set CalcularCapacidadePostos = New Collection
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO PRIVADA: ObterDadosParaCapacidade
' PROPÓSITO: Obter array de OPs filtrado por período
'--------------------------------------------------------------------------------
Private Function ObterDadosParaCapacidade(pTabela As String, pDataInicio As Date, pDataFim As Date) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    Dim i As Long, totalLinhas As Long
    Dim resultados() As Variant
    Dim contResultado As Long
    
    On Error GoTo ErroObterDados
    
    If pTabela = "TabelaOPsSimulacao" Then
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("BD_OPsSimulacao")
        On Error GoTo 0
        
        If ws Is Nothing Then
            ObterDadosParaCapacidade = CVErr(xlErrRef)
            Exit Function
        End If
        
        Set tbl = Nothing
        On Error Resume Next
        Set tbl = ws.ListObjects("TabelaOPsSimulacao")
        On Error GoTo 0
        
        If tbl Is Nothing Then
            ObterDadosParaCapacidade = CVErr(xlErrRef)
            Exit Function
        End If
    Else
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("BD_OPs")
        On Error GoTo 0
        
        If ws Is Nothing Then
            ObterDadosParaCapacidade = CVErr(xlErrRef)
            Exit Function
        End If
        
        Set tbl = Nothing
        On Error Resume Next
        Set tbl = ws.ListObjects("TabelaOPs")
        On Error GoTo 0
        
        If tbl Is Nothing Then
            ObterDadosParaCapacidade = CVErr(xlErrRef)
            Exit Function
        End If
    End If
    
    dados = tbl.Range.Value
    totalLinhas = UBound(dados, 1)
    
    ReDim resultados(1 To totalLinhas, 1 To 12)
    contResultado = 0
    
    For i = 2 To totalLinhas
        Dim inicioOP As Date
        Dim fimOP As Date
        
        If IsDate(dados(i, 5)) Then inicioOP = CDate(dados(i, 5))
        If IsDate(dados(i, 6)) Then fimOP = CDate(dados(i, 6))
        
        ' Verifica sobreposição com o período
        If fimOP >= pDataInicio And inicioOP <= pDataFim Then
            contResultado = contResultado + 1
            Dim j As Long
            For j = 1 To 12
                resultados(contResultado, j) = dados(i, j)
            Next j
        End If
    Next i
    
    If contResultado = 0 Then
        ObterDadosParaCapacidade = Array()
    Else
        ReDim Preserve resultados(1 To contResultado, 1 To 12)
        ObterDadosParaCapacidade = resultados
    End If
    
    Exit Function
    
ErroObterDados:
    ObterDadosParaCapacidade = CVErr(xlErrRef)
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO PRIVADA: ObterDadosEquipamentos
' PROPÓSITO: Ler todos os equipamentos da TabelaEquipamentos
' RETORNO: Variant Array (1-based)
'--------------------------------------------------------------------------------
Private Function ObterDadosEquipamentos() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterEquipamentos
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_Equipamentos")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ObterDadosEquipamentos = CVErr(xlErrRef)
        Exit Function
    End If
    
    Set tbl = Nothing
    On Error Resume Next
    Set tbl = ws.ListObjects("TabelaEquipamentos")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        ObterDadosEquipamentos = CVErr(xlErrRef)
        Exit Function
    End If
    
    dados = tbl.Range.Value
    ObterDadosEquipamentos = dados
    
    Exit Function
    
ErroObterEquipamentos:
    ObterDadosEquipamentos = CVErr(xlErrRef)
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO PRIVADA: BuildMapaPostos
' PROPÓSITO: Construir dicionário de postos a partir dos dados de equipamentos
'--------------------------------------------------------------------------------
Private Function BuildMapaPostos(pDados As Variant) As Object
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim totalLinhas As Long
    If IsEmpty(pDados) Then
        Set BuildMapaPostos = mapa
        Exit Function
    End If
    totalLinhas = UBound(pDados, 1)
    
    For i = 2 To totalLinhas
        Dim idEq As String
        Dim nomeEq As String
        Dim statusManut As String
        
        idEq = Trim(CStr(pDados(i, 1)))
        nomeEq = Trim(CStr(pDados(i, 2)))
        statusManut = Trim(CStr(pDados(i, 4)))
        
        If idEq <> "" Then
            Dim item As New clsCapacidadePosto
            item.ID_Equipamento = idEq
            item.NomePosto = IIf(nomeEq <> "", nomeEq, idEq)
            item.StatusManutencao = statusManut
            mapa(idEq) = item
        End If
    Next i
    
    Set BuildMapaPostos = mapa
End Function
