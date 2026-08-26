Attribute VB_Name = "modRelatorios"
'================================================================================
' MÓDULO: modRelatorios
' DESCRIÇÃO: Geração de relatórios de produção somente leitura
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: GerarRelatorioProducao
' PROPÓSITO: Gerar relatório de produção com indicadores agregados
' PARÂMETROS: pDataInicio, pDataFim, pPosto, pProduto, pStatus
' RETORNO: clsRelatorioProducao
'--------------------------------------------------------------------------------
Public Function GerarRelatorioProducao(pDataInicio As Date, pDataFim As Date, _
                                       Optional pPosto As String = "", _
                                       Optional pProduto As String = "", _
                                       Optional pStatus As String = "") As clsRelatorioProducao
    Dim relatorio As New clsRelatorioProducao
    Dim dadosOPs As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroGerar
    
    relatorio.PeriodoInicio = pDataInicio
    relatorio.PeriodoFim = pDataFim
    relatorio.FiltroPosto = pPosto
    relatorio.FiltroProduto = pProduto
    relatorio.FiltroStatus = pStatus
    
    dadosOPs = ObterDadosOPsEmArray()
    
    If IsError(dadosOPs) Then
        Set GerarRelatorioProducao = relatorio
        Exit Function
    End If
    
    totalLinhas = UBound(dadosOPs, 1)
    
    Dim mapaPostos As Object
    Set mapaPostos = CreateObject("Scripting.Dictionary")
    
    For i = 2 To totalLinhas
        Dim inicioOP As Date
        Dim fimOP As Date
        Dim equipamento As String
        Dim produtoOP As String
        Dim statusOP As String
        Dim quantidadeOP As Long
        Dim duracaoOP As Double
        
        If IsDate(dadosOPs(i, 5)) Then inicioOP = CDate(dadosOPs(i, 5))
        If IsDate(dadosOPs(i, 6)) Then fimOP = CDate(dadosOPs(i, 6))
        equipamento = Trim(CStr(dadosOPs(i, 3)))
        produtoOP = Trim(CStr(dadosOPs(i, 2)))
        statusOP = Trim(CStr(dadosOPs(i, 8)))
        
        If IsNumeric(dadosOPs(i, 4)) Then quantidadeOP = CLng(dadosOPs(i, 4))
        If IsNumeric(dadosOPs(i, 7)) Then duracaoOP = CDbl(dadosOPs(i, 7))
        
        ' Verifica sobreposição com o período
        If fimOP < pDataInicio Or inicioOP > pDataFim Then
            GoTo ProximaOP
        End If
        
        ' Aplica filtros
        If pPosto <> "" And StrComp(equipamento, pPosto, vbTextCompare) <> 0 Then
            GoTo ProximaOP
        End If
        
        If pProduto <> "" And StrComp(produtoOP, pProduto, vbTextCompare) <> 0 Then
            GoTo ProximaOP
        End If
        
        If pStatus <> "" And StrComp(statusOP, pStatus, vbTextCompare) <> 0 Then
            GoTo ProximaOP
        End If
        
        ' Contabiliza indicadores
        relatorio.TotalOPs = relatorio.TotalOPs + 1
        relatorio.QuantidadePlanejada = relatorio.QuantidadePlanejada + quantidadeOP
        relatorio.HorasProgramadas = relatorio.HorasProgramadas + duracaoOP
        
        Select Case Trim(statusOP)
            Case "Planejada": relatorio.OPsPlanejadas = relatorio.OPsPlanejadas + 1
            Case "Em Andamento": relatorio.OPsEmAndamento = relatorio.OPsEmAndamento + 1
            Case "Concluído": relatorio.OPsConcluidas = relatorio.OPsConcluidas + 1
            Case "Atrasado": relatorio.OPsAtrasadas = relatorio.OPsAtrasadas + 1
        End Select
        
        ' Contabiliza postos únicos
        If equipamento <> "" Then
            If Not mapaPostos.Exists(equipamento) Then
                mapaPostos.Add equipamento, 1
            End If
        End If
        
ProximaOP:
    Next i
    
    relatorio.QuantidadePostos = mapaPostos.Count
    
    ' Calcula percentual de ocupação
    Dim horasDisponiveis As Double
    horasDisponiveis = (pDataFim - pDataInicio) * 24
    
    If horasDisponiveis > 0 Then
        relatorio.PercentualOcupacao = (relatorio.HorasProgramadas / horasDisponiveis) * 100
    Else
        relatorio.PercentualOcupacao = 0
    End If
    
    Set GerarRelatorioProducao = relatorio
    
Sair:
    Exit Function
    
ErroGerar:
    MsgBox "Erro ao gerar relatório: " & Err.Description, vbCritical, "APS PURAN"
    Set GerarRelatorioProducao = relatorio
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterListaPostos
' PROPÓSITO: Obter lista de postos/equipamentos distintos
' RETORNO: Collection de strings
'--------------------------------------------------------------------------------
Public Function ObterListaPostos() As Collection
    Dim resultados As New Collection
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroObter
    
    dados = ObterDadosTabelaEmArray("BD_Equipamentos", "TabelaEquipamentos")
    
    If IsError(dados) Then
        Set ObterListaPostos = resultados
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        Dim nomeEq As String
        nomeEq = Trim(CStr(dados(i, 2)))
        If nomeEq <> "" Then
            On Error Resume Next
            resultados.Add nomeEq, nomeEq
            On Error GoTo ErroObter
        End If
    Next i
    
    Set ObterListaPostos = resultados
    
Sair:
    Exit Function
    
ErroObter:
    Set ObterListaPostos = resultados
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterListaProdutos
' PROPÓSITO: Obter lista de produtos distintos
' RETORNO: Collection de strings
'--------------------------------------------------------------------------------
Public Function ObterListaProdutos() As Collection
    Dim resultados As New Collection
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroObter
    
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then
        Set ObterListaProdutos = resultados
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        Dim produto As String
        produto = Trim(CStr(dados(i, 2)))
        If produto <> "" Then
            On Error Resume Next
            resultados.Add produto, produto
            On Error GoTo ErroObter
        End If
    Next i
    
    Set ObterListaProdutos = resultados
    
Sair:
    Exit Function
    
ErroObter:
    Set ObterListaProdutos = resultados
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterListaStatus
' PROPÓSITO: Obter lista de status distintos
' RETORNO: Collection de strings
'--------------------------------------------------------------------------------
Public Function ObterListaStatus() As Collection
    Dim resultados As New Collection
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroObter
    
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then
        Set ObterListaStatus = resultados
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        Dim status As String
        status = Trim(CStr(dados(i, 8)))
        If status <> "" Then
            On Error Resume Next
            resultados.Add status, status
            On Error GoTo ErroObter
        End If
    Next i
    
    Set ObterListaStatus = resultados
    
Sair:
    Exit Function
    
ErroObter:
    Set ObterListaStatus = resultados
    Resume Sair
End Function
