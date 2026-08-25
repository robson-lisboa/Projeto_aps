Attribute VB_Name = "modGantt"
'================================================================================
' MÓDULO: modGantt
' DESCRIÇÃO: Motor de posicionamento temporal dos cards de produção
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' CONSTANTES DE ESCALA
'--------------------------------------------------------------------------------
Private Const FATOR_BASE_PIXELS_POR_HORA As Double = 100
Private Const ZOOM_MIN As Double = 0.5
Private Const ZOOM_MAX As Double = 3.0
Private Const ALTURA_LINHA_POSTO As Single = 80
Private Const MARGEM_ESQUERDA As Single = 120

'--------------------------------------------------------------------------------
' TIPO: EscalaTemporal
' PROPÓSITO: Armazenar configurações de escala temporal
'--------------------------------------------------------------------------------
Public Type EscalaTemporal
    DataInicio As Date
    DataFim As Date
    PixelsPorHora As Double
    Zoom As Double
    LarguraTotal As Double
End Type

'--------------------------------------------------------------------------------
' FUNÇÃO: CalcularEscalaTemporal
' PROPÓSITO: Calcular escala temporal baseada em período e zoom
' PARÂMETROS: DataInicio, DataFim, Zoom
' RETORNO: EscalaTemporal
'--------------------------------------------------------------------------------
Public Function CalcularEscalaTemporal(pDataInicio As Date, pDataFim As Date, Optional pZoom As Double = 1.0) As EscalaTemporal
    Dim escala As EscalaTemporal
    
    escala.DataInicio = pDataInicio
    escala.DataFim = pDataFim
    escala.Zoom = Application.WorksheetFunction.Min(Application.WorksheetFunction.Max(pZoom, ZOOM_MIN), ZOOM_MAX)
    escala.PixelsPorHora = FATOR_BASE_PIXELS_POR_HORA * escala.Zoom
    
    Dim totalHoras As Double
    totalHoras = (pDataFim - pDataInicio) * 24
    
    escala.LarguraTotal = totalHoras * escala.PixelsPorHora
    
    CalcularEscalaTemporal = escala
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ConverterDataParaX
' PROPÓSITO: Converter data/hora para posição X na timeline
' PARÂMETROS: DataHora, EscalaTemporal
' RETORNO: Posição X em pixels
'--------------------------------------------------------------------------------
Public Function ConverterDataParaX(pDataHora As Date, pEscala As EscalaTemporal) As Double
    Dim horasDesdeInicio As Double
    horasDesdeInicio = (pDataHora - pEscala.DataInicio) * 24
    
    ConverterDataParaX = MARGEM_ESQUERDA + (horasDesdeInicio * pEscala.PixelsPorHora)
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: CalcularLarguraCard
' PROPÓSITO: Calcular largura do card baseada na duração
' PARÂMETROS: DataInicio, DataFim, EscalaTemporal
' RETORNO: Largura em pixels
'--------------------------------------------------------------------------------
Public Function CalcularLarguraCard(pDataInicio As Date, pDataFim As Date, pEscala As EscalaTemporal) As Double
    Dim duracaoHoras As Double
    duracaoHoras = (pDataFim - pDataInicio) * 24
    
    If duracaoHoras <= 0 Then
        CalcularLarguraCard = 50
    Else
        CalcularLarguraCard = duracaoHoras * pEscala.PixelsPorHora
    End If
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterIndicePosto
' PROPÓSITO: Obter índice do posto na lista ordenada
' PARÂMETROS: NomePosto, ColecaoPostos
' RETORNO: Índice (0-based)
'--------------------------------------------------------------------------------
Public Function ObterIndicePosto(pNomePosto As String, pColecaoPostos As Collection) As Long
    Dim i As Long
    For i = 1 To pColecaoPostos.Count
        If pColecaoPostos(i) = pNomePosto Then
            ObterIndicePosto = i - 1
            Exit Function
        End If
    Next i
    
    pColecaoPostos.Add pNomePosto
    ObterIndicePosto = pColecaoPostos.Count - 1
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: CalcularTopoCard
' PROPÓSITO: Calcular posição Y do card baseada no posto
' PARÂMETROS: IndicePosto
' RETORNO: Posição Y em pixels
'--------------------------------------------------------------------------------
Public Function CalcularTopoCard(pIndicePosto As Long) As Single
    CalcularTopoCard = 40 + (pIndicePosto * ALTURA_LINHA_POSTO)
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: DetectarConflitos
' PROPÓSITO: Detectar conflitos de horário por posto
' PARÂMETROS: ColecaoOPs, ColecaoPostos
' RETORNO: Colecao de conflitos
'--------------------------------------------------------------------------------
Public Function DetectarConflitos(pColecaoOPs As Collection, pColecaoPostos As Collection) As Collection
    Dim conflitos As New Collection
    Dim i As Long, j As Long
    
    For i = 1 To pColecaoOPs.Count
        Dim opI As Object
        Set opI = pColecaoOPs(i)
        
        Dim indicePostoI As Long
        indicePostoI = ObterIndicePosto(opI.Equipamento, pColecaoPostos)
        
        For j = i + 1 To pColecaoOPs.Count
            Dim opJ As Object
            Set opJ = pColecaoOPs(j)
            
            Dim indicePostoJ As Long
            indicePostoJ = ObterIndicePosto(opJ.Equipamento, pColecaoPostos)
            
            If indicePostoI = indicePostoJ Then
                If opI.DataFim > opJ.DataInicio And opJ.DataFim > opI.DataInicio Then
                    Dim conflito As New Collection
                    conflito.Add opI
                    conflito.Add opJ
                    conflitos.Add conflito
                End If
            End If
        Next j
    Next i
    
    Set DetectarConflitos = conflitos
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterPostosDistintos
' PROPÓSITO: Extrair lista de postos únicos das OPs
' PARÂMETROS: Array de OPs
' RETORNO: Collection de postos
'--------------------------------------------------------------------------------
Public Function ObterPostosDistintos(pDados As Variant) As Collection
    Dim postos As New Collection
    Dim i As Long
    Dim totalLinhas As Long
    
    totalLinhas = UBound(pDados, 1)
    
    For i = 2 To totalLinhas
        Dim equipamento As String
        equipamento = CStr(pDados(i, 3))
        
        If Trim(equipamento) <> "" Then
            Dim existe As Boolean
            existe = False
            Dim j As Long
            For j = 1 To postos.Count
                If postos(j) = equipamento Then
                    existe = True
                    Exit For
                End If
            Next j
            
            If Not existe Then
                postos.Add equipamento
            End If
        End If
    Next i
    
    Set ObterPostosDistintos = postos
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ConverterPeriodoParaTexto
' PROPÓSITO: Retornar texto descritivo do período
' PARÂMETROS: DataInicio, DataFim
' RETORNO: String
'--------------------------------------------------------------------------------
Public Function ConverterPeriodoParaTexto(pDataInicio As Date, pDataFim As Date) As String
    Dim dias As Long
    dias = DateDiff("d", pDataInicio, pDataFim) + 1
    
    If dias = 1 Then
        ConverterPeriodoParaTexto = Format(pDataInicio, "dd/mm/yyyy")
    ElseIf dias <= 7 Then
        ConverterPeriodoParaTexto = Format(pDataInicio, "dd/mm") & " - " & Format(pDataFim, "dd/mm/yyyy")
    Else
        ConverterPeriodoParaTexto = Format(pDataInicio, "dd/mm/yyyy") & " - " & Format(pDataFim, "dd/mm/yyyy")
    End If
End Function

'--------------------------------------------------------------------------------
' SUBROTINA PÚBLICA: CarregarGantt
' PROPÓSITO: Renderizar visualização completa do planejamento (timeline + cards)
' PARÂMETROS: pContainer As MSForms.Frame, pHScroll As MSForms.ScrollBar (opcional)
'--------------------------------------------------------------------------------
Public Sub CarregarGantt(pContainer As MSForms.Frame, Optional pHScroll As MSForms.ScrollBar = Nothing, _
                         Optional pDataInicio As Date = 0, Optional pDataFim As Date = 0, _
                         Optional pZoom As Double = 1#, Optional pFiltroStatus As String = "Todos")
    On Error GoTo ErroCarregarGantt
    
    Dim dados() As Variant
    Dim i As Long, j As Long
    Dim totalLinhas As Long
    Dim numOPs As Long
    
    Dim postos As New Collection
    Dim escala As EscalaTemporal
    
    '--- 1. Leitura de dados ----------------------------------------------------
    dados = ObterDadosOPsEmArray()
    If IsError(dados) Then
        Err.Raise vbObjectError + 400, "CarregarGantt", "Não foi possível carregar dados da TabelaOPs."
    End If
    
    totalLinhas = UBound(dados, 1)
    numOPs = totalLinhas - 1
    
    If numOPs <= 0 Then
        Exit Sub
    End If
    
    '--- 2. Extrai postos distintos ---------------------------------------------
    Dim colPostos As New Collection
    Dim postosArray() As String
    ReDim postosArray(1 To totalLinhas - 1)
    
    Dim idxPosto As Long
    idxPosto = 0
    Dim existePosto As Boolean
    
    For i = 2 To totalLinhas
        Dim nomePosto As String
        nomePosto = Trim(CStr(dados(i, 3)))
        If nomePosto <> "" Then
            existePosto = False
            For j = 1 To colPostos.Count
                If colPostos(j) = nomePosto Then
                    existePosto = True
                    Exit For
                End If
            Next j
            If Not existePosto Then
                colPostos.Add nomePosto
            End If
            idxPosto = idxPosto + 1
            postosArray(idxPosto) = nomePosto
        End If
    Next i
    
    '--- 3. Monta array final de postos únicos ----------------------------------
    Dim numPostos As Long
    numPostos = colPostos.Count
    If numPostos <= 0 Then Exit Sub
    Dim arrPostos() As String
    ReDim arrPostos(1 To numPostos)
    For i = 1 To numPostos
        arrPostos(i) = colPostos(i)
    Next i
    
    '--- 4. Calcula escala temporal ---------------------------------------------
    Dim dataInicio As Date
    Dim dataFim As Date
    
    dataInicio = pDataInicio
    dataFim = pDataFim
    
    If dataFim <= dataInicio Then
        dataFim = DateAdd("d", 1, dataInicio)
    End If
    
    escala = CalcularEscalaTemporal(dataInicio, dataFim, pZoom)
    
    '--- 5. Limpa container -----------------------------------------------------
    Dim ctrl As MSForms.Control
    For Each ctrl In pContainer.Controls
        pContainer.Controls.Remove ctrl.Name
    Next ctrl
    
    '--- 6. Cria cabeçalho do módulo --------------------------------------------
    Dim titulo As MSForms.Label
    Set titulo = pContainer.Controls.Add("Forms.Label.1", "lblGanttTitulo", True)
    With titulo
        .Caption = "PLANEJAMENTO DE PRODUÇÃO – " & ConverterPeriodoParaTexto(dataInicio, dataFim)
        .Left = 0
        .Top = 0
        .Width = pContainer.Width
        .Height = 28
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_CABECALHO
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- 7. Cria timeline (rótulos de tempo) ------------------------------------
    Dim topoTimeline As Single
    topoTimeline = 32
    
    Dim margemEsquerda As Single
    margemEsquerda = MARGEM_ESQUERDA
    
    Dim larguraTimeline As Double
    larguraTimeline = escala.LarguraTotal
    
    ' Fundo da timeline
    Dim fundoTimeline As MSForms.Label
    Set fundoTimeline = pContainer.Controls.Add("Forms.Label.1", "lblGanttFundoTimeline", True)
    With fundoTimeline
        .Left = margemEsquerda
        .Top = topoTimeline
        .Width = larguraTimeline
        .Height = 28
        .BackColor = COR_CABECALHO
    End With
    
    ' Marcadores de tempo na timeline
    Dim intervaloMarcador As Double
    Dim unidade As String
    Dim totalDias As Long
    totalDias = DateDiff("d", dataInicio, dataFim) + 1
    
    If totalDias = 1 Then
        intervaloMarcador = 1 / 24
        unidade = "hora"
    ElseIf totalDias <= 7 Then
        intervaloMarcador = 1
        unidade = "dia"
    ElseIf totalDias <= 31 Then
        intervaloMarcador = 1
        unidade = "dia"
    Else
        intervaloMarcador = 7
        unidade = "semana"
    End If
    
    Dim passo As Double
    passo = intervaloMarcador
    Dim contadorMarcador As Long
    contadorMarcador = 0
    
    Do While dataInicio + passo <= dataFim Or (contadorMarcador = 0)
        Dim posMarcador As Double
        Dim dataMarcador As Date
        dataMarcador = dataInicio + passo
        
        posMarcador = ConverterDataParaX(dataMarcador, escala)
        
        Dim textoMarcador As String
        If unidade = "hora" Then
            textoMarcador = Format(dataMarcador, "HH:MM")
        ElseIf unidade = "dia" Then
            textoMarcador = Format(dataMarcador, "dd/mm")
        Else
            textoMarcador = Format(dataMarcador, "'sem' dd/mm")
        End If
        
        Dim lblMarcador As MSForms.Label
        Set lblMarcador = pContainer.Controls.Add("Forms.Label.1", "lblGanttMarc_" & contadorMarcador, True)
        With lblMarcador
            .Caption = textoMarcador
            .Left = posMarcador
            .Top = topoTimeline + 4
            .Width = 60
            .Height = 20
            .Font.Size = 8
            .ForeColor = COR_TEXTO_CLARO
            .BackColor = COR_CABECALHO
            .TextAlign = fmTextAlignCenter
        End With
        
        ' Linha vertical de grade
        Dim linhaGrade As MSForms.Label
        Set linhaGrade = pContainer.Controls.Add("Forms.Label.1", "lblGanttGrade_" & contadorMarcador, True)
        With linhaGrade
            .Left = posMarcador
            .Top = topoTimeline + 28
            .Width = 1
            .Height = numPostos * ALTURA_LINHA_POSTO + 40
            .BackColor = COR_CINZA_LINHA
        End With
        
        passo = passo + intervaloMarcador
        contadorMarcador = contadorMarcador + 1
    Loop
    
    '--- 8. Cria área de fundo dos postos ----------------------------------------
    Dim topoPostos As Single
    topoPostos = topoTimeline + 28
    
    Dim fundoPostos As MSForms.Label
    Set fundoPostos = pContainer.Controls.Add("Forms.Label.1", "lblGanttFundoPostos", True)
    With fundoPostos
        .Left = 0
        .Top = topoPostos
        .Width = pContainer.Width
        .Height = numPostos * ALTURA_LINHA_POSTO + 40
        .BackColor = COR_FUNDO_TIMELINE
    End With
    
    '--- 9. Cria labels dos postos (eixo Y) --------------------------------------
    Dim topoCard As Single
    Dim nomePostoAtual As String
    
    For i = 1 To numPostos
        nomePostoAtual = arrPostos(i)
        topoCard = topoPostos + 8 + ((i - 1) * ALTURA_LINHA_POSTO)
        
        Dim lblPosto As MSForms.Label
        Set lblPosto = pContainer.Controls.Add("Forms.Label.1", "lblGanttPosto_" & i, True)
        With lblPosto
            .Caption = nomePostoAtual
            .Left = 4
            .Top = topoCard
            .Width = margemEsquerda - 8
            .Height = 24
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = COR_FUNDO_TIMELINE
            .TextAlign = fmTextAlignRight
        End With
        
        ' Linha separadora horizontal
        Dim linhaH As MSForms.Label
        Set linhaH = pContainer.Controls.Add("Forms.Label.1", "lblGanttLinhaH_" & i, True)
        With linhaH
            .Left = margemEsquerda
            .Top = topoCard + 26
            .Width = larguraTimeline
            .Height = 1
            .BackColor = COR_CINZA_LINHA
        End With
    Next i
    
    '--- 10. Detecta conflitos --------------------------------------------------
    Dim colecaoOPs As New Collection
    Dim arrConflitos() As Boolean
    ReDim arrConflitos(1 To numOPs)
    
    Dim opI As clsCardProducao
    Dim opJ As clsCardProducao
    
    Dim indicePostoI As Long
    Dim indicePostoJ As Long
    
    Dim posArrayOP As Long
    posArrayOP = 2
    
    ' Primeiro cria todos os objetos OP
    Dim ops() As clsCardProducao
    ReDim ops(1 To numOPs)
    
    For i = 1 To numOPs
        Set ops(i) = New clsCardProducao
        With ops(i)
            .ID_OP = CStr(dados(posArrayOP, 1))
            .Produto = CStr(dados(posArrayOP, 2))
            .Equipamento = CStr(dados(posArrayOP, 3))
            .Quantidade = CLng(dados(posArrayOP, 4))
            If IsDate(dados(posArrayOP, 5)) Then .DataInicio = CDate(dados(posArrayOP, 5))
            If IsDate(dados(posArrayOP, 6)) Then .DataFim = CDate(dados(posArrayOP, 6))
            .Duracao = CDbl(dados(posArrayOP, 7))
            .Status = CStr(dados(posArrayOP, 8))
        End With
        colecaoOPs.Add ops(i)
        posArrayOP = posArrayOP + 1
    Next i
    
    ' Detecta conflitos
    Dim conflitos As Collection
    Set conflitos = DetectarConflitos(colecaoOPs, colPostos)
    
    ' Marca OPs com conflito
    Dim conflitoMarcado() As Boolean
    ReDim conflitoMarcado(1 To numOPs)
    Dim conflitoCount As Long
    conflitoCount = 0
    
    For i = 1 To conflitos.Count
        Dim par As Collection
        Set par = conflitos(i)
        If par.Count >= 2 Then
            Dim idxA As Long, idxB As Long
            idxA = 0
            idxB = 0
            For j = 1 To numOPs
                If ops(j).ID_OP = par(1).ID_OP Then idxA = j
                If ops(j).ID_OP = par(2).ID_OP Then idxB = j
            Next j
            If idxA > 0 Then conflitoMarcado(idxA) = True
            If idxB > 0 Then conflitoMarcado(idxB) = True
        End If
    Next i
    
    '--- 11. Renderiza cards das OPs ---------------------------------------------
    Dim alturaCard As Single
    alturaCard = ALTURA_LINHA_POSTO - 8
    
    Dim uniqueId As Long
    uniqueId = 0
    
    Dim opAtual As clsCardProducao
    Dim indicePosto As Long
    Dim posX As Double
    Dim larguraCard As Double
    Dim corCard As Long
    Dim corBorda As Long
    Dim corTexto As Long
    
    For i = 1 To numOPs
        Set opAtual = ops(i)
        
        ' Aplica filtro de status
        If pFiltroStatus <> "Todos" Then
            If Trim(opAtual.Status) <> Trim(pFiltroStatus) Then
                GoTo ProximaOP
            End If
        End If
        
        indicePosto = ObterIndicePosto(opAtual.Equipamento, colPostos)
        posX = ConverterDataParaX(opAtual.DataInicio, escala)
        larguraCard = CalcularLarguraCard(opAtual.DataInicio, opAtual.DataFim, escala)
        topoCard = topoPostos + 8 + (indicePosto * ALTURA_LINHA_POSTO)
        
        uniqueId = uniqueId + 1
        
        ' Define cores conforme status e conflito
        corTexto = COR_TEXTO_ESCURO
        Select Case Trim(opAtual.Status)
            Case "Atrasado"
                corBorda = COR_ALERTA
                corCard = COR_CINZA_ATRASO
            Case "Em Andamento", "Concluído"
                corBorda = COR_AZUL
                corCard = COR_CINZA_CARD
            Case Else
                corBorda = COR_HEADER
                corCard = COR_CINZA_CARD
        End Select
        
        If conflitoMarcado(i) Then
            corBorda = COR_ALERTA
        End If
        
        ' Card container
        Dim card As MSForms.Label
        Set card = pContainer.Controls.Add("Forms.Label.1", "cardGantt_" & uniqueId, True)
        With card
            .Caption = ""
            .Tag = opAtual.ID_OP
            .Left = posX
            .Top = topoCard
            .Width = larguraCard
            .Height = alturaCard
            .BackColor = corCard
            .BorderStyle = fmBorderStyleSingle
        End With
        
        ' Borda colorida (esquerda)
        Dim bordaEsq As MSForms.Label
        Set bordaEsq = pContainer.Controls.Add("Forms.Label.1", "cardGanttBorda_" & uniqueId & "_E", True)
        With bordaEsq
            .Left = posX
            .Top = topoCard
            .Width = 4
            .Height = alturaCard
            .BackColor = corBorda
        End With
        
        ' ID OP
        Dim lblID As MSForms.Label
        Set lblID = pContainer.Controls.Add("Forms.Label.1", "cardGanttID_" & uniqueId, True)
        With lblID
            .Caption = opAtual.ID_OP
            .Left = posX + 8
            .Top = topoCard + 4
            .Width = larguraCard - 12
            .Height = 16
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = corTexto
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Produto
        Dim lblProd As MSForms.Label
        Set lblProd = pContainer.Controls.Add("Forms.Label.1", "cardGanttProd_" & uniqueId, True)
        With lblProd
            .Caption = Left(opAtual.Produto, 25)
            .Left = posX + 8
            .Top = topoCard + 22
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .ForeColor = corTexto
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Período
        Dim lblPer As MSForms.Label
        Set lblPer = pContainer.Controls.Add("Forms.Label.1", "cardGanttPer_" & uniqueId, True)
        With lblPer
            If IsDate(opAtual.DataInicio) And IsDate(opAtual.DataFim) Then
                .Caption = Format(opAtual.DataInicio, "HH:MM") & " - " & Format(opAtual.DataFim, "HH:MM")
            Else
                .Caption = "--:-- - --:--"
            End If
            .Left = posX + 8
            .Top = topoCard + 38
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .ForeColor = corTexto
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Duração
        Dim lblDur As MSForms.Label
        Set lblDur = pContainer.Controls.Add("Forms.Label.1", "cardGanttDur_" & uniqueId, True)
        With lblDur
            .Caption = Format(opAtual.Duracao, "0.0") & "h"
            .Left = posX + 8
            .Top = topoCard + 54
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .ForeColor = corTexto
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Status
        Dim lblStat As MSForms.Label
        Set lblStat = pContainer.Controls.Add("Forms.Label.1", "cardGanttStat_" & uniqueId, True)
        With lblStat
            .Caption = opAtual.Status
            .Left = posX + 8
            .Top = topoCard + 70
            .Width = larguraCard - 12
            .Height = 14
            .Font.Size = 8
            .Font.Bold = True
            .ForeColor = corBorda
            .BackColor = corCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Indicador de conflito (triângulo)
        If conflitoMarcado(i) Then
            Dim lblAlerta As MSForms.Label
            Set lblAlerta = pContainer.Controls.Add("Forms.Label.1", "cardGanttAlerta_" & uniqueId, True)
            With lblAlerta
                .Caption = "!"
                .Left = posX + larguraCard - 20
                .Top = topoCard + 4
                .Width = 16
                .Height = 16
                .Font.Size = 10
                .Font.Bold = True
                .ForeColor = vbWhite
                .BackColor = COR_ALERTA
                .TextAlign = fmTextAlignCenter
            End With
        End If
ProximaOP:
    Next i
    
    '--- 12. Linha AGORA --------------------------------------------------------
    Dim agora As Date
    agora = Now
    
    If agora >= dataInicio And agora <= dataFim Then
        Dim posAgora As Double
        posAgora = ConverterDataParaX(agora, escala)
        
        Dim linhaAgora As MSForms.Label
        Set linhaAgora = pContainer.Controls.Add("Forms.Label.1", "lblGanttAgora", True)
        With linhaAgora
            .Left = posAgora
            .Top = topoPostos
            .Width = 2
            .Height = numPostos * ALTURA_LINHA_POSTO + 40
            .BackColor = COR_ALERTA
        End With
        
        Dim lblAgora As MSForms.Label
        Set lblAgora = pContainer.Controls.Add("Forms.Label.1", "lblGanttAgoraTxt", True)
        With lblAgora
            .Caption = "AGORA"
            .Left = posAgora + 4
            .Top = topoPostos
            .Width = 50
            .Height = 16
            .Font.Size = 8
            .Font.Bold = True
            .ForeColor = COR_ALERTA
            .BackColor = COR_FUNDO_TIMELINE
            .TextAlign = fmTextAlignLeft
        End With
    End If
    
    ' Ajusta scroll vertical do container
    pContainer.ScrollHeight = numPostos * ALTURA_LINHA_POSTO + 40 + 28
    
    '--- 13. Configura scroll horizontal -----------------------------------------
    If Not pHScroll Is Nothing Then
        Dim scrollMax As Long
        scrollMax = larguraTimeline + margemEsquerda
        If scrollMax < pContainer.Width Then scrollMax = 0
        pHScroll.Max = scrollMax
        pHScroll.LargeChange = pContainer.Width / 2
        pHScroll.SmallChange = 50
        pHScroll.Value = 0
    End If
    
Sair:
    Exit Sub
    
ErroCarregarGantt:
    MsgBox "Erro ao carregar planejamento: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Planejamento"
    Resume Sair
End Sub
