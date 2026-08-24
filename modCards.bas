Attribute VB_Name = "modCards"
'================================================================================
' MÓDULO: modCards
' DESCRIÇÃO: Renderização de Cards de Produção (Visão Kanban) para APS PURAN
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' CONSTANTES VISUAIS
'--------------------------------------------------------------------------------
Private Const COR_FUNDO_CONTAINER As Long = 14211288
Private Const COR_CABECALHO As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_AZUL As Long = 15773696
Private Const COR_ALERTA As Long = 255
Private Const COR_CINZA_CARD As Long = 16777215
Private Const COR_CINZA_ATRASO As Long = 10092543

Private Const CARD_LARGURA As Single = 220
Private Const CARD_ALTURA As Single = 140
Private Const CARD_MARGEM As Single = 12
Private Const CARDS_POR_LINHA As Long = 3
Private Const TOPO_FILTROS As Single = 50

'--------------------------------------------------------------------------------
' SUBROTINA PÚBLICA: CarregarCards
' PROPÓSITO: Renderizar mini-cartões Kanban de OPs dentro do container fornecido
' PARÂMETROS: pContainer As MSForms.Frame – frame de conteúdo do frmPrincipal
'--------------------------------------------------------------------------------
Public Sub CarregarCards(pContainer As MSForms.Frame)
    On Error GoTo ErroCarregarCards
    
    Dim dados() As Variant
    Dim i As Long
    Dim totalLinhas As Long
    Dim numOPs As Long
    
    '--- 1. Leitura otimizada via array em memória ------------------------------
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then
        Err.Raise vbObjectError + 500, "CarregarCards", _
            "Não foi possível carregar dados da TabelaOPs."
    End If
    
    totalLinhas = UBound(dados, 1)
    numOPs = totalLinhas - 1
    
    '--- 2. Limpa conteúdo anterior do container --------------------------------
    Dim ctrl As MSForms.Control
    For Each ctrl In pContainer.Controls
        pContainer.Controls.Remove ctrl.Name
    Next ctrl
    
    '--- 3. Cabeçalho do módulo -------------------------------------------------
    Dim titulo As MSForms.Label
    Set titulo = pContainer.Controls.Add("Forms.Label.1", "lblCardsTitulo", True)
    With titulo
        .Caption = "CARDS DE PRODUÇÃO – VISÃO KANBAN"
        .Left = 10
        .Top = 8
        .Width = pContainer.Width - 20
        .Height = 28
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COR_TEXTO_CLARO
        .BackColor = COR_CABECALHO
        .TextAlign = fmTextAlignCenter
    End With
    
    '--- 4. Renderiza cartões em grade ------------------------------------------
    Dim cardIndex As Long
    cardIndex = 0
    
    Dim posArray As Long
    posArray = 2
    
    Dim k As Long
    For k = 1 To numOPs
        Dim idOP As String
        Dim produto As String
        Dim equipamento As String
        Dim inicio As Date
        Dim fim As Date
        Dim status As String
        
        idOP = CStr(dados(posArray, 1))
        produto = CStr(dados(posArray, 2))
        equipamento = CStr(posArray, 3)
        inicio = CDate(dados(posArray, 5))
        fim = CDate(dados(posArray, 6))
        status = CStr(dados(posArray, 8))
        
        Dim corBorda As Long
        Dim corFundoCard As Long
        
        Select Case Trim(status)
            Case "Atrasado"
                corBorda = COR_ALERTA
                corFundoCard = COR_CINZA_ATRASO
            Case "Em Andamento", "Concluído"
                corBorda = COR_AZUL
                corFundoCard = COR_CINZA_CARD
            Case Else
                corBorda = COR_CABECALHO
                corFundoCard = COR_CINZA_CARD
        End Select
        
        ' Posicionamento em grade
        Dim col As Long
        Dim linha As Long
        col = cardIndex Mod CARDS_POR_LINHA
        linha = cardIndex \ CARDS_POR_LINHA
        
        Dim cardX As Single
        Dim cardY As Single
        cardX = CARD_MARGEM + col * (CARD_LARGURA + CARD_MARGEM)
        cardY = TOPO_FILTROS + CARD_MARGEM + linha * (CARD_ALTURA + CARD_MARGEM)
        
        ' Frame container do cartão
        Dim cardFrame As MSForms.Label
        Set cardFrame = pContainer.Controls.Add("Forms.Label.1", "cardFrame_" & k, True)
        With cardFrame
            .Left = cardX
            .Top = cardY
            .Width = CARD_LARGURA
            .Height = CARD_ALTURA
            .BackColor = corFundoCard
            .BorderStyle = fmBorderStyleSingle
        End With
        
        ' Título: OP
        Dim cardTitulo As MSForms.Label
        Set cardTitulo = pContainer.Controls.Add("Forms.Label.1", "cardTitulo_" & k, True)
        With cardTitulo
            .Caption = "OP " & idOP
            .Left = cardX + 8
            .Top = cardY + 8
            .Width = CARD_LARGURA - 16
            .Height = 22
            .Font.Size = 11
            .Font.Bold = True
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corFundoCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Produto
        Dim cardProd As MSForms.Label
        Set cardProd = pContainer.Controls.Add("Forms.Label.1", "cardProd_" & k, True)
        With cardProd
            .Caption = produto
            .Left = cardX + 8
            .Top = cardY + 32
            .Width = CARD_LARGURA - 16
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corFundoCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Equipamento
        Dim cardEquip As MSForms.Label
        Set cardEquip = pContainer.Controls.Add("Forms.Label.1", "cardEquip_" & k, True)
        With cardEquip
            .Caption = equipamento
            .Left = cardX + 8
            .Top = cardY + 52
            .Width = CARD_LARGURA - 16
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corFundoCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Período
        Dim cardPeriodo As MSForms.Label
        Set cardPeriodo = pContainer.Controls.Add("Forms.Label.1", "cardPeriodo_" & k, True)
        With cardPeriodo
            .Caption = Format(inicio, "dd/mm/yy") & " - " & Format(fim, "dd/mm/yy")
            .Left = cardX + 8
            .Top = cardY + 72
            .Width = CARD_LARGURA - 16
            .Height = 18
            .Font.Size = 9
            .ForeColor = COR_TEXTO_ESCURO
            .BackColor = corFundoCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Status com indicador de alerta
        Dim cardStatus As MSForms.Label
        Set cardStatus = pContainer.Controls.Add("Forms.Label.1", "cardStatus_" & k, True)
        With cardStatus
            .Caption = status
            .Left = cardX + 8
            .Top = cardY + 92
            .Width = CARD_LARGURA - 30
            .Height = 18
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = corBorda
            .BackColor = corFundoCard
            .TextAlign = fmTextAlignLeft
        End With
        
        ' Barra de alerta lateral (vermelha se atrasado)
        Dim barraAlerta As MSForms.Label
        Set barraAlerta = pContainer.Controls.Add("Forms.Label.1", "cardBarra_" & k, True)
        With barraAlerta
            .Left = cardX + CARD_LARGURA - 12
            .Top = cardY
            .Width = 8
            .Height = CARD_ALTURA
            .BackColor = corBorda
        End With
        
        cardIndex = cardIndex + 1
        posArray = posArray + 1
    Next k
    
Sair:
    Exit Sub
    
ErroCarregarCards:
    MsgBox "Erro ao carregar Cards de Produção: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Cards"
    Resume Sair
End Sub
