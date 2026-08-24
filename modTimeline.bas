Attribute VB_Name = "modTimeline"
'================================================================================
' MÓDULO: modTimeline
' DESCRIÇÃO: Renderização da Timeline de Produção (Gantt simplificado) para APS PURAN
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' CONSTANTES VISUAIS
'--------------------------------------------------------------------------------
Private Const COR_FUNDO_TIMELINE As Long = 14211288
Private Const COR_CABECALHO As Long = 3355443
Private Const COR_TEXTO_CLARO As Long = 16777215
Private Const COR_TEXTO_ESCURO As Long = 0
Private Const COR_AZUL As Long = 15773696
Private Const COR_ALERTA As Long = 255
Private Const COR_CINZA_LINHA As Long = 16777215
Private Const COR_CINZA_ATRASO As Long = 10092543

Private Const COLUNA_LARGURA As Single = 120
Private Const COLUNA_ALTURA As Single = 22
Private Const LINHA_ALTURA As Single = 40
Private Const MARGEM_ESQ As Single = 20
Private Const MARGEM_SUP As Single = 70

'--------------------------------------------------------------------------------
' SUBROTINA PÚBLICA: CarregarTimeline
' PROPÓSITO: Popular grade visual da Timeline no frame de conteúdo do formulário
'            principal, com formatação condicional por status
'--------------------------------------------------------------------------------
Public Sub CarregarTimeline(pContainer As MSForms.Frame)
    On Error GoTo ErroCarregarTimeline
    
    Dim dados() As Variant
    Dim i As Long
    Dim totalLinhas As Long
    Dim numOPs As Long
    
    '--- 1. Leitura otimizada via array em memória ------------------------------
    dados = ObterDadosOPsEmArray()
    
    If IsError(dados) Then
        Err.Raise vbObjectError + 300, "CarregarTimeline", _
            "Não foi possível carregar dados da TabelaOPs."
    End If
    
    totalLinhas = UBound(dados, 1)
    numOPs = totalLinhas - 1
    
    '--- 2. Limpa conteúdo anterior do container --------------------------------
    Dim ctrl As MSForms.Control
    For Each ctrl In pContainer.Controls
        pContainer.Controls.Remove ctrl.Name
    Next ctrl
    
    '--- 3. Cria cabeçalho do módulo --------------------------------------------
    Dim titulo As MSForms.Label
    Set titulo = pContainer.Controls.Add("Forms.Label.1", "lblTimelineTitulo", True)
    With titulo
        .Caption = "TIMELINE DE PRODUÇÃO – ORDENS DE PRODUÇÃO"
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
    
    '--- 4. Define posicionamento das colunas -----------------------------------
    Dim colID As Single: colID = MARGEM_ESQ
    Dim colProd As Single: colProd = colID + COLUNA_LARGURA
    Dim colEquip As Single: colEquip = colProd + COLUNA_LARGURA
    Dim colInicio As Single: colInicio = colEquip + COLUNA_LARGURA
    Dim colFim As Single: colFim = colInicio + COLUNA_LARGURA
    Dim colStatus As Single: colStatus = colFim + COLUNA_LARGURA
    
    Dim cabecalhos As Variant
    cabecalhos = Array("ID OP", "Produto", "Equipamento", "Início", "Fim", "Status")
    
    Dim posicoes As Variant
    posicoes = Array(colID, colProd, colEquip, colInicio, colFim, colStatus)
    
    '--- 5. Renderiza cabeçalhos das colunas ------------------------------------
    Dim j As Long
    For j = LBound(cabecalhos) To UBound(cabecalhos)
        Dim hdr As MSForms.Label
        Set hdr = pContainer.Controls.Add("Forms.Label.1", "lblTimelineHdr" & j, True)
        With hdr
            .Caption = cabecalhos(j)
            .Left = posicoes(j)
            .Top = MARGEM_SUP
            .Width = COLUNA_LARGURA
            .Height = COLUNA_ALTURA
            .Font.Size = 9
            .Font.Bold = True
            .ForeColor = COR_TEXTO_CLARO
            .BackColor = COR_CABECALHO
            .TextAlign = fmTextAlignCenter
            .BorderStyle = fmBorderStyleSingle
        End With
    Next j
    
    '--- 6. Popula linhas de dados com formatação condicional -------------------
    Dim linhaY As Single
    linhaY = MARGEM_SUP + COLUNA_ALTURA + 2
    
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
        equipamento = CStr(dados(posArray, 3))
        inicio = CDate(dados(posArray, 5))
        fim = CDate(dados(posArray, 6))
        status = CStr(dados(posArray, 8))
        
        Dim corFundo As Long
        Dim corStatus As Long
        
        Select Case Trim(status)
            Case "Atrasado"
                corFundo = COR_CINZA_ATRASO
                corStatus = COR_ALERTA
            Case "Em Andamento", "Concluído"
                corFundo = COR_CINZA_LINHA
                corStatus = COR_AZUL
            Case Else
                corFundo = COR_CINZA_LINHA
                corStatus = COR_TEXTO_ESCURO
        End Select
        
        Dim valores As Variant
        valores = Array(idOP, produto, equipamento, Format(inicio, "dd/mm/yy"), Format(fim, "dd/mm/yy"), status)
        
        Dim m As Long
        For m = LBound(valores) To UBound(valores)
            Dim celula As MSForms.Label
            Set celula = pContainer.Controls.Add("Forms.Label.1", "lblTimeline_" & k & "_" & m, True)
            With celula
                .Caption = valores(m)
                .Left = posicoes(m)
                .Top = linhaY
                .Width = COLUNA_LARGURA
                .Height = LINHA_ALTURA - 2
                .Font.Size = 9
                .ForeColor = COR_TEXTO_ESCURO
                .BackColor = corFundo
                .TextAlign = fmTextAlignCenter
                .BorderStyle = fmBorderStyleSingle
            End With
        Next m
        
        ' Indicador visual de status (bolinha colorida)
        Dim indicador As MSForms.Label
        Set indicador = pContainer.Controls.Add("Forms.Label.1", "lblTimelineInd_" & k, True)
        With indicador
            .Left = posicoes(5) + COLUNA_LARGURA + 8
            .Top = linhaY + 8
            .Width = 14
            .Height = 14
            .BackColor = corStatus
            .BorderStyle = fmBorderStyleSingle
        End With
        
        linhaY = linhaY + LINHA_ALTURA
        posArray = posArray + 1
    Next k
    
Sair:
    Exit Sub
    
ErroCarregarTimeline:
    MsgBox "Erro ao carregar Timeline: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN – Timeline"
    Resume Sair
End Sub
