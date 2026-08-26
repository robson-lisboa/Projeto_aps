Attribute VB_Name = "modComparacao"
'================================================================================
' MÓDULO: modComparacao
' DESCRIÇÃO: Lógica de comparação entre duas simulações de produção
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: CompararSimulacoes
' PROPÓSITO: Comparar duas simulações e retornar lista de diferenças
' PARÂMETROS: pID_SimulacaoA, pID_SimulacaoB
' RETORNO: Collection de clsComparacaoOP
'--------------------------------------------------------------------------------
Public Function CompararSimulacoes(pID_SimulacaoA As String, pID_SimulacaoB As String) As Collection
    Dim resultados As New Collection
    Dim dadosA As Variant
    Dim dadosB As Variant
    Dim mapaA As Object
    Dim mapaB As Object
    Dim i As Long
    Dim chave As String
    
    On Error GoTo ErroComparar
    
    dadosA = ObterOPsSimulacaoEmArray(pID_SimulacaoA)
    dadosB = ObterOPsSimulacaoEmArray(pID_SimulacaoB)
    
    If IsError(dadosA) Then dadosA = Array()
    If IsError(dadosB) Then dadosB = Array()
    
    Set mapaA = BuildMapaOPs(dadosA)
    Set mapaB = BuildMapaOPs(dadosB)
    
    Dim todasChaves As Object
    Set todasChaves = BuildTodasChaves(mapaA, mapaB)
    
    For Each chave In todasChaves
        Dim item As clsComparacaoOP
        Set item = New clsComparacaoOP
        item.ID_OP = chave
        
        If mapaA.Exists(chave) And mapaB.Exists(chave) Then
            item.ExisteA = True
            item.ExisteB = True
            Call CompararOP(item, mapaA(chave), mapaB(chave))
        ElseIf mapaA.Exists(chave) Then
            item.ExisteA = True
            item.ExisteB = False
            Call PreencherDadosA(item, mapaA(chave))
            item.TipoDiferenca = "SOMENTE_A"
            item.Observacao = "Existente apenas na Simulação A"
        Else
            item.ExisteA = False
            item.ExisteB = True
            Call PreencherDadosB(item, mapaB(chave))
            item.TipoDiferenca = "SOMENTE_B"
            item.Observacao = "Existente apenas na Simulação B"
        End If
        
        resultados.Add item
    Next chave
    
    Set CompararSimulacoes = resultados
    
Sair:
    Exit Function
    
ErroComparar:
    MsgBox "Erro ao comparar simulações: " & Err.Description, vbCritical, "APS PURAN"
    Set CompararSimulacoes = New Collection
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO PRIVADA: BuildMapaOPs
' PROPÓSITO: Construir dicionário de OPs a partir do array de dados
'--------------------------------------------------------------------------------
Private Function BuildMapaOPs(pDados As Variant) As Object
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim totalLinhas As Long
    If IsEmpty(pDados) Then
        Set BuildMapaOPs = mapa
        Exit Function
    End If
    totalLinhas = UBound(pDados, 1)
    
    For i = 1 To totalLinhas
        Dim idOP As String
        idOP = CStr(pDados(i, 2))
        If idOP <> "" Then
            Dim dadosOP As Variant
            dadosOP = Array( _
                CStr(pDados(i, 4)), _ ' Produto
                CStr(pDados(i, 5)), _ ' Equipamento
                CStr(pDados(i, 6)), _ ' Quantidade
                IIf(IsDate(pDados(i, 7)), CDate(pDados(i, 7)), 0), _
                IIf(IsDate(pDados(i, 8)), CDate(pDados(i, 8)), 0), _
                CDbl(pDados(i, 9)), _ ' Duracao
                CStr(pDados(i, 10)) _ ' Status
            )
            mapa(idOP) = dadosOP
        End If
    Next i
    
    Set BuildMapaOPs = mapa
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO PRIVADA: BuildTodasChaves
' PROPÓSITO: Unir chaves dos dois mapas sem duplicatas
'--------------------------------------------------------------------------------
Private Function BuildTodasChaves(pMapaA As Object, pMapaB As Object) As Collection
    Dim todas As New Collection
    Dim chave As Variant
    
    On Error Resume Next
    For Each chave In pMapaA.Keys
        todas.Add chave, CStr(chave)
    Next chave
    For Each chave In pMapaB.Keys
        todas.Add chave, CStr(chave)
    Next chave
    On Error GoTo 0
    
    Set BuildTodasChaves = todas
End Function

'--------------------------------------------------------------------------------
' SUBROTINA PRIVADA: CompararOP
' PROPÓSITO: Comparar campos de uma OP existente em ambas as simulações
'--------------------------------------------------------------------------------
Private Sub CompararOP(pItem As clsComparacaoOP, pDadosA As Variant, pDadosB As Variant)
    Dim houveAlteracao As Boolean
    houveAlteracao = False
    
    pItem.EquipamentoA = pDadosA(1)
    pItem.EquipamentoB = pDadosB(1)
    pItem.DataInicioA = pDadosA(3)
    pItem.DataInicioB = pDadosB(3)
    pItem.DataFimA = pDadosA(4)
    pItem.DataFimB = pDadosB(4)
    pItem.DuracaoA = pDadosA(5)
    pItem.DuracaoB = pDadosB(5)
    pItem.StatusA = pDadosA(6)
    pItem.StatusB = pDadosB(6)
    
    Dim partes As String
    partes = ""
    
    If StrComp(pItem.EquipamentoA, pItem.EquipamentoB, vbTextCompare) <> 0 Then
        partes = partes & "Posto;"
        houveAlteracao = True
    End If
    
    If DateDiff("s", pItem.DataInicioA, pItem.DataInicioB) <> 0 Then
        partes = partes & "Início;"
        houveAlteracao = True
    End If
    
    If DateDiff("s", pItem.DataFimA, pItem.DataFimB) <> 0 Then
        partes = partes & "Fim;"
        houveAlteracao = True
    End If
    
    If Abs(pItem.DuracaoA - pItem.DuracaoB) > 0.01 Then
        partes = partes & "Duração;"
        houveAlteracao = True
    End If
    
    If StrComp(pItem.StatusA, pItem.StatusB, vbTextCompare) <> 0 Then
        partes = partes & "Status;"
        houveAlteracao = True
    End If
    
    If houveAlteracao Then
        pItem.TipoDiferenca = "ALTERADA"
        If Len(partes) > 0 Then
            If Right(partes, 1) = ";" Then partes = Left(partes, Len(partes) - 1)
        End If
        pItem.Observacao = "Alterado(s): " & partes
    Else
        pItem.TipoDiferenca = "IGUAL"
        pItem.Observacao = "Sem alterações"
    End If
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA PRIVADA: PreencherDadosA
' PROPÓSITO: Preencher item com dados apenas da simulação A
'--------------------------------------------------------------------------------
Private Sub PreencherDadosA(pItem As clsComparacaoOP, pDados As Variant)
    pItem.EquipamentoA = pDados(1)
    pItem.DataInicioA = pDados(3)
    pItem.DataFimA = pDados(4)
    pItem.DuracaoA = pDados(5)
    pItem.StatusA = pDados(6)
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA PRIVADA: PreencherDadosB
' PROPÓSITO: Preencher item com dados apenas da simulação B
'--------------------------------------------------------------------------------
Private Sub PreencherDadosB(pItem As clsComparacaoOP, pDados As Variant)
    pItem.EquipamentoB = pDados(1)
    pItem.DataInicioB = pDados(3)
    pItem.DataFimB = pDados(4)
    pItem.DuracaoB = pDados(5)
    pItem.StatusB = pDados(6)
End Sub
