Attribute VB_Name = "modEngine"
'================================================================================
' MÓDULO: modEngine
' DESCRIÇÃO: Motor de performance em arrays para leitura/escrita otimizada de OPs
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterDadosOPsEmArray
' PROPÓSITO: Ler toda a TabelaOPs para memória RAM (Variant Array) otimizando
'            a velocidade de preenchimento de ListBoxes e demais controles
' RETORNO: Variant Array (1-based) contendo todos os dados da tabela
'--------------------------------------------------------------------------------
Public Function ObterDadosOPsEmArray() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dados As Variant
    
    On Error GoTo ErroObterDados
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    dados = tbl.Range.Value
    
    ObterDadosOPsEmArray = dados
    
Sair:
    Exit Function
    
ErroObterDados:
    ObterDadosOPsEmArray = CVErr(xlErrRef)
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: SalvarNovaOP
' PROPÓSITO: Inserir uma nova Ordem de Produção na última linha livre da TabelaOPs
' PARÂMETROS: Dados da OP a ser persistida
'--------------------------------------------------------------------------------
Public Sub SalvarNovaOP(pID As String, _
                        pProduto As String, _
                        pEquipamento As String, _
                        pQtd As Long, _
                        pInicio As Date, _
                        pFim As Date, _
                        pDuracao As Double, _
                        pStatus As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error GoTo ErroSalvar
    
    ' Validação de campos obrigatórios
    If Trim(pID) = "" Or Trim(pProduto) = "" Or Trim(pEquipamento) = "" Or pQtd <= 0 Then
        Err.Raise vbObjectError + 100, "SalvarNovaOP", _
            "Campos obrigatórios não foram preenchidos corretamente."
    End If
    
    If pFim < pInicio Then
        Err.Raise vbObjectError + 101, "SalvarNovaOP", _
            "A Data de Fim não pode ser anterior à Data de Início."
    End If
    
    Set ws = ThisWorkbook.Worksheets("BD_OPs")
    Set tbl = ws.ListObjects("TabelaOPs")
    
    Application.ScreenUpdating = False
    
    ' Localiza a última linha livre e insere nova linha
    Set novaLinha = tbl.ListRows.Add
    
    With novaLinha.Range
        .Cells(1, tbl.ListColumns("ID_OP").Index).Value = pID
        .Cells(1, tbl.ListColumns("Produto").Index).Value = pProduto
        .Cells(1, tbl.ListColumns("Equipamento").Index).Value = pEquipamento
        .Cells(1, tbl.ListColumns("Quantidade").Index).Value = pQtd
        .Cells(1, tbl.ListColumns("Data_Inicio").Index).Value = pInicio
        .Cells(1, tbl.ListColumns("Data_Fim").Index).Value = pFim
        .Cells(1, tbl.ListColumns("Duracao_Horas").Index).Value = pDuracao
        .Cells(1, tbl.ListColumns("Status").Index).Value = pStatus
    End With
    
Sair:
    Application.ScreenUpdating = True
    Exit Sub
    
ErroSalvar:
    MsgBox "Erro ao salvar nova OP: " & Err.Description, _
           vbCritical + vbOKOnly, "APS PURAN - Engine"
    Resume Sair
End Sub
