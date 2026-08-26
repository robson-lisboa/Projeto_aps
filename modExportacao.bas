Attribute VB_Name = "modExportacao"
'================================================================================
' MÓDULO: modExportacao
' DESCRIÇÃO: Exportação e impressão de relatórios de produção
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' SUBROTINA: ExportarRelatorioExcel
' PROPÓSITO: Exportar relatório para arquivo Excel
'--------------------------------------------------------------------------------
Public Sub ExportarRelatorioExcel(pRelatorio As clsRelatorioProducao, _
                                  pPeriodoInicio As Date, _
                                  pPeriodoFim As Date, _
                                  pPostoFiltro As String, _
                                  pProdutoFiltro As String, _
                                  pStatusFiltro As String)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim linha As Long
    Dim dadosOPs As Variant
    Dim nomeArquivo As String
    
    On Error GoTo ErroExportarExcel
    
    ' Cria nova pasta de trabalho
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Relatório Produção"
    
    ' Título
    ws.Range("A1").Value = "APS PURAN — RELATÓRIO DE PRODUÇÃO"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Resize(1, 8).Merge
    ws.Range("A1").RowHeight = 30
    
    ' Informações do relatório
    linha = 3
    ws.Cells(linha, 1).Value = "Data de geração:"
    ws.Cells(linha, 2).Value = Format(Now, "dd/mm/yyyy HH:MM")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 4
    ws.Cells(linha, 1).Value = "Período:"
    ws.Cells(linha, 2).Value = Format(pPeriodoInicio, "dd/mm/yyyy") & " a " & Format(pPeriodoFim, "dd/mm/yyyy")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 5
    ws.Cells(linha, 1).Value = "Posto:"
    ws.Cells(linha, 2).Value = IIf(pPostoFiltro = "", "Todos", pPostoFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 6
    ws.Cells(linha, 1).Value = "Produto:"
    ws.Cells(linha, 2).Value = IIf(pProdutoFiltro = "", "Todos", pProdutoFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 7
    ws.Cells(linha, 1).Value = "Status:"
    ws.Cells(linha, 2).Value = IIf(pStatusFiltro = "", "Todos", pStatusFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Indicadores
    linha = 9
    ws.Cells(linha, 1).Value = "INDICADORES"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 8).Merge
    
    linha = 11
    ws.Cells(linha, 1).Value = "Total de OPs:"
    ws.Cells(linha, 2).Value = pRelatorio.TotalOPs
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 12
    ws.Cells(linha, 1).Value = "Quantidade Planejada:"
    ws.Cells(linha, 2).Value = pRelatorio.QuantidadePlanejada
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 13
    ws.Cells(linha, 1).Value = "Horas Programadas:"
    ws.Cells(linha, 2).Value = pRelatorio.HorasProgramadas
    ws.Cells(linha, 2).NumberFormat = "0.00"
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 14
    ws.Cells(linha, 1).Value = "OPs Planejadas / Em Andamento / Concluídas / Atrasadas:"
    ws.Cells(linha, 2).Value = pRelatorio.OPsPlanejadas & " / " & pRelatorio.OPsEmAndamento & " / " & pRelatorio.OPsConcluidas & " / " & pRelatorio.OPsAtrasadas
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 15
    ws.Cells(linha, 1).Value = "Postos Utilizados:"
    ws.Cells(linha, 2).Value = pRelatorio.QuantidadePostos
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 16
    ws.Cells(linha, 1).Value = "Percentual de Ocupação:"
    ws.Cells(linha, 2).Value = pRelatorio.PercentualOcupacao
    ws.Cells(linha, 2).NumberFormat = "0.0%"
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Tabela detalhada de OPs
    linha = 18
    ws.Cells(linha, 1).Value = "DETALHAMENTO DAS OPs"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 8).Merge
    
    linha = 20
    ws.Cells(linha, 1).Value = "ID_OP"
    ws.Cells(linha, 2).Value = "Produto"
    ws.Cells(linha, 3).Value = "Equipamento"
    ws.Cells(linha, 4).Value = "Quantidade"
    ws.Cells(linha, 5).Value = "Data Início"
    ws.Cells(linha, 6).Value = "Data Fim"
    ws.Cells(linha, 7).Value = "Duração (h)"
    ws.Cells(linha, 8).Value = "Status"
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Font.Bold = True
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Interior.Color = RGB(220, 220, 220)
    
    ' Carrega dados detalhados
    dadosOPs = ObterDadosOPsEmArray()
    
    If Not IsError(dadosOPs) Then
        Dim totalLinhas As Long
        totalLinhas = UBound(dadosOPs, 1)
        
        linha = 21
        Dim j As Long
        For i = 2 To totalLinhas
            Dim inicioOP As Date
            Dim fimOP As Date
            
            If IsDate(dadosOPs(i, 5)) Then inicioOP = CDate(dadosOPs(i, 5))
            If IsDate(dadosOPs(i, 6)) Then fimOP = CDate(dadosOPs(i, 6))
            
            ' Verifica sobreposição com período e filtros
            If fimOP < pPeriodoInicio Or inicioOP > pPeriodoFim Then
                GoTo ProximaLinha
            End If
            
            If pPostoFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 3))), pPostoFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinha
            End If
            
            If pProdutoFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 2))), pProdutoFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinha
            End If
            
            If pStatusFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 8))), pStatusFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinha
            End If
            
            ws.Cells(linha, 1).Value = CStr(dadosOPs(i, 1))
            ws.Cells(linha, 2).Value = CStr(dadosOPs(i, 2))
            ws.Cells(linha, 3).Value = CStr(dadosOPs(i, 3))
            ws.Cells(linha, 4).Value = IIf(IsNumeric(dadosOPs(i, 4)), CLng(dadosOPs(i, 4)), 0)
            ws.Cells(linha, 5).Value = IIf(IsDate(dadosOPs(i, 5)), CDate(dadosOPs(i, 5)), "")
            ws.Cells(linha, 6).Value = IIf(IsDate(dadosOPs(i, 6)), CDate(dadosOPs(i, 6)), "")
            ws.Cells(linha, 7).Value = IIf(IsNumeric(dadosOPs(i, 7)), CDbl(dadosOPs(i, 7)), 0)
            ws.Cells(linha, 8).Value = CStr(dadosOPs(i, 8))
            
            If IsDate(dadosOPs(i, 5)) Then ws.Cells(linha, 5).NumberFormat = "dd/mm/yyyy HH:MM"
            If IsDate(dadosOPs(i, 6)) Then ws.Cells(linha, 6).NumberFormat = "dd/mm/yyyy HH:MM"
            ws.Cells(linha, 7).NumberFormat = "0.00"
            
            linha = linha + 1
            
ProximaLinha:
        Next i
    End If
    
    ' Formatação
    ws.Columns("A:H").AutoFit
    ws.Range("A1").Resize(linha - 1, 8).Borders.LineStyle = xlContinuous
    ws.Range("A1").Resize(linha - 1, 8).HorizontalAlignment = xlCenter
    ws.Range("A1").Resize(linha - 1, 8).VerticalAlignment = xlCenter
    
    ' Configura página
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$20"
    
    ' Salva arquivo
    nomeArquivo = GerarNomeArquivo("APS_PURAN_Relatorio", "xlsx")
    
    If nomeArquivo <> "" Then
        wb.SaveAs Filename:=nomeArquivo, FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
        MsgBox "Relatório exportado com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    Else
        wb.Close SaveChanges:=False
    End If
    
Sair:
    Exit Sub
    
ErroExportarExcel:
    MsgBox "Erro ao exportar Excel: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExportarRelatorioPDF
' PROPÓSITO: Exportar relatório para PDF
'--------------------------------------------------------------------------------
Public Sub ExportarRelatorioPDF(pRelatorio As clsRelatorioProducao, _
                                pPeriodoInicio As Date, _
                                pPeriodoFim As Date, _
                                pPostoFiltro As String, _
                                pProdutoFiltro As String, _
                                pStatusFiltro As String)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim linha As Long
    Dim nomeArquivo As String
    
    On Error GoTo ErroExportarPDF
    
    ' Cria pasta de trabalho temporária
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Relatório Produção"
    
    ' Título
    ws.Range("A1").Value = "APS PURAN — RELATÓRIO DE PRODUÇÃO"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Resize(1, 8).Merge
    ws.Range("A1").RowHeight = 30
    
    ' Informações
    linha = 3
    ws.Cells(linha, 1).Value = "Data de geração:"
    ws.Cells(linha, 2).Value = Format(Now, "dd/mm/yyyy HH:MM")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 4
    ws.Cells(linha, 1).Value = "Período:"
    ws.Cells(linha, 2).Value = Format(pPeriodoInicio, "dd/mm/yyyy") & " a " & Format(pPeriodoFim, "dd/mm/yyyy")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 5
    ws.Cells(linha, 1).Value = "Posto:"
    ws.Cells(linha, 2).Value = IIf(pPostoFiltro = "", "Todos", pPostoFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 6
    ws.Cells(linha, 1).Value = "Produto:"
    ws.Cells(linha, 2).Value = IIf(pProdutoFiltro = "", "Todos", pProdutoFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 7
    ws.Cells(linha, 1).Value = "Status:"
    ws.Cells(linha, 2).Value = IIf(pStatusFiltro = "", "Todos", pStatusFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Indicadores
    linha = 9
    ws.Cells(linha, 1).Value = "INDICADORES"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 8).Merge
    
    linha = 11
    ws.Cells(linha, 1).Value = "Total de OPs:"
    ws.Cells(linha, 2).Value = pRelatorio.TotalOPs
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 12
    ws.Cells(linha, 1).Value = "Quantidade Planejada:"
    ws.Cells(linha, 2).Value = pRelatorio.QuantidadePlanejada
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 13
    ws.Cells(linha, 1).Value = "Horas Programadas:"
    ws.Cells(linha, 2).Value = pRelatorio.HorasProgramadas
    ws.Cells(linha, 2).NumberFormat = "0.00"
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 14
    ws.Cells(linha, 1).Value = "OPs Planejadas / Em Andamento / Concluídas / Atrasadas:"
    ws.Cells(linha, 2).Value = pRelatorio.OPsPlanejadas & " / " & pRelatorio.OPsEmAndamento & " / " & pRelatorio.OPsConcluidas & " / " & pRelatorio.OPsAtrasadas
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 15
    ws.Cells(linha, 1).Value = "Postos Utilizados:"
    ws.Cells(linha, 2).Value = pRelatorio.QuantidadePostos
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 16
    ws.Cells(linha, 1).Value = "Percentual de Ocupação:"
    ws.Cells(linha, 2).Value = pRelatorio.PercentualOcupacao
    ws.Cells(linha, 2).NumberFormat = "0.0%"
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Tabela detalhada
    linha = 18
    ws.Cells(linha, 1).Value = "DETALHAMENTO DAS OPs"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 8).Merge
    
    linha = 20
    ws.Cells(linha, 1).Value = "ID_OP"
    ws.Cells(linha, 2).Value = "Produto"
    ws.Cells(linha, 3).Value = "Equipamento"
    ws.Cells(linha, 4).Value = "Quantidade"
    ws.Cells(linha, 5).Value = "Data Início"
    ws.Cells(linha, 6).Value = "Data Fim"
    ws.Cells(linha, 7).Value = "Duração (h)"
    ws.Cells(linha, 8).Value = "Status"
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Font.Bold = True
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Interior.Color = RGB(220, 220, 220)
    
    ' Preenche dados
    Dim dadosOPs As Variant
    dadosOPs = ObterDadosOPsEmArray()
    
    If Not IsError(dadosOPs) Then
        Dim totalLinhas As Long
        totalLinhas = UBound(dadosOPs, 1)
        
        linha = 21
        Dim i As Long
        For i = 2 To totalLinhas
            Dim inicioOP As Date
            Dim fimOP As Date
            
            If IsDate(dadosOPs(i, 5)) Then inicioOP = CDate(dadosOPs(i, 5))
            If IsDate(dadosOPs(i, 6)) Then fimOP = CDate(dadosOPs(i, 6))
            
            If fimOP < pPeriodoInicio Or inicioOP > pPeriodoFim Then
                GoTo ProximaLinhaPDF
            End If
            
            If pPostoFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 3))), pPostoFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinhaPDF
            End If
            
            If pProdutoFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 2))), pProdutoFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinhaPDF
            End If
            
            If pStatusFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 8))), pStatusFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinhaPDF
            End If
            
            ws.Cells(linha, 1).Value = CStr(dadosOPs(i, 1))
            ws.Cells(linha, 2).Value = CStr(dadosOPs(i, 2))
            ws.Cells(linha, 3).Value = CStr(dadosOPs(i, 3))
            ws.Cells(linha, 4).Value = IIf(IsNumeric(dadosOPs(i, 4)), CLng(dadosOPs(i, 4)), 0)
            ws.Cells(linha, 5).Value = IIf(IsDate(dadosOPs(i, 5)), CDate(dadosOPs(i, 5)), "")
            ws.Cells(linha, 6).Value = IIf(IsDate(dadosOPs(i, 6)), CDate(dadosOPs(i, 6)), "")
            ws.Cells(linha, 7).Value = IIf(IsNumeric(dadosOPs(i, 7)), CDbl(dadosOPs(i, 7)), 0)
            ws.Cells(linha, 8).Value = CStr(dadosOPs(i, 8))
            
            If IsDate(dadosOPs(i, 5)) Then ws.Cells(linha, 5).NumberFormat = "dd/mm/yyyy HH:MM"
            If IsDate(dadosOPs(i, 6)) Then ws.Cells(linha, 6).NumberFormat = "dd/mm/yyyy HH:MM"
            ws.Cells(linha, 7).NumberFormat = "0.00"
            
            linha = linha + 1
            
ProximaLinhaPDF:
        Next i
    End If
    
    ' Formatação
    ws.Columns("A:H").AutoFit
    ws.Range("A1").Resize(linha - 1, 8).Borders.LineStyle = xlContinuous
    
    ' Configura página para PDF
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$20"
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    ' Exporta PDF
    nomeArquivo = GerarNomeArquivo("APS_PURAN_Relatorio", "pdf")
    
    If nomeArquivo <> "" Then
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nomeArquivo, _
                               Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False, OpenAfterPublish:=True
        wb.Close SaveChanges:=False
        MsgBox "Relatório PDF gerado com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    Else
        wb.Close SaveChanges:=False
    End If
    
Sair:
    Exit Sub
    
ErroExportarPDF:
    MsgBox "Erro ao gerar PDF: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ImprimirRelatorio
' PROPÓSITO: Imprimir relatório atual
'--------------------------------------------------------------------------------
Public Sub ImprimirRelatorio(pRelatorio As clsRelatorioProducao, _
                             pPeriodoInicio As Date, _
                             pPeriodoFim As Date, _
                             pPostoFiltro As String, _
                             pProdutoFiltro As String, _
                             pStatusFiltro As String)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim linha As Long
    Dim dadosOPs As Variant
    
    On Error GoTo ErroImprimir
    
    ' Cria pasta de trabalho temporária
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Relatório Produção"
    
    ' Título
    ws.Range("A1").Value = "APS PURAN — RELATÓRIO DE PRODUÇÃO"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Resize(1, 8).Merge
    ws.Range("A1").RowHeight = 30
    
    ' Informações
    linha = 3
    ws.Cells(linha, 1).Value = "Data de geração:"
    ws.Cells(linha, 2).Value = Format(Now, "dd/mm/yyyy HH:MM")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 4
    ws.Cells(linha, 1).Value = "Período:"
    ws.Cells(linha, 2).Value = Format(pPeriodoInicio, "dd/mm/yyyy") & " a " & Format(pPeriodoFim, "dd/mm/yyyy")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 5
    ws.Cells(linha, 1).Value = "Posto:"
    ws.Cells(linha, 2).Value = IIf(pPostoFiltro = "", "Todos", pPostoFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 6
    ws.Cells(linha, 1).Value = "Produto:"
    ws.Cells(linha, 2).Value = IIf(pProdutoFiltro = "", "Todos", pProdutoFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 7
    ws.Cells(linha, 1).Value = "Status:"
    ws.Cells(linha, 2).Value = IIf(pStatusFiltro = "", "Todos", pStatusFiltro)
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Indicadores
    linha = 9
    ws.Cells(linha, 1).Value = "INDICADORES"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 8).Merge
    
    linha = 11
    ws.Cells(linha, 1).Value = "Total de OPs:"
    ws.Cells(linha, 2).Value = pRelatorio.TotalOPs
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 12
    ws.Cells(linha, 1).Value = "Quantidade Planejada:"
    ws.Cells(linha, 2).Value = pRelatorio.QuantidadePlanejada
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 13
    ws.Cells(linha, 1).Value = "Horas Programadas:"
    ws.Cells(linha, 2).Value = pRelatorio.HorasProgramadas
    ws.Cells(linha, 2).NumberFormat = "0.00"
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 14
    ws.Cells(linha, 1).Value = "OPs Planejadas / Em Andamento / Concluídas / Atrasadas:"
    ws.Cells(linha, 2).Value = pRelatorio.OPsPlanejadas & " / " & pRelatorio.OPsEmAndamento & " / " & pRelatorio.OPsConcluidas & " / " & pRelatorio.OPsAtrasadas
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 15
    ws.Cells(linha, 1).Value = "Postos Utilizados:"
    ws.Cells(linha, 2).Value = pRelatorio.QuantidadePostos
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 16
    ws.Cells(linha, 1).Value = "Percentual de Ocupação:"
    ws.Cells(linha, 2).Value = pRelatorio.PercentualOcupacao
    ws.Cells(linha, 2).NumberFormat = "0.0%"
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Tabela detalhada
    linha = 18
    ws.Cells(linha, 1).Value = "DETALHAMENTO DAS OPs"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 8).Merge
    
    linha = 20
    ws.Cells(linha, 1).Value = "ID_OP"
    ws.Cells(linha, 2).Value = "Produto"
    ws.Cells(linha, 3).Value = "Equipamento"
    ws.Cells(linha, 4).Value = "Quantidade"
    ws.Cells(linha, 5).Value = "Data Início"
    ws.Cells(linha, 6).Value = "Data Fim"
    ws.Cells(linha, 7).Value = "Duração (h)"
    ws.Cells(linha, 8).Value = "Status"
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Font.Bold = True
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Interior.Color = RGB(220, 220, 220)
    
    ' Preenche dados
    dadosOPs = ObterDadosOPsEmArray()
    
    If Not IsError(dadosOPs) Then
        Dim totalLinhas As Long
        totalLinhas = UBound(dadosOPs, 1)
        
        linha = 21
        Dim i As Long
        For i = 2 To totalLinhas
            Dim inicioOP As Date
            Dim fimOP As Date
            
            If IsDate(dadosOPs(i, 5)) Then inicioOP = CDate(dadosOPs(i, 5))
            If IsDate(dadosOPs(i, 6)) Then fimOP = CDate(dadosOPs(i, 6))
            
            If fimOP < pPeriodoInicio Or inicioOP > pPeriodoFim Then
                GoTo ProximaLinhaPrint
            End If
            
            If pPostoFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 3))), pPostoFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinhaPrint
            End If
            
            If pProdutoFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 2))), pProdutoFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinhaPrint
            End If
            
            If pStatusFiltro <> "" And StrComp(Trim(CStr(dadosOPs(i, 8))), pStatusFiltro, vbTextCompare) <> 0 Then
                GoTo ProximaLinhaPrint
            End If
            
            ws.Cells(linha, 1).Value = CStr(dadosOPs(i, 1))
            ws.Cells(linha, 2).Value = CStr(dadosOPs(i, 2))
            ws.Cells(linha, 3).Value = CStr(dadosOPs(i, 3))
            ws.Cells(linha, 4).Value = IIf(IsNumeric(dadosOPs(i, 4)), CLng(dadosOPs(i, 4)), 0)
            ws.Cells(linha, 5).Value = IIf(IsDate(dadosOPs(i, 5)), CDate(dadosOPs(i, 5)), "")
            ws.Cells(linha, 6).Value = IIf(IsDate(dadosOPs(i, 6)), CDate(dadosOPs(i, 6)), "")
            ws.Cells(linha, 7).Value = IIf(IsNumeric(dadosOPs(i, 7)), CDbl(dadosOPs(i, 7)), 0)
            ws.Cells(linha, 8).Value = CStr(dadosOPs(i, 8))
            
            If IsDate(dadosOPs(i, 5)) Then ws.Cells(linha, 5).NumberFormat = "dd/mm/yyyy HH:MM"
            If IsDate(dadosOPs(i, 6)) Then ws.Cells(linha, 6).NumberFormat = "dd/mm/yyyy HH:MM"
            ws.Cells(linha, 7).NumberFormat = "0.00"
            
            linha = linha + 1
            
ProximaLinhaPrint:
        Next i
    End If
    
    ' Formatação
    ws.Columns("A:H").AutoFit
    ws.Range("A1").Resize(linha - 1, 8).Borders.LineStyle = xlContinuous
    
    ' Configura página para impressão
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$20"
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    ' Imprime
    ws.PrintOut Copies:=1, Collate:=True
    
    wb.Close SaveChanges:=False
    MsgBox "Relatório enviado para impressora.", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroImprimir:
    MsgBox "Erro ao imprimir: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: GerarNomeArquivo
' PROPÓSITO: Gerar nome de arquivo seguro com data
' PARÂMETROS: pPrefixo, pExtensao
' RETORNO: String com caminho completo ou vazio se cancelado
'--------------------------------------------------------------------------------
Private Function GerarNomeArquivo(pPrefixo As String, pExtensao As String) As String
    Dim nomeBase As String
    Dim caminho As String
    Dim nomeArquivo As String
    Dim i As Long
    
    nomeBase = pPrefixo & "_" & Format(Now, "yyyy-mm-dd")
    caminho = ThisWorkbook.Path & "\"
    nomeArquivo = caminho & nomeBase & "." & pExtensao
    
    ' Verifica se arquivo já existe
    i = 1
    Do While Dir(nomeArquivo) <> ""
        nomeArquivo = caminho & nomeBase & "_" & Format(i, "00") & "." & pExtensao
        i = i + 1
    Loop
    
    ' Solicita confirmação do caminho
    On Error Resume Next
    Dim dialogo As FileDialog
    Set dialogo = Application.FileDialog(msoFileDialogSaveAs)
    
    If dialogo Is Nothing Then
        GerarNomeArquivo = nomeArquivo
        Exit Function
    End If
    
    With dialogo
        .InitialFileName = nomeArquivo
        .Title = "Salvar " & UCase(pExtensao)
        .FilterIndex = 1
    End With
    
    If dialogo.Show = -1 Then
        GerarNomeArquivo = dialogo.SelectedItems(1)
    Else
        GerarNomeArquivo = ""
    End If
    
    On Error GoTo 0
End Function
