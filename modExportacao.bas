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

'--------------------------------------------------------------------------------
' FUNÇÃO: EscreverDadosPlanejamento
' PROPÓSITO: Escrever dados do planejamento em uma planilha
' PARÂMETROS: ws, periodoInicio, periodoFim, postoFiltro, produtoFiltro, statusFiltro
' RETORNO: Próxima linha disponível
'--------------------------------------------------------------------------------
Private Function EscreverDadosPlanejamento(ws As Worksheet, _
                                           pPeriodoInicio As Date, _
                                           pPeriodoFim As Date, _
                                           pFiltroStatus As String, _
                                           pTextoBusca As String, _
                                           pOrdenarPor As String) As Long
    Dim linha As Long
    Dim dadosOPs As Variant
    Dim ops As Collection
    Dim i As Long
    Dim op As clsCardProducao
    
    On Error GoTo ErroEscrever
    
    ' Título
    ws.Range("A1").Value = "APS PURAN — PLANEJAMENTO DE PRODUÇÃO"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Resize(1, 9).Merge
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
    ws.Cells(linha, 1).Value = "Filtro de Status:"
    ws.Cells(linha, 2).Value = IIf(pFiltroStatus = "", "Todos", pFiltroStatus)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 6
    ws.Cells(linha, 1).Value = "Busca:"
    ws.Cells(linha, 2).Value = IIf(pTextoBusca = "", "(nenhuma)", pTextoBusca)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 7
    ws.Cells(linha, 1).Value = "Ordenação:"
    ws.Cells(linha, 2).Value = IIf(pOrdenarPor = "", "(padrão)", pOrdenarPor)
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Cabeçalho da tabela
    linha = 9
    ws.Cells(linha, 1).Value = "DETALHAMENTO DAS OPs"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 9).Merge
    
    linha = 11
    ws.Cells(linha, 1).Value = "ID_OP"
    ws.Cells(linha, 2).Value = "Produto"
    ws.Cells(linha, 3).Value = "Equipamento"
    ws.Cells(linha, 4).Value = "Quantidade"
    ws.Cells(linha, 5).Value = "Data Início"
    ws.Cells(linha, 6).Value = "Data Fim"
    ws.Cells(linha, 7).Value = "Duração (h)"
    ws.Cells(linha, 8).Value = "Status"
    ws.Cells(linha, 9).Value = "Conflito"
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 9)).Font.Bold = True
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 9)).Interior.Color = RGB(220, 220, 220)
    
    Set ops = ObterDadosPlanejamentoProcessados(pPeriodoInicio, pPeriodoFim, 1#, pFiltroStatus, pTextoBusca, pOrdenarPor)
    
    If ops.Count <= 0 Then
        linha = linha + 1
        ws.Cells(linha, 1).Value = "Nenhuma OP encontrada para os filtros selecionados."
        EscreverDadosPlanejamento = linha + 1
        Exit Function
    End If
    
    linha = 12
    For i = 1 To ops.Count
        Set op = ops(i)
        
        ws.Cells(linha, 1).Value = op.ID_OP
        ws.Cells(linha, 2).Value = op.Produto
        ws.Cells(linha, 3).Value = op.Equipamento
        ws.Cells(linha, 4).Value = op.Quantidade
        ws.Cells(linha, 5).Value = IIf(IsDate(op.DataInicio), op.DataInicio, "")
        ws.Cells(linha, 6).Value = IIf(IsDate(op.DataFim), op.DataFim, "")
        ws.Cells(linha, 7).Value = op.Duracao
        ws.Cells(linha, 8).Value = op.Status
        ws.Cells(linha, 9).Value = IIf(op.TemConflito, "Sim", "Não")
        
        If IsDate(op.DataInicio) Then ws.Cells(linha, 5).NumberFormat = "dd/mm/yyyy HH:MM"
        If IsDate(op.DataFim) Then ws.Cells(linha, 6).NumberFormat = "dd/mm/yyyy HH:MM"
        ws.Cells(linha, 7).NumberFormat = "0.00"
        
        linha = linha + 1
    Next i
    
    ' Formatação
    ws.Columns("A:I").AutoFit
    ws.Range("A1").Resize(linha - 1, 9).Borders.LineStyle = xlContinuous
    ws.Range("A1").Resize(linha - 1, 9).HorizontalAlignment = xlCenter
    ws.Range("A1").Resize(linha - 1, 9).VerticalAlignment = xlCenter
    
    EscreverDadosPlanejamento = linha
    
Sair:
    Exit Function
    
ErroEscrever:
    Err.Raise Err.Number, "EscreverDadosPlanejamento", Err.Description
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: ExportarPlanejamentoExcel
' PROPÓSITO: Exportar planejamento para Excel
'--------------------------------------------------------------------------------
Public Sub ExportarPlanejamentoExcel(pPeriodoInicio As Date, _
                                     pPeriodoFim As Date, _
                                     pFiltroStatus As String = "Todos", _
                                     pTextoBusca As String = "", _
                                     pOrdenarPor As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim nomeArquivo As String
    
    On Error GoTo ErroExportarExcel
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Planejamento Produção"
    
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$11"
    
    Call EscreverDadosPlanejamento(ws, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca, pOrdenarPor)
    
    nomeArquivo = GerarNomeArquivo("APS_PURAN_Planejamento", "xlsx")
    
    If nomeArquivo <> "" Then
        wb.SaveAs Filename:=nomeArquivo, FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
        MsgBox "Planejamento exportado com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    Else
        wb.Close SaveChanges:=False
    End If
    
Sair:
    Exit Sub
    
ErroExportarExcel:
    MsgBox "Erro ao exportar planejamento: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExportarPlanejamentoPDF
' PROPÓSITO: Exportar planejamento para PDF
'--------------------------------------------------------------------------------
Public Sub ExportarPlanejamentoPDF(pPeriodoInicio As Date, _
                                   pPeriodoFim As Date, _
                                   pFiltroStatus As String = "Todos", _
                                   pTextoBusca As String = "", _
                                   pOrdenarPor As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim nomeArquivo As String
    
    On Error GoTo ErroExportarPDF
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Planejamento Produção"
    
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$11"
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    Call EscreverDadosPlanejamento(ws, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca, pOrdenarPor)
    
    nomeArquivo = GerarNomeArquivo("APS_PURAN_Planejamento", "pdf")
    
    If nomeArquivo <> "" Then
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nomeArquivo, _
                               Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False, OpenAfterPublish:=True
        wb.Close SaveChanges:=False
        MsgBox "Planejamento PDF gerado com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    Else
        wb.Close SaveChanges:=False
    End If
    
Sair:
    Exit Sub
    
ErroExportarPDF:
    MsgBox "Erro ao gerar PDF do planejamento: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ImprimirPlanejamento
' PROPÓSITO: Imprimir planejamento atual
'--------------------------------------------------------------------------------
Public Sub ImprimirPlanejamento(pPeriodoInicio As Date, _
                                pPeriodoFim As Date, _
                                pFiltroStatus As String = "Todos", _
                                pTextoBusca As String = "", _
                                pOrdenarPor As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    
    On Error GoTo ErroImprimir
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Planejamento Produção"
    
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$11"
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    Call EscreverDadosPlanejamento(ws, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca, pOrdenarPor)
    
    ws.PrintOut Copies:=1, Collate:=True
    wb.Close SaveChanges:=False
    MsgBox "Planejamento enviado para impressora.", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroImprimir:
    MsgBox "Erro ao imprimir planejamento: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: EscreverDadosSimulacao
' PROPÓSITO: Escrever dados da simulação em uma planilha
' PARÂMETROS: ws, pID_Simulacao, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca
' RETORNO: Próxima linha disponível
'--------------------------------------------------------------------------------
Private Function EscreverDadosSimulacao(ws As Worksheet, _
                                        pID_Simulacao As String, _
                                        pPeriodoInicio As Date, _
                                        pPeriodoFim As Date, _
                                        pFiltroStatus As String = "Todos", _
                                        pTextoBusca As String = "") As Long
    Dim linha As Long
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    Dim nomeSimulacao As String
    
    On Error GoTo ErroEscrever
    
    ' Título
    ws.Range("A1").Value = "APS PURAN — SIMULAÇÃO DE PRODUÇÃO"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Resize(1, 9).Merge
    ws.Range("A1").RowHeight = 30
    
    ' Informações
    linha = 3
    ws.Cells(linha, 1).Value = "Data de geração:"
    ws.Cells(linha, 2).Value = Format(Now, "dd/mm/yyyy HH:MM")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 4
    ws.Cells(linha, 1).Value = "Simulação:"
    ws.Cells(linha, 2).Value = pID_Simulacao
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 5
    ws.Cells(linha, 1).Value = "Período:"
    ws.Cells(linha, 2).Value = Format(pPeriodoInicio, "dd/mm/yyyy") & " a " & Format(pPeriodoFim, "dd/mm/yyyy")
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 6
    ws.Cells(linha, 1).Value = "Filtro de Status:"
    ws.Cells(linha, 2).Value = IIf(pFiltroStatus = "", "Todos", pFiltroStatus)
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = 7
    ws.Cells(linha, 1).Value = "Busca:"
    ws.Cells(linha, 2).Value = IIf(pTextoBusca = "", "(nenhuma)", pTextoBusca)
    ws.Cells(linha, 1).Font.Bold = True
    
    ' Cabeçalho da tabela
    linha = 9
    ws.Cells(linha, 1).Value = "DETALHAMENTO DAS OPs DA SIMULAÇÃO"
    ws.Cells(linha, 1).Font.Size = 12
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 1).Resize(1, 9).Merge
    
    linha = 11
    ws.Cells(linha, 1).Value = "ID_OP"
    ws.Cells(linha, 2).Value = "Produto"
    ws.Cells(linha, 3).Value = "Equipamento"
    ws.Cells(linha, 4).Value = "Quantidade"
    ws.Cells(linha, 5).Value = "Data Início"
    ws.Cells(linha, 6).Value = "Data Fim"
    ws.Cells(linha, 7).Value = "Duração (h)"
    ws.Cells(linha, 8).Value = "Status"
    ws.Cells(linha, 9).Value = "Sequência"
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 9)).Font.Bold = True
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 9)).Interior.Color = RGB(220, 220, 220)
    
    dados = ObterOPsSimulacaoEmArray(pID_Simulacao)
    
    If IsError(dados) Then
        linha = linha + 1
        ws.Cells(linha, 1).Value = "Nenhuma OP encontrada nesta simulação."
        EscreverDadosSimulacao = linha + 1
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    Dim textoBuscaLower As String
    
    If pTextoBusca <> "" Then
        textoBuscaLower = LCase(Trim(pTextoBusca))
    End If
    
    linha = 12
    Dim linhasEscritas As Long
    linhasEscritas = 0
    
    For i = 1 To totalLinhas
        Dim idOP As String
        Dim produto As String
        Dim equipamento As String
        Dim quantidade As Long
        Dim dataInicio As Date
        Dim dataFim As Date
        Dim duracao As Double
        Dim status As String
        Dim sequencia As Long
        
        idOP = CStr(dados(i, 2))
        produto = CStr(dados(i, 4))
        equipamento = CStr(dados(i, 5))
        quantidade = CLng(dados(i, 6))
        If IsDate(dados(i, 7)) Then dataInicio = CDate(dados(i, 7))
        If IsDate(dados(i, 8)) Then dataFim = CDate(dados(i, 8))
        duracao = CDbl(dados(i, 9))
        status = CStr(dados(i, 10))
        sequencia = CLng(dados(i, 3))
        
        ' Aplica filtro de status
        If pFiltroStatus <> "Todos" Then
            If Trim(status) <> Trim(pFiltroStatus) Then
                GoTo ProximaLinhaSim
            End If
        End If
        
        ' Aplica filtro de busca
        If pTextoBusca <> "" Then
            If LCase(Trim(idOP)) <> textoBuscaLower And _
               LCase(Trim(produto)) <> textoBuscaLower Then
                GoTo ProximaLinhaSim
            End If
        End If
        
        ws.Cells(linha, 1).Value = idOP
        ws.Cells(linha, 2).Value = produto
        ws.Cells(linha, 3).Value = equipamento
        ws.Cells(linha, 4).Value = quantidade
        ws.Cells(linha, 5).Value = IIf(IsDate(dataInicio), dataInicio, "")
        ws.Cells(linha, 6).Value = IIf(IsDate(dataFim), dataFim, "")
        ws.Cells(linha, 7).Value = duracao
        ws.Cells(linha, 8).Value = status
        ws.Cells(linha, 9).Value = sequencia
        
        If IsDate(dataInicio) Then ws.Cells(linha, 5).NumberFormat = "dd/mm/yyyy HH:MM"
        If IsDate(dataFim) Then ws.Cells(linha, 6).NumberFormat = "dd/mm/yyyy HH:MM"
        ws.Cells(linha, 7).NumberFormat = "0.00"
        
        linha = linha + 1
        linhasEscritas = linhasEscritas + 1
        
ProximaLinhaSim:
    Next i
    
    If linhasEscritas = 0 Then
        linha = linha + 1
        ws.Cells(linha, 1).Value = "Nenhuma OP encontrada para os filtros selecionados."
    End If
    
    ' Formatação
    ws.Columns("A:I").AutoFit
    ws.Range("A1").Resize(linha - 1, 9).Borders.LineStyle = xlContinuous
    ws.Range("A1").Resize(linha - 1, 9).HorizontalAlignment = xlCenter
    ws.Range("A1").Resize(linha - 1, 9).VerticalAlignment = xlCenter
    
    EscreverDadosSimulacao = linha
    
Sair:
    Exit Function
    
ErroEscrever:
    Err.Raise Err.Number, "EscreverDadosSimulacao", Err.Description
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: ExportarSimulacaoExcel
' PROPÓSITO: Exportar simulação para Excel
'--------------------------------------------------------------------------------
Public Sub ExportarSimulacaoExcel(pID_Simulacao As String, _
                                  pPeriodoInicio As Date, _
                                  pPeriodoFim As Date, _
                                  pFiltroStatus As String = "Todos", _
                                  pTextoBusca As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim nomeArquivo As String
    
    On Error GoTo ErroExportarExcel
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Simulação Produção"
    
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$11"
    
    Call EscreverDadosSimulacao(ws, pID_Simulacao, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca)
    
    nomeArquivo = GerarNomeArquivo("APS_PURAN_Simulacao_" & pID_Simulacao, "xlsx")
    
    If nomeArquivo <> "" Then
        wb.SaveAs Filename:=nomeArquivo, FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
        MsgBox "Simulação exportada com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    Else
        wb.Close SaveChanges:=False
    End If
    
Sair:
    Exit Sub
    
ErroExportarExcel:
    MsgBox "Erro ao exportar simulação: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExportarSimulacaoPDF
' PROPÓSITO: Exportar simulação para PDF
'--------------------------------------------------------------------------------
Public Sub ExportarSimulacaoPDF(pID_Simulacao As String, _
                                pPeriodoInicio As Date, _
                                pPeriodoFim As Date, _
                                pFiltroStatus As String = "Todos", _
                                pTextoBusca As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim nomeArquivo As String
    
    On Error GoTo ErroExportarPDF
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Simulação Produção"
    
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$11"
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    Call EscreverDadosSimulacao(ws, pID_Simulacao, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca)
    
    nomeArquivo = GerarNomeArquivo("APS_PURAN_Simulacao_" & pID_Simulacao, "pdf")
    
    If nomeArquivo <> "" Then
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nomeArquivo, _
                               Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False, OpenAfterPublish:=True
        wb.Close SaveChanges:=False
        MsgBox "Simulação PDF gerada com sucesso!" & vbCrLf & nomeArquivo, vbInformation, "APS PURAN"
    Else
        wb.Close SaveChanges:=False
    End If
    
Sair:
    Exit Sub
    
ErroExportarPDF:
    MsgBox "Erro ao gerar PDF da simulação: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ImprimirSimulacao
' PROPÓSITO: Imprimir simulação atual
'--------------------------------------------------------------------------------
Public Sub ImprimirSimulacao(pID_Simulacao As String, _
                             pPeriodoInicio As Date, _
                             pPeriodoFim As Date, _
                             pFiltroStatus As String = "Todos", _
                             pTextoBusca As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    
    On Error GoTo ErroImprimir
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Simulação Produção"
    
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.PaperSize = xlPaperA4
    ws.PageSetup.Margins.Left = 0.5
    ws.PageSetup.Margins.Right = 0.5
    ws.PageSetup.Margins.Top = 0.75
    ws.PageSetup.Margins.Bottom = 0.75
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = False
    ws.PageSetup.PrintTitleRows = "$1:$11"
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    Call EscreverDadosSimulacao(ws, pID_Simulacao, pPeriodoInicio, pPeriodoFim, pFiltroStatus, pTextoBusca)
    
    ws.PrintOut Copies:=1, Collate:=True
    wb.Close SaveChanges:=False
    MsgBox "Simulação enviada para impressora.", vbInformation, "APS PURAN"
    
Sair:
    Exit Sub
    
ErroImprimir:
    MsgBox "Erro ao imprimir simulação: " & Err.Description, vbCritical, "APS PURAN"
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Resume Sair
End Sub
