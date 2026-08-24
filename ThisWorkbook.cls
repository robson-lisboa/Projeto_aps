VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ThisWorkbook"
'================================================================================
' OBJETO: ThisWorkbook
' DESCRIÇÃO: Configuração de inicialização do aplicativo APS PURAN
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' EVENTO: Workbook_Open
' PROPÓSITO: Rotina de startup corporativa do APS PURAN
'--------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error GoTo ErroInicializacao
    
    Dim principalForm As frmPrincipal
    
    ' Performance: desativa atualização de tela durante inicialização
    Application.ScreenUpdating = False
    
    ' Interface: oculta elementos visuais desnecessários e maximiza a janela
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    Application.WindowState = xlMaximized
    
    ' Integridade: garante que o banco de dados oculto esteja íntegro
    modDatabase.CriarEstruturasDeDados
    
    ' Apresentação: exibe o formulário principal
    Set principalForm = New frmPrincipal
    principalForm.Show vbModeless
    
Sair:
    ' Segurança: garante que a atualização de tela seja reativada
    Application.ScreenUpdating = True
    Exit Sub
    
ErroInicializacao:
    MsgBox "Erro crítico na inicialização do APS PURAN." & vbCrLf & _
           "Descrição: " & Err.Description & vbCrLf & _
           "O aplicativo será encerrado.", _
           vbCritical + vbOKOnly, "APS PURAN - Inicialização"
    Application.ScreenUpdating = True
    Application.Quit
End Sub
