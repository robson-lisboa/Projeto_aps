Attribute VB_Name = "modConfig"
'================================================================================
' MÓDULO: modConfig
' DESCRIÇÃO: Persistência de preferências do usuário via BD_Config/TabelaConfig
' VERSÃO: 1.0
'================================================================================
Option Explicit

'--------------------------------------------------------------------------------
' CONSTANTES
'--------------------------------------------------------------------------------
Private Const PLANILHA_CONFIG As String = "BD_Config"
Private Const TABELA_CONFIG As String = "TabelaConfig"

'--------------------------------------------------------------------------------
' FUNÇÃO: ObterConfiguracao
' PROPÓSITO: Ler um valor de configuração pelo nome da chave
' PARÂMETROS: pChave As String
' RETORNO: String com o valor ou "" se não encontrado
'--------------------------------------------------------------------------------
Public Function ObterConfiguracao(pChave As String) As String
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    
    On Error GoTo ErroObter
    
    dados = ObterDadosTabelaEmArray(PLANILHA_CONFIG, TABELA_CONFIG)
    If IsError(dados) Then
        ObterConfiguracao = ""
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        If Trim(CStr(dados(i, 1))) = Trim(pChave) Then
            ObterConfiguracao = CStr(dados(i, 2))
            Exit Function
        End If
    Next i
    
    ObterConfiguracao = ""
    
Sair:
    Exit Function
    
ErroObter:
    ObterConfiguracao = ""
    Resume Sair
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: SalvarConfiguracao
' PROPÓSITO: Salvar ou atualizar um valor de configuração (upsert)
' PARÂMETROS: pChave As String, pValor As String
'--------------------------------------------------------------------------------
Public Sub SalvarConfiguracao(pChave As String, pValor As String)
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    Dim chaveEncontrada As Boolean
    
    On Error GoTo ErroSalvar
    
    chaveEncontrada = False
    
    dados = ObterDadosTabelaEmArray(PLANILHA_CONFIG, TABELA_CONFIG)
    If Not IsError(dados) Then
        totalLinhas = UBound(dados, 1)
        
        For i = 2 To totalLinhas
            If Trim(CStr(dados(i, 1))) = Trim(pChave) Then
                chaveEncontrada = True
                Call AtualizarLinhaTabela(PLANILHA_CONFIG, TABELA_CONFIG, "Chave", pChave, Array(pChave, pValor))
                Exit For
            End If
        Next i
    End If
    
    If Not chaveEncontrada Then
        Call InserirLinhaTabela(PLANILHA_CONFIG, TABELA_CONFIG, Array(pChave, pValor))
    End If
    
Sair:
    Exit Sub
    
ErroSalvar:
    Resume Sair
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ExcluirConfiguracao
' PROPÓSITO: Remover uma chave de configuração
' PARÂMETROS: pChave As String
'--------------------------------------------------------------------------------
Public Sub ExcluirConfiguracao(pChave As String)
    On Error Resume Next
    Call ExcluirLinhaTabela(PLANILHA_CONFIG, TABELA_CONFIG, "Chave", pChave)
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: CarregarTodasConfiguracoes
' PROPÓSITO: Carregar todas as configurações em um Dictionary para acesso rápido
' RETORNO: Object (Scripting.Dictionary) com chave=>valor
'--------------------------------------------------------------------------------
Public Function CarregarTodasConfiguracoes() As Object
    Dim dados As Variant
    Dim i As Long
    Dim totalLinhas As Long
    Dim mapa As Object
    
    On Error GoTo ErroCarregar
    
    Set mapa = CreateObject("Scripting.Dictionary")
    
    dados = ObterDadosTabelaEmArray(PLANILHA_CONFIG, TABELA_CONFIG)
    If IsError(dados) Then
        Set CarregarTodasConfiguracoes = mapa
        Exit Function
    End If
    
    totalLinhas = UBound(dados, 1)
    
    For i = 2 To totalLinhas
        Dim chave As String
        chave = Trim(CStr(dados(i, 1)))
        If chave <> "" Then
            mapa(chave) = CStr(dados(i, 2))
        End If
    Next i
    
    Set CarregarTodasConfiguracoes = mapa
    
Sair:
    Exit Function
    
ErroCarregar:
    Set CarregarTodasConfiguracoes = CreateObject("Scripting.Dictionary")
    Resume Sair
End Function
