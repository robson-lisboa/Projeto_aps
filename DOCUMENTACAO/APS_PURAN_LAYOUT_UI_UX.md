# APS PURAN - Layout UI/UX e Arquitetura Visual

## 1. Objetivo e Arquitetura Visual

O APS PURAN é um sistema de planejamento de produção industrial desenvolvido em Excel VBA. A interface segue o padrão de layout **sidebar + content area**, comum em sistemas industriais modernos, garantindo usabilidade, organização clara das informações e responsividade.

### Princípios de Design

- **Profissionalismo industrial**: cores sóbrias, alto contraste, tipografia legível
- **Responsividade real**: o layout se adapta ao redimensionamento manual da janela
- **Clareza hierárquica**: menu lateral fixo, área de conteúdo dinâmica
- **Performance**: controles criados dinamicamente apenas quando necessário
- **Compatibilidade**: funciona em Office 32-bit e 64-bit

### Estrutura Visual

```
┌─────────────────────────────────────────────────────────────┐
│ [─] [□] [×]  APS PURAN – Sistema de Planejamento...        │
├──────┬──────────────────────────────────────────────────────┤
│      │  INDICADORES DE PRODUÇÃO – DASHBOARD                  │
│      ├──────────────────────────────────────────────────────┤
│ MENU │  ┌─────────────────────┐  ┌──────────────────┐      │
│      │  │ Total de Produções  │  │ Produções        │      │
│ Dash │  │   Planejadas       │  │ Concluídas       │      │
│      │  │        45          │  │       12         │      │
│ Time │  └─────────────────────┘  └──────────────────┘      │
│      │                                                       │
│ Cards│  ┌─────────────────────┐  ┌──────────────────┐      │
│      │  │ Produções em        │  │ Produções        │      │
│ Produ│  │   Andamento        │  │ Atrasadas        │      │
│      │  │        23          │  │        3         │      │
│ Event│  └─────────────────────┘  └──────────────────┘      │
│      │                                                       │
│ Conf │  ┌─────────────────────┐  ┌──────────────────┐      │
│      │  │ Total de Horas     │  │ Total de Horas   │      │
│      │  │   Planejadas       │  │ Realizadas       │      │
│      │  │    1,234.50        │  │      987.25      │      │
│      │  └─────────────────────┘  └──────────────────┘      │
└──────┴──────────────────────────────────────────────────────┘
```

---

## 2. frmPrincipal

### Propriedades do UserForm

| Propriedade | Valor | Descrição |
|-------------|-------|-----------|
| `Width` | 1200 | Largura padrão da janela |
| `Height` | 800 | Altura padrão da janela |
| `ClientWidth` | 1188 | Área útil interna (sem bordas) |
| `ClientHeight` | 756 | Área útil interna (sem bordas) |
| `BorderStyle` | 3 | Sizable (redimensionável) |
| `StartUpPosition` | 1 | CenterScreen (centraliza na tela) |
| `ScaleMode` | 3 | Twips (unidade padrão VBA) |
| `BackColor` | 4474559 | Verde escuro corporativo (RGB 68,68,68) |
| `KeyPreview` | -1 | True (captura teclas no formulário) |
| `MinWidth` | 600 | Tamanho mínimo de largura |
| `MinHeight` | 500 | Tamanho mínimo de altura |

### Controles do UserForm

| Controle | Tipo | Propósito |
|----------|------|-----------|
| `lblTitulo` | Label | Cabeçalho com título do sistema |
| `btnMinimizar` | CommandButton | Minimiza a janela |
| `btnMaximizar` | CommandButton | Maximiza/restaura a janela |
| `btnFechar` | CommandButton | Fecha o formulário |
| `fraMenu` | Frame | Container do menu lateral |
| `fraConteudo` | Frame | Container principal de conteúdo |
| `btnDashboard` | CommandButton | Navegação para Dashboard |
| `btnTimeline` | CommandButton | Navegação para Timeline |
| `btnCards` | CommandButton | Navegação para Cards |
| `btnProducao` | CommandButton | Navegação para Cadastro de OP |
| `btnEventos` | CommandButton | Navegação para Eventos |
| `btnConfig` | CommandButton | Navegação para Configurações |

### Eventos Principais

- `UserForm_Initialize`: Inicializa controles e exibe Dashboard
- `UserForm_Resize`: Recalcula layout durante redimensionamento
- `UserForm_QueryClose`: Confirmação de saída

---

## 3. Cabeçalho (Header)

### lblTitulo

| Propriedade | Valor |
|-------------|-------|
| `Left` | 12 |
| `Top` | 12 |
| `Width` | Me.ClientWidth - 24 (responsivo) |
| `Height` | 48 |
| `BackColor` | 3355443 (RGB 51,51,51) |
| `ForeColor` | 16777215 (branco) |
| `Font.Size` | 14 |
| `Font.Bold` | True |
| `TextAlign` | fmTextAlignCenter |

### Botões de Controle da Janela

Todos os botões estão posicionados no canto superior direito do cabeçalho.

| Botão | Left | Top | Width | Height | Cor Fundo | Ícone |
|-------|------|-----|-------|--------|-----------|-------|
| `btnMinimizar` | ClientWidth - 90 | 18 | 24 | 24 | 3355443 | − (U+2212) |
| `btnMaximizar` | ClientWidth - 60 | 18 | 24 | 24 | 3355443 | □ (U+25A1) / ⧉ (U+25C9) |
| `btnFechar` | ClientWidth - 30 | 18 | 24 | 24 | 255 (vermelho) | × (U+00D7) |

**Estilo:**
- `BorderStyle = fmBorderStyleNone`
- `Font.Size = 12` (minimizar/maximizar) ou `14` (fechar)
- `Font.Bold = True`
- Reposicionados automaticamente em `UserForm_Resize`

---

## 4. Menu Lateral (fraMenu)

### Propriedades do Frame

| Propriedade | Valor |
|-------------|-------|
| `Left` | 12 |
| `Top` | 72 |
| `Width` | 160 |
| `Height` | Me.ClientHeight - 96 (responsivo) |
| `BackColor` | 14211288 (RGB 216,216,216) |
| `BorderStyle` | fmBorderStyleSingle |

### Botões de Navegação

Criados dinamicamente em `UserForm_Initialize`. Posicionados verticalmente dentro do `fraMenu`.

| Botão | Left | Top | Width | Height | Caption |
|-------|------|-----|-------|--------|---------|
| `btnDashboard` | 12 | 12 | 124 | 36 | Dashboard |
| `btnTimeline` | 12 | 52 | 124 | 36 | Timeline |
| `btnCards` | 12 | 92 | 124 | 36 | Cards |
| `btnProducao` | 12 | 132 | 124 | 36 | Produção |
| `btnEventos` | 12 | 172 | 124 | 36 | Eventos |
| `btnConfig` | 12 | 212 | 124 | 36 | Configurações |

**Estilo dos botões:**
- `BackColor = 3355443` (header)
- `ForeColor = 16777215` (branco)
- `Font.Size = 10`
- `Font.Bold = True`
- Espaçamento entre botões: 4px (`ESPACAMENTO = 8`, com altura 36)
- Largura do botão: `LARGURA_BOTAO - (2 * ESPACAMENTO) = 124`

### Eventos

Cada botão usa `clsButtonEvents` para capturar cliques via `WithEvents`, garantindo funcionamento correto mesmo com controles criados dinamicamente.

---

## 5. Painel de Conteúdo (fraConteudo)

### Propriedades do Frame

| Propriedade | Valor |
|-------------|-------|
| `Left` | 184 |
| `Top` | 72 |
| `Width` | Me.ClientWidth - 196 (responsivo) |
| `Height` | Me.ClientHeight - 96 (responsivo) |
| `BackColor` | 14211288 |
| `BorderStyle` | fmBorderStyleSingle |
| `ScrollBars` | fmScrollBarsVertical |
| `ScrollHeight` | 2000 |

### Cálculo Dinâmico

```vba
fraConteudo.Left = margem + menuLargura + margem  ' 12 + 160 + 12 = 184
fraConteudo.Width = Me.ClientWidth - fraConteudo.Left - margem
fraConteudo.Height = fraMenu.Height
```

**Garantia de integridade:**
- `fraConteudo.Width < 400` → força para 400px mínimo
- Nunca ultrapassa a área direita do menu lateral

---

## 6. Dashboard

### Estrutura

O Dashboard é o módulo padrão exibido na inicialização. Renderiza 6 indicadores (KPIs) organizados em duas colunas dentro do `fraConteudo`.

### Título do Dashboard

| Propriedade | Valor |
|-------------|-------|
| `Name` | lblDashTitulo |
| `Caption` | INDICADORES DE PRODUÇÃO – DASHBOARD |
| `Left` | 20 |
| `Top` | 20 |
| `Width` | fraConteudo.Width - 40 |
| `Height` | 40 |
| `Font.Size` | 16 |
| `Font.Bold` | True |
| `TextAlign` | fmTextAlignCenter |

### Fonte de Dados

Os KPIs são calculados a partir da `TabelaOPs` na planilha `BD_OPs`, através da função `ObterDadosOPsEmArray()` em `modEngine.bas`.

---

## 7. Duas Colunas de KPI

### Proporções

| Coluna | Proporção | Largura Mínima | Largura Máxima |
|--------|-----------|----------------|----------------|
| Coluna 1 (maior) | 65% do espaço disponível | 300 px | 700 px |
| Coluna 2 (menor) | 30% do espaço disponível | 200 px | 400 px |
| Espaço entre colunas | 40 px | 30 px | 60 px |

### Posicionamento

```
COLUNA1_X = 20
ESPACO_ENTRE_COLUNAS = 40
COLUNA2_X = COLUNA1_X + LARGURA_KPI + ESPACO_ENTRE_COLUNAS
```

### Cards de KPI

Cada KPI é composto por 3 Labels:

| Camada | Nome | Propósito |
|--------|------|-----------|
| Fundo | lblKPI_{uniqueID} | Card branco com borda |
| Título | lblKPI_{uniqueID}_T | Texto do indicador |
| Valor | lblKPI_{uniqueID}_V | Valor numérico |

**Propriedades do card:**

| Propriedade | Valor |
|-------------|-------|
| `BackColor` | vbWhite |
| `BorderStyle` | fmBorderStyleSingle |
| `Altura` | 100 px |
| `Padding` | 10 px |

**Propriedades do título:**

| Propriedade | Valor |
|-------------|-------|
| `Font.Size` | 10 |
| `Font.Bold` | True |
| `ForeColor` | COR_TEXTO_ESCURO (0) |
| `Height` | 30 |
| `Top` | pTop + 10 |

**Propriedades do valor:**

| Propriedade | Valor |
|-------------|-------|
| `Font.Size` | 20 |
| `Font.Bold` | True |
| `Height` | 40 |
| `Top` | pTop + 45 |

### Cores por Status

| Status | Cor |
|--------|-----|
| Concluído | vbGreen |
| Em Andamento | COR_HEADER (3355443) |
| Atrasado | COR_ALERTA (255) |
| Planejadas/Total | COR_TEXTO_ESCURO (0) |

---

## 8. Responsividade

### Tamanho Mínimo

```vba
If Me.ClientWidth < 600 Then Me.ClientWidth = 600
If Me.ClientHeight < 500 Then Me.ClientHeight = 500
```

### Comportamento Responsivo

- `fraMenu` sempre fixo em 160px de largura, nunca redimensiona
- `fraConteudo` ocupa automaticamente todo o espaço restante
- Largura das colunas de KPI recalculada proporcionalmente
- Espaçamento vertical reduzido se altura for pequena
- Botões de controle da janela reposicionados no canto superior direito

### Cálculo de Colunas Responsivo

```vba
larguraDisponivel = fraConteudo.Width - 40
LARGURA_KPI = larguraDisponivel * 0.65
LARGURA_KPI_MENOR = larguraDisponivel * 0.30

' Ajuste se não couber
If COLUNA1_X + LARGURA_KPI + espacoEntreColunas + LARGURA_KPI_MENOR > larguraDisponivel Then
    LARGURA_KPI = (larguraDisponivel - espacoEntreColunas) * 0.65
    LARGURA_KPI_MENOR = (larguraDisponivel - espacoEntreColunas) * 0.35
End If
```

---

## 9. UserForm_Resize

### Funcionamento

O evento `UserForm_Resize` é disparado automaticamente pelo VBA quando o UserForm é redimensionado. Ele recalcula:

1. Tamanho mínimo (600x500)
2. Altura do `fraMenu`
3. Posição e largura do `fraConteudo`
4. Largura do `lblTitulo`
5. Posição dos botões de controle da janela
6. Layout dos KPIs (se Dashboard estiver ativo)

### Fluxo de Execução

```vba
Private Sub UserForm_Resize()
    ' 1. Aplica tamanho mínimo
    If Me.ClientWidth < 600 Then Me.ClientWidth = 600
    If Me.ClientHeight < 500 Then Me.ClientHeight = 500
    
    ' 2. Ajusta menu lateral
    fraMenu.Height = Me.ClientHeight - 72 - margem
    
    ' 3. Ajusta painel de conteúdo
    fraConteudo.Left = margem + menuLargura + margem
    fraConteudo.Width = Me.ClientWidth - fraConteudo.Left - margem
    If fraConteudo.Width < 400 Then fraConteudo.Width = 400
    fraConteudo.Height = fraMenu.Height
    
    ' 4. Ajusta título
    lblTitulo.Width = Me.ClientWidth - margem * 2
    
    ' 5. Reposiciona botões de controle
    btnMinimizar.Left = Me.ClientWidth - 90
    btnMaximizar.Left = Me.ClientWidth - 60
    btnFechar.Left = Me.ClientWidth - 30
    
    ' 6. Recalcula KPIs se necessário
    If m_PainelAtivo = "Dashboard" And Not IsEmpty(m_DadosKPIs) Then
        RenderizarKPIs
    End If
End Sub
```

---

## 10. Redimensionamento Manual

### Implementação

O redimensionamento manual é habilitado via API do Windows, aplicando o estilo `WS_THICKFRAME` ao hWnd do UserForm.

```vba
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
    (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
    (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000

Private Sub AplicarEstiloRedimensionavel()
    Dim hWnd As LongPtr
    Dim estilo As LongPtr
    
    hWnd = Me.hWnd
    estilo = GetWindowLongPtr(hWnd, GWL_STYLE)
    estilo = estilo Or WS_THICKFRAME
    SetWindowLongPtr hWnd, GWL_STYLE, estilo
End Sub
```

### Chamado em

`UserForm_Initialize` → `AplicarEstiloRedimensionavel`

### Resultado

O usuário pode redimensionar a janela arrastando:
- Bordas superior, inferior, esquerda e direita
- Cantos superior esquerdo, superior direito, inferior esquerdo e inferior direito

---

## 11. Minimizar

### Implementação

```vba
Private Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Const SW_MINIMIZE As Long = 6

Private Sub btnMinimizar_Click()
    Dim hWnd As LongPtr
    hWnd = Me.hWnd
    ShowWindow hWnd, SW_MINIMIZE
End Sub
```

### Comportamento

- Minimiza a janela para a barra de tarefas
- Mantém o aplicativo Excel em execução
- Restauração ao clicar no ícone da barra de tarefas

---

## 12. Maximizar

### Implementação

```vba
Private Const SW_MAXIMIZE As Long = 3

Private Sub btnMaximizar_Click()
    Dim hWnd As LongPtr
    hWnd = Me.hWnd
    
    If m_Maximizado Then
        ' Restaura
        ShowWindow hWnd, SW_RESTORE
        Me.Width = m_TamanhoNormalLargura
        Me.Height = m_TamanhoNormalAltura
        Me.Left = m_TamanhoNormalLeft
        Me.Top = m_TamanhoNormalTop
        m_Maximizado = False
        btnMaximizar.Caption = ChrW(9633) ' □
    Else
        ' Maximiza
        m_TamanhoNormalLargura = Me.Width
        m_TamanhoNormalAltura = Me.Height
        m_TamanhoNormalLeft = Me.Left
        m_TamanhoNormalTop = Me.Top
        ShowWindow hWnd, SW_MAXIMIZE
        m_Maximizado = True
        btnMaximizar.Caption = ChrW(9634) ' ⧉
    End If
End Sub
```

### Estado Armazenado

| Variável | Propósito |
|----------|-----------|
| `m_TamanhoNormalLargura` | Largura antes da maximização |
| `m_TamanhoNormalAltura` | Altura antes da maximização |
| `m_TamanhoNormalLeft` | Posição X antes da maximização |
| `m_TamanhoNormalTop` | Posição Y antes da maximização |
| `m_Maximizado` | Flag de estado atual |

---

## 13. Restaurar

### Implementação

A restauração é feita no mesmo procedimento `btnMaximizar_Click`, quando `m_Maximizado = True`.

```vba
If m_Maximizado Then
    ShowWindow hWnd, SW_RESTORE
    Me.Width = m_TamanhoNormalLargura
    Me.Height = m_TamanhoNormalAltura
    Me.Left = m_TamanhoNormalLeft
    Me.Top = m_TamanhoNormalTop
    m_Maximizado = False
    btnMaximizar.Caption = ChrW(9633)
End If
```

### Comportamento

- Restaura para o tamanho e posição exatos anteriores à maximização
- Atualiza o ícone do botão para □ (quadrado)
- Dispara `UserForm_Resize` para recalcular layout

---

## 14. Fechar

### Implementação

```vba
Private Sub btnFechar_Click()
    UserForm_QueryClose 0, vbFormControlMenu
End Sub
```

### Comportamento

- Aciona o evento `UserForm_QueryClose`
- Exibe diálogo de confirmação: "Deseja realmente sair do APS PURAN?"
- Se confirmado: `Application.Visible = True` e fecha o formulário
- Se cancelado: mantém o formulário aberto

---

## 15. Timeline

### Módulo

`modTimeline.bas`

### Função Principal

```vba
Public Sub CarregarTimeline(pContainer As MSForms.Frame)
```

### Funcionalidade

- Exibe ordem de produção em formato de timeline visual
- Carrega dados da `TabelaOPs`
- Renderiza barras de progresso representando duração das OPs
- Cores por status: verde (Concluído), azul (Em Andamento), vermelho (Atrasado)

### Integração

Chamado por `frmPrincipal.ExibirPainel("Timeline")` que passa `fraConteudo` como container.

---

## 16. Cards de Produção

### Módulo

`modCards.bas`

### Função Principal

```vba
Public Sub CarregarCards(pContainer As MSForms.Frame)
```

### Funcionalidade

- Exibe OPs em formato de cards
- Informações por card: ID_OP, Produto, Equipamento, Quantidade, Datas, Status
- Layout em grid adaptativo
- Cores distintas por status

### Integração

Chamado por `frmPrincipal.ExibirPainel("Cards")` que passa `fraConteudo` como container.

---

## 17. Cadastro de OPs

### Formulário

`frmCadastroOP.frm`

### Propriedades

| Propriedade | Valor |
|-------------|-------|
| `Width` | 560 |
| `Height` | 500 |
| `BorderStyle` | 3 (Sizable) |
| `StartUpPosition` | 1 (CenterScreen) |

### Campos

| Campo | Tipo | Rótulo |
|-------|------|--------|
| `txtID_OP` | TextBox | ID OP: |
| `txtProduto` | TextBox | Produto: |
| `cboEquipamento` | ComboBox | Equipamento: |
| `txtQuantidade` | TextBox | Quantidade: |
| `txtData_Inicio` | TextBox | Data Início: |
| `txtData_Fim` | TextBox | Data Fim: |
| `cboStatus` | ComboBox | Status: |

### Botões

| Botão | Ação |
|-------|------|
| `btnSalvar` | Valida e persiste a OP via `SalvarNovaOP` em `modEngine.bas` |
| `btnCancelar` | Fecha o formulário sem salvar |

### Validações

- ID_OP obrigatório e único
- Produto obrigatório
- Equipamento obrigatório
- Quantidade > 0 e numérica
- Datas válidas e Data_Fim >= Data_Inicio
- Cálculo automático de duração em horas

### Exibição

Modal: `frmCadastroOP.Show vbModal`

---

## 18. Eventos

### Formulário

`frmEventos.frm`

### Propriedades

| Propriedade | Valor |
|-------------|-------|
| `Width` | 560 |
| `Height` | 500 |
| `BorderStyle` | 3 (Sizable) |
| `StartUpPosition` | 1 (CenterScreen) |

### Campos

| Campo | Tipo | Rótulo |
|-------|------|--------|
| `txtID_Evento` | TextBox | ID Evento: |
| `cboTipo` | ComboBox | Tipo: |
| `cboEquipamento` | ComboBox | Equipamento: |
| `txtInicio` | TextBox | Início: |
| `txtFim` | TextBox | Fim: |
| `txtMotivo` | TextBox | Motivo: |

### Botões

| Botão | Ação |
|-------|------|
| `btnSalvar` | Valida e persiste o evento via `SalvarEvento` |
| `btnCancelar` | Fecha o formulário sem salvar |

### Validações

- ID_Evento obrigatório e único
- Tipo obrigatório
- Equipamento obrigatório
- Motivo obrigatório
- Datas válidas e Data_Fim >= Data_Inicio

### Exibição

Modal: `frmEventos.Show vbModal`

---

## 19. Configurações

### Módulo

Painel genérico exibido por `ExibirPainelGenerico`.

### Funcionalidade

- Placeholder para futuras configurações do sistema
- Exibe Label centralizada com texto "Painel Ativo: Configurações"

### Integração

Chamado por `frmPrincipal.ExibirPainel("Configurações")`.

---

## 20. Padrão Visual

### Cores do Sistema

| Elemento | RGB | Decimal | Uso |
|----------|-----|---------|-----|
| Fundo principal | (68, 68, 68) | 4474559 | `frmPrincipal.BackColor` |
| Fundo painel | (216, 216, 216) | 14211288 | `fraMenu`, `fraConteudo`, formulários secundários |
| Header | (51, 51, 51) | 3355443 | `lblTitulo.BackColor`, botões do menu |
| Texto claro | (255, 255, 255) | 16777215 | Títulos, botões claros |
| Texto escuro | (0, 0, 0) | 0 | Texto de KPIs, labels |
| Alerta/Erro | (255, 0, 0) | 255 | Botão fechar, KPIs atrasados |
| Azul | (0, 120, 215) | 15773696 | Botões de ação (Salvar) |
| Verde | (0, 255, 0) | 65280 | KPIs concluídos |

### Tipografia

| Elemento | Fonte | Tamanho | Bold |
|----------|-------|---------|------|
| Título principal | Segoe UI | 14 | True |
| Título Dashboard | Segoe UI | 16 | True |
| Botões menu | Segoe UI | 10 | True |
| Título KPI | Segoe UI | 10 | True |
| Valor KPI | Segoe UI | 20 | True |
| Campos de texto | Segoe UI | 9 | False |

### Espaçamentos

| Elemento | Valor |
|----------|-------|
| Margem padrão | 12 px |
| Espaçamento entre botões | 8 px |
| Padding interno de cards | 10 px |
| Espaço entre colunas de KPI | 40 px |
| Espaço vertical entre linhas de KPI | 140 px |

### Bordas

| Elemento | Estilo |
|----------|--------|
| UserForm principal | BorderStyle = 3 (Sizable) |
| Formulários secundários | BorderStyle = 3 (Sizable) |
| Frames | fmBorderStyleSingle |
| Labels (cards) | fmBorderStyleSingle |
| Botões de controle | fmBorderStyleNone |
| Campos de texto | fmBorderStyleSingle |

---

## 21. Regras de Desenvolvimento

### 1. Não alterar Motor APS
- `modEngine.bas` não deve ser modificado para ajustes de layout
- Cálculos de duração, status e regras de negócio são intocáveis

### 2. Controles Dinâmicos
- Todos os controles do `frmPrincipal` são criados em tempo de execução via `Controls.Add`
- Não usar a janela de design do VBA para arrastar controles no `frmPrincipal`
- Usar `clsButtonEvents` para capturar eventos de controles dinâmicos

### 3. Responsividade
- Sempre usar `Me.ClientWidth` e `Me.ClientHeight` para cálculos
- Nunca hardcodar posições que dependam do tamanho da janela
- Manter tamanho mínimo de 600x500

### 4. API do Windows
- Usar `#If VBA7 Then` para compatibilidade 32/64-bit
- Preferir `Me.hWnd` quando disponível (VBA7)
- Fallback para `GetActiveWindow()` em VBA6

### 5. Eventos
- Não duplicar eventos de botões
- Usar `clsButtonEvents` para controles criados dinamicamente
- Manter `UserForm_Resize` como único ponto de recalculo de layout

### 6. Cores
- Usar constantes nomeadas (`COR_FUNDO`, `COR_PAINEL`, etc.)
- Não hardcodar valores RGB diretamente nos controles

### 7. Nomenclatura
- Formulários: `frm` + Nome (ex: `frmPrincipal`)
- Módulos: `mod` + Nome (ex: `modEngine`)
- Classes: `cls` + Nome (ex: `clsButtonEvents`)
- Controles: prefixo + Nome (ex: `btnSalvar`, `lblTitulo`)

### 8. Tratamento de Erros
- Sempre usar `On Error GoTo` com label de saída
- Sempre restaurar estado (`Application.ScreenUpdating = True`) no erro
- Nunca usar `On Error Resume Next` a menos que absolutamente necessário

---

## 22. Critérios Finais de Aceitação

### Funcionalidade

- [ ] O formulário abre centralizado na tela
- [ ] O menu lateral exibe todos os 6 botões de navegação
- [ ] Cada botão do menu abre o módulo correspondente
- [ ] O Dashboard exibe 6 KPIs em duas colunas
- [ ] Os cálculos dos KPIs estão corretos conforme regras do APS
- [ ] O cadastro de OP persiste dados na `TabelaOPs`
- [ ] O registro de eventos persiste dados na `TabelaEventos`

### Layout e Design

- [ ] Nenhum botão do menu é coberto pelo painel de conteúdo
- [ ] A coluna menor de KPIs nunca ultrapassa a área do `fraConteudo`
- [ ] Não há controles sobrepostos em qualquer tamanho de janela
- [ ] O título do formulário está sempre visível e com texto completo
- [ ] Os botões de controle da janela estão sempre no canto superior direito

### Responsividade

- [ ] A janela pode ser redimensionada arrastando bordas e cantos
- [ ] O tamanho mínimo de 600x500 é respeitado
- [ ] `fraMenu` permanece fixo à esquerda durante redimensionamento
- [ ] `fraConteudo` ocupa automaticamente o espaço restante
- [ ] As colunas de KPI se adaptam proporcionalmente
- [ ] `UserForm_Resize` é disparado durante redimensionamento manual

### Controles de Janela

- [ ] Botão minimizar minimiza a janela para barra de tarefas
- [ ] Botão maximizar expande a janela para área cliente do Excel
- [ ] Botão restaurar volta ao tamanho anterior exato
- [ ] Botão fechar exibe confirmação e fecha corretamente
- [ ] Não há conflitos entre redimensionamento manual e botões de controle

### Compatibilidade

- [ ] Funciona em Office 32-bit
- [ ] Funciona em Office 64-bit
- [ ] Não há dependências de bibliotecas externas
- [ ] O projeto compila sem erros

### Performance

- [ ] Abertura do formulário em menos de 2 segundos
- [ ] Redimensionamento sem flicker perceptível
- [ ] Navegação entre módulos fluida
- [ ] Sem vazamento de memória ao fechar/reabrir

---

## 23. Arquitetura de Arquivos

```
Projeto_aps/
├── frmPrincipal.frm          # Interface principal (menu + dashboard)
├── frmCadastroOP.frm         # Cadastro de Ordem de Produção
├── frmEventos.frm            # Registro de Eventos e Paradas
├── clsButtonEvents.cls       # Classe para eventos de botões dinâmicos
├── modEngine.bas             # Motor APS (regras de negócio)
├── modTimeline.bas           # Módulo Timeline
├── modCards.bas              # Módulo Cards de Produção
├── modDatabase.bas           # Módulo de banco de dados
├── ThisWorkbook.cls          # Inicialização do aplicativo
├── DOCUMENTACAO/
│   └── APS_PURAN_LAYOUT_UI_UX.md  # Este documento
└── [Planilhas ocultas]
    ├── BD_OPs/BD_Equipamentos/BD_Eventos (Worksheets)
    └── TabelaOPs/TabelaEquipamentos/TabelaEventos (ListObjects)
```

---

## 24. Referências

- [VBA UserForm Object](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/userforms)
- [MSForms.CommandButton](https://learn.microsoft.com/en-us/office/vba/api/msforms.commandbutton)
- [Windows API SetWindowLongPtr](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlongptra)
- [ShowWindow function](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow)
