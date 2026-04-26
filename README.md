# Auditoria de Folha — Desktop App

Aplicativo desktop em Python para auditoria de folha de pagamento, com confronto entre folha calculada (mês-alvo) e folhas implantadas (baseline), análise de outliers por critérios estatísticos de mercado, drill-down por funcionário e por verba, e exportação para Excel.

Desenvolvido como ferramenta interna para análise de payroll, segue o padrão visual Igarapé Digital e implementa critérios estatísticos consagrados em literatura técnica de auditoria.

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Critérios estatísticos](#critérios-estatísticos)
- [Estrutura do CSV de entrada](#estrutura-do-csv-de-entrada)
- [Instalação](#instalação)
- [Uso](#uso)
- [Abas do aplicativo](#abas-do-aplicativo)
- [Filtros e parâmetros](#filtros-e-parâmetros)
- [Exportação](#exportação)
- [Atalhos e ergonomia](#atalhos-e-ergonomia)
- [Decisões técnicas](#decisões-técnicas)
- [Limitações conhecidas](#limitações-conhecidas)
- [Estrutura do código](#estrutura-do-código)
- [Licença e autoria](#licença-e-autoria)

---

## Funcionalidades

- **Confronto folha calculada vs implantadas** com cards de PGTO, DESC, Líquido e Verba 9950, exibindo variação percentual e absoluta.
- **Detecção automática de desligados** via campo "Data Rescisão" comparado ao último dia do mês-alvo, com opção de inclusão/exclusão na análise.
- **Cálculo de líquido por funcionário** (PGTO − DESC) com classificação de status: AUSENTE, NEGATIVO, ZERO_SUSPEITO, EXTREMA, ALTA, Z_2SIGMA, MATERIAL, NOVO_FUNC, OK, SEM_DADOS, DESLIGADO.
- **Quatro métodos estatísticos selecionáveis** para detecção de outliers: Shewhart 3-sigma, Tukey IQR, MAD modificado e modo legacy heurístico.
- **Materialidade ISA 320** parametrizável por slider de 0% a 5% da folha total.
- **Análise de Headcount** via verba 0020 (Armazena Salário) com HC mensal, salário médio e total.
- **Top 15 outliers** por impacto absoluto em reais, com identificação de status e variação.
- **Dispersão Baseline vs Folha calculada** em escala logarítmica, com banda de tolerância de ±30% e numeração dos top 5 críticos com legenda lateral.
- **Tabela detalhada** com filtros por status, ordenação configurável e busca por nome/matrícula.
- **Drill-down por funcionário** com gráfico de evolução mensal do líquido e tabela de verbas movimentadas com filtros por C.R., Classe, Processo e busca textual.
- **Detecção de verbas regulares zeradas** que apareceram em ≥3 meses da baseline (≥R$ 500) e zeraram no mês-alvo.
- **Análise de verbas com multi-seleção** e busca textual, agregando totais e calculando HC de matrículas únicas em qualquer das verbas selecionadas.
- **Top 20 verbas por impacto absoluto** ordenadas por classe (PGTO → DESC → OUTRO), com destaque visual para verbas excluídas do cálculo.
- **Filtro de verbas** via diálogo modal, com proteção da verba 0020 e ações em massa (incluir/excluir/inverter visíveis).
- **Exportação para Excel** com 5 abas: Resumo_Mensal, Líquido_Funcionario, Verbas_PGTO_Zeradas, Verbas_DESC_Zeradas e Parametros (rastreabilidade do método estatístico utilizado).
- **Texto selecionável** em cards e labels críticos, permitindo seleção parcial com mouse e Ctrl+C nativo.
- **Cópia de tabelas como TSV** via Ctrl+C ou menu de contexto, colável diretamente no Excel.

---

## Critérios estatísticos

A detecção de variações anômalas usa quatro métodos selecionáveis no painel lateral. Cada relatório exportado registra o método utilizado na aba `Parametros` para rastreabilidade.

### Shewhart 3-sigma + Materialidade ISA 320 (default recomendado)

Padrão SPC industrial de Walter A. Shewhart (1931), formalizado nas Western Electric Rules (1956). Calcula z-score `(x − μ) / σ` sobre os meses de baseline:

- `|z| ≥ 2` → status ALTA (≈ 5% esperado em distribuição normal)
- `|z| ≥ 3` → status EXTREMA (≈ 0,3% esperado)
- Sobreposto: variação absoluta ≥ materialidade da folha → status MATERIAL

Materialidade segue ISA 320 (IFAC) e NBC TA 320 (CFC): de 0,5% a 5% da base, com default de 1% da folha líquida total do mês-alvo.

### Tukey IQR (boxplot clássico)

John W. Tukey, *Exploratory Data Analysis* (1977). Robusto a distribuições não-gaussianas. Calcula Q1, Q3 e IQR sobre a distribuição global de VAR_ABS:

- Fora de `Q1 − 1,5×IQR` ou `Q3 + 1,5×IQR` → status ALTA_IQR
- Fora de `Q1 − 3×IQR` ou `Q3 + 3×IQR` → status EXTREMA_IQR

### MAD modificado (Iglewicz-Hoaglin)

Iglewicz, B. & Hoaglin, D.C. (1993), *How to Detect and Handle Outliers*. Mais robusto que sigma quando há poucos meses de baseline ou alta volatilidade individual. Calcula `z* = 0,6745 × (x − mediana) / MAD` por funcionário:

- `|z*| ≥ 2,5` → status ALTA
- `|z*| ≥ 3,5` → status EXTREMA (threshold do paper original)

### Anderson Legacy (98%/30%)

Heurística original mantida para compatibilidade com auditorias históricas:

- `|VAR_PCT| ≥ 30%` e `|VAR_ABS| > R$ 200` → ALTA
- `|VAR_PCT| ≥ 98%` e `|VAR_ABS| > R$ 500` → EXTREMA

### Regras universais (aplicadas a todos os métodos)

- Funcionário com data de rescisão ≤ último dia do mês-alvo → DESLIGADO
- Sem baseline mas com líquido no mês-alvo → NOVO_FUNC
- Líquido = 0 com baseline > 0 → AUSENTE
- Líquido < −R$ 100 → NEGATIVO
- Líquido < R$ 500 com baseline ≥ R$ 1.000 → ZERO_SUSPEITO

---

## Estrutura do CSV de entrada

O aplicativo lê arquivos CSV no padrão de exportação ADP "Confere - Códigos por Período":

- Encoding: `latin-1`
- Separador: `;`
- Decimal: vírgula com ponto como separador de milhares
- Colunas obrigatórias: `Empresa`, `Matrícula`, `Nome`, `Código`, `Descrição`, `Clas.`, `C.R.`, `Processo`, `Data Rescisão`
- Colunas mensais: padrão `MMM/AA - Valor` (exemplo: `JAN/26 - Valor`, `FEV/26 - Valor`)

Os meses são detectados automaticamente a partir das colunas `* - Valor` presentes no arquivo.

---

## Instalação

### Pré-requisitos

- Python 3.10 ou superior
- Sistema operacional: Windows, macOS ou Linux

### Dependências

Instale via pip:

```
pip install customtkinter pandas numpy matplotlib openpyxl
```

Versões testadas:

- customtkinter ≥ 5.2.0
- pandas ≥ 2.0
- numpy ≥ 1.24
- matplotlib ≥ 3.7
- openpyxl ≥ 3.1

### Para distribuição em ambiente externo

Empacotamento via cx_Freeze (padrão Win32GUI). Crie um `setup.py` na pasta do projeto:

```
from cx_Freeze import setup, Executable
import sys

base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="AuditoriaFolha",
    version="1.0",
    options={"build_exe": {
        "packages": ["customtkinter", "pandas", "numpy", "matplotlib", "openpyxl", "tkinter"],
        "include_files": [],
    }},
    executables=[Executable("auditoria_folha_ctk.py", base=base, target_name="AuditoriaFolha.exe")]
)
```

Build: `python setup.py build`

---

## Uso

Execute diretamente:

```
python auditoria_folha_ctk.py
```

Fluxo recomendado:

1. Clique em **Carregar arquivo CSV** no painel lateral e selecione o arquivo exportado do ADP
2. Os filtros de **Empresa** e **Processo** são populados automaticamente; ajuste se necessário
3. O **mês-alvo** vem definido como o último mês detectado; ajuste no combo se desejar comparar outro mês
4. Os meses de **baseline** vêm marcados por padrão (todos exceto o alvo); desmarque os que não quer incluir
5. Selecione o método estatístico em **Parâmetros de outlier**
6. Ajuste a **Materialidade ISA 320** no slider, se aplicável
7. Opcionalmente, clique em **Filtrar verbas...** para excluir verbas específicas do cálculo
8. Clique em **Reprocessar auditoria** se mudar qualquer parâmetro
9. Navegue pelas abas para análise
10. Aba **Exportar** gera o arquivo Excel completo

---

## Abas do aplicativo

### Visão Geral

- 4 cards: PGTO, DESC, Líquido, Verba 9950 com variação vs baseline
- Aviso de desligados detectados
- Gráfico de evolução mensal: barras de PGTO e DESC, linha de Líquido, marcadores diamante para Verba 9950
- Conciliação no mês-alvo (PGTO, DESC, Líquido, Verba 9950)
- Distribuição de status (barras horizontais)

### Headcount

- 3 cards: HC verba 0020, Salário médio, Total verba 0020
- Gráfico de barras de HC mensal com linha de salário médio em eixo Y secundário
- Linha de referência do HC médio baseline

### Outliers

- Dispersão Baseline × Folha calculada em escala log, com banda ±30% e diagonal y=x
- Top 5 outliers numerados no plot com legenda lateral (Mat, Nome, status, impacto, var %)
- Top 15 outliers em barras horizontais com rótulos completos

### Tabela & Drill-down

- Tabela superior: filtros por status, ordenação (impacto, var %, líquido, baseline) e busca textual; multi-seleção com Ctrl+Click
- Drill-down inferior ao selecionar um funcionário:
  - Gráfico de evolução do líquido com baseline como linha de referência
  - Tabela de verbas movimentadas no mês-alvo com filtros por C.R., Classe, Processo e busca de código/descrição
  - Cabeçalhos das colunas dos 3 meses anteriores atualizados dinamicamente

### Verbas Zeradas

- Verbas regulares de PGTO e DESC que aparecem em ≥3 meses da baseline (≥R$ 500) e zeraram no mês-alvo
- Top 15 ranqueadas por média baseline em barras horizontais

### Análise de Verbas

- **Top 20 verbas por impacto absoluto** ordenadas em PGTO → DESC → OUTRO; verbas excluídas em cinza com hachura `//`
- **Listbox multi-seleção** com busca textual em tempo real e botões de seleção rápida por classe
- **4 cards** de HC, Total agregado, Valor médio e Status (incluída/excluída)
- **Gráfico HC + valor médio** mensal das verbas selecionadas, similar à análise da verba 0020
- HC = matrículas únicas com valor ≠ 0 em qualquer das verbas selecionadas

### Exportar

- Botão de salvar relatório Excel com 5 abas
- Status visual da última exportação

---

## Filtros e parâmetros

### Sidebar

- **Empresa(s)**: multi-seleção por checkbox quando há mais de uma
- **Processo (tipo de folha)**: multi-seleção por checkbox quando há mais de um (MENSAL, FERIAS, RESCIS, RESCOM, ADIPAR, ADIQUI, DECTER, etc.)
- **Mês-alvo**: combo com todos os meses detectados
- **Meses de baseline**: checkboxes; default = todos exceto o alvo
- **Excluir desligados da análise**: switch on/off
- **Filtrar verbas...**: abre diálogo modal de gerenciamento
- **Parâmetros de outlier**: combo com os 4 métodos estatísticos
- **Materialidade ISA 320**: slider de 0% a 5% (default 1%)

### Diálogo "Filtrar verbas"

- Tabela ordenada PGTO → DESC → OUTRO + valor decrescente
- Coluna "Inc?" com toggle por clique (Sim/Não/—)
- Verba 0020 protegida (não pode ser excluída pois é usada para HC)
- Filtro por classe e busca textual
- Ações em massa: Incluir todas (visíveis), Excluir todas (visíveis), Inverter (visíveis)
- Status no rodapé: quantas excluídas + impacto absoluto em reais

---

## Exportação

O arquivo Excel gerado contém 5 abas:

| Aba | Conteúdo |
|---|---|
| Resumo_Mensal | PGTO, DESC, Líquido, Verba 9950 e número de funcionários por mês |
| Liquido_Funcionario | Cada matrícula com valores mensais, baseline, var abs/pct, z-score, status, auditoria textual e data de rescisão |
| Verbas_PGTO_Zeradas | Verbas pagas regulares que zeraram no mês-alvo |
| Verbas_DESC_Zeradas | Descontos regulares que zeraram no mês-alvo |
| Parametros | Método estatístico, materialidade, Q1/Q3/IQR, folha total, mês-alvo, baseline, timestamp de geração |

A aba `Parametros` garante rastreabilidade da auditoria: em qualquer momento futuro é possível verificar qual critério foi aplicado naquele relatório.

---

## Atalhos e ergonomia

- **Ctrl+C** em qualquer tabela copia a(s) linha(s) selecionada(s) como TSV (cola direto no Excel mantendo colunas)
- **Botão direito** em tabelas: menu com "Copiar linha selecionada" e "Copiar tabela inteira (visível)"
- **Botão direito** em cards e labels: menu "Copiar"
- **Texto selecionável** em valores de cards, label de drill-down, stats da sidebar e aviso da aba Análise: arraste para selecionar trecho parcial e Ctrl+C
- **Ctrl+Click** e **Shift+Click** funcionam para multi-seleção em listboxes e treeviews
- **Espaço** no diálogo de filtro de verbas alterna a linha selecionada
- **Cursor I-beam** em campos selecionáveis sinaliza interatividade

---

## Decisões técnicas

### Escolha de bibliotecas

- **CustomTkinter** para a UI: nativo, sem dependências de browser, baixa latência em desktop, fácil empacotamento via cx_Freeze
- **Matplotlib (TkAgg)** para gráficos: integração direta no canvas, controle fino de rótulos e anotações
- **Pandas + NumPy** para o pipeline de dados: agrupamentos, agregações e cálculos vetorizados
- **OpenPyXL** para exportação Excel: suporte completo a múltiplas abas e tipos numéricos

### Trade-offs assumidos

- Streamlit foi descartado em favor de CustomTkinter para distribuição offline em ambiente corporativo sem dependência de browser ou servidor
- Plotly interativo substituído por matplotlib estático: perde hover dinâmico, ganha empacotamento simples
- Tooltip nos pontos do scatter não foi implementado: substituído por numeração dos top 5 + legenda lateral, mais legível em apresentações
- Filtro de baseline mínima R$ 100 na dispersão: evita que matrículas com resíduos microscópicos esmaguem a escala log

### Padrão visual Igarapé Digital

Cores oficiais (constantes no topo do código):

- Dominante `#0083CA` (azul)
- Apoio `#003C64` (azul escuro), `#005A64` (verde escuro), `#6EB4DC` (azul claro)
- Acentos `#7D0041` (vinho), `#8C321E` (laranja queimado)
- Texto `#1A1A1A` (alto contraste), `#646464` (secundário)
- Grid `#CCCCCC`
- Fundo `#FFFFFF`

Fontes: Rufina (cabeçalhos, títulos de cards e gráficos) e Roboto (corpo, rótulos, controles).

### Verba 0020 protegida

A verba 0020 (Armazena Salário) é a referência canônica de HC ativo na folha. Por isso:

- Não pode ser excluída via filtro de verbas (ignorada se vier marcada)
- É a base do cálculo de HC mensal e salário médio na aba Headcount
- Aparece no diálogo de filtro com indicador visual diferenciado

### Texto selecionável

`tk.Text` em `state="disabled"` bloqueia também a seleção em algumas versões do Tk em diferentes plataformas. A solução adotada mantém `state="normal"` e intercepta todas as teclas via `bind("<Key>")`, permitindo apenas Ctrl+C, Ctrl+A e teclas de navegação. Comportamento robusto cross-platform.

---

## Limitações conhecidas

- A aba "Análise de Verbas" sempre mostra todas as verbas (inclusive excluídas), porque é uma aba de análise, não de operação. Verbas excluídas afetam os cálculos das outras abas mas continuam visíveis aqui para auditoria.
- A análise de "verbas zeradas" examina o universo completo de verbas, ignorando o filtro de exclusão (o objetivo é justamente detectar verbas que sumiram).
- Tabela detalhada exibe no máximo 1.000 linhas por vez; o total continua disponível no rótulo de contagem e na exportação Excel.
- Tooltip por hover em pontos do scatter não está implementado (matplotlib estático).
- Ordenação por clique em coluna do Treeview não está implementada por padrão; a ordenação é controlada pelo combo "Ordenar".

---

## Estrutura do código

Arquivo único `auditoria_folha_ctk.py` (≈2900 linhas), organizado em blocos:

1. **Imports e constantes** — paleta de cores, status críticos, parâmetros estatísticos de mercado
2. **Utilitários de formatação** — `fmt_brl`, `fmt_pct`, `fmt_int`, `fmt_eixo_brl`, `ultimo_dia_mes`
3. **Camada de dados** — `carregar_csv`, `detectar_desligados`, `resumo_macro`, `salario_verba20_por_mes`, `hc_e_total_por_verba`, `impacto_por_verba`, `liquido_por_funcionario`, `verbas_zeradas`
4. **Helpers de UI** — `cor_status`, `estilo_eixos`, `_copy_to_clipboard`, `make_label_copiable`, `bind_treeview_copy`
5. **Widgets selecionáveis** — `SelectableEntry`, `SelectableText`
6. **Componentes** — `MetricCard`, `ChartContainer`
7. **Diálogo modal** — `FiltroVerbasDialog`
8. **Aplicação principal** — `AuditoriaFolhaApp` com métodos `_build_*` para layout e `_render_*` para renderização

Ponto de entrada: `if __name__ == "__main__":` ao final do arquivo.

---

## Licença e autoria

Aplicativo desenvolvido por **Anderson Marinho | Igarapé Digital** como ferramenta de auditoria de folha. Padrão visual Igarapé Digital aplicado integralmente.

Uso interno corporativo. Para distribuição externa, consultar o autor.

Anderson Marinho | Igarapé Digital
