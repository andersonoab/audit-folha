# Validação — seleção e totalização

Validação executada em Chromium 143, headless, em abertura `file://`, com acesso HTTP/HTTPS bloqueado nos testes. As bibliotecas JavaScript foram carregadas da pasta `vendor` real, sem substituição por mocks. A versão empacotada também foi extraída e executada.

## Resultado

**16 grupos de validação de integração aprovados**, além dos testes unitários do modelo e das verificações de sintaxe.

| Cenário | Resultado observado |
| --- | --- |
| Clique simples e segundo clique | Seleciona e desmarca apenas a linha; não abre investigação |
| Seleção múltipla | Cumulativa, sem Ctrl |
| Checkbox | Uma alteração por ação; não propaga clique para a linha |
| INSS + PGBL | −R$ 988,07 − R$ 270,00 = **−R$ 1.258,07** |
| Salário + INSS | R$ 9.000,00 − R$ 988,07 = **R$ 8.011,93** |
| Desmarcação | Retira somente a contribuição desmarcada |
| Valor zero | Acrescenta uma linha ao contador sem mudar o total |
| Valores negativos na fonte | Descontos já negativos mantêm o sinal; não ocorre dupla inversão |
| Registros duplicados | Duas ocorrências de R$ 0,01 têm IDs distintos e totalizam R$ 0,02 |
| CNPJs diferentes | Mesma matrícula/nome são agregados distintos; investigação mantém a empresa correta |
| Naturezas distintas | Folha, bases, provisões e líquido informado usam subtotais próprios |
| Precisão | Conversão decimal por dígitos e soma BigInt; inclui 0,1 + 0,2, notação científica e arredondamento de 1,005 |
| Destaque | Persistente após sair com o mouse; hover usa outro fundo |
| Teclado | Espaço seleciona; Enter abre detalhe; sem mudança de seleção ao investigar |
| Duplo clique | Abre composição/detalhe; seleção e total finais permanecem idênticos |
| Ordenação | Crescente/decrescente preservam IDs, checkboxes e totais |
| Paginação | Seleções de páginas diferentes acumulam; voltar restaura os checkboxes |
| Filtros | Remove somente IDs que saíram do resultado e exibe aviso |
| Resultado vazio | Seleção vazia e total zero, inclusive sem empresas selecionadas |
| Holerite | Soma independente de Calc., busca e classes respeitadas |
| Conferência | Seleção numérica não modifica as conferências salvas |
| Análise de verbas | Soma mista por fontes; limpeza das verbas esvazia seleção antiga |
| Excel | Workbook comparado antes/depois da seleção: mesmas abas, população e valores, excluindo somente horário de geração |
| Reprocessamento | Preserva IDs quando o contexto permanece; limpa quando o filtro exclui todos |
| Importação | Um mês e vários meses funcionam; nova importação reinicia a seleção |
| Fonte original | Valores reimportados idênticos antes e depois de selecionar |
| Recibo auxiliar | Lançamentos, comparativos, visão geral, consulta e divergências usam o mesmo componente |
| Pacote completo | Index e recibo abrem localmente; não exigem CDN para cálculos, gráficos ou Excel |

## Limites de cobertura

- Foi usada uma base fictícia de 753 ocorrências e 124 agregados de colaboradores, com duas empresas e duas competências, além da variante com uma competência. Não foi disponibilizada uma base operacional real para reconciliação de produção.
- O clique simples de linha é confirmado após **360 ms**, pois o navegador primeiro emite eventos de clique e depois informa o duplo clique. Checkboxes e Espaço são imediatos. Em duplo clique lento reconhecido pelo sistema após esse intervalo, o componente restaura a seleção anterior antes de investigar.
- A seleção numérica não é persistida entre recargas/importações; o mecanismo original de conferências continua separado.
- Interfaces Python anteriores estão preservadas e não receberam a função. Não foram executadas neste ambiente.
- Não foram validados instaladores EXE, impressão física, navegadores antigos ou desempenho com milhões de registros. A interface requer navegador atual com suporte a BigInt e dialog.
- Fontes Google opcionais podem usar fallback sem internet; CSV, gráficos e Excel estão disponíveis localmente.

Os testes estão em `tests/selection-model.test.cjs` e `tests/browser.test.cjs`. As bases para reprodução estão em `examples/`.
