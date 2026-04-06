# UX Handoff

## Tese visual

Painel financeiro com cara de mesa de controle, nao de mosaico genérico: base clara, contraste forte, acento petróleo, leitura densa e hierarquia editorial.

## Plano de conteúdo

- `Base`: orientacao, carga e governanca da sessao.
- `Configurações`: decisao operacional de regras e CFOP.
- `Versões`: memoria auditavel do trabalho.
- `Análises`: leitura executiva e drill-down.

## Tese de interação

- Navegação lateral persistente com foco em fluxo, nao em menu genérico.
- Transição de estado por contexto: upload -> configuracao -> snapshot -> analise.
- Drill-down por modal, mantendo a pagina principal limpa.

## Arquivo Figma recomendado

Criar um arquivo dedicado com as paginas:

- `00_Audit`
- `01_Flows`
- `02_Wireframes`
- `03_High-Fidelity`
- `04_Components`

## Fluxos que precisam estar desenhados

1. Nova carga manual.
2. Carga + aplicacao do padrão global.
3. Reclassificacao manual de CFOP.
4. Salvamento de snapshot.
5. Abertura de versão congelada.
6. Duplicacao de versão para nova revisão.

## Wireframes obrigatórios

- `Base / sem carga`
- `Base / carga concluída`
- `Configurações / regras`
- `Configurações / tabela de CFOP`
- `Versões / lista`
- `Análises / visão executiva`
- `Análises / drill-down modal`

## Componentes para o design system

- Nav lateral com estado ativo
- Hero de contexto da página
- Chips de estado
- Painel de filtros
- Cartões de KPI
- Tabela com ações inline
- Callouts de status
- Modal de drill-down

## Regras de composição

- Primeira dobra sempre orienta o operador sobre onde ele está no fluxo.
- `Configurações` precisa parecer área de decisão, nao um rodape secundario.
- `Versões` precisa enfatizar segurança histórica e nao apenas listagem.
- `Análises` deve priorizar leitura e decisão; tabelas entram depois da síntese.

## Tipografia e tom

- Tipografia principal: sans geométrica ou humanista com peso forte.
- Tipografia auxiliar: mono para labels, ids e metadados.
- Tom visual: sóbrio, operacional, premium e sem elementos decorativos gratuitos.

## Observação de execução

A conta Figma conectada nesta sessão está em modo de visualização. O handoff acima está pronto para materialização assim que houver acesso de edição no arquivo destino.
