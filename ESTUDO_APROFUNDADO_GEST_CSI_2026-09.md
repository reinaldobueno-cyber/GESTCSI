# Estudo aprofundado de evolução do GEST CSI

> Plano de execução relacionado: [Plano de evolução do GEST CSI](./PLANO_EVOLUCAO_BACKEND_GEST_CSI.md), identificado como `GESTCSI-BACKEND-V1-20260914`.

## Resumo executivo

O GEST CSI não deve passar por uma retirada ampla de funcionalidades neste momento. A evidência disponível mostra um produto com valor operacional real, mas com três problemas que se reforçam: não existe medição central de uso por tela, a arquitetura expõe complexidade técnica ao usuário e a navegação distribui funções relacionadas em quinze entradas de primeiro nível. Retirar módulos agora confundiria baixo uso com baixa encontrabilidade ou baixa confiabilidade.

A prioridade correta é estabilizar, medir e consolidar. O desenho recomendado reduz as quinze entradas atuais para cinco áreas de trabalho, preservando as capacidades existentes; adota a tela de Bonificação como sistema visual de referência; e substitui a cadeia de sincronizações visíveis e fallbacks no navegador por uma camada de dados materializada, monitorada e automática.

Há também uma conclusão importante sobre confiabilidade. O histórico técnico anterior registrou que 42 de 54 alterações recentes eram de sustentação, correção ou refinamento, evidenciando custo elevado de manutenção.^1 Contudo, os registros agregados da planilha mostram melhora expressiva: no histórico completo, 2.172 de 2.604 cargas terminaram literalmente como `ok` (83,4%), enquanto 305 foram registradas como `falha` (11,7%); desde 15/08/2026, 343 de 344 cargas registradas terminaram como `ok` (99,7%).^2 O sistema melhorou, mas sua experiência ainda comunica fragilidade porque mantém botões manuais de sincronização, diversas rotas de contingência e mensagens técnicas dispersas.

O caminho para entregar a mesma percepção de excelência do Gestão de Atendimento não é copiar apenas o tema escuro. A referência possui backend dedicado, armazenamento persistente, tarefas agendadas, observabilidade habilitada, indicadores de saúde e uma suíte identificada de 352 testes. No GEST CSI publicado, a suíte possui 37 testes e a camada central permanece concentrada em um HTML de aproximadamente 1,13 MB e um Apps Script de aproximadamente 438 KB.^3 A diferença essencial é operacional: no Gestão de Atendimento, o produto explica quando está saudável, quando está degradado e o que fará automaticamente; no GEST CSI, o usuário ainda participa demais da manutenção do dado.

### Decisão recomendada

1. **Não excluir funcionalidades durante os próximos 30 dias.** Instrumentar uso real por rota e ação.
2. **Reduzir imediatamente a complexidade percebida**, agrupando as quinze entradas em cinco áreas sem apagar capacidades.
3. **Tratar segurança e integração como prioridade zero**, especialmente o endpoint Apps Script indisponível no teste externo, as leituras públicas de planilha e as rotas mutáveis sem autenticação uniforme.
4. **Usar Bonificação como padrão interno de design**, com hierarquia editorial, resumo executivo, detalhamento auditável e impressão consistente.
5. **Migrar gradualmente a integração para um backend operacional**, mantendo Google Sheets como entrada administrativa e fonte de conferência, não como API pública consultada diretamente pelo navegador.

## 1. Escopo, evidências e limitações

Este estudo cobre o GEST CSI, sua integração com Google Sheets, Apps Script, ClickUp e CMAX, e uma comparação técnica e visual com o Gestão de Atendimento disponível no mesmo ambiente. Foram examinados o levantamento técnico revisado, o repositório local, a referência remota correspondente à versão publicada, os testes existentes, as abas técnicas acessíveis por consulta agregada e a documentação oficial das plataformas.

O ambiente local não representa sozinho a produção. A branch aberta estava em `6386e0d`, com alterações não publicadas, 492 commits atrás e 50 commits à frente de `gestcsi/main`. A versão pública identificada foi `2026-09-11-clickup-activity-v4`, correspondente à referência `gestcsi/main` em `11e7c2f`; o arquivo local principal ainda declarava `2026-08-05-portfolio-195-v1`.^3 Por isso, as conclusões de arquitetura e produto tomam a versão publicada como referência principal e tratam o worktree local apenas como evidência complementar.

O diagnóstico de uso possui uma limitação decisiva: o painel registra cargas, desempenho local e atualizações, mas `goTab()` não registra visualizações de telas ou cliques de funções em uma base central.^3 Assim, há evidência de uso do sistema e de algumas operações que persistem dados, mas não há base para afirmar com rigor que Dashboard, Executivo, Logística, Vendedores ou qualquer outra tela seja pouco usada. Toda recomendação de retirada definitiva fica condicionada à telemetria proposta neste relatório.

## 2. O que é usado e o que ainda não pode ser medido

### 2.1 Evidência objetiva disponível

| Evidência | Resultado observado | Leitura correta |
|---|---:|---|
| Atualizações registradas desde 22/04/2026 | 2.604 | O painel tem uso recorrente, mas cada registro representa uma carga, não uma sessão única |
| Rótulos de ator/dispositivo no histórico | 119 | Indica variedade de acessos; não equivale a 119 pessoas, pois há rótulos de navegador/dispositivo |
| Cargas manuais | 305 | A operação ainda recorre com frequência à atualização explícita |
| Cargas automáticas | 2.235 | A automação já existe e é o caminho predominante |
| Monitor automático | 64 | A reconciliação automática mais recente está ativa, mas ainda é pequena no histórico total |
| Acompanhamentos persistidos | 161 | A função foi realmente usada, não é apenas código disponível |
| Distribuição dos acompanhamentos | 96 em maio; 43 em junho; 21 em julho; 1 em agosto | Há queda forte de adoção ou mudança de processo; precisa ser validada com o time antes de retirar |
| Atividade ClickUp materializada | 12 linhas; atualização até 14/09/2026 | O controle administrativo de adoção está sendo alimentado recentemente |

Fonte dos agregados: consultas `count`, `min` e `max` nas abas técnicas `PANEL_UPDATE_HISTORY`, `ACOMPANHAMENTOS_PROJETOS` e `CLICKUP_USER_ACTIVITY`, sem coleta de nomes, clientes ou conteúdo individual.^2

### 2.2 Matriz preliminar de decisão

Esta matriz avalia valor e sobreposição funcional, não popularidade. “Consolidar” significa retirar a entrada isolada do menu e preservar a capacidade dentro de uma área mais coerente.

| Capacidade atual | Valor provável | Problema atual | Decisão preliminar | Condição de revisão após 30 dias |
|---|---|---|---|---|
| Dashboard | Muito alto | Sobreposição com Executivo e Gestão | **Manter como início** | Deve concentrar saúde, prioridades e atalhos de ação |
| Executivo | Alto | Duplica síntese do Dashboard | **Consolidar no Dashboard** | Manter seção ou modo executivo se houver uso recorrente por liderança |
| Gestão / produtividade ClickUp | Alto para administração | Mistura governança técnica com gestão diária | **Mover para Administração / Saúde** | Manter destaque apenas para administradores e gestores |
| Acompanhamento | Alto, com uso comprovado | Adoção caiu de 96 registros em maio para 1 em agosto | **Manter e redesenhar o fluxo** | Investigar se queda veio de fricção, processo encerrado ou substituição por ClickUp |
| Diárias CMAX | Muito alto | É fonte da Bonificação, mas aparece como domínio separado | **Manter em Agenda e Fechamentos** | Medir consultas, filtros e correções manuais |
| Fechamento | Muito alto | Fragmentado de Fechamento de Projeto e Bonificação | **Consolidar em Fechamentos** | Preservar filas e decisões distintas como subabas |
| Fechamento de Projeto | Muito alto | Nome truncado e navegação isolada | **Consolidar em Fechamentos** | Medir tempo para concluir e taxa de reabertura |
| Projetos | Essencial | Tela central compete com várias visões derivadas | **Manter como base operacional** | Tornar a ficha de projeto o ponto de verdade e ação |
| Consultores | Médio/alto | É uma dimensão da carteira, não um fluxo completo | **Transformar em visão/filtro de Pessoas** | Manter entrada própria somente se houver recorrência e tarefas exclusivas |
| Vendedores | Médio | É uma dimensão da carteira e da origem comercial | **Transformar em visão/filtro de Pessoas** | Retirar a tela dedicada se o uso e decisões forem baixos |
| Avulso/Novo | Médio | Segmento de projeto, não domínio de navegação | **Transformar em filtro salvo de Projetos** | Manter atalho contextual se representar volume relevante |
| Logística | Médio/alto | Pode ser valiosa, mas está isolada do planejamento | **Manter em Agenda e Logística** | Retirar como tela autônoma se não gerar decisões ou correções |
| Prazos | Alto | Repete alertas e indicadores do Dashboard | **Consolidar em Pendências** | Manter filtros e regras; eliminar duplicação de tela |
| Alertas | Alto | Repete Prazos e sinais do Dashboard | **Consolidar em Pendências** | Medir resolução por alerta, não apenas visualização |
| Bonificação | Muito alto | Melhor padrão visual, porém dependente de três fontes | **Manter e usar como padrão de design** | Medir fechamento, impressão, divergências e retrabalho |

### 2.3 Regra para retirar uma função

Uma função somente deve ser descontinuada quando satisfizer simultaneamente quatro critérios durante uma janela mínima de 30 dias:

- menos de 5% dos usuários elegíveis a utilizam;
- não participa de um fluxo regulatório, financeiro, de auditoria ou de fechamento;
- existe outro caminho mais simples que entrega o mesmo resultado;
- sua retirada não aumenta trabalho manual, risco ou dependência de planilha.

Função pouco acessada, mas responsável por fechamento mensal, exceção crítica ou auditoria, deve ser simplificada e mantida. Função acessada muitas vezes porque falha ou exige repetição não é um sucesso de adoção; é retrabalho. A telemetria precisa distinguir ambos.

## 3. Diagnóstico da arquitetura e das integrações

### 3.1 Arquitetura atual

```text
Navegador / GitHub Pages
    ├── snapshot materializado via Google Sheets GViz
    ├── Apps Script por JSONP (fonte ALL e por mês)
    ├── leitura direta de 12 abas mensais como contingência
    ├── cache e diagnóstico no localStorage
    └── ações manuais de sincronização
             │
             ▼
Google Apps Script
    ├── ClickUp API + webhooks + triggers
    ├── CMAX API + cache + histórico
    ├── autenticação e usuários
    ├── acompanhamentos, decisões e bonificação
    └── gravação em várias abas do Google Sheets
```

Essa arquitetura contém boas proteções: snapshots completos são validados, cargas parciais não substituem a carteira íntegra, processos longos foram divididos em lotes e há retomada por cursor. A suíte publicada confirma invariantes de integridade, concorrência e preservação de dados; os 37 testes passaram na execução deste estudo.^3

O problema é que a resiliência foi construída acumulando caminhos. A carga principal tenta uma aba materializada, depois snapshot privado compactado, Apps Script `ALL`, doze chamadas mensais e cache local. Em caso de falha, o sistema preserva dados, mas a quantidade de rotas torna comportamento, diagnóstico e manutenção difíceis. A confiabilidade interna melhorou; a explicabilidade operacional não acompanhou.

### 3.2 Latência e volume da leitura pública

A aba materializada respondeu em aproximadamente 2,2 segundos e 198 KB no teste pontual. As abas mensais com dados chegaram a aproximadamente 1,53 MB e apresentaram latências individuais entre 2,99 e 11,86 segundos; janeiro, fevereiro e março responderam em 9,66 s, 11,86 s e 9,83 s, respectivamente.^4 Isso explica por que abrir doze abas diretamente no navegador é uma contingência cara e sujeita a timeouts, especialmente em rede corporativa ou móvel.

O Google recomenda minimizar chamadas a serviços externos, usar operações em lote e cache. Também limita execuções comuns de Apps Script a seis minutos, além de impor cotas de `URL Fetch` e tempo total de triggers.^5 O código atual já reage a esses limites com lotes, cache e jobs retomáveis; o próximo salto de qualidade é deixar de fazer o navegador orquestrar tantas fontes.

### 3.3 ClickUp

O ClickUp limita requisições por token conforme o plano; em Free, Unlimited e Business, o limite documentado é de 100 requisições por minuto, com erro HTTP 429 quando excedido.^6 O GEST CSI já contém tratamento para 429, atraso entre projetos, leitura resumida e fallbacks de lista para view. Isso reduz falhas, mas ainda existe risco quando sincronização geral, atividade de usuários, fechamento e consultas individuais competem pelo mesmo token.

O desenho de webhooks é parcialmente automático: eventos entram em uma fila de planilha e um trigger os processa. Entretanto, o receptor valida um token compartilhado na query string e não a assinatura criptográfica do webhook; se a propriedade de token estiver ausente, a função aceita a requisição. O próprio README reconhece a limitação de cabeçalhos do Apps Script e recomenda migrar o receptor para um backend capaz de validar `X-Signature`.^3 Webhooks devem continuar sendo a fonte de atualização incremental, mas com assinatura obrigatória, idempotência, reprocessamento e fila de erro.

### 3.4 CMAX

A integração CMAX já renova autenticação, materializa uma visão, grava histórico e usa cache de até seis horas. O principal risco não é ausência de automação, mas ausência de um contrato observável: o usuário não vê de forma única a última coleta, o período coberto, a quantidade esperada versus recebida e se o painel está mostrando dado atual, cache ou contingência. Como CMAX alimenta Diárias e Bonificação, sua saúde deve estar visível no fechamento e em uma Central de Integrações.

### 3.5 Segurança e governança

Foram identificados quatro riscos prioritários:

1. **Endpoint publicado indisponível no teste externo.** A URL de Apps Script embutida na versão pública retornou HTTP 404 em 14/09/2026. Isso precisa de validação E2E no navegador autenticado, pois pode indicar implantação removida, permissão ou redirecionamento incompatível com o cliente de teste.^4
2. **Planilha acessível diretamente.** As abas mensais e técnicas responderam por GViz sem tela de login. Mesmo que o painel tenha autenticação visual, isso permite contornar a interface e consultar a origem. A planilha não deve ser API pública de dados operacionais.^4
3. **Autorização inconsistente.** Na referência publicada, `syncProject`, `processDirty`, `setProjectClosingDecision`, `setConsultantSeniority` e algumas leituras sensíveis não aplicam `requireUser_()` ou `requireAdmin_()` no roteador principal. Rotas que alteram estado precisam de autorização uniforme e, idealmente, não devem usar GET/JSONP.^3
4. **Sessão em CacheService.** Tokens ficam em cache por até seis horas e são renovados por atividade. É funcional, mas não oferece o mesmo controle, auditoria e revogação de uma sessão persistida em backend dedicado.^3

Esses riscos justificam uma frente imediata antes de qualquer expansão funcional.

## 4. Por que o produto parece espalhado

### 4.1 Navegação orientada por relatórios, não por trabalho

As quinze entradas laterais refletem a ordem em que funcionalidades foram adicionadas: Dashboard, Executivo, Gestão, Acompanhamento, CMAX, dois fechamentos, Projetos, três cortes de carteira, Logística, Prazos, Alertas e Bonificação. Para o usuário, várias opções respondem partes da mesma pergunta.

Exemplos de sobreposição:

- Dashboard, Executivo e Gestão sintetizam situação da carteira;
- Projetos, Consultores, Vendedores e Avulso/Novo são cortes do mesmo conjunto;
- Prazos e Alertas expõem pendências que também aparecem no Dashboard;
- Fechamento, Fechamento de Projeto e Bonificação são etapas do mesmo ciclo financeiro;
- CMAX é simultaneamente agenda, evidência de execução e fonte de remuneração.

O resultado é alto custo de escolha. O usuário precisa conhecer a implementação interna para decidir onde procurar, em vez de escolher a tarefa que deseja concluir.

### 4.2 Nova arquitetura de informação

Reduzir de quinze para cinco áreas diminui em 67% as escolhas de primeiro nível, sem retirar capacidade.

| Nova área | Conteúdo consolidado | Pergunta respondida |
|---|---|---|
| **Visão Geral** | Dashboard + resumo executivo + saúde das integrações | O que exige atenção agora? |
| **Operação** | Projetos + Acompanhamento + Pendências (Prazos/Alertas) | O que precisa ser feito e por quem? |
| **Pessoas** | Consultores + Vendedores + capacidade/adoção ClickUp | Como carteira, execução e carga se distribuem? |
| **Agenda e Logística** | CMAX + calendário + mapa + formato/localização | Quando e onde o trabalho acontecerá? |
| **Fechamentos** | Marcos + Projetos + Bonificação + relatórios | O que precisa ser validado, pago e auditado? |

Administração, usuários, diagnóstico e configuração devem ficar em um menu de perfil, visível somente para quem administra. Avulso/Novo vira filtro salvo dentro de Projetos. Prazos e Alertas viram uma fila única com filtros, severidade, responsável, vencimento e ação recomendada.

### 4.3 Bonificação como sistema de design

A Bonificação é a melhor referência interna porque possui uma narrativa completa:

- cabeçalho com contexto e competência;
- total executivo destacado;
- agrupamento por natureza financeira;
- detalhamento por pessoa;
- justificativas e exceções;
- base administrativa auditável;
- modos de impressão e assinatura.

Esse padrão deve ser transformado em componentes reutilizáveis: `PageHeader`, `SummaryStrip`, `MetricGroup`, `ActionQueue`, `EvidenceTable`, `EmptyState`, `SourceStatus` e `ReportView`. O objetivo não é fazer todas as telas parecerem uma folha de pagamento; é reutilizar a hierarquia clara entre resumo, decisão, evidência e relatório.

### 4.4 Princípios visuais

- Um único tema padrão consistente; modo escuro permanece opcional, não como solução estética principal.
- Largura e espaçamento definidos por escala, evitando regras específicas acumuladas por página.
- No máximo seis indicadores na primeira dobra; demais métricas aparecem por aprofundamento.
- Cor reservada para estado e prioridade, não para decorar categorias.
- Um botão primário por contexto; sincronização e diagnóstico ficam fora da navegação principal.
- Tabelas extensas ganham cabeçalho fixo, colunas essenciais e painel lateral de detalhe.
- Toda métrica deve indicar fonte, competência, atualização e regra de cálculo.
- Estados vazio, carregando, degradado e erro devem ser componentes do sistema, não mensagens improvisadas.

## 5. O que significa excelência comparável ao Gestão de Atendimento

A análise visual e técnica da referência mostra cinco características transferíveis:

1. **Navegação em camadas.** O menu principal organiza Visão Geral, Operação, Pendências, Equipe, Resultados e Histórico; subabas aparecem apenas dentro do domínio escolhido.
2. **Saúde visível.** A interface exibe atualização automática, status “ao vivo”, última sincronização, contagem de falhas e painéis de saúde.
3. **Detalhamento contextual.** Métricas abrem explicações que mostram quais itens elevam a média e levam diretamente ao atendimento correspondente.
4. **Automação no servidor.** A configuração local possui Worker, observabilidade, crons a cada um e dois minutos, KV e D1.^7
5. **Cobertura de regressão.** Foram identificados 22 arquivos e 352 casos de teste abrangendo desempenho, estabilidade, persistência, concorrência, alertas, reconciliação e regras de negócio.^7

O Gestão de Atendimento também tem dívida técnica: seu HTML local passa de 5 MB e o Worker de 760 KB. Portanto, ele não deve ser copiado estruturalmente. O que merece ser replicado é o contrato operacional: atualização automática, saúde explícita, persistência confiável, drill-down acionável e testes de cenários reais.

## 6. Arquitetura-alvo

```text
Interface GEST CSI
    │  uma API autenticada, JSON, versionada
    ▼
Backend operacional
    ├── autenticação, papéis e auditoria
    ├── API de leitura materializada
    ├── comandos idempotentes
    ├── health checks e métricas
    └── fila de eventos e reprocessamento
          │             │             │
          ▼             ▼             ▼
       ClickUp         CMAX       Google Sheets
      webhooks       coleta       entrada/admin
      assinados      agendada     e conferência
          └─────────────┬─────────────┘
                        ▼
             snapshot atômico versionado
```

### Responsabilidade de cada camada

- **Interface:** renderizar, filtrar e enviar comandos; nunca conciliar doze fontes nem decidir qual fallback técnico promover.
- **Backend:** aplicar autenticação, limites, retries, circuit breaker, idempotência e regras de negócio.
- **Banco persistente:** guardar estado operacional, decisões, telemetria, sessões e histórico de sincronização.
- **Cache/snapshot:** entregar a leitura do painel rapidamente e continuar servindo a última versão íntegra durante indisponibilidade.
- **Google Sheets:** continuar útil para cadastro, conferência e exportação, mas deixar de ser endpoint público primário.
- **Webhooks e agendadores:** manter a base fresca sem depender de um usuário abrir a tela ou clicar em Sync.

Cloudflare Workers é uma opção natural porque a organização já opera esse modelo no Gestão de Atendimento, e a plataforma oferece métricas, logs e traces integrados.^8 A decisão tecnológica deve considerar governança e capacidade do time; o requisito obrigatório é o contrato operacional acima, não um fornecedor específico.

## 7. Telemetria necessária para decidir o que retirar

### 7.1 Eventos mínimos

| Evento | Campos principais | Decisão suportada |
|---|---|---|
| `session_started` | papel, versão, dispositivo pseudônimo | usuários ativos e versões em uso |
| `route_viewed` | área, subaba, origem da navegação | uso real de cada tela |
| `action_started` / `action_completed` | ação, duração, sucesso, motivo de falha | valor versus retrabalho |
| `filter_applied` | tela, tipo de filtro | filtros que merecem atalhos |
| `detail_opened` | entidade e origem | quais indicadores geram investigação |
| `decision_saved` | tipo, fluxo, tempo até concluir | eficiência de fechamentos e acompanhamento |
| `export_generated` | relatório, competência | valor de relatórios e Bonificação |
| `data_state_shown` | fonte, atual/cache/degradado, idade | impacto real das integrações |

Não registrar nomes de clientes, textos de acompanhamento, senhas, tokens ou conteúdo de tarefas. IDs operacionais devem ser pseudonimizados ou agregados. Retenção recomendada: eventos detalhados por 90 dias e agregados mensais por 13 meses.

### 7.2 Painel de adoção

Para cada capacidade, medir:

- usuários elegíveis e usuários ativos;
- sessões e dias distintos de uso;
- ações concluídas e taxa de erro;
- tempo mediano e p95 para concluir a tarefa;
- quantidade de cliques manuais de atualização;
- recorrência semanal e mensal;
- origem do acesso: menu, alerta, Dashboard ou link direto;
- resultado gerado: decisão, correção, exportação ou apenas consulta.

## 8. Roadmap priorizado

### Fase 0 — contenção e verdade operacional (0–5 dias)

| Ação | Resultado esperado | Critério de aceite |
|---|---|---|
| Definir repositório, branch e versão canônicos | Eliminar divergência entre local, remotos e produção | Uma origem documentada; versão visível no app e no deploy |
| Verificar/republicar o Apps Script | Restaurar API usada pelo painel | Health endpoint autenticável e E2E de login/leitura aprovado |
| Proteger todas as rotas mutáveis | Fechar alteração anônima e GET mutável | Matriz de autorização testada; zero comando sem papel adequado |
| Restringir acesso direto às planilhas | Impedir contorno da autenticação | Dados operacionais servidos somente pelo backend |
| Criar Central de Saúde mínima | Tornar fonte, idade e falha visíveis | ClickUp, CMAX, Sheets e snapshot com status e última execução |
| Congelar novas telas | Concentrar esforço em estabilidade | Nenhuma nova área antes de métricas e arquitetura aprovadas |

### Fase 1 — medição e simplificação (1–2 semanas)

- Instrumentar os eventos mínimos, com consentimento e minimização de dados.
- Implementar a navegação de cinco áreas por trás de flag de recurso.
- Aplicar componentes visuais derivados da Bonificação ao Dashboard, Operação e Fechamentos.
- Unificar Prazos e Alertas em uma fila acionável.
- Transformar Consultores, Vendedores e Avulso/Novo em visões da carteira.
- Manter URLs profundas para que favoritos e links existentes continuem funcionando.

### Fase 2 — automação confiável (2–6 semanas)

- Criar backend autenticado e versionado.
- Materializar snapshot único e atômico de Projetos.
- Receber webhooks ClickUp com assinatura obrigatória.
- Implementar fila persistente, idempotência, backoff e fila de erro.
- Agendar CMAX e reconciliações no servidor.
- Substituir JSONP e comandos GET por API JSON com métodos adequados.
- Centralizar métricas de latência, erro, volume, frescor e cobertura.

### Fase 3 — decisão de portfólio (após 30 dias medidos)

- Classificar cada capacidade em investir, manter, consolidar ou retirar.
- Remover inicialmente apenas entradas de navegação redundantes.
- Executar entrevistas curtas com três perfis: administração, coordenação e consultoria.
- Descontinuar funcionalidade somente após confirmar ausência de obrigação e alternativa equivalente.
- Publicar changelog e período de transição para qualquer retirada.

### Fase 4 — excelência contínua (6–12 semanas)

- Separar o HTML monolítico em módulos por domínio e componentes de design.
- Extrair regras de negócio do Apps Script para módulos testáveis.
- Elevar testes de integração, autorização, falha parcial, concorrência e recuperação.
- Criar homologação isolada com dados fictícios.
- Adotar liberação gradual, métricas antes/depois e rollback sem perda de dados.

## 9. Metas e indicadores de sucesso

| Dimensão | Meta inicial |
|---|---:|
| Disponibilidade da leitura | ≥ 99,5% ao mês |
| Cargas íntegras | ≥ 99,5% |
| Frescor ClickUp | 95% dos projetos atualizados em até 10 minutos após evento relevante |
| Frescor CMAX | competência e última coleta visíveis; atraso crítico alertado automaticamente |
| Ações manuais de Sync | redução de 80% em 60 dias |
| Erros não classificados | zero; toda falha com categoria, correlação e ação recomendada |
| Tempo p95 da tela inicial | ≤ 2,5 s com snapshot íntegro |
| Segurança | zero rota mutável sem autenticação e autorização |
| Navegação | cinco áreas de primeiro nível; tarefa crítica acessível em até dois níveis |
| Testes | toda regra financeira, mudança de status, retry e autorização com teste comportamental |
| Adoção | ≥ 70% dos usuários elegíveis usando semanalmente pelo menos um fluxo acionável |

## 10. Riscos e controles

| Risco | Probabilidade | Impacto | Controle |
|---|---|---|---|
| Retirar algo útil por falta de telemetria | Alta | Alto | Janela de 30 dias e entrevistas por perfil |
| Migração quebrar Bonificação ou fechamento | Média | Muito alto | Snapshot paralelo, reconciliação e testes de valores por competência |
| Continuar expondo planilha e comandos | Alta | Muito alto | Prioridade zero de segurança e backend autenticado |
| Copiar complexidade do Gestão de Atendimento | Média | Alto | Reutilizar padrões operacionais, não o monólito |
| Manter dois backends indefinidamente | Média | Alto | Migração por domínio com data de encerramento do caminho antigo |
| Telemetria coletar dados excessivos | Média | Alto | Pseudonimização, minimização, retenção e acesso restrito |
| Métricas melhorarem sem percepção do usuário | Média | Médio | Estado de saúde visível, mensagens simples e testes de tarefa com usuários |

## 11. Backlog recomendado por prioridade

### P0 — antes de redesenhar

1. Corrigir e monitorar o endpoint Apps Script publicado.
2. Uniformizar autenticação e autorização de todas as ações.
3. Deixar de expor dados operacionais diretamente por GViz.
4. Definir uma única origem canônica e um único snapshot íntegro.
5. Instrumentar uso por rota e ação.

### P1 — maior impacto percebido

1. Navegação com cinco áreas.
2. Dashboard inicial com prioridades, saúde e atalhos.
3. Prazos + Alertas em fila única.
4. Fechamento + Fechamento de Projeto + Bonificação em um ciclo único.
5. Centro de Integrações com ClickUp, CMAX e Sheets.
6. Padrão visual derivado da Bonificação.

### P2 — redução de manutenção

1. Backend dedicado e API JSON.
2. Webhook assinado, fila persistente e dead-letter.
3. Módulos de front-end e regras extraídas do Apps Script.
4. Suíte de integração e segurança.
5. Homologação isolada e liberação gradual.

### P3 — decisão de retirada

1. Avaliar Executivo como tela separada.
2. Avaliar telas dedicadas de Vendedores e Consultores.
3. Avaliar Logística como tela autônoma.
4. Avaliar Gestão ClickUp fora da administração.
5. Retirar somente funções que falharem nos quatro critérios da seção 2.3.

## Conclusão

O GEST CSI não está sem valor; está com sua complexidade exposta. As proteções recentes aumentaram muito a integridade das cargas, e o sistema já contém fluxos sofisticados de carteira, acompanhamento, ClickUp, CMAX e fechamento. A percepção de produto ultrapassado nasce da combinação de navegação fragmentada, sincronização visível, múltiplas contingências e ausência de uma saúde operacional clara.

A mudança decisiva é transformar o painel de um conjunto de relatórios e botões de sincronização em um produto orientado a tarefas, com dados materializados, automação no servidor e evidência de confiabilidade. Bonificação fornece o melhor vocabulário visual interno; Gestão de Atendimento fornece a melhor referência de contrato operacional. A evolução recomendada preserva o que já funciona, reduz a superfície aparente e cria dados objetivos para decidir o que realmente pode ser retirado.

## Fontes

1. `Levantamento_Tecnico_GEST_CSI_GEST_ATENDIMENTO_CHURN_Revisado.pdf`, seções 3, 7, 8 e Anexo Técnico; documento local, emitido em 15/08/2026.
2. Google Sheets, abas técnicas `PANEL_UPDATE_HISTORY`, `ACOMPANHAMENTOS_PROJETOS` e `CLICKUP_USER_ACTIVITY`; consultas agregadas realizadas em 14/09/2026, sem coleta de conteúdo individual. Acesso privado/local documentado neste estudo.
3. Repositório GEST CSI: referência `gestcsi/main` em `11e7c2f`, arquivos `index.html`, `apps_script/ClickUpSync.gs`, `apps_script/README.md` e `tests/*.test.mjs`; inspeção e execução local em 14/09/2026. Página pública verificada em [GEST CSI](https://reinaldobueno-cyber.github.io/GESTCSI/).
4. Verificações HTTP de disponibilidade, latência e tamanho realizadas em 14/09/2026 contra a página pública, o endpoint Apps Script embutido e consultas GViz das abas configuradas. Testes pontuais, não equivalentes a monitoramento contínuo.
5. Google, [Apps Script best practices](https://developers.google.com/apps-script/guides/support/best-practices) e [Quotas for Google Services](https://developers.google.com/apps-script/guides/services/quotas), consultados em 14/09/2026.
6. ClickUp, [Rate Limits](https://developer.clickup.com/docs/rate-limits) e [Webhooks](https://developer.clickup.com/docs/webhooks), consultados em 14/09/2026.
7. Repositório local Gestão de Atendimento, referência de trabalho em `e12006a`, arquivos `wrangler.jsonc`, `src/worker.js`, `index.html`, `redesign.css` e `test/*.test.js`; inspeção em 14/09/2026. O worktree continha alterações não publicadas, portanto a comparação descreve a referência local, não garante o estado exato de produção.
8. Cloudflare, [Workers Observability](https://developers.cloudflare.com/workers/observability/), consultado em 14/09/2026.
