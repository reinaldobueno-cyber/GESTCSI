# Plano de evolução do GEST CSI

## Identificação e regra de status

**Plano:** `GESTCSI-BACKEND-V1-20260914`

Este é um plano vivo. Cada entrega contém um campo **Critério de OK**. O status somente deve mudar para `[x] OK` depois que o critério for comprovado e a evidência for registrada no próprio item.

- `[x] OK` — concluído, testado e com evidência registrada.
- `[~] EM ANDAMENTO` — implementação iniciada, ainda sem aceite completo.
- `[ ] PENDENTE` — ainda não iniciado ou sem evidência suficiente.
- `[!] BLOQUEADO` — depende de decisão, acesso ou correção externa explicitada.

Não é correto marcar uma implementação futura como OK apenas porque foi planejada. Neste documento, as entregas de diagnóstico e desenho já realizadas estão marcadas como OK; as entregas de produto permanecem pendentes até sua execução e validação.

## Resultado esperado

Ao final deste plano, o GEST CSI deverá:

- abrir rapidamente usando um snapshot íntegro e versionado;
- atualizar ClickUp e CMAX automaticamente no servidor;
- mostrar claramente a saúde e a idade de cada fonte;
- não expor planilhas operacionais como API pública;
- não permitir comandos sem autenticação e autorização;
- reduzir as quinze entradas atuais para cinco áreas de trabalho;
- manter Bonificação e fechamentos auditáveis durante toda a migração;
- registrar uso real para decidir, com dados, quais capacidades manter ou retirar;
- possuir homologação, testes e publicação controlada comparáveis ao Gestão de Atendimento.

## Tudo que foi percebido

### 1. Pontos fortes existentes

1. O sistema resolve problemas reais de operação, não é apenas um painel visual.
2. A carteira possui proteções contra cargas parciais e duplicidades.
3. Existe snapshot local e preservação da última base íntegra.
4. O ClickUp já possui sincronização em lote, execução em segundo plano, cursor de retomada e tratamentos de erros recuperáveis.
5. O CMAX já possui renovação automática de autenticação, histórico materializado e cache.
6. A Bonificação possui boa hierarquia, rastreabilidade, justificativas, relatórios individuais, impressão e assinatura.
7. Métricas importantes permitem abrir detalhes e auditar os projetos usados no cálculo.
8. A versão publicada possui 37 testes que cobrem integridade da carteira, geolocalização, sessões, ClickUp, CMAX e regras de fechamento; todos passaram durante o estudo.
9. O processo de publicação já contém verificação e proteção contra rollback que sobrescreva uma publicação concorrente.
10. A confiabilidade recente das cargas melhorou: 343 de 344 registros desde 15/08/2026 terminaram como `ok`.

### 2. Uso e adoção

1. O painel possui uso recorrente: há 2.604 registros de atualização desde abril.
2. Há 305 cargas manuais registradas, sinal de que o usuário ainda sente necessidade de intervir.
3. Há 2.235 cargas automáticas e 64 reconciliações do monitor mais recente; a automação existe, mas não elimina a percepção de dependência manual.
4. Acompanhamento foi efetivamente usado 161 vezes.
5. O uso persistido de Acompanhamento caiu de 96 registros em maio para 43 em junho, 21 em julho e 1 em agosto.
6. A queda de Acompanhamento pode significar abandono, fricção, mudança de processo ou substituição por ClickUp; a evidência atual não permite escolher entre essas hipóteses.
7. O painel não registra centralmente `route_viewed`, cliques ou conclusão de tarefas. Portanto, não existe ranking confiável de uso das quinze telas.
8. Os 119 rótulos encontrados no histórico não devem ser interpretados como 119 usuários únicos, pois podem incluir dispositivos e navegadores.
9. Antes de retirar qualquer capacidade, é necessário medir usuários elegíveis, usuários ativos, ações concluídas, recorrência e falhas por 30 dias.

### 3. Experiência e design

1. A navegação foi crescendo por adição de relatórios e hoje possui quinze entradas de primeiro nível.
2. Dashboard, Executivo e Gestão respondem partes semelhantes da pergunta “como está a operação?”.
3. Projetos, Consultores, Vendedores e Avulso/Novo são cortes da mesma carteira.
4. Prazos e Alertas exibem pendências que também aparecem no Dashboard.
5. Fechamento, Fechamento de Projeto e Bonificação pertencem ao mesmo ciclo de validação e pagamento.
6. Diárias CMAX é simultaneamente agenda, evidência de execução e origem financeira, mas aparece como domínio isolado.
7. A quantidade de opções obriga o usuário a conhecer a arquitetura interna.
8. Há muitos controles globais no topo: tema, histórico, usuários, sincronização, atualização e perfil.
9. O produto mistura comandos administrativos com trabalho operacional.
10. A Bonificação é a tela mais coerente porque conduz do resumo para composição, exceção, evidência e relatório.
11. O modo escuro do Gestão de Atendimento contribui para identidade, mas não é a causa principal da percepção de excelência.
12. O maior diferencial da referência é mostrar atualização automática, saúde, falhas, última execução e detalhes acionáveis.

### 4. Dados e integrações

1. O navegador ainda participa da orquestração de fontes.
2. A carga pode percorrer aba materializada, snapshot compactado, Apps Script `ALL`, doze chamadas mensais, leitura GViz direta e cache local.
3. Essas contingências protegem dados, mas tornam o comportamento difícil de prever e diagnosticar.
4. A aba materializada respondeu em cerca de 2,2 segundos no teste pontual.
5. Abas mensais chegaram a 1,53 MB e demoraram até 11,86 segundos individualmente.
6. A leitura direta de doze abas não é um bom caminho primário nem uma contingência leve.
7. O Apps Script trabalha próximo de limitações naturais para processamento longo e chamadas externas; o próprio código precisou criar lotes, cursores, timeouts e retomadas.
8. Sync geral, atividade de usuários, fechamento, consultas individuais e webhook competem pelo mesmo token e cota do ClickUp.
9. O ClickUp aplica limite por token e pode responder 429; o sistema trata o erro, mas falta uma fila central que controle concorrência e backoff.
10. A fila atual de webhook usa planilha como armazenamento e trigger temporal como consumidor.
11. O webhook usa token compartilhado na URL e não valida obrigatoriamente a assinatura oficial.
12. Se o token compartilhado não estiver configurado, a função atual aceita a requisição.
13. A saúde do CMAX está distribuída entre cache, histórico e mensagens de tela, sem contrato único de atualidade e cobertura.
14. As regras financeiras dependem de ClickUp, CMAX e decisões persistidas; essa composição precisa de snapshot e versionamento de regra.
15. A arquitetura escreve resultados em várias abas e, em alguns pontos, célula a célula, aumentando latência e risco de estado intermediário.

### 5. Segurança e privacidade

1. O endpoint Apps Script embutido na produção retornou HTTP 404 no teste externo de 14/09/2026 e precisa de validação E2E.
2. As abas mensais e técnicas responderam por GViz sem tela de autenticação.
3. A autenticação do painel protege a interface, mas não protege uma origem pública acessada diretamente.
4. Há comandos no roteador publicado sem aplicação uniforme de `requireUser_()` ou `requireAdmin_()`.
5. Entre os casos observados estão sincronização individual, processamento da fila, decisão de fechamento e alteração de senioridade.
6. Comandos que alteram estado são acionados por GET/JSONP, o que dificulta proteção adequada contra repetição e requisições indevidas.
7. Tokens de sessão ficam em CacheService e no `localStorage` do navegador.
8. Falta uma trilha central e imutável de quem executou cada comando administrativo.
9. A telemetria futura não deve registrar nomes de clientes, textos de tarefas, observações, credenciais ou tokens.

### 6. Código e manutenção

1. A versão publicada concentra cerca de 18,3 mil linhas e 754 funções no `index.html`.
2. O Apps Script publicado concentra cerca de 10,6 mil linhas e 462 funções.
3. O HTML mistura estrutura, estilos, estado, regras, integrações e renderização.
4. O Apps Script mistura API pública, autenticação, ClickUp, CMAX, usuários, fechamento, bonificação e persistência.
5. Há 173 handlers inline no HTML publicado, aumentando acoplamento entre marcação e lógica.
6. O número de caminhos de fallback elevou o custo de regressão.
7. A raiz local aberta não é a produção: estava 492 commits atrás e 50 à frente de `gestcsi/main`.
8. Há dois remotos com histórias diferentes, `gestcsi` e `origin`, aumentando risco de publicar a origem errada.
9. O worktree raiz contém alterações e muitos artefatos não rastreados; não deve ser usado diretamente para a migração.
10. O Gestão de Atendimento também possui arquivos grandes; ele deve servir de referência operacional, não de modelo para copiar monólitos.

### 7. Operação e governança

1. A versão do sistema está visível no código, mas a diferença entre local, remoto e publicado não é resolvida automaticamente.
2. O histórico de atualização mede carga de dados, não saúde completa do produto.
3. Diagnósticos de desempenho ficam principalmente no navegador e não formam uma visão central.
4. Não existe SLO formal para disponibilidade, frescor, latência ou erro.
5. Não existe um catálogo único de fontes, responsáveis, credenciais, periodicidade e regra de contingência.
6. O botão Sync ClickUp no topo faz o usuário sentir que a integridade depende dele.
7. Uma falha de integração precisa gerar alerta e reprocessamento automático, não depender de descoberta visual.
8. A excelência percebida exige que o sistema informe “atual”, “atualizando”, “degradado” ou “indisponível”, sempre com horário e efeito sobre os números.

## Decisões de arquitetura

### Componentes do primeiro ciclo

| Componente | Uso no GEST CSI | Decisão |
|---|---|---|
| Cloudflare Worker | API autenticada, comandos, integrações e health | Adotar |
| D1 | usuários, sessões, auditoria, jobs, decisões, acompanhamentos e metadados de snapshot | Adotar como estado transacional |
| KV | último snapshot publicado, cache de leitura e configuração não sensível | Adotar como cache, nunca como livro contábil |
| Cloudflare Queue | ClickUp, CMAX, webhooks, retry e dead-letter | Adotar |
| Cron Triggers | reconciliação, atualização de segurança e verificação de jobs presos | Adotar |
| Workers Observability | logs estruturados, erros, métricas e correlação | Adotar |
| R2 | arquivos históricos e exportações volumosas | Adiar; só incluir quando houver necessidade real |
| Durable Objects | coordenação fortemente consistente por entidade | Não usar inicialmente; D1 + Queue atende o escopo atual |
| Google Sheets | entrada administrativa, conferência e exportação | Manter, mas retirar da leitura pública do navegador |
| Apps Script | ponte temporária durante migração | Descontinuar gradualmente como backend principal |

### Princípios obrigatórios

- Uma fonte canônica por domínio.
- Toda resposta informa `schemaVersion`, `generatedAt`, `source` e `freshness`.
- Escritas são autenticadas, autorizadas, auditadas e idempotentes.
- Nenhum comando mutável usa GET.
- Falha parcial nunca publica snapshot híbrido.
- O último snapshot íntegro permanece disponível durante falha de origem.
- Retry possui limite, backoff e fila de erro.
- Credenciais ficam em secrets do ambiente, nunca no código, URL ou planilha pública.
- Toda entrega possui teste automatizado e evidência de aceite.

## Contrato inicial da API

| Método e rota | Finalidade | Papel mínimo |
|---|---|---|
| `GET /api/v1/health` | Estado resumido sem dados operacionais | Público ou monitor |
| `GET /api/v1/health/integrations` | Saúde detalhada de Sheets, ClickUp e CMAX | Gestor |
| `POST /api/v1/auth/login` | Criar sessão segura | Público |
| `POST /api/v1/auth/logout` | Revogar sessão | Autenticado |
| `GET /api/v1/me` | Validar sessão e papel | Autenticado |
| `GET /api/v1/snapshot` | Snapshot íntegro e versionado do painel | Autenticado |
| `GET /api/v1/projects` | Consulta paginada e filtrada | Autenticado |
| `GET /api/v1/projects/:id` | Ficha e evidências do projeto | Autenticado |
| `POST /api/v1/projects/:id/followups` | Registrar acompanhamento | Autenticado |
| `PATCH /api/v1/followups/:id` | Atualizar acompanhamento/kanban | Autenticado |
| `POST /api/v1/closing/milestones/:id/decision` | Decisão de marco | Gestor/admin |
| `POST /api/v1/closing/projects/:id/decision` | Decisão de fechamento | Gestor/admin |
| `GET /api/v1/bonus/:competence` | Consolidado auditável | Gestor/admin |
| `POST /api/v1/sync/clickup` | Enfileirar sync administrativo | Admin |
| `POST /api/v1/sync/cmax` | Enfileirar sync administrativo | Admin |
| `GET /api/v1/jobs/:id` | Progresso de job | Papel que iniciou/admin |
| `POST /api/v1/webhooks/clickup` | Receber evento assinado | Assinatura ClickUp |
| `POST /api/v1/telemetry/events` | Uso pseudonimizado em lote | Autenticado |

## Modelo mínimo de dados

### D1

- `users`, `sessions`, `roles` e `user_role_assignments`;
- `audit_events` para comandos e mudanças administrativas;
- `integration_runs`, `integration_health` e `job_attempts`;
- `projects`, `project_sources` e `project_snapshot_versions`;
- `project_followups` e `project_kanban_state`;
- `milestone_closing_decisions` e `project_closing_decisions`;
- `consultant_compensation` e `bonus_indications`;
- `telemetry_daily` e, se necessário, eventos detalhados com retenção curta.

### KV

- `snapshot:current`;
- `snapshot:manifest`;
- caches de leitura por competência;
- flags de recurso e configuração pública versionada.

### Queue

- `clickup-events`;
- `clickup-full-sync`;
- `cmax-sync`;
- `snapshot-build`;
- `telemetry-ingest`;
- dead-letter por integração.

## Plano de entregas com aceite

### Etapa A — diagnóstico e preparação

#### E00 — estudo técnico e inventário

- **Status:** [x] OK
- **Entrega:** diagnóstico de uso, design, integrações, segurança, código, governança e comparação com Gestão de Atendimento.
- **Critério de OK:** relatório e plano disponíveis no repositório, com fontes e limitações explícitas.
- **Evidência:** `ESTUDO_APROFUNDADO_GEST_CSI_2026-09.md` e este documento.

#### E01 — origem canônica de desenvolvimento

- **Status:** [x] OK
- **Entrega:** definir remoto, branch, SHA publicado e worktree exclusivo para a migração.
- **Critério de OK:** documento registra uma única origem; worktree nasce de `gestcsi/main`; `git status` está limpo; nenhuma mudança do worktree raiz foi copiada implicitamente.
- **Testes:** conferir SHA local, remoto e artefato público.
- **Risco:** publicar a história errada por confusão entre `origin` e `gestcsi`.
- **Evidência:** [`ORIGEM_CANONICA_GEST_CSI.md`](./ORIGEM_CANONICA_GEST_CSI.md) e [`VALIDACAO_ESTABILIDADE_E01.md`](./VALIDACAO_ESTABILIDADE_E01.md); branch `codex/backend-evolution`; worktree `.worktrees/backend-evolution`; SHA-base `11e7c2f`; 37/37 testes aprovados; smoke local e artefato público HTTP 200 na versão `2026-09-11-clickup-activity-v4`.

#### E02 — baseline funcional e financeiro

- **Status:** [x] OK
- **Entrega:** capturar resultados atuais de carteira, CMAX, marcos, fechamento de projeto e Bonificação em competências de referência anonimizadas.
- **Critério de OK:** arquivo de baseline contém totais e hashes; 37 testes atuais passam; divergências conhecidas estão registradas.
- **Testes:** suíte atual, smoke E2E e reconciliação de pelo menos duas competências.
- **Evidência:** [`BASELINE_FUNCIONAL_FINANCEIRA_E02.md`](./BASELINE_FUNCIONAL_FINANCEIRA_E02.md) e [`baselines/e02-2026-09-14.json`](./baselines/e02-2026-09-14.json); JUL/2026 e AGO/2026 reconciliados; SHA agregado `7dabb942...e014`; 48/48 testes aprovados, incluindo os 37 anteriores, validação de privacidade e smoke HTTP.

#### E03 — matriz de fontes e responsáveis

- **Status:** [ ] PENDENTE
- **Entrega:** catálogo de Sheets, ClickUp, CMAX, propriedades, tokens, periodicidade, owner, contingência e dados consumidores.
- **Critério de OK:** 100% das métricas do painel possuem fonte, owner, SLA de frescor e regra de fallback.
- **Testes:** revisão por responsável operacional e técnico.

### Etapa B — segurança e observabilidade imediatas

#### E04 — health check do Apps Script atual

- **Status:** [ ] PENDENTE
- **Entrega:** endpoint mínimo de saúde e verificação da implantação usada em produção.
- **Critério de OK:** responde 200 em monitor externo; informa versão sem expor dados; alerta após falhas consecutivas.
- **Testes:** público, autenticado, timeout e implantação removida.

#### E05 — autorização uniforme no legado

- **Status:** [ ] PENDENTE
- **Entrega:** proteger todas as rotas de leitura sensível e mutação antes da migração.
- **Critério de OK:** inventário de rotas revisado; zero comando mutável anônimo; papéis testados; GET não altera estado.
- **Testes:** matriz anônimo/consultor/gestor/admin por rota.

#### E06 — proteção da origem Google Sheets

- **Status:** [ ] PENDENTE
- **Entrega:** retirar dependência do navegador em abas operacionais públicas.
- **Critério de OK:** painel funciona com planilha privada; nenhuma aba com cliente ou operação responde anonimamente por GViz.
- **Testes:** tentativa anônima, painel autenticado e contingência.

#### E07 — telemetria mínima

- **Status:** [ ] PENDENTE
- **Entrega:** eventos de sessão, rota, ação, resultado, duração e estado de dados.
- **Critério de OK:** eventos aparecem agregados por tela e papel; nenhuma PII ou conteúdo operacional é coletado; retenção definida.
- **Testes:** payload permitido, payload rejeitado, batching e indisponibilidade da telemetria sem quebrar o painel.

#### E08 — Central de Saúde v1

- **Status:** [ ] PENDENTE
- **Entrega:** estado visível de Sheets, Apps Script, ClickUp, CMAX e snapshot.
- **Critério de OK:** cada fonte mostra estado, última execução, última base válida, idade, cobertura, erro e próxima ação.
- **Testes:** saudável, atrasado, degradado, falha e recuperado.

### Etapa C — fundação do backend

#### E09 — Worker e ambientes isolados

- **Status:** [ ] PENDENTE
- **Entrega:** projeto Worker com desenvolvimento, homologação e produção separados.
- **Critério de OK:** nomes, bindings, secrets e bancos não se misturam; homologação usa somente dados fictícios ou anonimizados; observabilidade está habilitada.
- **Testes:** smoke local, homologação e prova de isolamento.

#### E10 — esquema D1 e migrações

- **Status:** [ ] PENDENTE
- **Entrega:** tabelas mínimas, índices, constraints, migrações versionadas e seed fictício.
- **Critério de OK:** migração sobe do zero, reaplica sem dano, faz rollback lógico documentado e rejeita duplicidades críticas.
- **Testes:** banco vazio, upgrade, concorrência e constraints.

#### E11 — autenticação e sessões

- **Status:** [ ] PENDENTE
- **Entrega:** login, logout, sessão persistida, revogação e papéis.
- **Critério de OK:** cookie `HttpOnly`, `Secure` e `SameSite`; senha não trafega nem persiste em claro; logout revoga; alteração de papel invalida acesso incompatível.
- **Testes:** força bruta limitada, expiração, revogação, CSRF e matriz de papéis.

#### E12 — auditoria de comandos

- **Status:** [ ] PENDENTE
- **Entrega:** trilha para toda mutação administrativa e financeira.
- **Critério de OK:** ator, ação, entidade, antes/depois permitido, horário, correlação e resultado ficam registrados; segredos e textos sensíveis são omitidos.
- **Testes:** sucesso, falha, retry e concorrência.

#### E13 — API de snapshot v1

- **Status:** [ ] PENDENTE
- **Entrega:** leitura rápida, íntegra, versionada e autenticada.
- **Critério de OK:** contrato contém manifesto por mês, total, hash, geração, fonte e frescor; snapshot parcial nunca substitui o atual.
- **Testes:** íntegro, duplicado, parcial, versão antiga e indisponibilidade de origem.

### Etapa D — integrações automáticas

#### E14 — pipeline Google Sheets

- **Status:** [ ] PENDENTE
- **Entrega:** backend lê Sheets, normaliza, valida e publica snapshot atômico.
- **Critério de OK:** navegador realiza uma leitura de snapshot; doze abas não são consultadas pelo cliente; divergência gera alerta sem publicação parcial.
- **Testes:** nova linha, aba lenta, aba ausente, cabeçalho alterado e duplicidade.

#### E15 — ClickUp incremental

- **Status:** [ ] PENDENTE
- **Entrega:** webhook assinado, Queue, consumidor idempotente e atualização por projeto.
- **Critério de OK:** evento válido atualiza o projeto uma vez; assinatura inválida é rejeitada; duplicata não duplica efeito; falha faz retry e depois dead-letter.
- **Testes:** assinatura, duplicidade, ordem invertida, 429, 5xx e job preso.

#### E16 — ClickUp reconciliação completa

- **Status:** [ ] PENDENTE
- **Entrega:** Cron de segurança para corrigir eventos perdidos e inventário divergente.
- **Critério de OK:** cobertura esperada versus recebida é medida; reconciliação não compete descontroladamente com eventos; limite de API é respeitado.
- **Testes:** evento perdido, cursor retomado, limite 429 e snapshot preservado.

#### E17 — CMAX automático

- **Status:** [ ] PENDENTE
- **Entrega:** coleta agendada, renovação de token, histórico e materialização.
- **Critério de OK:** competência, cobertura, última coleta e origem ficam visíveis; resposta vazia nunca zera histórico válido; falha é reprocessada automaticamente.
- **Testes:** credencial expirada, resposta vazia, resposta parcial, duplicidade e virada de competência.

#### E18 — motor de snapshot consolidado

- **Status:** [ ] PENDENTE
- **Entrega:** unir carteira, ClickUp, CMAX, acompanhamentos e decisões em leitura consistente.
- **Critério de OK:** toda publicação possui versão de regra e hashes das fontes; Bonificação reconcilia com o baseline; não existe janela de leitura vazia.
- **Testes:** concorrência, falha de uma fonte, reprocessamento e comparação financeira.

### Etapa E — migração funcional

#### E19 — leitura paralela e comparador

- **Status:** [ ] PENDENTE
- **Entrega:** frontend ou rotina de homologação compara legado e backend sem trocar a origem exibida.
- **Critério de OK:** 30 dias de comparação ou volume equivalente; divergências classificadas; zero divergência financeira sem explicação.
- **Testes:** carteira, prazos, alertas, CMAX, marcos, projetos e bônus.

#### E20 — acompanhamentos e decisões no backend

- **Status:** [ ] PENDENTE
- **Entrega:** migrar gravações de acompanhamento, kanban, marcos e fechamento.
- **Critério de OK:** escrita idempotente, auditoria, leitura consistente e exportação de conferência para Sheets.
- **Testes:** edição concorrente, repetição de clique, autorização e histórico.

#### E21 — Bonificação no backend

- **Status:** [ ] PENDENTE
- **Entrega:** cálculo e relatório usando snapshot versionado e regras testáveis.
- **Critério de OK:** valores batem com competências de baseline; justificativas, assinaturas e impressão são preservadas; cada número aponta para evidência.
- **Testes:** diárias parciais, especiais, marcos aprovados/reprovados, fechamento de projeto, indicação e competência.

#### E22 — corte do Apps Script como API principal

- **Status:** [ ] PENDENTE
- **Entrega:** frontend usa somente a API nova; Apps Script permanece temporariamente apenas como ponte administrativa necessária.
- **Critério de OK:** tráfego do painel para Apps Script e GViz é zero; monitoramento da API nova cumpre SLO por sete dias; rollback de leitura foi ensaiado sem perda.
- **Testes:** E2E, falha de ClickUp, falha de CMAX, API degradada e snapshot anterior.

### Etapa F — produto e design

#### E23 — navegação de cinco áreas

- **Status:** [ ] PENDENTE
- **Entrega:** Visão Geral, Operação, Pessoas, Agenda e Logística, Fechamentos.
- **Critério de OK:** quinze capacidades preservadas; no máximo cinco opções principais; URLs antigas redirecionam; tarefas críticas ficam a até dois níveis.
- **Testes:** desktop, notebook, tablet, teclado, histórico do navegador e permissões.

#### E24 — sistema visual baseado em Bonificação

- **Status:** [ ] PENDENTE
- **Entrega:** componentes compartilhados de cabeçalho, resumo, grupos de métricas, fila, evidências, estados e relatórios.
- **Critério de OK:** Dashboard, Operação e Fechamentos usam os mesmos componentes e escala visual; não há regressão na impressão.
- **Testes:** claro/escuro, estados longos, vazio, erro, carregando, impressão A4 e contraste.

#### E25 — fila unificada de Pendências

- **Status:** [ ] PENDENTE
- **Entrega:** juntar Prazos e Alertas com severidade, responsável, vencimento e ação.
- **Critério de OK:** nenhuma regra desaparece; duplicatas são consolidadas; cada item pode ser aberto e resolvido a partir da fila.
- **Testes:** filtros, contagens, resolução, permissão e retorno ao contexto.

#### E26 — carteira orientada por visões

- **Status:** [ ] PENDENTE
- **Entrega:** Projetos como base, com visões salvas para Consultores, Vendedores e Avulso/Novo.
- **Critério de OK:** todos os totais atuais são reproduzidos; usuário alterna visão sem trocar de domínio; filtros podem ser compartilhados por URL.
- **Testes:** contagens, busca, filtros combinados e links profundos.

### Etapa G — medição e retirada

#### E27 — painel de adoção por capacidade

- **Status:** [ ] PENDENTE
- **Entrega:** usuários elegíveis/ativos, sessões, ações, conclusão, erro, duração e origem por módulo.
- **Critério de OK:** 30 dias íntegros, cobertura conhecida e dados suficientes para classificar cada capacidade.
- **Testes:** agregação, fuso, anonimização, retenção e papéis.

#### E28 — decisão manter/consolidar/retirar

- **Status:** [ ] PENDENTE
- **Entrega:** decisão executiva baseada em telemetria, obrigação, alternativa e entrevistas.
- **Critério de OK:** cada capacidade recebe decisão, evidência, impacto, comunicação e plano de reversão.
- **Regra:** retirar apenas se uso for inferior a 5% dos elegíveis, não houver obrigação/auditoria, houver alternativa equivalente e não aumentar trabalho manual.

#### E29 — desativação controlada do legado

- **Status:** [ ] PENDENTE
- **Entrega:** remover rotas, triggers, planilhas públicas e botões antigos somente após corte comprovado.
- **Critério de OK:** dependências zeradas, backup lógico validado, observação por 30 dias, documentação atualizada e nenhum consumidor conhecido.
- **Testes:** busca de dependências, tráfego, jobs, exportações e recuperação documentada.

## Sequência de migração sem parada

```text
1. Medir e proteger o legado
2. Criar backend isolado
3. Espelhar dados sem alterar produção
4. Comparar backend e legado
5. Migrar leitura para grupos pequenos
6. Migrar escritas por domínio
7. Migrar Bonificação por último
8. Observar SLOs
9. Encerrar caminhos antigos
10. Decidir retiradas por telemetria
```

Bonificação deve ser a última escrita migrada, porque reúne maior risco financeiro. A leitura pode ser espelhada cedo; decisões e cálculo somente mudam de origem depois da reconciliação completa.

## Ordem prática dos primeiros ciclos

### Ciclo 1 — segurança e base confiável

E01, E02, E03, E04, E05 e E06.

### Ciclo 2 — fundação observável

E07, E08, E09, E10, E11, E12 e E13.

### Ciclo 3 — integração automática

E14, E15, E16, E17 e E18.

### Ciclo 4 — migração de operação

E19, E20, E21 e E22.

### Ciclo 5 — experiência e simplificação

E23, E24, E25 e E26.

### Ciclo 6 — decisão de portfólio

E27, E28 e E29.

## SLOs de aceite global

| Indicador | Meta para considerar o programa OK |
|---|---:|
| Disponibilidade mensal da API de leitura | ≥ 99,5% |
| Snapshot íntegro | ≥ 99,5% das publicações |
| p95 da tela inicial com snapshot | ≤ 2,5 segundos |
| Projetos atualizados após evento ClickUp | 95% em até 10 minutos |
| Falha de integração não classificada | 0 |
| Comando mutável sem autenticação/autorização | 0 |
| Dependência de Sync manual | redução de 80% |
| Divergência financeira sem explicação | 0 |
| Escolhas no primeiro nível | 5 |
| Fluxo crítico | alcançável em até 2 níveis |

## Evidência obrigatória para marcar uma entrega como OK

Cada item concluído deve receber, logo abaixo do status:

```text
- Evidência de código: <commit/arquivos>
- Evidência de teste: <comando e resultado>
- Evidência funcional: <cenário validado>
- Evidência operacional: <métrica/log/health>
- Limitações restantes: <lista ou nenhuma>
- Validado em: <data e ambiente>
```

Sem esses seis registros, a entrega permanece pendente ou em andamento.

## Fontes técnicas

- [Estudo aprofundado do GEST CSI](./ESTUDO_APROFUNDADO_GEST_CSI_2026-09.md).
- [Cloudflare Workers](https://developers.cloudflare.com/workers/).
- [Cloudflare D1](https://developers.cloudflare.com/d1/).
- [Cloudflare KV](https://developers.cloudflare.com/kv/).
- [Cloudflare Queues](https://developers.cloudflare.com/queues/).
- [Cron Triggers](https://developers.cloudflare.com/workers/configuration/cron-triggers/).
- [Workers Observability](https://developers.cloudflare.com/workers/observability/).
- [Google Apps Script best practices](https://developers.google.com/apps-script/guides/support/best-practices).
- [Google Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas).
- [ClickUp API rate limits](https://developer.clickup.com/docs/rate-limits).
- [ClickUp webhooks](https://developer.clickup.com/docs/webhooks).
