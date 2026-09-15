# E05 — auditoria de autorização do roteador legado (15/09/2026)

Status: em execução na branch `codex/backend-evolution`; **não implantado**. Esta matriz descreve a guarda de homologação, não a segurança do endpoint de produção @286.

O catálogo E03 contém 56 ações HTTP no `doGet`; há ainda entrada sem `action` com `mes`, aliases `log_update=1` e `history=1`, além do webhook `doPost` e funções internas de menu/trigger. A guarda E05 é aplicada antes do despacho HTTP, com política default-deny para ações desconhecidas. `health` e `login` permanecem públicos; todas as demais ações conhecidas exigem sessão; administração de usuários, bonificações e jobs exige admin. Consultor e coordenador são papéis autenticados, mas coordenador não recebe privilégio de admin.

| Papel mínimo | Ações HTTP |
| --- | --- |
| Público | `health`, `login`, consulta sem `action` e sem parâmetros sensíveis |
| Autenticado | `getMonthlyProjects`, entrada com `mes`, `getClickUpInventory`, `getClickUpMilestoneClosing`, `diagnoseProjectClosing`, `getProjectClosingDecisions`, `getProjectClosingCandidates`, `getProjectClosingSyncStatus`, `getCmaxDailyEvents`, `getConsultantCompensation`, `getCmaxDailyHistoryStatus`, `logPanelUpdate`, `getPanelUpdateHistory`, `me`, `logProjectFollowup`, `getProjectFollowups`, `setProjectFollowupStatus`, `setProjectKanbanStage` |
| Admin | `syncProject`, `syncAll`, `startProjectSyncBackground`, `getProjectSyncBackgroundStatus`, `startProjectClosingSync`, `stopProjectClosingSync`, `processDirty`, `validateConfig`, `diagnoseClickUpMilestoneTask`, `syncClickUpMilestoneTask`, `diagnoseProjectClosingCandidateCounts`, `setProjectClosingDecision`, `startClickUpMilestoneClosingBackground`, `restoreMilestoneClosingFromMonthlyHistory`, `syncClickUpMilestoneRecent`, `confirmClickUpMilestoneStatuses`, `syncClickUpClosedMilestones`, `syncClickUpApprovedMilestones`, `syncClickUpRejectedMilestones`, `stopLegacyClickUpMilestoneAudit`, `syncClickUpUserActivity`, `startClickUpUserActivityBackground`, `getClickUpUserActivity`, `getClickUpUserActivityStatus`, `setConsultantSeniority`, `getBonusSalesIndications`, `saveBonusSalesIndication`, `deleteBonusSalesIndication`, `syncCmaxDailyEvents`, `startCmaxDailyHistoryBackground`, `continueCmaxDailyHistoryBatch`, `listUsers`, `createUser`, `setUserEnabled`, `resetUserPassword`, `setUserSeniority`, `deleteProjectFollowup` |

`getClickUpInventory_` e `getClickUpMilestoneClosing_` já fazem validação de sessão internamente. As demais validações internas continuam em vigor: a camada HTTP não substitui autorização por projeto (carteira), nem autenticação do webhook.

## Semântica de método e efeitos colaterais

| Classe | Exemplos auditados | Condição para aceite |
| --- | --- | --- |
| Consulta de estado sem avanço | `health`, `getClickUpInventory`, `getProjectSyncBackgroundStatus`, `getClickUpUserActivityStatus`, `getProjectClosingSyncStatus` (corrigida na branch) | GET pode retornar somente estado/dados, sem persistir nem agendar trabalho |
| Consulta com escrita/agendamento condicional | `getMonthlyProjects` (`lean` ausente/stale agenda snapshot), `getCmaxDailyEvents` (view ausente agenda build), `getProjectClosingCandidates` (`refresh=1` grava snapshot), `diagnoseProjectClosing` (persiste diagnóstico) | Desacoplar o comando; GET não pode dispará-lo, mesmo por parâmetro |
| Comando explicitamente mutável | `syncProject`, `syncAll`, `processDirty`, `validateConfig` (gera diagnóstico), ações `start*`, `stop*`, `sync*`, `confirm*`, `restore*`, `set*`, `save*`, `delete*`, `createUser`, `resetUserPassword`, `logPanelUpdate`, `logProjectFollowup` | POST/PATCH/DELETE com sessão, papel, idempotência e auditoria; GET rejeitado |
| Criação de sessão | `login` | POST sem hash/credencial na URL, proteção contra repetição e limitação de tentativas |

As “consultas” que chamam `getOrCreateSheet_`, `ensureHeaders_` ou renovam a sessão no `CacheService` ainda precisam de revisão de efeito observável. O aceite não pode ser inferido apenas pelo verbo `get`.

O contrato executável [`governance/e05-http-method-contract.json`](./governance/e05-http-method-contract.json) classifica as 56 actions sem duplicatas. `npm run audit:e05` aponta atualmente **36 comandos por GET**, **3 GETs com efeito condicional** e **17 candidatos de leitura ainda sem prova de efeito zero**. `ready=false` até zerar os três grupos pendentes. O teste do plano falha se alguém marcar E05 como OK sem limpar essas categorias.

Foi corrigido na branch um erro adicional de sessão: `syncClickUpMilestoneTask_` agora repassa `auth_token` ao consultar o resultado protegido depois de gravar o marco. Sem isso, a chamada podia devolver erro após ter modificado a base.

## Pendências que impedem o aceite E05

1. O painel publicado ainda chama sem `auth_token` a carteira mensal, CMAX, status CMAX, histórico global, registro de atualização e `syncProject`; ativar a guarda integral agora quebraria consultas e bootstrap antes do login. Adaptar os consumidores e smoke autenticado antes de publicar a guarda. Não usar um token de admin fixo no navegador.
2. Várias ações mutáveis são chamadas por GET/JSONP (`syncProject`, jobs, decisões, bonificações, cadastros, followups, histórico). `diagnoseProjectClosing` persiste diagnóstico; `getProjectClosingCandidates` com `refresh=1` grava snapshot; leituras de portfólio e CMAX também podem agendar materializações. Nesta branch, `getProjectClosingSyncStatus` já foi convertido em consulta pura: deixou de avançar o job por polling. Para atender literalmente “GET não altera estado”, separar os demais caminhos de consulta/avanço/agendamento e migrar comandos para POST autenticado, com proteção contra repetição. JSONP não pode transportar POST; exige transporte compatível antes do corte.
3. `login` via GET/JSONP transmite hash de senha na URL e cria sessão. Migrar para POST antes do aceite.
4. O webhook `doPost` precisa de testes específicos de assinatura, corpo malformado e autorização da fila, sem confundi-lo com o POST de comandos de usuários.
5. A guarda de sessão não filtra automaticamente o payload de `getMonthlyProjects`, CMAX ou os históricos compartilhados por projeto. Hoje a tela dá visão geral a todo usuário autenticado, enquanto `getClickUpInventory_` filtra o histórico por consultor; o contrato de escopo está inconsistente. É necessária decisão operacional entre visão geral autenticada e carteira individual antes de mudar o filtro, sempre aplicada no servidor. Testar uma sessão real de consultor/coordenador/admin em homologação. Se a carteira for filtrada, a validação de integridade hoje exige o piso global de 201 projetos e o contrato precisa ser versionado para não gerar falsa carga parcial. Os testes atuais são uma matriz unitária da guarda com sessões simuladas, não prova E2E autenticada.

Aceite E05: **pendente**. A suíte da branch testa as 56 ações em quatro perfis, aliases e default-deny; nenhum bloqueio novo foi promovido a produção enquanto os consumidores e a semântica de GET estiverem incompletos.

## Decisão de transporte para o corte

A documentação oficial confirma que web apps Apps Script recebem POST por `doPost`, mas respostas do `ContentService` são redirecionadas a `script.googleusercontent.com`. Mais importante, o guia oficial diz para usar JSONP somente para leitura de informação não sensível: qualquer página pode embutir a tag e capturar a resposta. Portanto, converter apenas os comandos para POST no Apps Script não resolve sozinho as leituras sensíveis do painel em JSONP; um `fetch` cross-origin diretamente para o `ContentService` também não pode ser presumido funcional sem smoke de navegador. A inferência técnica é que o corte E05 precisa de um transporte de mesma origem ou de um gateway autenticado, com testes de CORS/redirect e sessão, antes de desligar as actions GET. Isso antecipa uma pequena parte da fundação prevista em E09 ou exige ponte temporária equivalente; não autoriza implantar nova infraestrutura sem revisão do desenho.

Fontes primárias: [guia de web apps](https://developers.google.com/apps-script/guides/web), [guia de Content Service e alerta sobre JSONP](https://developers.google.com/apps-script/guides/content), [referência do redirecionamento TextOutput](https://developers.google.com/apps-script/reference/content/text-output).
