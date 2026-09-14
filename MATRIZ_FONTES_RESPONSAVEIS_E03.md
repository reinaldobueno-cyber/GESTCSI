# Matriz de fontes e responsáveis — E03

Inventário concluído em 14/09/2026 para a iniciativa `GESTCSI-BACKEND-V1-20260914`.

## Resultado

**Status: [x] OK**

O contrato versionado está em [`governance/data-contract.json`](./governance/data-contract.json), com SHA-256 `461d7392aba2dbafe9dd74ec6cc43db48af971f373ec84a832afc6075a9578a9`.

Cobertura validada:

- 20 fontes de dados e estado.
- 52 métricas funcionais e financeiras.
- 15 áreas atuais de navegação cobertas.
- 55 ações expostas pelo roteador do Apps Script inventariadas.
- 14 classes de configuração, segredo, endpoint, cache e estado de execução.
- 6 papéis de responsabilidade.
- 0 valor de credencial armazenado no catálogo.

## Modelo de responsabilidade

| Papel | Responsabilidade |
|---|---|
| Gestão CSI / PMO | Regra operacional, qualidade e aceite dos dados de carteira, acompanhamento e entrega. |
| Administração do painel | Usuários, permissões, sessões e comandos sensíveis. |
| Administração ClickUp | Workspace, listas, campos, webhooks, mapeamentos e token ClickUp. |
| Administração CMAX | Credenciais, agenda, histórico e qualidade dos eventos CMAX. |
| Financeiro / Diretoria | Contratos, taxas, indicações e aceite do fechamento financeiro. |
| Engenharia GEST CSI | Integrações, disponibilidade, observabilidade, segurança técnica e evolução. |

Os owners são papéis, não nomes pessoais. Isso mantém o contrato válido quando pessoas mudam. A designação nominal deve ocorrer no ambiente administrativo futuro, sem publicar dados pessoais no repositório.

## Matriz consolidada das fontes

| Fonte | Sistema | Classificação | Cadência/SLA | Owner operacional | Contingência | Destino planejado |
|---|---|---|---|---|---|---|
| Abas JAN–DEZ | Google Sheets | Interno | mudança visível em até 15 min | Gestão CSI | último snapshot íntegro | D1 + import job |
| `PANEL_MONTHLY_BROWSER` | Google Sheets | Interno | materialização em até 15 min | Gestão CSI | snapshot Apps Script ou leitura mensal | KV + D1 |
| Snapshot comprimido mensal | Script Properties | Interno | TTL atual de 5 min | Gestão CSI | aba materializada | KV versionado |
| `contracts_mrr.json` | Artefato estático | Confidencial agregado | mensal | Financeiro | campo da carteira ou “não localizado” | D1 restrito |
| Projetos/tarefas | ClickUp API | Interno | webhook/fila 5 min; reconciliação 6 h | Administração ClickUp | JSON mensal válido + retentativa | Queue + D1 |
| Configuração/inventário ClickUp | Sheets | Confidencial | por mudança estrutural | Administração ClickUp | redescoberta controlada | D1 mappings |
| Marcos | ClickUp API | Confidencial | incremental; reconciliação 12 h | Gestão CSI | consolidado e histórico mensal | Queue + D1 |
| `CLICKUP_FECHAMENTO_MARCOS` | Sheets | Confidencial financeiro | até 30 min | Gestão CSI | histórico mensal | D1 |
| Decisões/candidatos de projeto | Sheets | Confidencial financeiro | até 5/30 min | Gestão CSI | última decisão persistida | D1 |
| Atividade de usuários | ClickUp API | Confidencial | a cada 6 h; SLA 8 h | Administração ClickUp | última leitura com idade | Queue + D1 |
| `ACOMPANHAMENTOS_PROJETOS` | Sheets | Confidencial | escrita até 30 s | Gestão CSI | última leitura, sem inventar evento | D1 |
| `ACOMPANHAMENTOS_KANBAN` | Sheets | Interno | escrita até 30 s | Gestão CSI | estado derivado sinalizado | D1 |
| Agenda CMAX | CMAX API | Confidencial financeiro | mês corrente até 6 h | Administração CMAX | última visão materializada | Queue + D1 |
| `CMAX_DIARIAS*` | Sheets | Confidencial financeiro | após sincronização | Administração CMAX | não aceitar resposta vazia | D1 + KV |
| `CONSULTORES_REMUNERACAO` | Sheets | Restrito financeiro | alteração até 5 min | Financeiro | sem taxa significa bloquear pagamento | D1 |
| `BONUS_INDICACOES` | Sheets | Restrito financeiro | lançamento até 5 min | Financeiro | último lançamento persistido | D1 |
| `PAINEL_USUARIOS` | Sheets | Restrito segurança | alteração até 1 min | Administração do painel | negar acesso sem confirmação | D1 + sessão assinada |
| Histórico/diagnósticos | Sheets | Interno | evento até 1/5 min | Engenharia | telemetria local parcial | D1 + Logs |
| Nominatim/OSM | Serviço público | Público | sob demanda serializada | Gestão CSI | aliases IBGE e cache validado | KV |
| `localStorage` | Navegador | Local confidencial | imediato, sem multi-dispositivo | Administração do painel | fonte remota ou padrão | apenas preferências |

Os detalhes completos de consumidores, fallback e destino de cada fonte permanecem no contrato JSON.

## Configurações e segredos

O catálogo registra somente os **nomes e classes**, nunca valores:

| Classe | Exemplos | Destino |
|---|---|---|
| Identificadores | `SHEET_ID`, `CLICKUP_TEAM_ID` | configuração isolada por ambiente |
| Segredos permanentes | `CLICKUP_TOKEN`, `CLICKUP_WEBHOOK_TOKEN`, credenciais CMAX | Workers Secrets |
| Segredo efêmero | `CMAX_JWT_TOKEN` | secret/cache de curta duração |
| Endpoints | webhook ClickUp e autenticação/refresh CMAX | configuração por ambiente |
| Tuning | batch size, timeout e delay | variáveis versionadas |
| Estado de job | active, cursor, processed, error, timestamps | D1 `job_runs` + Queue |
| Estado de cache | manifestos de carteira e visão CMAX | KV com versão e TTL |
| Sessão do painel | token atualmente no `localStorage` | cookie `HttpOnly`, `Secure`, `SameSite` |

## Cobertura das métricas

As 52 métricas foram agrupadas em:

- Carteira e entrega: 9.
- Gestão, retenção e MRR: 6.
- Acompanhamento: 4.
- Adoção e confiabilidade ClickUp: 4.
- CMAX e seu frescor: 6.
- Fechamento de marcos e projetos: 8.
- Bonificação: 5.
- Pessoas e comercial: 3.
- Logística: 1.
- Prazos: 2.
- Alertas: 1.
- Confiabilidade geral: 2.
- Usuários e acesso: 1.

Cada entrada possui fontes existentes, owner operacional, owner técnico, SLA de frescor, fallback e telas consumidoras. O teste impede que uma referência a fonte ou owner inexistente seja aceita.

## Decisões de contingência

1. Uma resposta parcial nunca pode substituir um snapshot íntegro maior.
2. Um fallback deve ser identificado na interface, com origem e idade.
3. Ausência de taxa de senioridade bloqueia pagamento; não assume valor zero como decisão financeira.
4. Ausência de MRR é “não localizado”, não receita zero confirmada.
5. Falha de autenticação ou leitura de usuário segue `fail closed`.
6. Escritas não podem cair silenciosamente em estado somente local.
7. A fila pode retentar integrações; comandos financeiros exigem idempotência e auditoria.

## Problemas que a matriz tornou explícitos

- As mesmas métricas dependem de Sheets, Apps Script, cache local e ClickUp, dificultando determinar qual resposta é a vigente.
- Algumas planilhas classificadas como confidenciais ou restritas ainda são lidas por GViz público. O tratamento está previsto na E06.
- Sessão no `localStorage` e autorização desigual entre ações permanecem riscos das E05 e E11.
- Estado de execução, configuração e segredo convivem em Script Properties; deverão ser separados no backend.
- Telemetria está dividida entre planilha e navegador, sem monitor central. E04, E07 e E08 tratam essa lacuna.
- O Apps Script expõe 55 ações em um único roteador, incluindo leitura, escrita, diagnóstico e jobs administrativos.

## Validação

Foram adicionados testes que:

- conferem unicidade e completude das 20 fontes;
- exigem owner, SLA, fallback e consumidor para todas as 52 métricas;
- cobrem as 15 áreas atuais de navegação;
- comparam automaticamente as 55 ações catalogadas com o roteador real do Apps Script;
- rejeitam configuração que tente persistir valores de segredo no catálogo.

Resultado final desta entrega: **53 testes aprovados, 0 falhas**.
