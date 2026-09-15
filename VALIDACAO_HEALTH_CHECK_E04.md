# E04 — Health check do Apps Script

Data da validação: 15/09/2026

Branch: `codex/backend-evolution`

Estado: **[x] OK — implantado, testado e monitorado.**

## Resultado entregue

O Apps Script possui a ação pública `health`, com contrato próprio e versão `2026-09-15-health-v1`. A resposta pública contém somente:

- disponibilidade lógica;
- nome do serviço;
- estado `ready` ou `degraded`;
- versão do contrato e da implantação;
- instante da verificação;
- nível de visibilidade.

Ela não lê planilhas, não chama ClickUp ou CMAX e não retorna identificadores, métricas operacionais, nomes, tokens, senhas ou valores de propriedades. Apenas a presença de `SHEET_ID` e `CLICKUP_TOKEN` influencia o estado público.

Com `detail=1`, a rota exige uma sessão de administrador por `requireAdmin_()`. O detalhe informa somente presença de configuração e estado resumido dos jobs de projetos, atividade e histórico CMAX; valores sensíveis continuam ausentes.

## Monitor externo

O comando abaixo consulta a implantação, segue o redirecionamento do Google, aplica timeout, valida serviço e versão e exige estado `ready`:

```powershell
npm run health:apps-script
```

O monitor realiza três tentativas. Apenas depois de três falhas consecutivas retorna erro. O workflow `.github/workflows/apps-script-health.yml` executa esse comando a cada 15 minutos no branch principal, abre uma única issue de incidente e encerra a issue quando o serviço se recupera.

## Cenários cobertos

| Cenário | Resultado |
|---|---|
| Endpoint público válido | Aceita contrato exato sem credencial ou dado de negócio |
| Detalhe autenticado | Envia sessão e exige resposta `authenticated` com dependências |
| Timeout | Interrompe a chamada e classifica o tempo excedido |
| Implantação removida/HTTP 404 | Repete três vezes e gera um único incidente |
| Serviço degradado | Recusa a resposta para permitir alerta operacional |
| Versão antiga | Recusa HTTP 200 com serviço ou versão incompatível |

## Publicação

URL consumida pelo painel:

`https://script.google.com/macros/s/AKfycbxtbpwBZEHDHCtU7lovmFHQ6R-MDgff-CB-yAyH6DfRwo0SjR9WXU6B4EYrgCcza6Kj/exec`

Para impedir regressão após a indisponibilidade observada em 15/09/2026, a publicação partiu da versão imutável `283`, que havia restaurado login e acesso ao storage. Somente a constante de versão, a rota `health` e sua função de resposta foram adicionadas. A versão resultante foi registrada como `285`.

Antes da promoção, o deployment de homologação foi atualizado para `285` e validado. Depois, o deployment de produção existente foi atualizado para a mesma versão, preservando a URL consumida pelo painel e a versão `283` como rollback.

## Evidências pós-publicação

- três respostas públicas consecutivas: HTTP 200, serviço `gestcsi-apps-script`, estado `ready` e versão `2026-09-15-health-v1`;
- login/JSONP: HTTP 200 e callback válido, sem erro de acesso ao Apps Script ou storage;
- histórico do painel: HTTP 200, `ok=true`, 2.635 registros disponíveis;
- manifesto do portfólio: HTTP 200, `ok=true`, 234 projetos materializados;
- histórico CMAX: HTTP 200, `ok=true`, competências desde 2024-01;
- página pública do Gestão CSI: HTTP 200;
- suíte automatizada: `60/60` testes aprovados;
- monitor CLI: resposta aprovada contra a URL de produção;
- workflow no branch principal: execução manual aprovada e agenda de 15 minutos ativa.

## Rollback

Em caso de regressão, o deployment de produção pode ser novamente apontado para a versão `283`, sem troca da URL pública. Nenhum dado operacional foi migrado ou alterado pela E04.
