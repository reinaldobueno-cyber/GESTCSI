# E04 — Health check do Apps Script

Data da validação: 14/09/2026

Branch: `codex/backend-evolution`

Estado: implementação pronta; aguardando nova implantação do Web App e ativação do workflow no branch principal.

## Resultado entregue

O Apps Script agora possui a ação pública `health`, com contrato próprio e versão `2026-09-14-health-v1`. A resposta pública contém somente:

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

O monitor realiza três tentativas. Apenas depois de três falhas consecutivas retorna erro. O workflow `.github/workflows/apps-script-health.yml` executará esse comando a cada 15 minutos no branch principal, abrirá uma única issue de incidente e encerrará a issue quando o serviço se recuperar.

## Cenários cobertos

| Cenário | Resultado |
|---|---|
| Endpoint público válido | Aceita contrato exato sem credencial ou dado de negócio |
| Detalhe autenticado | Envia sessão e exige resposta `authenticated` com dependências |
| Timeout | Interrompe a chamada e classifica o tempo excedido |
| Implantação removida/HTTP 404 | Repete três vezes e gera um único incidente |
| Serviço degradado | Recusa a resposta para permitir alerta operacional |
| Versão antiga | Recusa HTTP 200 com serviço ou versão incompatível |

Suíte completa após a implementação: `60/60` testes aprovados, sem regressões.

## Verificação da implantação atual

URL consumida pelo painel:

`https://script.google.com/macros/s/AKfycbxtbpwBZEHDHCtU7lovmFHQ6R-MDgff-CB-yAyH6DfRwo0SjR9WXU6B4EYrgCcza6Kj/exec`

Em 14/09/2026, `?action=health` respondeu HTTP 200, mas com o contrato genérico antigo:

- serviço retornado: `clickup-sync`;
- versão esperada pelo health check: `2026-09-14-health-v1`;
- resultado do monitor: três falhas consecutivas `unexpected_service`.

Isso comprova que disponibilidade HTTP isolada não é suficiente e evita declarar saudável um Web App antigo.

## Condição restante para marcar OK

O repositório não contém `.clasp.json`, manifesto `appsscript.json` nem credencial de implantação automatizada. Portanto, a publicação do Web App precisa ser feita por quem administra o projeto no Google Apps Script:

1. substituir o conteúdo de `ClickUpSync.gs` pela versão deste branch;
2. em **Implantar > Gerenciar implantações**, editar a implantação existente e selecionar **Nova versão**;
3. preservar a URL usada pelo painel e o acesso já configurado;
4. executar `npm run health:apps-script`;
5. confirmar HTTP 200, serviço `gestcsi-apps-script`, estado `ready` e versão `2026-09-14-health-v1`;
6. integrar o branch ao principal para ativar o monitor agendado e confirmar uma execução manual do workflow.

Somente após essas verificações a caixa da E04 deve ser alterada para `[x] OK`.
