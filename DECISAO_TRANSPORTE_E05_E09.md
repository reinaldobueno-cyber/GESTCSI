# Transporte do corte E05 e dependência de E09

Status em 15/09/2026: **decisão técnica proposta, não implantada**. O primeiro lote de comandos POST existe apenas na branch `codex/backend-evolution`. O painel publicado ainda usa GET/JSONP; desligar essas rotas agora quebraria login e sincronização.

## Decisão proposta

Servir o painel e `/api/v1/*` no **mesmo domínio controlado** por um Worker com Static Assets. A página pode então chamar `fetch('/api/v1/...', { method: 'POST' })` sem depender de CORS do Apps Script ou de JSONP. Um Worker hospedado em outro domínio, enquanto o painel permanece no GitHub Pages, **não** fornece mesma origem. A migração de domínio e o mecanismo de sessão precisam de homologação antes de qualquer corte do legado.

O Worker não deve ser um proxy aberto para o endpoint Apps Script. Durante a ponte, cada rota deve ter uma allowlist explícita de método, ação e papel; uma sessão emitida no servidor com cookie `HttpOnly`, `Secure` e `SameSite`; proteção CSRF nos comandos; limitação de login; idempotência e auditoria para escrita; limites de corpo e resposta; timeout; e erro fechado quando o upstream não puder ser confirmado. Nenhum token Apps Script, ClickUp, CMAX ou de admin deve ser exposto ao navegador, à URL, à configuração versionada ou aos logs. Webhooks precisam de rota separada, com validação da assinatura do provedor, e não devem compartilhar a autenticação de usuário.

O endpoint `ContentService` do Apps Script redireciona a resposta para `script.googleusercontent.com`. O Worker pode seguir esse redirecionamento no lado servidor, mas isto exige smoke real com o deployment de homologação: os probes externos de GET observados foram inconclusivos e não comprovam que POST/redirect funciona. A ponte também não corrige, por si, o acesso anônimo às abas por GViz; E06 continua sendo uma entrega independente.

## Portões antes de ativar

1. Definir o domínio de homologação sob controle da equipe, separado da produção, e servir nele o mesmo build do painel e a API. Não reutilizar a URL atual do GitHub Pages como se ela pudesse receber uma rota Worker local.
2. Implantar E09 com bindings e secrets separados por ambiente, dados fictícios/anonimizados em homologação e observabilidade sem PII.
3. Implantar E11/E12 ou controles equivalentes para sessão, revogação, papéis, CSRF, limite de tentativas, idempotência e auditoria. Um gateway que apenas repassa o `auth_token` do navegador não atende ao objetivo de E05.
4. Smoke de navegador em homologação: login POST, `me`, consulta histórica paginada, sync autorizado, 403 de consultor para comando admin, 405 em comando GET, redirect Apps Script, sessão expirada, CSRF e upstream indisponível.
5. Conferir integridade da carteira por papel: em visão geral autenticada, os 337 projetos históricos e 215 de 2026 devem produzir 552 quando a fonte estiver íntegra; timeout deve manter 2026 visível sem apresentar zero histórico como ausência real. Se a política for carteira individual, versionar e testar os totais filtrados em todas as fontes.
6. Adaptar o painel para a API de mesma origem, retirar chamadas JSONP sensíveis e só então publicar a guarda E05 no Apps Script. O rollback deve restaurar o último build íntegro, sem abrir novamente mutações anônimas.

## Limite desta entrega

O protótipo local [`gateway/`](./gateway/) agora contém uma ponte de homologação desligada por padrão, com login/sessão opaca, logout com CSRF e leitura histórica paginada. Isso não cria conta, domínio, banco, secrets nem deployment; tampouco marca E05 ou E09 como OK. O primeiro lote POST, a ponte e os testes simulados são preparação para homologação, não prova de navegador nem correção publicada dos projetos `<=2025`.

Fontes primárias: [Cloudflare Static Assets e binding](https://developers.cloudflare.com/workers/static-assets/binding/), [Cloudflare Workers Best Practices](https://developers.cloudflare.com/workers/best-practices/workers-best-practices/), [Google Apps Script Content Service](https://developers.google.com/apps-script/guides/content/).
