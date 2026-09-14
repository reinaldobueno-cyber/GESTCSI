# Baseline funcional e financeira — E02

Captura executada em 14/09/2026 para a iniciativa `GESTCSI-BACKEND-V1-20260914`.

## Resultado

**Status: [x] OK**

A baseline foi capturada a partir da produção `gestcsi/main`, SHA `11e7c2fdc9af2b1f36f1217ae90a061773f19490`, versão pública `2026-09-11-clickup-activity-v4`. O arquivo contém somente métricas agregadas, metadados técnicos e hashes SHA-256. Nenhum nome de consultor, cliente, projeto, usuário ou registro individual foi armazenado.

Arquivo selado: [`baselines/e02-2026-09-14.json`](./baselines/e02-2026-09-14.json)

SHA-256 agregado da captura: `7dabb94282f27fc5948c1a46653ae1c7ef8fbf49e53b993ba238780abf9eb014`

## Competências reconciliadas

| Métrica | JUL/2026 | AGO/2026 |
|---|---:|---:|
| Carteira na competência | 29 | 25 |
| Carteira materializada total | 234 | 234 |
| Eventos CMAX da fonte | 363 | 414 |
| Eventos CMAX elegíveis | 197 | 172 |
| Diárias equivalentes pagáveis | 153,5 | 126,9999 |
| Diárias equivalentes especiais 1,5x | 16,5 | 13,5 |
| Consultores elegíveis sem taxa | 0 | 0 |
| Valor das diárias | R$ 14.536,25 | R$ 12.938,74 |
| Marcos encerrados na competência | 189 | 191 |
| Marcos aprovados para pagamento | 153 | 165 |
| Marcos reprovados na competência de validação | 14 | 52 |
| Valor dos marcos | R$ 4.590,00 | R$ 4.950,00 |
| Fechamentos de projeto aprovados | 0 | 1 |
| Valor dos fechamentos de projeto | R$ 0,00 | R$ 80,00 |
| Indicações convertidas | 2 | 0 |
| Valor das indicações | R$ 100,00 | R$ 0,00 |
| **Total financeiro calculado** | **R$ 19.226,25** | **R$ 17.968,74** |

O valor exato não formatado de AGO/2026 é `17968.7415`. O painel apresenta moeda com duas casas decimais; o arquivo preserva quatro casas para evitar perda na comparação durante a migração.

## Regras exercitadas

- JUL/2026 inclui a exceção dos três eventos transferidos de JUN/2026, identificados pelas chaves já definidas no frontend.
- Um consultor não operacional para pagamento, já excluído pelo frontend, permanece fora do cálculo.
- Somente eventos CMAX positivos e de modalidades elegíveis contam como diária.
- Diárias são limitadas a 1 equivalente por consultor/dia.
- Fins de semana e feriados configurados recebem multiplicador financeiro de 1,5x.
- Marcos são pagos na competência efetiva de validação, inclusive a regra de validação no primeiro dia do mês.
- Fechamento de projeto permanece separado de marco e usa a base de decisões.
- Indicações usam o mês da venda convertida.
- Taxas de senioridade foram aplicadas sem ausência nas duas competências.

## Reconciliação

| Verificação | JUL/2026 | AGO/2026 |
|---|---|---|
| Soma dos componentes igual ao total financeiro | OK | OK |
| Total declarado da carteira igual às linhas recebidas | OK | OK |
| Todos os consultores elegíveis possuem taxa | OK | OK |
| Fonte CMAX materializada respondeu HTTP 200 | OK | OK |
| Fonte de marcos respondeu HTTP 200 | OK | OK |
| Hash de cada fonte registrado | OK | OK |

## Evidências técnicas

| Fonte | Registros | SHA-256 |
|---|---:|---|
| Carteira materializada | 234 | `157272d38f1b9b9921c15588c37e933e9c25c1f6f748ef8df0a4396aa949be6b` |
| Fechamento de marcos | 2.227 | `0f2a75ebd56cdd04546a54e53077652116f7838ae159b5e18824d7f75ae08f7b` |
| Decisões de fechamento | 1 | `54ae2670cdf173bc1edd802f670a4e370330321ba14d2b076b23a6833ee5f462` |
| Indicações | 2 | `ec9ec3d49a5ae9c8b6ad79f5a7b9b78b17aa47b39287dc768db2e554e52fa585` |
| Remuneração por senioridade | 18 | `fe02742660928d9de2d93914fb3f9fd97dabb9d0323b7ff280f950f9d37eed33` |
| CMAX JUN/2026 | 372 | `0529232f5cd7d68d5c678b56d661abe762846d2c06b92d6f9c45d077f01b6a81` |
| CMAX JUL/2026 | 363 | `25c0df4d665dac33e5ff43b6cd6ad1262bc3a9a35f490160d90bb54cf7ed61e8` |
| CMAX AGO/2026 | 414 | `866a575441cbfd6f3d7b91e0c0ae1972452451ea70585aea51001d3c6ef12549` |

## Estabilidade verificada

- Os 37 testes anteriores continuam aprovados.
- Foram adicionados 11 testes de baseline, privacidade, reconciliação e smoke HTTP.
- Resultado final: **48 testes aprovados, 0 falhas**.
- O smoke inicia o servidor real de preview, solicita `/` e `/src/portfolio-integrity.js`, confere a versão e exige HTTP 404 para uma rota inexistente.
- `git diff --check` não apresentou erro.
- Todas as leituras da baseline são `GET`; nenhuma sincronização ou escrita foi executada.

## Divergências e limitações conhecidas

1. O teste histórico usa um fixture mínimo de 201 projetos; a fonte materializada atual declara e entrega 234. O nome do teste foi corrigido para não tratar 201 como total atual.
2. Chamadas CMAX paralelas produziram HTTP 404 durante a primeira captura e uma repetição sequencial também falhou após retentativas. A execução diagnóstica seguinte reproduziu integralmente o SHA agregado. O capturador usa sequência, redirecionamento explícito e retentativa, mas a origem ainda precisa de monitoramento.
3. Os dados CMAX das competências históricas informaram `synced_at = 2026-09-11T14:31:41.330Z`. A captura ocorreu em 14/09/2026; portanto a baseline comprova o snapshot materializado disponível, não uma sincronização forçada naquele instante.
4. Os hashes mudarão quando as fontes operacionais forem atualizadas. O arquivo versionado é a referência imutável para comparação da migração.
5. Observações administrativas guardadas somente no navegador não integram a baseline financeira.
6. A validação não executa login nem comandos administrativos de escrita; esses fluxos continuam protegidos da captura automatizada.

## Reprodução

```powershell
npm run baseline:capture
npm test
```

O primeiro comando imprime uma nova captura anonimizada. Ela deve ser comparada com o arquivo versionado; não deve substituí-lo automaticamente.
