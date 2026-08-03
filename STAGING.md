# Ambiente de homologação GESTCSI

Este worktree é isolado da produção e parte do mesmo commit publicado na `main`.

## Executar testes

```powershell
$env:PATH="$env:LOCALAPPDATA\CodexTools\node-v22;$env:PATH"
npm test
```

## Abrir homologação local

```powershell
$env:PATH="$env:LOCALAPPDATA\CodexTools\node-v22;$env:PATH"
npm run preview
```

Acesse `http://127.0.0.1:4173`. O servidor envia `Cache-Control: no-store` para impedir validações contra arquivos antigos.

## Regra de promoção

Uma mudança só deve seguir para produção quando:

1. Todos os testes estiverem verdes.
2. O total declarado conferir com o número de projetos recebido.
3. Os totais por mês conferirem com o manifesto.
4. Uma carga menor não substituir o último snapshot completo.
5. Impressão individual, impressão administrativa e mapa passarem no smoke test.
