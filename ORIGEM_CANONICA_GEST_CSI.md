# Origem canônica de desenvolvimento do GEST CSI

Decisão registrada em 14/09/2026 para a iniciativa `GESTCSI-BACKEND-V1-20260914`.

## Decisão

| Item | Origem canônica |
|---|---|
| Produto | GEST CSI |
| Repositório | `https://github.com/reinaldobueno-cyber/GESTCSI.git` |
| Remoto Git local | `gestcsi` |
| Branch-base de produção | `gestcsi/main` |
| SHA-base confirmado | `11e7c2fdc9af2b1f36f1217ae90a061773f19490` |
| Versão pública confirmada | `2026-09-11-clickup-activity-v4` |
| Branch da evolução | `codex/backend-evolution` |
| Worktree exclusivo | `.worktrees/backend-evolution` |
| Upstream da branch | `gestcsi/main` |

O remoto `origin`, atualmente associado a `https://github.com/reinaldobueno-cyber/CLAUDEE.git`, **não é origem do GEST CSI** e não deve ser usado para atualizar, comparar ou publicar esta evolução.

## Evidência da criação

- `git fetch gestcsi main --prune` confirmou `gestcsi/main` no SHA `11e7c2f`.
- A página pública `https://reinaldobueno-cyber.github.io/GESTCSI/` respondeu HTTP 200 em 14/09/2026.
- O artefato público declarou `PANEL_APP_VERSION = 2026-09-11-clickup-activity-v4`, a mesma versão contida no `index.html` do SHA-base.
- O cabeçalho público informou última modificação em `Fri, 11 Sep 2026 18:07:46 GMT`.
- O worktree foi criado diretamente de `gestcsi/main`, sem copiar arquivos ou alterações do diretório raiz.
- No momento da criação, `codex/backend-evolution` apontava exatamente para `11e7c2f` e rastreava `gestcsi/main`.

## Isolamento obrigatório

O diretório raiz `GESTCSI` não será usado para desenvolver a migração. Ele contém alterações locais e artefatos não rastreados que pertencem a outros trabalhos. Na verificação de 14/09/2026, seu HEAD era `6386e0d`, com divergência de 50 commits exclusivos locais e 492 commits exclusivos em `gestcsi/main`.

Todo código novo da evolução deve ser criado em:

`C:\Users\Suporte2\Documents\GitHub\GESTCSI\.worktrees\backend-evolution`

Essa separação evita misturar alterações locais, branches antigas, arquivos de publicação e experimentos já existentes.

## Regra de atualização e publicação

1. Antes de iniciar uma entrega, atualizar a referência com `git fetch gestcsi main`.
2. Comparar a branch de evolução exclusivamente com `gestcsi/main`.
3. Desenvolver e testar em `codex/backend-evolution` ou em branches derivadas dela.
4. Nunca mesclar automaticamente arquivos do worktree raiz.
5. Publicar somente após aceite, testes e comparação explícita com a produção.
6. Nenhuma publicação ou alteração de produção faz parte da entrega E01.

## Verificação reproduzível

Executar dentro do worktree exclusivo:

```powershell
git rev-parse HEAD
git rev-parse gestcsi/main
git status --short --branch
git branch -vv
```

Para a E01 permanecer válida, os dois SHAs iniciais devem ser `11e7c2fdc9af2b1f36f1217ae90a061773f19490`, o status deve estar limpo e a branch deve rastrear `gestcsi/main`.
