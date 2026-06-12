# MVP de diarias CMAX

O painel consulta a agenda CMAX e armazena somente eventos com resultado
`Positivo` na aba `CMAX_DIARIAS`.

## Configuracao

1. No Apps Script, abra **Configuracoes do projeto**.
2. Em **Propriedades do script**, crie:
   - `CMAX_JWT_TOKEN`: token JWT atual do CMAX, sem escrever `JWT` no inicio.
   - `CMAX_DIARIA_BONUS`: valor por diaria. Se ausente, usa `50`.
3. Substitua o conteudo publicado de `ClickUpSync.gs` pela versao deste repositorio.
4. Crie uma nova implantacao do Web App.

O token CMAX expira. Quando isso ocorrer, o botao **Atualizar CMAX** mostrara
uma mensagem clara pedindo a atualizacao de `CMAX_JWT_TOKEN`.

## Uso

- Todos os usuarios podem consultar a guia **Diarias CMAX**.
- Apenas administradores podem executar **Atualizar CMAX**.
- A atualizacao substitui somente o mes selecionado e preserva os meses anteriores.
- O menu da planilha tambem oferece **Sincronizar diarias CMAX do mes atual**.
