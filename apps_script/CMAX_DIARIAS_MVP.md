# MVP de diarias CMAX

O painel consulta a agenda CMAX e armazena somente eventos com resultado
`Positivo` na aba `CMAX_DIARIAS`.

## Configuracao

1. No Apps Script, abra **Configuracoes do projeto**.
2. Em **Propriedades do script**, crie:
   - `CMAX_JWT_TOKEN`: token JWT atual do CMAX, sem escrever `JWT` no inicio.
3. Substitua o conteudo publicado de `ClickUpSync.gs` pela versao deste repositorio.
4. Crie uma nova implantacao do Web App.

O token CMAX expira. Quando isso ocorrer, o botao **Atualizar CMAX** mostrara
uma mensagem clara pedindo a atualizacao de `CMAX_JWT_TOKEN`.

## Uso

- Todos os usuarios podem consultar a guia **Diarias CMAX**.
- Apenas administradores podem executar **Atualizar CMAX**.
- A atualizacao substitui somente o mes selecionado e preserva os meses anteriores.
- O menu da planilha tambem oferece **Sincronizar diarias CMAX do mes atual**.
- O MVP apresenta somente quantidades. Valores de remuneracao serao tratados
  posteriormente conforme a senioridade de cada tecnico.
- Somente `TREINAMENTO ON LINE` e `TREINAMENTO IN LOCO` positivos contam como
  diaria. As demais modalidades continuam visiveis como atividades registradas.
- A tela apresenta quantitativo por modalidade e detalhamento por consultor,
  cliente, dia da semana, data e horario.
- A tela apresenta filtros em blocos por ano/mes, consultor e atividade.
- Somente pessoas que possuem ao menos um `TREINAMENTO ON LINE` ou
  `TREINAMENTO IN LOCO` no historico fazem parte da visao. Todas as atividades
  positivas dessas pessoas continuam disponiveis para auditoria.
