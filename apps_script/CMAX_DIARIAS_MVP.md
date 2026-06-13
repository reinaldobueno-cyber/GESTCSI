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
- Apenas administradores podem executar **Atualizar historico CMAX**.
- A atualizacao traz o mes atual imediatamente e continua em segundo plano,
  mes a mes, ate janeiro de 2023. Os anos e meses aparecem progressivamente.
- Enquanto um administrador mantiver a guia aberta, a propria tela acelera a
  carga historica em lotes de tres meses.
- Clicar novamente em atualizar preserva o ponto ja alcancado no historico.
- Todos os meses desde janeiro de 2023 permanecem visiveis; meses ainda nao
  processados aparecem tracejados.
- Gravacoes mensais usam trava exclusiva para impedir que atualizacoes
  simultaneas apaguem meses ja carregados.
- Para alterar o inicio do historico, configure `CMAX_HISTORY_START_MONTH` no
  formato `AAAA-MM`.
- O menu da planilha tambem oferece **Sincronizar diarias CMAX do mes atual**.
- O menu da planilha oferece **Sincronizar historico de diarias CMAX**.
- O MVP apresenta somente quantidades. Valores de remuneracao serao tratados
  posteriormente conforme a senioridade de cada tecnico.
- `TREINAMENTO ON LINE`, `TREINAMENTO IN LOCO`, `TREINAMENTO ON LINE AVULSO`
  e `TREINAMENTO IN LOCO AVULSO` positivos contam como diaria. Essas modalidades
  aparecem primeiro e destacadas. As demais continuam visiveis como atividades.
- A tela apresenta quantitativo por modalidade e detalhamento por consultor,
  cliente, dia da semana, data e horario.
- Os detalhes de cada consultor iniciam minimizados e podem ser abertos pela seta.
- Horarios sao formatados explicitamente no fuso `America/Sao_Paulo`.
- As colunas de horario sao lidas pelo valor exibido e gravadas como texto,
  evitando a conversao de horas do Google Sheets para datas de 1899.
- A tela apresenta filtros em blocos por ano/mes, consultor e atividade.
- Somente pessoas que possuem ao menos um dos quatro treinamentos contabilizaveis
  no historico fazem parte da visao. Todas as atividades positivas dessas pessoas
  continuam disponiveis para auditoria.
- Cada consultor possui uma visao calendario mensal com intensidade por volume,
  quantidade de diarias/atividades no dia e detalhamento ao clicar na data.
