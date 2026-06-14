# MVP de diarias CMAX

O painel consulta a agenda CMAX e armazena somente eventos com resultado
`Positivo` na aba `CMAX_DIARIAS`.

## Configuracao

1. No Apps Script, abra **Configuracoes do projeto**.
2. Em **Propriedades do script**, crie:
   - `CMAX_EMAIL`: email usado para entrar no CMAX.
   - `CMAX_PASSWORD`: senha usada para entrar no CMAX.
   - `CMAX_JWT_TOKEN`: opcional; o painel cria e renova automaticamente.
3. Substitua o conteudo publicado de `ClickUpSync.gs` pela versao deste repositorio.
4. Crie uma nova implantacao do Web App.

Quando o token CMAX expira, o Apps Script autentica novamente em
`/servicos/login/`, salva o novo JWT e continua a atualizacao sem intervencao.

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
- A trava cobre somente a escrita final na planilha; consultas ao CMAX nao
  bloqueiam a leitura da tela nem o botao de atualizacao.
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
- Celulas antigas convertidas para 1899 sao reparadas pelo `raw_json` original;
  a interface nunca exibe datas de 1899 como horario.
- A hora final usa tambem o campo original `end_time` do CMAX.
- Quando o nome do campo varia, o extrator localiza semanticamente campos de
  fim/final/ate/end/termino ou intervalos escritos como `09:00 - 12:00`.
- A tela apresenta filtros em blocos por ano/mes, consultor e atividade.
- Somente pessoas que possuem ao menos um dos quatro treinamentos contabilizaveis
  nos ultimos 6 meses fazem parte da visao. Todas as atividades positivas dessas
  pessoas continuam disponiveis para auditoria. O periodo pode ser alterado pela
  propriedade `CMAX_TRAINING_TEAM_MONTHS`.
- A consulta da tela ignora a coluna pesada `raw_json`, reduzindo o tempo de
  carregamento sem perder os dados de auditoria armazenados.
- A tela aguarda ate 90 segundos na primeira leitura e tenta novamente
  automaticamente caso a base esteja temporariamente ocupada.
- Cada consultor possui uma visao calendario mensal com intensidade por volume,
  quantidade de diarias/atividades no dia e detalhamento ao clicar na data.
