# pje-dashboard-sise-jt-juizos

O arquivo `juizos` diz qual juiz(a) assina cada final de processo em cada vara. Ele é **gerado automaticamente**
a partir do [cronograma de afastamentos](https://docs.google.com/spreadsheets/d/1ypqbgBT9jV3UR43s7ZesXTQGgB6kNRdx/edit)
(aba *Cronograma*). **Não edite `juizos` à mão**: a próxima execução sobrescreve. Altere a planilha.

## Quando muda

- **Saída:** às **10h do último dia útil anterior** à data inicial do afastamento.
- **Retorno:** às **8h do primeiro dia útil posterior** à data final.
- Afastamentos encadeados (ex.: compensação seguida de férias) viram um só período.

O GitHub Actions (`.github/workflows/atualiza-juizos.yml`) roda de hora em hora nos dias úteis, das 8h05 às 18h05,
baixa a planilha, recalcula e faz commit só se algo mudou. Para rodar na hora: aba **Actions → Atualiza juizos pelo
cronograma → Run workflow**.

## Regras de redistribuição

- Os finais do ausente vão para o(s) **titular(es) presente(s)** da mesma vara; havendo dois, por paridade
  (ex.: 9 → quem tem ímpares, 0 → quem tem pares).
- Sem titular presente, vão para quem estiver presente (ex.: juiz da simetria).
- Sem ninguém presente, vão para quem está na coluna **"Quem assina quando não há substituto"** da planilha.

## Configuração

- `config/base.json` — divisão fixa de finais por vara (titular/simetria) e lista de Postos Avançados.
- `config/feriados.json` — feriados (dias não úteis além de fins de semana). Revise todo ano.
- Teste local: `python scripts/atualiza_juizos.py --xlsx cronograma.xlsx --agora "2026-12-07 08:00" --saida /tmp/juizos`
