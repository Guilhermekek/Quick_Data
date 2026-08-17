# 25. Declaração de Cobertura

## O que foi analisado

- Estrutura completa das 28 planilhas (nomes, ordem, conteúdo de células relevante).
- Código-fonte completo dos 70 componentes do projeto VBA (192 procedimentos), decompilado diretamente do binário `vbaProject.bin`.
- Os 3 UserForms (código-behind completo; ver limitação sobre definição visual em "o que não foi analisado").
- Os 166 nomes definidos do workbook, com sua classificação própria de erro/exclusão.
- Referências de bibliotecas do projeto VBA (extraídas do stream `VBA/dir`).
- Evidência de proteção do projeto VBA (campos `CMG`/`DPB`/`GC`).
- Rótulos e nomes de shapes/botões das principais telas de interação (via `drawing*.xml`).
- Credenciais e strings de conexão SQL Server usadas pelo sistema (tratadas conforme regra de segurança 5 — nunca reproduzidas neste documento).

## O que NÃO foi analisado (ou não foi possível analisar)

- **Comportamento em tempo de execução** — nenhuma macro, consulta ou conexão foi executada; toda a documentação de fluxo é inferência de leitura estática de código.
- **Proteção de planilha/workbook** — não decodificável pelas ferramentas usadas.
- **Definição visual binária dos UserForms** (`.frx`) — controles inferidos por uso no código, não lidos diretamente.
- **Conteúdo binário de tabelas dinâmicas, links externos e conexões de workbook** — existência confirmada, conteúdo interno não decodificado.
- **Formatação condicional** de forma sistemática.
- **Mapeamento CodeName↔nome de aba** para parte das 28 planilhas — resolvido só onde havia evidência textual indireta no código.
- **Ligação exata entre cada shape/botão e sua macro** (`OnAction`) — confirmados apenas os rótulos de texto dos botões.

## Quantidade de componentes identificados (contagem mecânica, alta confiança)

| Item | Quantidade |
|---|---|
| Planilhas | 28 |
| Componentes de código VBA (módulos + forms + ThisWorkbook) | 70 |
| Procedimentos (Sub/Function/Property) | 192 |
| Nomes definidos | 166 (42 sinalizados para exclusão) |
| Regras de negócio catalogadas | 112 (RN-001 a RN-115, com gaps de numeração reservados não usados) |
| Riscos classificados | 46 (consolidado dos 4 clusters — ver seção 18) |
| Fluxos funcionais documentados | 8 (F-01 a F-08) |
| Cenários de teste derivados | 17 (CT-01 a CT-17) |
| Fontes de dados de entrada | 9 (1 SQL automatizada + 7 arquivo/pasta manuais + 1 entrada interna) |
| Referências de biblioteca VBA | 5 (incluindo 1 referência quebrada) |

## Limitações que podem afetar esta documentação

1. Este documento é fruto de **análise estática** de um arquivo protegido por senha no VBE, extraído por ferramenta externa — não houve confirmação interativa de nenhum comportamento.
2. Vários itens permanecem `[VALIDAR COM O NEGÓCIO]` (seção 23) — este documento **não deve ser tratado como definitivo** até essas 21 pendências serem resolvidas com alguém do time de negócio ou com acesso interativo ao arquivo.
3. A extração de código foi feita uma única vez, em 2026-08-14 — qualquer alteração no arquivo após essa data não está refletida aqui.
4. O trabalho de catalogação de procedimentos (seção 10) e regras de negócio (seção 12) foi produzido por 4 frentes de análise em paralelo, cada uma cobrindo um subconjunto de módulos — a convenção de formatação varia levemente entre as 4 partes (ver nota no início de cada subseção de 10 e 12), embora o conteúdo técnico siga o mesmo padrão de 15 campos/10 campos exigido.
5. Por não ter sido possível validar node a node o comportamento em runtime, riscos classificados como "crítico"/"alto" nesta análise refletem a leitura do código-fonte (o que ele *pode* causar), não incidentes confirmados em produção — exceto onde a própria planilha `tk_Lista_de_erros` já registra ocorrências reais (seção 16.4).

---

*Documento gerado por engenharia reversa assistida do arquivo `Quick Data 3.23.xlsb`. Ver [PREENCHER] nos campos administrativos da capa antes de circular oficialmente.*
