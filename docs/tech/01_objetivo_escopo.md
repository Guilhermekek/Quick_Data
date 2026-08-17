# 4. Objetivo e Escopo

## 4.1 Finalidade do documento

Documentar o funcionamento técnico e funcional completo do arquivo `Quick Data 3.23.xlsb` para permitir que desenvolvedores e a equipe de sustentação: entendam o sistema sem depender do conhecimento informal de quem o construiu; façam correções e manutenção com segurança; implementem novas funcionalidades sem quebrar dependências ocultas; investiguem erros, lentidão e corrupção do arquivo com uma referência técnica confiável; e, futuramente, usem este documento como base de conhecimento para a reescrita do sistema como aplicação standalone (ver `02_melhorias_e_recomendacoes.md` para a avaliação de melhorias — este documento não trata de modernização, apenas do estado atual).

## 4.2 Público-alvo

Desenvolvedores e analistas de sustentação com conhecimento técnico de Excel/VBA e de conceitos de ETL/modelagem de dados. Também serve de insumo técnico para quem for reconstruir o sistema em outra tecnologia.

## 4.3 Metodologia e componentes efetivamente analisados

Esta análise **não foi feita com o arquivo aberto interativamente no Excel/VBE**. Foi feita por engenharia reversa do pacote do arquivo `.xlsb` (que é um contêiner ZIP/OOXML binário) e do projeto VBA nele embutido, usando ferramentas de parsing programático:

- **Estrutura do workbook e conteúdo de células:** extraído via `pyxlsb` (leitor de células BIFF12) — permitiu ler nomes de abas, ordem das abas, e valores/fórmulas de células específicas.
- **Código VBA (módulos padrão, módulos de classe, código de planilha, código de `ThisWorkbook`, código-behind de UserForms):** extraído via `oletools`/`olevba`, que decompila o projeto VBA diretamente do stream OLE `vbaProject.bin`. **Evidência de que o projeto está protegido por senha:** o stream `PROJECT` do `vbaProject.bin` contém os campos `CMG=`, `DPB=` e `GC=`, que o formato VBA usa para armazenar o hash de senha e o estado de bloqueio do projeto. Isso normalmente impede a abertura do editor VBA (VBE) sem a senha — mas **não impede a extração do código-fonte por ferramentas de parsing direto do arquivo**, que foi o método usado aqui. Ou seja: o código-fonte deste documento é fiel ao que está gravado no arquivo, mas não foi possível confirmar interativamente nenhum comportamento em tempo de execução (breakpoints, valores de variável durante a execução, etc.).
- **Nomes definidos:** extraídos de uma planilha de auditoria que o próprio sistema mantém (`ListDefinedNames`), que lista todos os 166 nomes definidos do arquivo com uma classificação própria de "erro"/"apagar".
- **Botões e shapes das telas principais:** extraídos dos arquivos `drawing*.xml` internos ao pacote (texto dos rótulos e nomes dos shapes).
- **Referências de bibliotecas do projeto VBA:** extraídas via decompressão do stream `VBA/dir` (algoritmo MS-OVBA) do `vbaProject.bin`.

## 4.4 Componentes NÃO acessíveis por este método (limitação declarada)

Os itens abaixo **não puderam ser inspecionados** com as ferramentas usadas nesta análise e estão marcados como **[NÃO ACESSÍVEL]** onde relevante ao longo do documento:

- **Proteção de planilha e de workbook** (se há senha em alguma aba/estrutura do workbook) — essa informação fica em registros BIFF12 que os leitores usados não decodificam.
- **Definição visual binária dos controles dos UserForms** (arquivo `.frx` — posição X/Y, largura, altura, propriedades visuais exatas de cada controle). Os controles foram inferidos pelo nome das variáveis e pelo uso no código (ex.: `.AddItem`, `.Selected` só existem em ListBox), com alta confiança, mas **não é uma leitura direta da definição do formulário**.
- **Conteúdo binário de tabelas dinâmicas** (`pivotCacheDefinition1.bin`, `pivotTable1.bin`, `pivotTable2.bin`) — confirma-se que existem 2 tabelas dinâmicas (associadas à aba `DP_Rateio`), mas os campos internos não foram decodificados; os nomes dos filtros/segmentações (slicers) associados foram obtidos indiretamente pelos nomes dos shapes (`EMPRESA`, `ABERTURA_1`, `LINHA_BD`, `ABERTURA_2`, `IFRS_CONTABIL`).
- **Conteúdo dos links externos** (`externalLink1.bin`, `externalLink2.bin`, `externalLink3.bin`) e do arquivo `connections.bin` — confirma-se a existência de 3 links externos e de definições de conexão a nível de workbook, mas o conteúdo binário não foi decodificado. As conexões a banco de dados usadas pelo VBA (SQL Server) foram documentadas a partir do código-fonte de `Conexoes.bas`, que é uma fonte mais confiável para esse propósito.
- **Formatação condicional** — não foi extraída de forma sistemática; apenas o que apareceu incidentalmente no código VBA (ex.: coloração de células por `Selection.Interior` em `UPDATE_aplicar_CDC_por_Referencia`) foi documentado.
- **Comportamento em tempo de execução** — nenhuma macro, consulta ou conexão externa foi executada durante esta análise (por definição do escopo: análise estática apenas). Qualquer afirmação sobre "o que acontece quando..." é uma inferência da leitura do código, não uma observação de execução real.

## 4.5 Limitações adicionais da análise

- O código-fonte VBA foi decompilado a partir do binário; comentários e nomes de variáveis são fiéis ao original, mas a formatação de indentação/espaçamento pode diferir do editor original.
- A correspondência entre o **CodeName** de uma planilha (usado pelo VBA, ex. `Sheet3`) e seu **nome de exibição** (aba visível, ex. "Base") **não pôde ser obtida de forma mecânica e completa** — essa informação fica em um registro BIFF12 (`BrtWsProp`) dentro de cada `sheetN.bin`, não exposto pelas bibliotecas usadas. Os mapeamentos apresentados na seção 7 foram reconstruídos por **evidência indireta** (o próprio código, em comentários e nomes de variáveis, referenciando o nome da aba logo perto de uma atribuição do tipo `Set Sh_X = SheetN`) e têm nível de confiança declarado por linha. Os mapeamentos não confirmados estão marcados **[NÃO IDENTIFICADO]**. **Um desenvolvedor com o arquivo aberto no VBE pode confirmar todos eles em poucos minutos** olhando a janela "Project Explorer" (cada planilha aparece como `SheetN (NomeDaAba)`) — ver seção 23.
- Esta é uma análise estática de código e estrutura. Regras de negócio documentadas como "confirmadas" são confirmadas **no código**, não necessariamente validadas contra o comportamento esperado pelo negócio hoje — itens como o status real do módulo `BackupCodigo_MainResults` (parece desativado, mas isso não pode ser confirmado sem alguém do time de negócio) permanecem como **[VALIDAR COM O NEGÓCIO]**.
- Este documento **não está completo por definição** enquanto os itens da seção 4.4 permanecerem [NÃO ACESSÍVEL] e os itens da seção 23 não forem esclarecidos — ele documenta o que foi possível confirmar com as ferramentas disponíveis nesta rodada.
