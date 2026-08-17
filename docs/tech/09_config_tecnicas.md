# 17. Configurações e Dependências Técnicas

**Método:** referências extraídas por decompressão do stream `VBA/dir` do `vbaProject.bin` (algoritmo MS-OVBA, via `oletools`), que armazena o registro binário de cada biblioteca referenciada pelo projeto.

## 17.1 Versão e arquitetura do Office

**[CONFIRMADO — evidência: caminho da referência de biblioteca]** O projeto referencia `C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL` ("Microsoft Office 16.0 Object Library") — confirma que o arquivo foi editado pela última vez (ou desde sempre) em uma versão do **Office 16.0** (Office 2016, 2019, 2021 ou Microsoft 365 desktop — o número "16.0" é compartilhado por todas essas versões, não dá pra diferenciar qual exatamente). Arquitetura (32 ou 64 bits) — **[NÃO IDENTIFICADO]**: o caminho `C:\Program Files\...` (sem `(x86)`) é consistente com Office 64 bits, mas não é uma confirmação definitiva sem inspecionar a máquina original.

## 17.2 Referências VBA (bibliotecas)

| Biblioteca | GUID | Caminho referenciado | Uso confirmado no código |
|---|---|---|---|
| Visual Basic For Applications | (implícita, padrão) | — | Base da linguagem |
| OLE Automation (`stdole`) | `{00020430-0000-0000-C000-000000000046}` | `C:\Windows\System32\stdole2.tlb` | Tipos automáticos padrão |
| Microsoft Office 16.0 Object Library | `{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}` | `...\OFFICE16\MSO.DLL` | `FileDialog` (`Application.FileDialog`), usado em `Aux_Leitura_Nome_Arqs.bas` e nos UserForms |
| Microsoft ActiveX Data Objects 2.8 Library | `{2A75196C-D9EB-4129-B803-931327F72D5C}` | `...\ado\msado28.tlb` | `ADODB.Connection`/`ADODB.Recordset`/`ADODB.Command`, usados em `Conexoes.bas` e todos os módulos que consultam SQL Server |
| Microsoft Forms 2.0 Object Library | `{0D452EE1-E08F-101A-852E-02608C4D0BB4}` | `C:\WINDOWS\system32\FM20.DLL` | Suporte aos 3 UserForms |
| **Referência quebrada/nula** | `{00000000-0000-0000-0000-000000000000}` | (vazio) | **[NÃO IDENTIFICADO]** — presença de uma entrada de referência com GUID totalmente zerado é uma anomalia; não foi possível determinar a que biblioteca ela deveria apontar originalmente. Ver risco na seção 18. |

**Evidência adicional (dado incidental, tratado com cautela — ver nota de privacidade abaixo):** o cache local de compilação do formulário (`MSForms.exd`) referencia um caminho de perfil de usuário do Windows na máquina onde o projeto foi compilado pela última vez. O identificador de usuário específico não é reproduzido aqui (ver Regra de Segurança 5 do escopo desta análise) — foi substituído por **[INFORMAÇÃO SENSÍVEL OMITIDA]**. Esse caminho é apenas cache local de compilação, não é um requisito funcional do sistema (não impede o arquivo de rodar em outra máquina).

## 17.3 Proteção do projeto VBA

**[CONFIRMADO]** O stream `PROJECT` contém os campos `CMG=`, `DPB=` e `GC=`, que o VBA usa para armazenar hash de senha e flags de bloqueio de visualização do projeto — **evidência de que o projeto VBA está (ou esteve) protegido por senha**. Isso bloqueia a abertura do editor VBA (Alt+F11) sem a senha correta, mas não impede a extração do código-fonte por ferramentas externas de parsing (usado nesta análise). **[VALIDAR COM O NEGÓCIO]**: confirmar se a senha do projeto é conhecida/documentada pela equipe atual — se não for, isso é, na prática, uma trava que impede qualquer edição direta do código dentro do próprio Excel até a senha ser recuperada ou o projeto ser desbloqueado.

## 17.4 ActiveX

**[INFERIDO]** O uso de `Microsoft Forms 2.0 Object Library` é o único indício de controles ActiveX/Forms — usado pelos UserForms (ListBox, TextBox, CommandButton, OptionButton, ComboBox, MultiPage, conforme catálogo da seção 10). Não há evidência de outros controles ActiveX incorporados diretamente em planilhas (ex.: um `CommandButton` ActiveX solto numa aba) nos módulos de código analisados — os botões das telas principais são **shapes/formas com macro atribuída** (confirmado pelos arquivos `drawing*.xml`), não controles ActiveX.

## 17.5 APIs e componentes do Windows

**[CONFIRMADO]** Uso de `Application.FileDialog` (msoFileDialogFilePicker / msoFileDialogFolderPicker) para diálogos nativos do Windows de seleção de arquivo/pasta. Não foi encontrada nenhuma declaração `Declare Function`/`Declare Sub` (chamada direta à API do Windows/DLL externa) em nenhum dos módulos lidos — **[CONFIRMADO POR AUSÊNCIA]**, mas não é uma varredura exaustiva de 100% do código-fonte, então trate como alta confiança, não certeza absoluta.

## 17.6 Caminhos de rede

**[CONFIRMADO — evidência: `LISTA_ARQ_AUX`, células de configuração em Sheet24]** O sistema depende de caminhos de rede fixos para os arquivos-fonte externos, documentados (parcialmente) na própria aba `LISTA_ARQ_AUX`:
- Hubble: `F:\P&C 2026\Data Management\Hubble\` (arquivo `Painel Hubble.xlsb`)
- De-Para: `F:\Nucleo Desenvolvimento\Resultado por Diretoria\Base\` (arquivo `Bases_DE_PARA.xlsx`)

Os caminhos das demais 6 fontes externas (1009, RGM, MOCKUP, Fixed Revenues, Other Income, Base Consolidada) **não são fixos no código** — são apontados manualmente pelo usuário a cada execução via `Aux_Leitura_Nome_Arqs.bas`, e o último caminho usado fica gravado em células da aba Extração (Sheet8).

## 17.7 Servidor de banco de dados

**[CONFIRMADO — evidência: `Conexoes.bas`]** SQL Server `SNEPDB24V`, bancos `BPAM` (principal) e `InfoGER` (fallback). Autenticação por usuário/senha SQL (não integrada ao Windows) — **credenciais em texto plano no código-fonte**, tratado como achado crítico de segurança na seção 18. `CommandTimeout` configurado para 1000 segundos (~16 min) em praticamente todas as consultas — indício de que as consultas são historicamente lentas.

## 17.8 Configurações regionais

**[INFERIDO]** Código e mensagens em português do Brasil. Uso de `xlPinYin` (ordenação chinesa) em ao menos uma rotina de ordenação (`Refresh_Sup_Linhas.bas`) — **[VALIDAR COM O NEGÓCIO]**: provável herança de configuração regional do Excel na máquina onde a macro foi originalmente gravada, não parece intencional para dados em PT-BR, mas funciona "por coincidência" (não há evidência de que produza resultado incorreto, só que é uma escolha estranha).

## 17.9 Outras dependências identificadas

- Todas as 8 fontes de extração de arquivo externo esperam arquivos `.xlsx`/`.xls` (não há suporte a `.csv` ou outros formatos identificado no código lido).
- Dependência de que o usuário tenha permissão de leitura nos caminhos de rede acima e de escrita/leitura no banco `BPAM`/`InfoGER` com o usuário SQL hardcoded.
- Dependência de que a extensão de arquivo selecionada no diálogo (`Application.GetOpenFilename`) seja de fato um Excel válido — o filtro usado é "Todos os Arquivos (*.*)" em `Aux_Leitura_Nome_Arqs.bas`, ou seja, **não há validação de extensão na seleção do arquivo** (ver riscos, seção 18).
