# 1. Capa

**Título:** Documentação Técnica e Funcional do Quick Data 3.23
**Sistema:** Quick Data 3.23
**Versão analisada:** 3.23 (conforme nome do arquivo `Quick Data 3.23.xlsb`; a checagem de versão interna do sistema, sub `Verifica_Versao` em `Auxiliar.bas`, compara contra a string local `"3.0"` no código — **[VALIDAR COM O NEGÓCIO]**: há uma divergência não explicada entre o número do arquivo (3.23) e o número usado no gate de versão do VBA (3.0); pode ser um controle de versão desatualizado/não mantido, ou os dois números terem significados diferentes)
**Data da análise:** 2026-08-14
**Responsável pelo preenchimento:** [PREENCHER]
**Status do documento:** Rascunho técnico — gerado por engenharia reversa assistida; pendente de revisão por alguém com acesso ao Excel/VBE (ver seção 4, "Limitações da análise", e seção 23, "Dúvidas e Pontos Pendentes")

---

# 2. Controle de Versões

| Versão do documento | Data | Autor | Descrição da alteração | Revisor |
|---|---|---|---|---|
| 0.1 | 2026-08-14 | Claude (assistente, via engenharia reversa automatizada) | Primeira versão funcional, alto nível (documento descontinuado — ver seção "Objetivo e Escopo") | [PREENCHER] |
| 1.0 | 2026-08-14 | Claude (assistente, via engenharia reversa automatizada) | Reescrita completa seguindo especificação de documentação técnica formal de 24 seções: catálogo de procedimentos, regras de negócio numeradas (RN-xxx), matriz de dependências e de rastreabilidade, cenários de teste | [PREENCHER] |

---

# 3. Sumário

> Estrutura de títulos pronta para gerar sumário automático no Word (Referências → Sumário → estilos de título aplicados nos números abaixo).

1. Capa
2. Controle de Versões
3. Sumário
4. Objetivo e Escopo
5. Visão Geral do Quick Data 3.23
6. Arquitetura da Solução
7. Inventário de Abas
8. Mapa "O Que Alimenta o Quê"
9. Inventário do Projeto VBA
10. Catálogo de Procedimentos VBA
11. Fluxos Funcionais e Operacionais
12. Catálogo de Regras de Negócio
13. Arquivos de Entrada
14. Arquivos e Dados de Saída
15. Fórmulas, Nomes Definidos e Conexões
16. Tratamento de Erros e Logs
17. Configurações e Dependências Técnicas
18. Riscos para Manutenção
19. Guia de Manutenção
20. Matriz de Rastreabilidade
21. Cenários de Teste
22. Glossário
23. Dúvidas e Pontos Pendentes
24. Anexos
25. Declaração de Cobertura
