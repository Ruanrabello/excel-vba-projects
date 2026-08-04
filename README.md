<p align="center">
  <img src="./assets/excel-vba-header.svg" width="100%" alt="Excel VBA Automation — macros e produtividade">
</p>

<p align="center">
  <strong>Macros e automações para relatórios, tratamento de dados e produtividade no Excel.</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="Microsoft Excel">
  <img src="https://img.shields.io/badge/VBA-172B4D?style=for-the-badge&logo=visualbasic&logoColor=white" alt="VBA">
  <img src="https://img.shields.io/badge/Automation-6C63FF?style=for-the-badge" alt="Automation">
  <img src="https://img.shields.io/badge/License-MIT-00A86B?style=for-the-badge" alt="MIT License">
</p>

## Sobre

Este repositório reúne soluções práticas em **VBA (Visual Basic for Applications)** para automatizar tarefas recorrentes no Microsoft Excel. Cada exemplo contém o módulo `.bas` para leitura e reutilização e, quando aplicável, uma planilha de demonstração.

O objetivo é demonstrar automação de processos, manipulação de dados, geração de relatórios e integração do Excel com o Outlook.

## Automações disponíveis

| Projeto | O que faz | Arquivos principais |
|---|---|---|
| Gráfico automático | Consolida dados, cria um gráfico de distribuição e exporta o relatório em PDF | [`AutomatizacaoGrafico/`](./AutomatizacaoGrafico) |
| E-mails pelo Outlook | Lê destinatário, assunto e mensagem da planilha e prepara e-mails individuais | [`EmailsAuto/`](./EmailsAuto) |
| Formatação de tabela | Limpa valores, converte textos em números e aplica uma tabela padronizada | [`Formatar/`](./Formatar) |
| Relatório por registro | Filtra registros aprovados, gera PDFs e prepara os anexos no Outlook | [`SRC/AutomatizacaoV1/`](./SRC/AutomatizacaoV1) |
| Reajuste de preços | Cria uma coluna com acréscimo de 5% para registros ativos | [`SRC/AutomatizacaoV2/`](./SRC/AutomatizacaoV2) |
| Busca de produtos | Automatiza uma consulta equivalente ao `PROCV` entre planilhas | [`SRC/Procv/`](./SRC/Procv) |

## Tecnologias e conceitos

- Microsoft Excel e VBA
- Manipulação de células, ranges e tabelas (`ListObject`)
- Filtros e tratamento de dados
- Criação de gráficos
- Exportação de planilhas para PDF
- Automação do Outlook
- Funções de busca e regras condicionais

## Como executar

### Usando uma planilha de exemplo

1. Baixe ou clone este repositório.
2. Abra um arquivo `.xlsm` no Microsoft Excel para Windows.
3. Revise o código antes de habilitar macros.
4. Ajuste nomes de planilhas, colunas, destinatários e pastas quando indicado.
5. Pressione `Alt + F8`, selecione a macro e clique em **Executar**.

### Importando um módulo `.bas`

1. Abra sua planilha e pressione `Alt + F11`.
2. No Editor do VBA, acesse **Arquivo > Importar arquivo**.
3. Selecione o módulo `.bas` desejado.
4. Adapte as referências de planilhas e intervalos para sua estrutura.
5. Salve o arquivo como **Pasta de Trabalho Habilitada para Macro (`.xlsm`)**.

## Cuidados antes de executar

Alguns exemplos contêm caminhos locais e endereços fictícios para demonstração. Antes de executar:

- substitua caminhos como `C:\Users\...` por uma pasta existente no seu computador;
- mantenha `.Display` durante os testes de e-mail para revisar a mensagem antes do envio;
- valide nomes de abas e cabeçalhos;
- trabalhe inicialmente com uma cópia da planilha;
- habilite macros apenas em arquivos cuja origem você conhece.

## Estrutura

```text
Excel-vba-projects/
├── AutomatizacaoGrafico/  # Gráfico e exportação para PDF
├── EmailsAuto/            # Integração com Outlook
├── Formatar/              # Limpeza e padronização de tabelas
├── SRC/                   # Estudos organizados por automação
└── Outros/                # Macros adicionais
```

## Próximas melhorias

- [ ] Adicionar imagens de antes e depois
- [ ] Substituir caminhos fixos por seletores de pasta
- [ ] Padronizar tratamento de erros
- [ ] Adicionar arquivos de dados fictícios para todos os exemplos
- [ ] Documentar entradas e saídas em cada módulo

## Autor

**Ruan Rabello** — estudante de Engenharia da Computação com foco em Dados, Automação e desenvolvimento de soluções.

[LinkedIn](https://www.linkedin.com/in/ruan-rabello-da-silva-9032b5274/) · [Portfólio](https://ruanportifolio.lovable.app) · [GitHub](https://github.com/Ruanrabello)

## Licença

Distribuído sob a [licença MIT](./LICENSE).
