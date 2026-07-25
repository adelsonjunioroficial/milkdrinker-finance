# MilkDrinker — Cash Cow Finance v1.1

Um sistema robusto de gestão financeira desenvolvido integralmente em VBA (Visual Basic for Applications) no Microsoft Excel. O projeto foi desenhado para operar como uma aplicação independente (Standalone), ocultando a interface tradicional de planilhas e oferecendo uma navegação baseada inteiramente em UserForms.

## Arquitetura Técnica

O banco de dados foi estruturado seguindo as melhores práticas de modelagem relacional:
* **Modelo Dimensional (Star Schema):** Divisão lógica entre Tabelas Fato (registros operacionais com 9 colunas) e Tabelas Dimensão (listas de apoio para cadastro).
* **Chaves Naturais:** Utilização de descrições textuais ao invés de IDs numéricos para simplificar a leitura de relatórios sem necessidade de cruzamentos complexos (JOINs).
* **Processamento em Memória:** Uso da estrutura `Scripting.Dictionary` para consolidação instantânea de dados no Dashboard, garantindo alta performance na renderização de resultados.

---

## Telas e Interface do Sistema

<!-- INSTRUÇÃO: Substitua o texto "link_da_imagem_aqui.png" pelo caminho da sua imagem no repositório. Exemplo: "docs/img/menu.png" -->

### Menu Principal
![Menu Principal](link_da_imagem_aqui.png)
> Central de navegação do usuário.

### Dashboard Analítico
![Dashboard Analítico](link_da_imagem_aqui.png)
> Motor de consolidação instantânea de totais, saldos e filtros operacionais.

### Lançamentos Financeiros (Avulsos e Parcelados)
![Lançamentos Financeiros](link_da_imagem_aqui.png)
> Interface paginada para captura de receitas e despesas.

### Gerenciamento de Categorias
![Gerenciamento de Categorias](link_da_imagem_aqui.png)
> Painel de edição e exclusão condicional de tabelas dimensão.

---

## Funcionalidades Principais

* **Motor de Parcelamento e Recorrência:** Módulo inteligente que decompõe valores de compras a prazo ou contas fixas recorrentes. Calcula automaticamente as parcelas e projeta as datas, meses e anos de vencimento futuros de forma sequencial.
* **Integridade Referencial Estrita:** Trava de segurança no módulo de exclusão. Se uma categoria, pagador ou forma de pagamento já possuir vínculo nas tabelas fato, o sistema bloqueia a exclusão para evitar a corrupção do histórico.
* **Atualização em Cascata (Cascading Update):** Ao renomear uma categoria na tabela dimensão, o sistema varre e atualiza automaticamente todos os lançamentos passados associados àquele item.
* **Auditoria Invisível (Log Automático):** Todo lançamento gravado no sistema recebe automaticamente um carimbo exato de data e hora de inserção na coluna primária.
* **Modo Desenvolvedor (Dev Mode):** Atalho de manutenção seguro para tornar a aplicação do Excel visível novamente, permitindo auditoria estrutural e ajustes diretamente nas tabelas.

---

## Pré-requisitos

* Microsoft Excel para Windows.
* Habilitar a execução de Macros (Conteúdo Ativo) no Excel.

## Como Executar

1. Faça o download do arquivo `MilkDrinker_v1.1.xlsm`.
2. Abra o arquivo no Microsoft Excel.
3. Aceite o aviso de segurança do Excel clicando em "Habilitar Conteúdo".
4. A interface das planilhas será automaticamente ocultada e o sistema inicializará no Menu Principal.

---

## Documentação Completa

Para detalhes aprofundados sobre a arquitetura de tabelas e mapeamento das colunas, consulte a documentação estrutural do projeto:

* [Dicionário de Dados](link_para_seu_dicionario.pdf)
* [Manual de Operações do Usuário](link_para_seu_manual.pdf)

> Desenvolvido como portfólio de programação VBA e modelagem de dados no Microsoft Excel.
