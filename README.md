# 💱 Automatização de Cotações BACEN para SAP

Este projeto em Excel + Power Query + VBA foi criado para automatizar o processo de obtenção de cotações de moedas estrangeiras (USD, GBP, EUR) a partir de dados públicos do Banco Central do Brasil (BACEN), formatando-os para inserção direta na transação OB08 do SAP.

---

## 🚀 Como usar

1. Abra a planilha `cotacao_bacen_para_sap.xlsx`.
2. Informe o intervalo de datas desejado.
3. Clique no botão "Atualizar Cotações".
4. Verifique a aba de saída com os dados já formatados.
5. Copie e cole a tabela diretamente no SAP (transação OB08).

---

## 🎯 Objetivo

No ambiente SAP onde atuo, a cotação de moedas precisa ser alimentada manualmente diariamente. Esta planilha:

- Elimina o trabalho manual de busca e formatação dos dados;
- Utiliza fontes oficiais: [Banco Central do Brasil - Dados de Cotações](https://www.bcb.gov.br) e [Calendário de Dias Úteis - ANBIMA](https://www.anbima.com.br);
- Garante que apenas datas úteis bancárias (D+1) sejam consideradas;
- Gera uma tabela pronta para colar no SAP.

---

## ⚙️ Funcionalidades

- Parâmetro de datas para consulta.
- Atualização automatizada via botão (VBA: `RefreshAll` + `ThisWorkbook.Save`).
- Consulta ao BACEN via Power Query (CSV).
- Validação de datas úteis com dados da ANBIMA (HTML).
- Tabela final formatada conforme estrutura exigida pelo SAP (OB08).

---

## 📊 Exemplo de Resultado

| CG | Val.Desde | Cot Indir | X | Fator(de) | De  | = | CotPrecos | X | Fator(para) | Para |
|----|-----------|-----------|---|------------|-----|---|------------|---|---------------|------|
| M  | 11.06.2025|           | X |            | USD |   | 5,32       | X |               | BRL  |

*Dados fictícios. Tabela gerada automaticamente na aba final da planilha.*

---

## 📂 Estrutura de Arquivos

- `planilha/cotacao_bacen_para_sap.xlsx`: arquivo Excel com toda a lógica implementada (Power Query + VBA).
- `assets/Selecao_datas.png`: visual da planilha para informar período de datas.
- `assets/Tabela_formatada_datas fic.png`: visual da planilha com retorno da tabela pronta (se atentar que não se trata do retorno segundo datas do outro print)
- `README.md`: explicação do projeto e seu valor de negócio.

---

## 🧠 Aprendizados e Conexão com BI

Apesar de simples visualmente, este projeto reúne várias competências:

- Conexão e tratamento de dados com **Power Query**;
- Automação com **VBA**;
- Integração com sistema corporativo (**SAP**) via formatação estratégica;
- Consumo de dados estruturados (**CSV**) e semi-estruturados (**HTML**);
- Consideração de calendário útil bancário (**ANBIMA**) — fundamental para negócios.

### ✳️ Relevância para BI

Este projeto demonstra a capacidade de:

- Automatizar rotinas críticas com ferramentas acessíveis;
- Tratar dados com lógica de negócio;
- Trabalhar com fontes públicas de dados;
- Organizar entregas que geram impacto direto na operação.

---
