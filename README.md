# Monitor de Fluxo de Ordens para Day Trade - Mini Índice

Este projeto é um **script em Python** desenvolvido para monitorar **Times & Trades (TT)** de operações no **mini índice (WIN)**, permitindo identificar padrões de fluxo de ordens em tempo real, como **absorção, agressão e passivos**. O programa gera alertas sonoros e exibe informações relevantes de entrada para auxiliar o trader na tomada de decisão rápida.

---

## ⚡ Objetivo

- Automatizar a **leitura de ordens do Excel** exportado via RTD ou planilha.
- Filtrar e processar **agressões significativas** de compradores e vendedores.
- Identificar:
  - **Compras passivas**: ordens grandes de venda absorvidas por compradores.
  - **Vendas passivas**: ordens grandes de compra absorvidas por vendedores.
- Emitir **alerta sonoro** e imprimir no terminal informações sobre oportunidades de entrada.
- Fornecer um **controle simples de parâmetros** para ajuste de sensibilidade e limites.

---

## 📂 Estrutura do Código

- `Gerenciador_excel`: classe responsável por carregar o Excel e filtrar novas ordens.
- Parâmetros configuráveis:
  - `SALDO_COMPRAS` / `SALDO_VENDAS`: limite para identificar desequilíbrio comprador/vendedor.
  - `MIN_PASSIVO_COMPRA` / `MIN_PASIVO_VENDA`: quantidade mínima de lotes para considerar ordens passivas.
  - `temp_avalicao`: intervalo de verificação (em segundos).
- Loop principal:
  - Carrega novas ordens.
  - Calcula saldo de compradores e vendedores.
  - Identifica **ordens passivas e agressoras**.
  - Emite alertas sonoros (`winsound.Beep`) e imprime no console informações detalhadas.

---

## ⚙️ Requisitos

- Python 3.11 ou superior
- Bibliotecas:
  - `pandas`
  - `openpyxl`
- Sistema operacional: Windows (para suporte ao `winsound`)
- Excel com planilha habilitada (`.xlsm`) contendo as colunas:
  - `Data`, `Hora`, `Compradora`, `Vendedora`, `Quantidade`, `Agressor`, `Valor`

---

## 🚀 Como usar

1. Clonar o repositório ou baixar os arquivos.
2. Sincronizar via RTD a planilha com a ferramenta Profit Pro. 
3. Configurar o caminho da planilha Excel na variável `ARQUIVO` e o nome da aba em `ABA`.
4. Ajustar os parâmetros de análise:
   ```python
   SALDO_COMPRAS = 300
   SALDO_VENDAS  = -300
   MIN_PASSIVO_COMPRA = 500
   MIN_PASIVO_VENDA = 500
   temp_avalicao = 10
   
