# Analise-de-estoque-e-custo-por-oportunidade
Script Python para gestão de estoque e custo de oportunidade. Utiliza lógica anti-falsa rotatividade baseada no tempo de investimento, protegendo a análise contra vendas isoladas de itens antigos. Inclui cálculo de margem real corrigida, categorização por faixas de lucratividade e dashboards formatados no padrão monetário brasileiro.

# Ambiente de Execução
Nota importante: Este script foi projetado especificamente para rodar no Google Colab. Ele utiliza bibliotecas nativas para autenticação com serviços Google (google.colab, gspread) e pressupõe uma estrutura de diretórios baseada em nuvem.

# Analise-de-estoque-e-custo-por-oportunidade

Este projeto é uma ferramenta de Business Intelligence voltada para a saúde financeira do estoque. Diferente de análises comuns, este script utiliza o Custo de Oportunidade e uma Lógica Anti-Falsa Rotatividade para identificar capital de giro que está sendo corroído pelo tempo, mesmo em itens que tiveram vendas recentes.

Diferenciais desta Versão

- Lógica de Investimento: Avalia o tempo que o capital está "preso" no item, impedindo que uma venda única mascare produtos estagnados há meses ou anos.
- Correção Financeira: Aplica juros compostos sobre o custo de aquisição para revelar a margem de lucro real atualizada.
- Dashboards Brasileiros: Gráficos de barras e dispersão com formatação de moeda (R$) e tratamento de discrepâncias visuais (trava de margens extremas).

Estrutura de Dados Necessária

O arquivo de entrada (CSV ou Excel) deve conter:

- SKU / Código: Identificador do produto.
- Estoque: Quantidade física disponível.
- Custo: Preço de compra original.
- Preço: Preço de venda atual.
- Ultima Compra: Data de entrada do lote.
- Ultima Saída: Data da última venda realizada.

Como Adaptar

Se o seu banco de dados utiliza nomes de colunas diferentes, basta ajustar o mapeamento na função processar_dados. O script foi desenhado para ser agnóstico ao setor, funcionando bem em varejos de construção, autopeças, vestuário e distribuidores.
