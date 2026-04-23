# -*- coding: utf-8 -*-
#------------------------------------------
# --- CÉLULA 1: IMPORTS E CONFIGURAÇÕES ---
#------------------------------------------

# 1.1 Imports e intalações
!pip install fpdf gspread oauth2client

import locale
import warnings
warnings.filterwarnings('ignore')
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from fpdf import FPDF
from google.colab import auth
import gspread
from google.auth import default
import logging
from typing import Optional, List, Tuple
import warnings
warnings.filterwarnings('ignore')
import os # Adicionado para manipulação de caminhos de arquivo

# 1.2 Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# 1.3 Configuração de Região
# Aplica separador de milhar "."
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, '')

# 1.4 Constantes de configuração
TAXA_JUROS_MENSAL = 0.01  # 1% ao mês - custo de oportunidade
MARGEM_ALVO = 0.30  # 30% de margem desejada
DIAS_POR_MES = 30.44  # Média de dias por mês para cálculos precisos

#---------------------------------------------
# --- CÉLULA 2: FUNÇÕES DE CORE BUSINESS ---
#---------------------------------------------

# 2.1 Carrega dados
def carregar_dados_excel(caminho_arquivo: str) -> pd.DataFrame:
    """
    Carrega dados do arquivo Excel ou CSV, dependendo da extensão.
    Args:
        caminho_arquivo: Caminho para o arquivo Excel ou CSV
    Returns:
        DataFrame com os dados carregados
    """
    try:
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo)
            logger.info(f"Dados CSV carregados com sucesso de: {caminho_arquivo}")
        elif caminho_arquivo.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb', '.odf', '.ods', '.odt')):
            df = pd.read_excel(caminho_arquivo)
            logger.info(f"Dados Excel carregados com sucesso de: {caminho_arquivo}")
        else:
            raise ValueError("Formato de arquivo não suportado. Por favor, forneça um arquivo Excel (.xls, .xlsx) ou CSV (.csv).")

        logger.info(f"Total de registros: {len(df)}")
        return df
    except FileNotFoundError:
        logger.error(f"Arquivo não encontrado: {caminho_arquivo}")
        raise
    except Exception as e:
        logger.error(f"Erro ao carregar arquivo: {str(e)}")
        raise

# 2.2 Ajusta datas e calcula tempo
def preparar_datas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converte colunas de data para datetime e calcula métricas de tempo.
    Args:
        df: DataFrame com colunas 'Ultima Compra' e 'Ultima Saída'
    Returns:
        DataFrame com datas convertidas e novas colunas de tempo
    """
    df = df.copy()
    hoje = datetime.now()

    # 2.2.1 Converte colunas de data de forma robusta
    for col in ['Ultima Entrada', 'Ultima Saída']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

            # Log de datas inválidas
            invalid_dates = df[col].isna().sum()
            if invalid_dates > 0:
                logger.warning(f"{invalid_dates} registros com data inválida na coluna '{col}'")

    # 2.2.3 Calcula dias desde a compra e última venda
    df['dias_investimento'] = (hoje - df['Ultima Entrada']).dt.days
    df['dias_sem_venda'] = (hoje - df['Ultima Saída']).dt.days

    # 2.2.4 Converte para meses (usando média de dias por mês para maior precisão)
    df['meses_investimento'] = (df['dias_investimento'] / DIAS_POR_MES).round(0).astype('Int64')
    df['meses_sem_venda'] = (df['dias_sem_venda'] / DIAS_POR_MES).round(0).astype('Int64')

    # 2.2.5 Trata valores negativos (datas futuras) como 0
    df['meses_investimento'] = df['meses_investimento'].clip(lower=0)
    df['meses_sem_venda'] = df['meses_sem_venda'].clip(lower=0)

    # 2.2.6 Preenche quaisquer valores NaN restantes em 'meses_investimento' e 'meses_sem_venda' com 0.
    df['meses_investimento'] = df['meses_investimento'].fillna(0)
    df['meses_sem_venda'] = df['meses_sem_venda'].fillna(0)

    logger.info("Datas processadas e métricas de tempo calculadas")
    return df

# 2.3 Calculos financeiros
def calcular_metricas_financeiras(
    df: pd.DataFrame,
    taxa_juros: float = TAXA_JUROS_MENSAL,
    margem_alvo: float = MARGEM_ALVO
) -> pd.DataFrame:
    """
    Calcula métricas financeiras baseadas no custo de oportunidade.
    Args:
        df: DataFrame com colunas 'Custo', 'Preço' e 'meses_investimento'
        taxa_juros: Taxa de juros mensal para cálculo do custo corrigido
        margem_alvo: Margem desejada para sugestão de preço
    Returns:
        DataFrame com novas colunas financeiras
    """
    df = df.copy()

    # 2.3.1 Certificar que 'Custo' e 'Preço' sejam numéricos e preencha os valores NaN com 0.
    df['Custo'] = pd.to_numeric(df['Custo'], errors='coerce').fillna(0)
    df['Preço'] = pd.to_numeric(df['Preço'], errors='coerce').fillna(0)

    # 2.3.2 Certificar que 'meses_investimento' ejam numéricos e preencha os valores NaN com 0.
    df['meses_investimento'] = pd.to_numeric(df['meses_investimento'], errors='coerce').fillna(0)

    # 2.3.3 Custo corrigido pelo tempo de investimento (juros compostos)
    df['custo_corrigido'] = df['Custo'] * ((1 + taxa_juros) ** df['meses_investimento'])
    df['custo_corrigido'] = df['custo_corrigido'].fillna(0)

    # 2.3.4 Cálculo da margem real atual com proteção contra divisão por zero
    df['margem_real_atual'] = np.where(
    df['Preço'] > 0,
    ((df['Preço'] - df['custo_corrigido']) / df['Preço']) * 100,
    -100.0 # Preço zero ou negativo define perda total (-100%)
    )

    # --- TRATAMENTO DE DISCREPÂNCIAS PARA GRÁFICOS ---
    # 2.3.5 Cria margem_visual travada em -30% para não quebrar a escala dos gráficos
    df['margem_visual'] = df['margem_real_atual'].clip(lower=-30, upper=100)

    # 2.3.6 Cria Faixas de Margem (Bucketing)
    bins = [-float('inf'), -0.01, 15, 30, 50, float('inf')]
    labels = ['Prejuízo Real', 'Margem Crítica (0-15%)', 'Margem Padrão (15-30%)', 'Margem Saudável (30-50%)', 'Alta Lucratividade']
    df['faixa_margem'] = pd.cut(df['margem_real_atual'], bins=bins, labels=labels)

    # 2.3.7 Preencher quaisquer NaNs restantes em margem_real_atual (por exemplo, se o custo_corrigido de alguma forma se tornou NaN) com -100,0
    df['margem_real_atual'] = df['margem_real_atual'].fillna(-100.0)

    # 2.3.8 Sugestão de preço para atingir a margem alvo
    if (1 - margem_alvo) == 0:
        df['sugestao_preco'] = np.inf # Tratar caso margem_alvo é 1,0
    else:
        df['sugestao_preco'] = (df['custo_corrigido'] / (1 - margem_alvo)).round(2)
    df['sugestao_preco'] = df['sugestao_preco'].fillna(0).replace([np.inf, -np.inf], 0) # Tratar inf/NaN aqui também.

    # 2.3.9 Prejuízo identificado
    df['em_prejuizo'] = np.where(
    (df['custo_corrigido'] > df['Preço']) |
    ((df['meses_investimento'] >= 12) & (df['margem_real_atual'] < 15)),
    True, False
)

    logger.info("Métricas financeiras calculadas com sucesso")
    return df

# 2.4 Define estratégia para os produtos
def definir_acao_final(row):
    """
    Define a estratégia comercial baseada no tempo de investimento, giro e margem real.
    Aplica lógica anti-falsa rotatividade para proteger o capital de giro.
    Args:
        row (pd.Series): Linha do DataFrame contendo as colunas 'Preço', 'custo_corrigido',
                         'meses_investimento', 'meses_sem_venda' e 'margem_real_atual'.
    Returns:
        str: Categoria da estratégia (ex: 'Liquidação Urgente', 'Preço Saudável', etc).
    """

    # 2.4.1 Parâmetros de Gestão
    margem_minima_seguranca = 15.0
    tempo_limite_investimento = 12 # 1 ano

    # 2.4.2 -1. LIQUIDAÇÃO URGENTE (Prejuízo nominal ou corrosão crítica de margem)
    if (row['Preço'] < row['custo_corrigido']) or \
       (row['meses_investimento'] >= tempo_limite_investimento and row['margem_real_atual'] < margem_minima_seguranca):
        return 'Liquidação Urgente'

    # 2.4.3 -2. ANTI-FALSA ROTATIVIDADE
    # Se o capital está preso há mais de 12 meses, mesmo com venda este mês,
    # ele NÃO é saudável. Precisa baixar estoque para recuperar o caixa.
    if row['meses_investimento'] >= tempo_limite_investimento and row['meses_sem_venda'] <= 1:
        return 'Monitorar Giro (Estoque Antigo)'

    # 2.4.4 -3. OBSOLESCÊNCIA (Parado há mais de 18 meses independente de qualquer fator)
    if row['meses_investimento'] >= 18:
        return 'Liquidação (Item Obsoleto)'

    # 2.4.5 -4. REVISÃO COMERCIAL (O produto ainda tem margem, mas está "dormindo")
    if row['meses_sem_venda'] > 3 and row['meses_investimento'] > 6:
        return 'Revisar Precificação'

    # 2.4.6 -5. AJUSTE DE ESTOQUE ANTIGO (Tem giro, mas o custo corrigido está subindo)
    if row['meses_investimento'] > 6 and row['meses_sem_venda'] <= 1:
        return 'Ajustar Margem (Estoque Antigo)'

    # 2.4.7 -6. PADRÃO
    return 'Preço Saudável'

# 2.5 Aplica a logica sobre os produtos
def aplicar_estrategias(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica a lógica de definição de estratégia e atualiza o status de prejuízo.
    Args:
        df: DataFrame com as métricas financeiras já calculadas.
    Returns:
        DataFrame com as colunas 'estrategia' e 'em_prejuizo' atualizadas.
    """
    df = df.copy()
    df['estrategia'] = df.apply(definir_estrategia_com_liquidez, axis=1)
    df['em_prejuizo'] = np.where(
    (df['custo_corrigido'] > df['Preço']) |
    ((df['meses_investimento'] >= 12) & (df['margem_real_atual'] < 15)),
    True, False
)
    logger.info("Estratégias de precificação aplicadas com sucesso.")
    return df

#--------------------------------------------------
# --- CÉLULA 3: FUNÇÕES DE LIMPEZA E FORMATAÇÃO ---
#--------------------------------------------------

# 3.1 Limpeza e formatação e dados
def limpar_para_exportacao(df: pd.DataFrame) -> List[List]:
    """
    Limpa o DataFrame para exportação (remove NaN, Inf, formata datas).
    Args:
        df: DataFrame a ser limpo
    Returns:
        Lista de listas pronta para exportação
    """
    df_export = df.copy()

    # 3.1.1 Formata colunas de data (mais robusto, pega qualquer datetime64)
    for col in df_export.select_dtypes(include=['datetime64']).columns:
        df_export[col] = df_export[col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')

    # 3.1.2 Identificar colunas monetárias para formatação
    currency_cols = ['Custo', 'Preço', 'custo_corrigido', 'sugestao_preco']

    for col in df_export.columns:
        if col in currency_cols and pd.api.types.is_numeric_dtype(df_export[col]):
            # Aplicar formatação de moeda para colunas monetárias
            # Substituir Inf/-Inf por NaN antes de formatar para que formatar_moeda lide com eles
            df_export[col] = df_export[col].replace([np.inf, -np.inf], np.nan)
            df_export[col] = df_export[col].apply(lambda x: formatar_moeda(x))
        elif pd.api.types.is_numeric_dtype(df_export[col]):
            # Lógica original para outras colunas numéricas
            df_export[col] = df_export[col].replace([np.inf, -np.inf], np.nan)
            df_export[col] = df_export[col].apply(lambda x: None if pd.isna(x) else x)

    # 3.1.3 Converte para lista de listas
    data_to_export = [df_export.columns.values.tolist()] + df_export.values.tolist()

    # 3.1.4 Limpeza final recursiva
    for r_idx, row in enumerate(data_to_export):
        for c_idx, val in enumerate(row):
            if isinstance(val, float) and (np.isnan(val) or np.isinf(val)):
                data_to_export[r_idx][c_idx] = None

    logger.info(f"Dados preparados para exportação: {len(data_to_export)-1} linhas")
    return data_to_export

# 3.2 Formata em moeda
def formatar_moeda(valor: float) -> str:
    """Formata valor para moeda brasileira."""
    if pd.isna(valor) or valor is None:
        return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

#-----------------------------------------------------
# --- CÉLULA 4: FUNÇÕES DE EXPORTAÇÃO E RELATÓRIOS ---
#-----------------------------------------------------

# 4.1 Exporta para o googlesheets (pode ser substituido para exportar para .xlsx)
def exportar_para_google_sheets(df: pd.DataFrame, nome_planilha: str) -> bool:
    """
    Exporta o DataFrame para o Google Sheets.
    Args:
        df: DataFrame a ser exportado
        nome_planilha: Nome da planilha no Google Sheets
    Returns:
        True se exportado com sucesso, False caso contrário
    """
    try:
        # 4.1.2 Autenticação
        auth.authenticate_user()
        creds, _ = default()
        gc = gspread.authorize(creds)

        # 4.1.3 Cria ou abre a planilha
        try:
            sh = gc.open(nome_planilha)
            logger.info(f"Planilha existente encontrada: {nome_planilha}")
        except gspread.SpreadsheetNotFound:
            sh = gc.create(nome_planilha)
            logger.info(f"Nova planilha criada: {nome_planilha}")

        # 4.1.4 Prepara e envia os dados
        worksheet = sh.sheet1
        data_to_export = limpar_para_exportacao(df)

        worksheet.clear()
        worksheet.update(data_to_export)

        logger.info(f"Dados exportados com sucesso para: {nome_planilha}")
        return True

    except Exception as e:
        logger.error(f"Erro ao exportar para Google Sheets: {str(e)}")
        return False

# 4.2 Cria relatório em PDF
def gerar_pdf_acoes(
    df_acoes: pd.DataFrame,
    nome_arquivo: str = "Relatorio_Acao_Imediata.pdf",
    titulo: str = "GUIA DE VENDAS E DESCONTOS - COMERCIAL AVENIDA",
    resumo_executivo: Optional[dict] = None
) -> bool:
    """
    Gera um PDF com os itens que necessitam de ação imediata.
    Args:
        df_acoes: DataFrame filtrado com produtos que precisam de ação
        nome_arquivo: Nome do arquivo PDF a ser gerado
        titulo: Título do relatório
        resumo_executivo: Dicionário com o resumo estatístico a ser adicionado.
    Returns:
        True se gerado com sucesso, False caso contrário
    """
    try:
        pdf = FPDF()
        pdf.add_page(orientation='L')

        # 4.2.1 Título
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(280, 10, titulo, ln=True, align='C')
        pdf.ln(5)

        # 4.2.2 Adicionar Resumo Executivo, se fornecido
        if resumo_executivo:
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, "Resumo Executivo", ln=True, align='L')
            pdf.set_font("Arial", '', 10)
            pdf.multi_cell(0, 6, f"Total de produtos analisados: {resumo_executivo['total_produtos']}")
            pdf.multi_cell(0, 6, f"Total de unidades em estoque: {resumo_executivo['total_estoque']:.0f}")
            pdf.multi_cell(0, 6, f"Valor do estoque a preço de custo: {formatar_moeda(resumo_executivo['valor_estoque_custo'])}")
            pdf.multi_cell(0, 6, f"Valor do estoque a preço de venda: {formatar_moeda(resumo_executivo['valor_estoque_venda'])}")
            pdf.multi_cell(0, 6, f"Valor do estoque com custo corrigido: {formatar_moeda(resumo_executivo['valor_estoque_custo_corrigido'])}")
            pdf.multi_cell(0, 6, f"Produtos em situação de prejuízo: {resumo_executivo['produtos_prejuizo']}")
            pdf.multi_cell(0, 6, f"Perda potencial estimada: {formatar_moeda(resumo_executivo['valor_perda_potencial'])}")
            pdf.multi_cell(0, 6, f"Média de margem real: {resumo_executivo['media_margem_real']:.1f}%")
            pdf.multi_cell(0, 6, f"Produtos que precisam de ação: {resumo_executivo['produtos_acao_imediata']}")
            pdf.ln(5)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, "Itens para Ação Imediata:", ln=True, align='L')
            pdf.ln(2)


        # 4.2.3 Cabeçalho da tabela
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(200, 200, 200)

        # 4.2.4 Ajuste das larguras para caber na página (total ~278mm para A4 paisagem)
        cabecalhos = ['SKU', 'Produto', 'Estoque', 'Investimento', 'Sem Venda', 'Margem Real', 'Status', 'Sugestão de Preço']
        larguras = [20, 70, 23, 28, 28, 25, 42, 42] # Ajustado 'Produto' para 70, adicionado 'SKU' com 20

        for i, cabecalho in enumerate(cabecalhos):
            pdf.cell(larguras[i], 10, cabecalho, border=1, align='C', fill=True)
        pdf.ln()

        # 4.2.5 Dados
        pdf.set_font("Arial", '', 9)
        for _, row in df_acoes.sort_values('meses_investimento', ascending=False).iterrows():
            # Truncar descrição se muito longa
            descricao = str(row['Descrição'])[:50]
            if len(str(row['Descrição'])) > 50:
                descricao += "..."

            # Adicionar linha
            pdf.cell(larguras[0], 8, str(row['SKU']), border=1, align='C')
            pdf.cell(larguras[1], 8, descricao.encode('latin-1', 'replace').decode('latin-1'), border=1)
            pdf.cell(larguras[2], 8, str(row['Estoque']), border=1, align='C')
            pdf.cell(larguras[3], 8, f"{row['meses_investimento']}m", border=1, align='C')
            pdf.cell(larguras[4], 8, f"{row['meses_sem_venda']}m", border=1, align='C')
            pdf.cell(larguras[5], 8, f"{row['margem_real_atual']:.1f}%", border=1, align='C') # Nova coluna de margem
            pdf.cell(larguras[6], 8, row['estrategia'], border=1, align='C')
            pdf.cell(larguras[7], 8, formatar_moeda(row['sugestao_preco']), border=1, align='R')
            pdf.ln()

        # 4.2.6 Rodapé com data de geração
        pdf.set_y(-15)
        pdf.set_font("Arial", 'I', 8)
        pdf.cell(280, 10, f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                ln=True, align='C')

        pdf.output(nome_arquivo)
        logger.info(f"PDF gerado com sucesso: {nome_arquivo}")
        return True

    except Exception as e:
        logger.error(f"Erro ao gerar PDF: {str(e)}")
        return False

# 4.3 Gera um resumo executivo
def gerar_resumo_estoque(df: pd.DataFrame) -> dict:
    """
    Gera um resumo estatístico do estoque.
    Args:
        df: DataFrame com os dados dos produtos
    Returns:
        Dicionário com métricas de resumo
    """
    resumo = {
        'total_produtos': len(df),
        'total_estoque': df['Estoque'].sum(),
        'valor_estoque_custo': (df['Estoque'] * df['Custo']).sum(),
        'valor_estoque_venda': (df['Estoque'] * df['Preço']).sum(),
        'valor_estoque_custo_corrigido': (df['Estoque'] * df['custo_corrigido']).sum(),
        'produtos_prejuizo': df['em_prejuizo'].sum(),
        'valor_perda_potencial': ((df['custo_corrigido'] - df['Preço']).clip(lower=0) * df['Estoque']).sum(),
        'media_margem_real': df['margem_real_atual'].mean(),
        'produtos_acao_imediata': len(df[df['estrategia'] != 'Preço Saudável'])
    }

    return resumo

# 4.4 Gera o cabeçalho para o PDF
def imprimir_resumo_estoque(resumo: dict):
    """
    Imprime os dados consolidados do resumo executivo no cabeçalho do PDF.
    Utiliza f-strings para extrair valores dinâmicos do dicionário de resumo.
    """
    print("\n" + "="*60)
    print("RESUMO EXECUTIVO - ANÁLISE DE ESTOQUE")
    print("="*60)
    print(f"Total de produtos analisados: {resumo['total_produtos']}")
    print(f"Total de unidades em estoque: {resumo['total_estoque']:.0f}")
    print(f"Valor do estoque a preço de custo: {formatar_moeda(resumo['valor_estoque_custo'])}")
    print(f"Valor do estoque a preço de venda: {formatar_moeda(resumo['valor_estoque_venda'])}")
    print(f"Valor do estoque com custo corrigido: {formatar_moeda(resumo['valor_estoque_custo_corrigido'])}")
    print(f"Produtos em situação de prejuízo: {resumo['produtos_prejuizo']}")
    print(f"Perda potencial estimada: {formatar_moeda(resumo['valor_perda_potencial'])}")
    print(f"Média de margem real: {resumo['media_margem_real']:.1f}%")
    print(f"Produtos que precisam de ação: {resumo['produtos_acao_imediata']}")
    print("="*60 + "\n")

# 4.5 Cria o grafico de dispersão
def gerar_grafico_analise(df: pd.DataFrame, salvar_imagem: bool = False):
    """
    Cria uma análise visual de dispersão entre Preço de Venda e Margem Real.
    Utiliza uma 'margem visual' limitada para evitar distorções de escala nos eixos.
        Args:
        df (pd.DataFrame): DataFrame processado.
        salvar_imagem (bool): Se True, exporta o gráfico como arquivo PNG.
    Returns:
        None: Exibe o gráfico ou salva o arquivo conforme parâmetro.
    """
    plt.figure(figsize=(14, 6))

    # 4.5.1 Cores personalizadas para cada estratégia
    cores = {
        'Preço Saudável': 'green',
        'Revisar Precificação / Queima': 'orange',
        'Liquidação Urgente': 'red',
        'Liquidação (Item Obsoleto)': 'brown',
        'Ajustar Margem (Estoque Antigo)': 'blue',
        'Otimizar Margem': 'purple'
    }

    # 4.5.2 Gráfico de dispersão
    for estrategia, cor in cores.items():
        dados_estrategia = df[df['estrategia'] == estrategia]
        if len(dados_estrategia) > 0:
            plt.scatter(
                dados_estrategia['meses_investimento'],
                dados_estrategia['margem_visual'],
                c=cor,
                label=estrategia,
                s=100,
                alpha=0.7,
                edgecolors='black',
                linewidth=1
            )

    # 4.5.3 Linha de referência para margem zero
    plt.axhline(0, color='black', linestyle='--', linewidth=1, alpha=0.7)

    # 4.5.4 Linha de referência para margem alvo
    plt.axhline(MARGEM_ALVO * 100, color='gray', linestyle=':', linewidth=1, alpha=0.7,
                label=f'Margem Alvo ({MARGEM_ALVO*100:.0f}%)')

    plt.title("Análise Correlacionada: Tempo de Compra vs. Margem Real", fontsize=14, fontweight='bold')
    plt.xlabel("Meses desde a Compra (Dinheiro Parado)", fontsize=12)
    plt.ylabel("Margem Real Atual (%)", fontsize=12)
    plt.legend(loc='best', fontsize=10)
    plt.grid(True, alpha=0.3)

    # 4.5.5 Adiciona anotações para pontos extremos
    pontos_extremos = df.nlargest(5, 'meses_investimento')
    for _, row in pontos_extremos.iterrows():
        plt.annotate(
            str(row['SKU']),
            (row['meses_investimento'], row['margem_visual']),
            xytext=(5, 5),
            textcoords='offset points',
            fontsize=8,
            alpha=0.7
        )

    plt.tight_layout()

    if salvar_imagem:
        plt.savefig('analise_estoque_scatter.png', dpi=300, bbox_inches='tight') # Renamed filename
        logger.info("Gráfico de dispersão salvo como 'analise_estoque_scatter.png'")

    plt.show()

    # 4.5.6 Gráfico de barras - Distribuição das estratégias
    plt.figure(figsize=(10, 6))
    estrategia_counts = df['estrategia'].value_counts()
    bars = plt.bar(estrategia_counts.index, estrategia_counts.values, color='steelblue', edgecolor='black')
    plt.title('Distribuição de Produtos por Estratégia', fontsize=14, fontweight='bold')
    plt.xlabel('Estratégia', fontsize=12)
    plt.ylabel('Quantidade de Produtos', fontsize=12)
    plt.xticks(rotation=45, ha='right')

    # 4.5.7 Adiciona valores nas barras
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height,
                f'{int(height)}', ha='center', va='bottom')

    plt.tight_layout()
    if salvar_imagem:
        plt.savefig('analise_estoque_bars.png', dpi=300, bbox_inches='tight') # Added save for second plot
        logger.info("Gráfico de barras salvo como 'analise_estoque_bars.png'")
    plt.show()

    logger.info("Gráficos gerados com sucesso")

# 4.6 Gera gráfico de feixa de produtos
def gerar_grafico_faixas(df):
    """Gera um gráfico de barras com formatação de moeda brasileira e milhar com ponto"""
    plt.figure(figsize=(12, 6))

    # 4.6.1 Agrupa o valor que está "parado" em estoque por faixa
    analise_faixa = df.groupby('faixa_margem')['custo_corrigido'].sum().reset_index()

    # 4.6.2 Capturando o retorno do gráfico na variável 'ax'
    ax = sns.barplot(data=analise_faixa, x='faixa_margem', y='custo_corrigido', palette='magma')

    plt.title('Capital Investido por Faixa de Margem (Visão Comercial Avenida)', fontsize=14, fontweight='bold')
    plt.ylabel('Total em Estoque (R$)', fontsize=12)
    plt.xlabel('Faixas de Margem Real', fontsize=12)
    plt.xticks(rotation=30)

    # 4.6.3 Formata os rótulos do eixo Y (lateral do gráfico)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')))

    # 4.6.4 Adiciona o valor em cima das barras
    for i, v in enumerate(analise_faixa['custo_corrigido']):
        # Formatação manual: milhar com ponto e decimal com vírgula
        valor_formatado = f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        plt.text(i, v + (v * 0.02), valor_formatado, ha='center', fontweight='bold', fontsize=10)

    plt.tight_layout()
    plt.show()

# 4.7 Cria PDF com gráficos
def gerar_pdf_graficos(
    nome_arquivo: str = "Relatorio_Graficos_Estoque.pdf",
    titulo: str = "ANÁLISE GRÁFICA DO ESTOQUE - COMERCIAL AVENIDA"
) -> bool:
    """
    Gera um PDF com os gráficos da análise.
    Args:
        nome_arquivo: Nome do arquivo PDF a ser gerado
        titulo: Título do relatório
    Returns:
        True se gerado com sucesso, False caso contrário
    """
    try:
        pdf = FPDF()
        pdf.add_page(orientation='L')

        pdf.set_font("Arial", 'B', 14)
        pdf.cell(280, 10, titulo, ln=True, align='C')
        pdf.ln(5)

        # 4.7.1 Adicionar o primeiro gráfico (dispersão)
        pdf.add_page(orientation='L')
        pdf.image('analise_estoque_scatter.png', x=10, y=20, w=270)
        pdf.ln(5)
        pdf.set_font("Arial", '', 10)
        pdf.cell(280, 10, "Gráfico 1: Análise Correlacionada: Tempo de Compra vs. Margem Real", ln=True, align='C')

        # 4.7.2 Adicionar o segundo gráfico (barras)
        pdf.add_page(orientation='L')
        pdf.image('analise_estoque_bars.png', x=10, y=20, w=270)
        pdf.ln(5)
        pdf.set_font("Arial", '', 10)
        pdf.cell(280, 10, "Gráfico 2: Distribuição de Produtos por Estratégia", ln=True, align='C')

        # 4.7.3 Rodapé com data de geração
        pdf.set_y(-15)
        pdf.set_font("Arial", 'I', 8)
        pdf.cell(280, 10, f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                ln=True, align='C')

        pdf.output(nome_arquivo)
        logger.info(f"PDF de gráficos gerado com sucesso: {nome_arquivo}")
        return True

    except Exception as e:
        logger.error(f"Erro ao gerar PDF de gráficos: {str(e)}")
        return False

#---------------------------------------------------------
# --- CÉLULA 5: FUNÇÃO PRINCIPAL (bota tudo pra rodar) ---
#---------------------------------------------------------

# 5.1 Função que executa todo o script
def main(caminho_arquivo: str = '/content/sample_data/base estoque parado.xlsx'):
    """
    Função principal que orquestra toda a análise.
    Args:
        caminho_arquivo: Caminho para o arquivo Excel com os dados
    """
    logger.info("="*60)
    logger.info("INICIANDO ANÁLISE DE PREÇO POR OPORTUNIDADE")
    logger.info("="*60)

    try:
        # 5.1.2 Carregamento dos dados
        df = carregar_dados_excel(caminho_arquivo)

        # 5.1.3 Padroniza o nome das colunas: remove espaços em branco e coloca tudo em minusculo
        df.columns = df.columns.str.strip().str.lower()
        logger.info(f"Colunas do DataFrame após normalização: {df.columns.tolist()}")

        # 5.1.4 Renomeia as colunas para relatório
        column_mapping = {
            'estoquefisico': 'Estoque',
            'custo': 'Custo',
            'preço': 'Preço',
            'preco': 'Preço', # Adicionado para robustez
            'ultima entrada': 'Ultima Entrada',
            'ultima saída': 'Ultima Saída',
            'descrição': 'Descrição',
            'descricao': 'Descrição', # Adicionado para robustez
            'sku': 'SKU'
        }

        # 5.1.5 Aplica o mapeamento de renomeação
        renamed_columns = {}
        for old_name_lower, new_name_capitalized in column_mapping.items():
            if old_name_lower in df.columns:
                # Somente renomeia se o nome capitalizado ainda não existe
                if new_name_capitalized not in df.columns:
                    renamed_columns[old_name_lower] = new_name_capitalized

        if renamed_columns:
            df = df.rename(columns=renamed_columns)
        logger.info(f"Colunas do DataFrame após renomeação final: {df.columns.tolist()}")

        # 5.1.6 Verificar se as colunas essenciais estão presentes após o renomeio
        required_cols = ['Estoque', 'Custo', 'Preço', 'Ultima Entrada', 'Ultima Saída', 'Descrição', 'SKU']
        for col in required_cols:
            if col not in df.columns:
                logger.error(f"A coluna essencial '{col}' não foi encontrada no DataFrame após o processamento. Colunas disponíveis: {df.columns.tolist()}")
                raise KeyError(f"Coluna '{col}' não encontrada após renomeação.")

        # 5.2 Preparação das datas e métricas de tempo
        df = preparar_datas(df)

        # 5.3 Cálculo das métricas financeiras
        df = calcular_metricas_financeiras(df)

        # 5.4 Aplicação das estratégias
        df = aplicar_estrategias(df)

        # 5.5 Geração de resumo executivo
        resumo = gerar_resumo_estoque(df)
        imprimir_resumo_estoque(resumo)

        # 5.6 Exportação para Google Sheets
        exportar_para_google_sheets(df, "Relatorio_Custo_Oportunidade_Avenida")

        # 5.7 Geração do PDF com ações imediatas
        df_acoes = df[(df['estrategia'] != 'Preço Saudável') & (df['Estoque'] > 0)]
        gerar_pdf_acoes(df_acoes, resumo_executivo=resumo)

        # 5.8. Geração dos gráficos e PDF de gráficos
        # Limita a margem para o intervalo de -50% a 100% apenas para o gráfico
        df_plot = df.copy()
        df_plot['margem_visual'] = df_plot['margem_visual'].clip(lower=-50, upper=100)
        gerar_grafico_analise(df, salvar_imagem=True)
        gerar_grafico_faixas(df)
        gerar_pdf_graficos()

        # 5.9 Exibição dos primeiros registros para conferência
        print("\n" + "="*60)
        print("PRÉVIA DOS DADOS PROCESSADOS (5 primeiros registros)")
        print("="*60)
        colunas_mostrar = ['SKU', 'Descrição', 'Estoque', 'meses_investimento',
                          'meses_sem_venda', 'margem_real_atual', 'sugestao_preco', 'estrategia']
        display(df[colunas_mostrar].head())

        logger.info("="*60)
        logger.info("ANÁLISE CONCLUÍDA COM SUCESSO!")
        logger.info("="*60)

        return df # Return the DataFrame

    except Exception as e:
        logger.error(f"Erro fatal na execução da análise: {str(e)}")
        raise

#---------------------------
# --- CÉLULA 6: EXECUÇÃO ---
#---------------------------

if __name__ == "__main__":
    # Executa a análise e armazena o DataFrame resultante em uma variável global 'df'
    df = main('/content/sample_data/base estoque parado - Página1.csv')

df_margem_negativa = df[df['margem_real_atual'] < -20]

if not df_margem_negativa.empty:
    gerar_pdf_acoes(
        df_margem_negativa,
        nome_arquivo="Relatorio_Produtos_Margem_Muito_Negativa.pdf",
        titulo="RELATÓRIO DE PRODUTOS COM MARGEM < -20%"
    )
    print("Relatório de produtos com margem muito negativa gerado com sucesso: Relatorio_Produtos_Margem_Muito_Negativa.pdf")
else:
    print("Nenhum produto encontrado com margem real menor que -20%.")
