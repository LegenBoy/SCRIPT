import streamlit as st
import pandas as pd
import os
import re
import shutil
import copy
import warnings
import tempfile
from openpyxl import load_workbook

# Ignorar avisos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================= CONFIGURAÇÕES DE LAYOUT =================
# O modelo deve estar na mesma pasta do script no GitHub
NOME_MODELO_PADRAO = "MODELO.xlsx" 

LINHA_INICIAL_DADOS = 7  
PASSO_ENTRE_ROTAS = 4    
MARGEM_COLUNAS = 2       

# ================= FUNÇÕES DE ESTILO (MANTIDAS) =================
def formatar_cidade(texto):
    if pd.isna(texto) or str(texto).strip() == "":
        return None
    texto = str(texto)
    padrao = r"^(\d+)-(.*?)[/]"
    match = re.search(padrao, texto)
    if match:
        return f"{match.group(2).strip()} - {match.group(1)}"
    return texto

def copiar_estilo(celula_origem, celula_destino):
    if celula_origem.has_style:
        celula_destino.font = copy.copy(celula_origem.font)
        celula_destino.border = copy.copy(celula_origem.border)
        celula_destino.fill = copy.copy(celula_origem.fill)
        celula_destino.number_format = copy.copy(celula_origem.number_format)
        celula_destino.protection = copy.copy(celula_origem.protection)
        celula_destino.alignment = copy.copy(celula_origem.alignment)

def replicar_bloco_formatacao(ws, linha_base, linha_nova_inicio, altura_bloco):
    max_col = ws.max_column
    for i in range(altura_bloco):
        src_row = linha_base + i
        dst_row = linha_nova_inicio + i
        for col in range(1, max_col + 1):
            src_cell = ws.cell(row=src_row, column=col)
            dst_cell = ws.cell(row=dst_row, column=col)
            copiar_estilo(src_cell, dst_cell)

    merges_origem = [
        rng for rng in ws.merged_cells.ranges
        if rng.min_row >= linha_base and rng.max_row < (linha_base + altura_bloco)
    ]
    for rng in merges_origem:
        offset = linha_nova_inicio - linha_base
        ws.merge_cells(
            start_row=rng.min_row + offset, start_column=rng.min_col,
            end_row=rng.max_row + offset, end_column=rng.max_col
        )

def preparar_estrutura_linhas(ws, total_rotas):
    ultima_linha_necessaria = LINHA_INICIAL_DADOS + (total_rotas * PASSO_ENTRE_ROTAS)
    max_row_excel = ws.max_row
    
    if ultima_linha_necessaria > max_row_excel:
        linha_atual_verificacao = LINHA_INICIAL_DADOS
        rotas_formatadas_existentes = 0
        while True:
            cell = ws.cell(row=linha_atual_verificacao, column=2)
            if not cell.border.left.style and linha_atual_verificacao > max_row_excel:
                break
            rotas_formatadas_existentes += 1
            linha_atual_verificacao += PASSO_ENTRE_ROTAS
        
        rotas_faltantes = total_rotas - rotas_formatadas_existentes + 5 
        linha_destino = linha_atual_verificacao
        
        for _ in range(rotas_faltantes):
            replicar_bloco_formatacao(ws, LINHA_INICIAL_DADOS, linha_destino, PASSO_ENTRE_ROTAS)
            linha_destino += PASSO_ENTRE_ROTAS

def obter_range_mesclado(ws, row, col):
    for merged_range in ws.merged_cells.ranges:
        if (row >= merged_range.min_row and row <= merged_range.max_row and
            col >= merged_range.min_col and col <= merged_range.max_col):
            return merged_range
    return None

def escrever_valor(ws, row, col_inicial, valor):
    col_atual = col_inicial
    while True:
        merged_range = obter_range_mesclado(ws, row, col_atual)
        if merged_range:
            if row == merged_range.min_row and col_atual == merged_range.min_col:
                celula = ws.cell(row, col_atual)
                celula.value = valor
                return merged_range.max_col + 1
            else:
                col_atual = merged_range.max_col + 1
                continue
        else:
            celula = ws.cell(row, col_atual)
            celula.value = valor
            return col_atual + 1

def limpar_sobras_total(ws, ultima_linha_usada, ultima_coluna_usada):
    linha_inicio_corte = ultima_linha_usada + PASSO_ENTRE_ROTAS
    max_row = ws.max_row
    if max_row >= linha_inicio_corte:
        qtd_linhas_apagar = (max_row - linha_inicio_corte) + 100
        ws.delete_rows(linha_inicio_corte, qtd_linhas_apagar)

    coluna_inicio_corte = ultima_coluna_usada + MARGEM_COLUNAS + 1
    max_col = ws.max_column
    if max_col >= coluna_inicio_corte:
        qtd_cols_apagar = (max_col - coluna_inicio_corte) + 50
        ws.delete_cols(coluna_inicio_corte, qtd_cols_apagar)

def ajustar_largura_colunas(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
            except:
                pass
        if max_length > 0:
            ws.column_dimensions[column].width = max_length + 3

def processar_arquivo(caminho_arquivo, dataframe):
    try:
        wb = load_workbook(caminho_arquivo)
        if "MODELO" in wb.sheetnames: ws = wb["MODELO"]
        elif "IMPRESSÃO" in wb.sheetnames: ws = wb["IMPRESSÃO"]
        else: ws = wb.active
    except Exception as e:
        st.error(f"Erro ao abrir modelo: {e}")
        return False

    preparar_estrutura_linhas(ws, len(dataframe))

    linha_atual = LINHA_INICIAL_DADOS
    max_col_global = 2 

    for index, row in dataframe.iterrows():
        coluna_cursor = 2
        for i in range(1, 13):
            col_nome = f"filial{i}/cubagem"
            if col_nome in dataframe.columns:
                valor = row[col_nome]
                cidade_fmt = formatar_cidade(valor)
                if cidade_fmt:
                    coluna_cursor = escrever_valor(ws, linha_atual, coluna_cursor, cidade_fmt)
        
        coluna_usada_nesta_linha = coluna_cursor - 1
        if coluna_usada_nesta_linha > max_col_global:
            max_col_global = coluna_usada_nesta_linha
            
        linha_atual += PASSO_ENTRE_ROTAS

    ultima_linha_real = linha_atual - PASSO_ENTRE_ROTAS
    limpar_sobras_total(ws, ultima_linha_real, max_col_global)
    ajustar_largura_colunas(ws)
    wb.save(caminho_arquivo)
    return True

# ================= INTERFACE STREAMLIT =================

def main():
    st.set_page_config(page_title="Gerador de Rotas", layout="centered")
    st.title("🚛 Gerador de Previsão de Descarga")
    st.markdown("Faça o upload da planilha base para gerar os arquivos formatados.")

    # 1. Verifica se o MODELO.xlsx existe na pasta do script
    if not os.path.exists(NOME_MODELO_PADRAO):
        st.error(f"ERRO CRÍTICO: O arquivo '{NOME_MODELO_PADRAO}' não foi encontrado na raiz do site.")
        st.info("Por favor, adicione o arquivo MODELO.xlsx ao repositório do GitHub.")
        return

    # 2. Upload do arquivo
    arquivo_upload = st.file_uploader("Selecione a planilha de dados (Excel)", type=["xlsx"])

    if arquivo_upload is not None:
        if st.button("Processar Arquivos"):
            with st.spinner('Processando... Aguarde.'):
                # Cria um diretório temporário para trabalhar
                with tempfile.TemporaryDirectory() as tmpdirname:
                    
                    # Salva o arquivo de upload no temp para o pandas ler
                    caminho_input = os.path.join(tmpdirname, "input.xlsx")
                    with open(caminho_input, "wb") as f:
                        f.write(arquivo_upload.getbuffer())

                    try:
                        df = pd.read_excel(caminho_input)
                        df.columns = [str(c).strip().lower() for c in df.columns]
                    except Exception as e:
                        st.error(f"Erro ao ler arquivo: {e}")
                        return

                    col_transp = next((c for c in df.columns if "transportadora" in c), None)
                    if not col_transp:
                        st.error("Erro: Coluna 'transportadora' não encontrada na planilha.")
                        return

                    # Lista para guardar os arquivos gerados para download
                    arquivos_gerados = []

                    # --- GERA O GERAL ---
                    caminho_geral = os.path.join(tmpdirname, "GERAL_ROTAS.xlsx")
                    shutil.copy(NOME_MODELO_PADRAO, caminho_geral) # Copia do modelo original
                    
                    if processar_arquivo(caminho_geral, df):
                        arquivos_gerados.append(("GERAL_ROTAS.xlsx", caminho_geral))

                    # --- GERA POR TRANSPORTADORA ---
                    for transp, dados in df.groupby(col_transp):
                        if pd.isna(transp): continue
                        nome_arquivo = str(transp).replace("/", "-").replace("\\", "").strip() + ".xlsx"
                        caminho_transp = os.path.join(tmpdirname, nome_arquivo)
                        
                        shutil.copy(NOME_MODELO_PADRAO, caminho_transp)
                        if processar_arquivo(caminho_transp, dados):
                            arquivos_gerados.append((nome_arquivo, caminho_transp))

                    st.success("Processamento concluído! Baixe os arquivos abaixo:")
                    st.divider()

                    # --- EXIBE BOTÕES DE DOWNLOAD ---
                    for nome_arq, caminho_completo in arquivos_gerados:
                        with open(caminho_completo, "rb") as file:
                            btn = st.download_button(
                                label=f"📥 Baixar {nome_arq}",
                                data=file,
                                file_name=nome_arq,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

if __name__ == "__main__":
    main()