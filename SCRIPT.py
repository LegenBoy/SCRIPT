import streamlit as st
import pandas as pd
import os
import re
import shutil
import copy
import warnings
import tempfile
import urllib.parse
from io import BytesIO
from openpyxl import load_workbook

# Ignorar avisos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================= CONFIGURAÇÕES DE LAYOUT =================
NOME_MODELO_PADRAO = "MODELO.xlsx" 
RELACAO_EMAILS = "Relação E-mail Transportadoras.xlsx"

LINHA_INICIAL_DADOS = 7 
PASSO_ENTRE_ROTAS = 4    
MARGEM_COLUNAS = 2        

# ================= FUNÇÕES DE APOIO =================
def formatar_cidade(texto):
    if pd.isna(texto) or str(texto).strip() == "":
        return None
    texto = str(texto)
    padrao = r"^(\d+)-(.*?)[/]"
    match = re.search(padrao, texto)
    if match:
        return f"{match.group(2).strip()} - {match.group(1)}"
    return texto

def carregar_dicionario_emails():
    """Lê a planilha de e-mails e agrupa por transportadora."""
    if not os.path.exists(RELACAO_EMAILS):
        return {}
    try:
        # Tenta ler a primeira aba do arquivo Excel
        df_em = pd.read_excel(RELACAO_EMAILS)
        df_em['TRANSPORTADORA'] = df_em['TRANSPORTADORA'].astype(str).str.strip().str.upper()
        df_em['E-MAIL'] = df_em['E-MAIL'].astype(str).str.strip()
        
        # Agrupa múltiplos e-mails da mesma transportadora separados por ponto e vírgula
        return df_em.groupby('TRANSPORTADORA')['E-MAIL'].apply(lambda x: ';'.join(x)).to_dict()
    except Exception as e:
        st.error(f"Erro ao carregar dicionário de e-mails: {e}")
        return {}

def gerar_link_outlook(transportadora, data_carregamento, destinatarios):
    """Gera o link de deeplink para o Outlook Web com codificação correta (%20)."""
    assunto = f"Zema - Previsão de Descarga {data_carregamento}"
    corpo = (
        "Prezados, Bom Dia!\n\n"
        "Segue previsão de descarga referente carregamento de hoje.\n\n"
        "Dúvidas estou à disposição.\n\n"
        "Obrigado."
    )
    
    # quote garante que espaços virem %20 e não o símbolo de '+'
    assunto_u = urllib.parse.quote(assunto)
    corpo_u = urllib.parse.quote(corpo)
    dest_u = urllib.parse.quote(destinatarios)
    
    # Montagem manual da URL para controle total da codificação
    link = (
        f"https://outlook.cloud.microsoft/mail/0/deeplink/compose?"
        f"path=%2Fmail%2Faction%2Fcompose&to={dest_u}&subject={assunto_u}&body={corpo_u}"
    )
    return link

# ================= FUNÇÕES DE ESTILO EXCEL =================
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
    linha_atual_verificacao = LINHA_INICIAL_DADOS
    rotas_formatadas_existentes = 0
    max_row_excel = ws.max_row
    
    while True:
        if linha_atual_verificacao > max_row_excel + 100: break
        cell = ws.cell(row=linha_atual_verificacao, column=2)
        if not cell.border.left.style: break
        rotas_formatadas_existentes += 1
        linha_atual_verificacao += PASSO_ENTRE_ROTAS
    
    if rotas_formatadas_existentes < total_rotas + 10:
        rotas_faltantes = (total_rotas - rotas_formatadas_existentes) + 50 
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
                ws.cell(row, col_atual).value = valor
                return merged_range.max_col + 1
            col_atual = merged_range.max_col + 1
        else:
            ws.cell(row, col_atual).value = valor
            return col_atual + 1

def limpar_sobras_total(ws, ultima_linha_usada, ultima_coluna_usada):
    linha_inicio_corte = ultima_linha_usada + PASSO_ENTRE_ROTAS
    if ws.max_row >= linha_inicio_corte:
        ws.delete_rows(linha_inicio_corte, (ws.max_row - linha_inicio_corte) + 200)
    coluna_inicio_corte = ultima_coluna_usada + MARGEM_COLUNAS + 1
    if ws.max_column >= coluna_inicio_corte:
        ws.delete_cols(coluna_inicio_corte, (ws.max_column - coluna_inicio_corte) + 50)

def ajustar_largura_colunas(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except: pass
        if max_length > 0:
            ws.column_dimensions[column].width = max_length + 3

def processar_arquivo(caminho_arquivo, dataframe):
    try:
        wb = load_workbook(caminho_arquivo)
        ws = wb["MODELO"] if "MODELO" in wb.sheetnames else (wb["IMPRESSÃO"] if "IMPRESSÃO" in wb.sheetnames else wb.active)
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
                cidade_fmt = formatar_cidade(row[col_nome])
                if cidade_fmt:
                    coluna_cursor = escrever_valor(ws, linha_atual, coluna_cursor, cidade_fmt)
        
        max_col_global = max(max_col_global, coluna_cursor - 1)
        linha_atual += PASSO_ENTRE_ROTAS

    limpar_sobras_total(ws, linha_atual - PASSO_ENTRE_ROTAS, max_col_global)
    ajustar_largura_colunas(ws)
    wb.save(caminho_arquivo)
    return True

# ================= INTERFACE STREAMLIT =================

def main():
    st.set_page_config(page_title="Gerador de Rotas", layout="centered")
    st.title("🚛 Previsão de Descarga")

    if not os.path.exists(NOME_MODELO_PADRAO):
        st.error(f"ERRO: '{NOME_MODELO_PADRAO}' não encontrado.")
        return

    # Inicializa estados de sessão
    if 'arquivos_prontos' not in st.session_state:
        st.session_state['arquivos_prontos'] = []
    if 'data_carregamento' not in st.session_state:
        st.session_state['data_carregamento'] = "Data"

    contatos_dict = carregar_dicionario_emails()
    arquivo_upload = st.file_uploader("Selecione a planilha de dados (Excel)", type=["xlsx"])

    if st.session_state['arquivos_prontos']:
        if st.button("🔄 Novo Processamento"):
            st.session_state['arquivos_prontos'] = []
            st.rerun()

    if arquivo_upload is not None and not st.session_state['arquivos_prontos']:
        if st.button("Processar Arquivos"):
            with st.spinner('Processando...'):
                with tempfile.TemporaryDirectory() as tmpdirname:
                    caminho_input = os.path.join(tmpdirname, "input.xlsx")
                    with open(caminho_input, "wb") as f:
                        f.write(arquivo_upload.getbuffer())

                    try:
                        df = pd.read_excel(caminho_input)
                        # Extrai a data da primeira linha para o assunto do e-mail
                        st.session_state['data_carregamento'] = str(df.iloc[0, 0])[:10]
                        df.columns = [str(c).strip().lower() for c in df.columns]
                    except:
                        st.error("Erro ao ler arquivo de entrada.")
                        return

                    col_transp = next((c for c in df.columns if "transportadora" in c), None)
                    if not col_transp:
                        st.error("Coluna 'transportadora' não encontrada.")
                        return

                    lista_arquivos = []
                    
                    # 1. Processar Arquivo Geral
                    caminho_geral = os.path.join(tmpdirname, "GERAL_ROTAS.xlsx")
                    shutil.copy(NOME_MODELO_PADRAO, caminho_geral)
                    if processar_arquivo(caminho_geral, df):
                        with open(caminho_geral, "rb") as f:
                            lista_arquivos.append({"nome": "GERAL_ROTAS.xlsx", "dados": f.read()})

                    # 2. Processar por Transportadora Individual
                    for transp, dados in df.groupby(col_transp):
                        if pd.isna(transp): continue
                        nome_limpo = str(transp).replace("/", "-").replace("\\", "").strip()
                        nome_arq = f"{nome_limpo}.xlsx"
                        caminho_t = os.path.join(tmpdirname, nome_arq)
                        shutil.copy(NOME_MODELO_PADRAO, caminho_t)
                        if processar_arquivo(caminho_t, dados):
                            with open(caminho_t, "rb") as f:
                                lista_arquivos.append({"nome": nome_arq, "dados": f.read()})

                    st.session_state['arquivos_prontos'] = lista_arquivos
                    st.success("Arquivos processados com sucesso!")
                    st.rerun()

    # Exibição de Download e Botão de E-mail
    if st.session_state['arquivos_prontos']:
        st.divider()
        st.subheader("📂 Arquivos Gerados:")
        
        for item in st.session_state['arquivos_prontos']:
            c1, c2 = st.columns([3, 2])
            nome_transp_key = item['nome'].replace(".xlsx", "")
            
            with c1:
                st.download_button(
                    label=f"📥 Baixar {item['nome']}",
                    data=item['dados'],
                    file_name=item['nome'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{item['nome']}"
                )
            
            with c2:
                # O botão de e-mail aparece para transportadoras individuais cadastradas
                if nome_transp_key != "GERAL_ROTAS":
                    destinatarios = contatos_dict.get(nome_transp_key.upper(), "")
                    if destinatarios:
                        link_mail = gerar_link_outlook(
                            nome_transp_key, 
                            st.session_state['data_carregamento'], 
                            destinatarios
                        )
                        st.markdown(f'''
                            <a href="{link_mail}" target="_blank" style="text-decoration: none;">
                                <div style="background-color: #0078d4; color: white; padding: 8px 16px; 
                                border-radius: 5px; text-align: center; font-weight: bold; font-size: 14px; 
                                cursor: pointer; transition: 0.3s;">
                                    📧 Preparar E-mail
                                </div>
                            </a>''', unsafe_allow_html=True)
                    else:
                        st.info("E-mail não cadastrado")

if __name__ == "__main__":
    main()
