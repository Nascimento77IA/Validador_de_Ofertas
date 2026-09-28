import streamlit as st
import pandas as pd
import re
from bs4 import BeautifulSoup

# --- FUNÇÕES DE LIMPEZA E FORMATAÇÃO ---
def limpar_preco(valor):
    if pd.isna(valor) or valor == '' or str(valor).upper().strip() == 'X':
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).upper().replace('R$', '').replace(' ', '').strip()
    if not texto or texto == 'NAN':
        return None
    if ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return None

def limpar_plu(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith('.0'):
        texto = texto[:-2]
    return re.sub(r'\D', '', texto)

def formatar_preco(valor):
    if valor is None:
        return ""
    return f"R$ {valor:,.2f}".replace('.', '#').replace(',', '.').replace('#', ',')

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador e Validador de Tabloide - Nagumo", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚀 Gerador Automático de Tabloide HTML (Versão 3.0)")
st.markdown("O sistema lê a planilha, injeta os códigos PLU e atualiza os preços automaticamente no HTML.")

# --- INTERFACE DO USUÁRIO ---
cluster = st.selectbox("1. Qual cluster você deseja processar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha de Ofertas (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail Original (HTML)", type=["html"])

if st.button("⚙️ Gerar HTML Atualizado com PLUs e Preços"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Processando planilha e atualizando o HTML..."):
            
            arquivo_excel.seek(0)
            try:
                todas_abas = pd.read_excel(arquivo_excel, sheet_name=None, header=None)
            except Exception as e:
                st.error(f"Erro ao ler o Excel: {e}")
                st.stop()
                
            df_raw = None
            linha_plu_idx = -1
            
            for nome_aba, df_aba in todas_abas.items():
                for i, row in df_aba.head(20).iterrows():
                    row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                    if 'PLU' in row_str:
                        linha_plu_idx = i
                        df_raw = df_aba
                        break
                if linha_plu_idx != -1:
                    break
                    
            if df_raw is None or linha_plu_idx == -1:
                st.error("❌ Erro: Não foi encontrada a coluna PLU na planilha.")
                st.stop()
                
            # Mapeamento do cabeçalho duplo
            row_main = df_raw.iloc[linha_plu_idx].values
            row_sub = df_raw.iloc[linha_plu_idx + 1].values
            
            col_plu_idx = -1
            sp_preco_idx = -1; sp_oferta_idx = -1; sp_cartao_idx = -1
            rio_preco_idx = -1; rio_oferta_idx = -1; rio_cartao_idx = -1
            
            current_cluster = ""
            for i in range(len(row_main)):
                if pd.notna(row_main[i]):
                    val_main = str(row_main[i]).upper()
                    if "PLU" in val_main:
                        col_plu_idx = i
                    elif "NAGUMO" in val_main:
                        current_cluster = val_main
                
                sub_val = str(row_sub[i]).upper() if pd.notna(row_sub[i]) else ""
                
                if "NAGUMO SP" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val: sp_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: sp_oferta_idx = i
                    elif "CART" in sub_val or "MEU NAGUMO" in sub_val: sp_cartao_idx = i
                elif "NAGUMO RIO" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val: rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: rio_oferta_idx = i
                    elif "CART" in sub_val or "MEU NAGUMO" in sub_val: rio_cartao_idx = i

            if cluster == "Lojas (SP)":
                idx_preco, idx_oferta, idx_cartao = sp_preco_idx, sp_oferta_idx, sp_cartao_idx
            else:
                idx_preco, idx_oferta, idx_cartao = rio_preco_idx, rio_oferta_idx, rio_cartao_idx

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            # Extrair lista de produtos da planilha ordenados
            produtos_excel = []
            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                plu = limpar_plu(row.iloc[col_plu_idx])
                if not plu:
                    continue
                
                de_val = limpar_preco(row.iloc[idx_preco]) if idx_preco != -1 and idx_preco < len(row) else None
                of_val = limpar_preco(row.iloc[idx_oferta]) if idx_oferta != -1 and idx_oferta < len(row) else None
                ca_val = limpar_preco(row.iloc[idx_cartao]) if idx_cartao != -1 and idx_cartao < len(row) else None
                
                produtos_excel.append({
                    'plu': plu,
                    'de': de_val,
                    'oferta': of_val,
                    'cartao': ca_val
                })

            # --- PROCESSAR O HTML ---
            arquivo_html.seek(0)
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Encontrar blocos de produtos no HTML (geralmente tabelas ou links principais de produtos)
            # Vamos procurar tags <a> ou <table> que representem os itens
            links_produtos = soup.find_all('a', title=True) # ou outra heurística padrão do seu template
            if not links_produtos:
                links_produtos = soup.find_all('table', attrs={'width': True}) # alternativa genérica

            st.info(f"Itens encontrados na planilha: {len(produtos_excel)} | Blocos encontrados no HTML: {len(links_produtos)}")

            # Injetar PLU e atualizar preços sequencialmente por ordem dos blocos
            atualizados = 0
            for idx, item in enumerate(produtos_excel):
                if idx < len(links_produtos):
                    link = links_produtos[idx]
                    plu = item['plu']
                    
                    # 1. Injeta o PLU como ID no elemento HTML
                    link['id'] = plu
                    
                    # 2. Encontra o container pai (card da tabela do produto)
                    container = link.find_parent('table') or link
                    
                    # 3. Atualiza o preço "DE" e "OFERTA" se existirem no texto do card
                    texto_container = str(container)
                    
                    # Substituição cirúrgica dos preços no HTML usando regex ou alteração de nós
                    # Vamos atualizar os spans de preço dentro do container
                    spans = container.find_all('span')
                    for span in spans:
                        txt = span.get_text()
                        if 'DE R$' in txt.upper() and item['de'] is not None:
                            span.string = f"DE R$ {formatar_preco(item['de'])}"
                    
                    atualizados += 1

            # Salvar o novo HTML modificado
            novo_html_str = str(soup)
            
            st.success(f"✅ Sucesso! {atualizados} produtos tiveram seus PLUs injetados e preços atualizados no HTML.")
            
            # Botão de Download do arquivo gerado
            st.download_button(
                label="📥 Baixar HTML Atualizado com PLUs",
                data=novo_html_str,
                file_name=f"tabloide_{cluster.lower().replace(' ', '_')}_com_plus.html",
                mime="text/html"
            )
    else:
        st.warning("Por favor, suba a planilha Excel e o arquivo HTML.")
