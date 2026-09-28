import streamlit as st
import pandas as pd
import re
from bs4 import BeautifulSoup

# --- FUNÇÕES DE LIMPEZA ---
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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador de Ofertas CRM - Nagumo/BWCA", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

# Título alterado para 2.1 para confirmar a atualização
st.title("📝 Validador de Ofertas CRM (Versão 2.1)")
st.markdown("Faça o upload dos arquivos da campanha para validar automaticamente os preços de Lojas (SP) ou Rio. Validação 100% ancorada pelo PLU.")

# --- INTERFACE DO USUÁRIO ---
cluster = st.selectbox("1. Qual cluster você deseja validar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha c/ itens da oferta (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail (HTML)", type=["html"])

if st.button("🔍 Validar Preços"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Analisando e cruzando PLUs..."):
            
            # --- LER EXCEL ---
            try:
                xls = pd.ExcelFile(arquivo_excel)
            except Exception as e:
                st.error(f"Erro ao ler o arquivo Excel: {e}")
                st.stop()
                
            df_raw = None
            linha_cab = -1
            aba_encontrada = ""
            
            # 1. Procura em todas as abas (sheets) pelo cabeçalho correto
            for sheet in xls.sheet_names:
                df_temp = pd.read_excel(xls, sheet_name=sheet, header=None)
                for i, row in df_temp.head(30).iterrows(): # Busca nas 30 primeiras linhas de cada aba
                    row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                    if 'PLU' in row_str and 'OFERTA' in row_str:
                        linha_cab = i
                        df_raw = df_temp
                        aba_encontrada = sheet
                        break
                if df_raw is not None:
                    break
                    
            if df_raw is None or linha_cab == -1:
                st.error("Erro: Não achei o cabeçalho com as palavras 'PLU' e 'OFERTA' em NENHUMA aba da planilha. Verifique a formatação do Excel.")
                st.stop()
                
            st.info(f"✅ Cabeçalho encontrado automaticamente na aba: **{aba_encontrada}** (Linha {linha_cab + 1})")
                
            # 2. Mapear colunas SP vs RIO de forma dinâmica
            columns = df_raw.iloc[linha_cab].fillna("").astype(str).str.upper().str.strip()
            df_excel = df_raw.iloc[linha_cab+1:].reset_index(drop=True)
            
            col_plu_idx = -1
            sp_preco_idx = -1; sp_oferta_idx = -1; sp_cartao_idx = -1
            rio_preco_idx = -1; rio_oferta_idx = -1; rio_cartao_idx = -1
            
            preco_count = 0; oferta_count = 0; cartao_count = 0
            
            for i, col_name in enumerate(columns):
                if 'PLU' in col_name and col_plu_idx == -1:
                    col_plu_idx = i
                elif 'PREÇO' in col_name or 'PRECO' in col_name:
                    if preco_count == 0: sp_preco_idx = i
                    elif preco_count == 1: rio_preco_idx = i
                    preco_count += 1
                elif 'OFERTA' in col_name and 'CART' not in col_name and 'NAGUMO' not in col_name:
                    if oferta_count == 0: sp_oferta_idx = i
                    elif oferta_count == 1: rio_oferta_idx = i
                    oferta_count += 1
                elif 'CART' in col_name or 'MEU NAGUMO' in col_name:
                    if cartao_count == 0: sp_cartao_idx = i
                    elif cartao_count == 1: rio_cartao_idx = i
                    cartao_count += 1
                        
            # Selecionar as colunas corretas baseado na escolha do usuário
            if cluster == "Lojas (SP)":
                idx_preco = sp_preco_idx
                idx_oferta = sp_oferta_idx
                idx_cartao = sp_cartao_idx
            else:
                idx_preco = rio_preco_idx
                idx_oferta = rio_oferta_idx
                idx_cartao = rio_cartao_idx

            # 3. Guardar os dados do Excel
            mapa_excel = {}
            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or pd.isna(row.iloc[col_plu_idx]):
                    continue
                    
                plu_limpo = limpar_plu(row.iloc[col_plu_idx])
                if not plu_limpo:
                    continue
                
                de_val = limpar_preco(row.iloc[idx_preco]) if idx_preco != -1 else None
                of_val = limpar_preco(row.iloc[idx_oferta]) if idx_oferta != -1 else None
                ca_val = limpar_preco(row.iloc[idx_cartao]) if idx_cartao != -1 else None
                
                mapa_excel[plu_limpo] = {
                    'de': de_val,
                    'oferta': of_val,
                    'cartao': ca_val
                }

            # --- LER HTML ---
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            produtos_html = []
            links_com_plu = soup.find_all(id=True)
            
            ids_processados = set()
            for el in links_com_plu:
                plu_html = limpar_plu(el.get('id'))
                if not plu_html or plu_html in ids_processados:
                    continue
                
                ids_processados.add(plu_html)
                nome_prod_html = el.get('title', el.get_text(strip=True)[:40] or 'Produto no HTML')
                
                container_pai = el.find_parent('table') or el.find_parent('td') or el
                html_item_str = str(container_pai)

                m_de = re.search(r'DE\s*R\$\s*([\d,]+)', html_item_str)
                de_preco = float(m_de.group(1).replace(',', '.')) if m_de else None

                m_of = re.search(r'background-color:#D50037[^>]*>.*?R\$\s*([\d,]+)', html_item_str, re.DOTALL)
                if not m_of:
                    m_of = re.search(r'R\$\s*([\d,]+)', html_item_str)
                oferta_preco = float(m_of.group(1).replace(',', '.')) if m_of else None

                m_card = re.search(r'Cartão Nagumo.*?R\$\s*([\d,]+)', html_item_str, re.DOTALL)
                if not m_card:
                    m_card = re.search(r'Preço Exclusivo.*?R\$\s*([\d,]+)', html_item_str, re.DOTALL)
                cartao_preco = float(m_card.group(1).replace(',', '.')) if m_card else None

                produtos_html.append({
                    'plu': plu_html,
                    'nome': nome_prod_html,
                    'de': de_preco,
                    'oferta': oferta_preco,
                    'cartao': cartao_preco
                })

            # --- CRUZAMENTO E VALIDAÇÃO ---
            st.markdown("### Resultados da Validação (Por PLU)")
            
            erros = []
            alertas = []
            validados_count = 0

            for p_html in produtos_html:
                plu_h = p_html['plu']
                nome_h = p_html['nome']
                
                if plu_h in mapa_excel:
                    dados_excel = mapa_excel[plu_h]
                    validados_count += 1
                    
                    if p_html['de'] is not None and dados_excel['de'] is not None and abs(p_html['de'] - dados_excel['de']) > 0.01:
                        erros.append(f"**ERRO PREÇO (DE)** | {nome_h} (PLU: {plu_h}) | HTML: R${p_html['de']:.2f} | Excel: R${dados_excel['de']:.2f}")
                    
                    if p_html['oferta'] is not None and dados_excel['oferta'] is not None and abs(p_html['oferta'] - dados_excel['oferta']) > 0.01:
                        erros.append(f"**ERRO OFERTA** | {nome_h} (PLU: {plu_h}) | HTML: R${p_html['oferta']:.2f} | Excel: R${dados_excel['oferta']:.2f}")
                    
                    if p_html['cartao'] is not None:
                        if dados_excel['cartao'] is not None:
                            if abs(p_html['cartao'] - dados_excel['cartao']) > 0.01:
                                erros.append(f"**ERRO CARTÃO** | {nome_h} (PLU: {plu_h}) | HTML: R${p_html['cartao']:.2f} | Excel: R${dados_excel['cartao']:.2f}")
                        else:
                            erros.append(f"**ERRO CARTÃO** | {nome_h} (PLU: {plu_h}) | HTML tem Cartão (R${p_html['cartao']:.2f}), mas no Excel está vazio.")
                else:
                    alertas.append(f"⚠️ **PLU {plu_h} não encontrado:** O HTML pediu o PLU {plu_h}, mas ele não está na planilha.")

            # --- EXIBIÇÃO ---
            if validados_count == 0:
                st.warning("⚠️ Nenhum PLU bateu. Confirme se os IDs estão corretos nas tags do HTML.")
            else:
                if not erros and not alertas:
                    st.success(f"✅ Perfeito! Todos os {validados_count} produtos batem rigorosamente com o Excel.")
                
                for e in erros:
                    st.error(e)
                    
                for a in alertas:
                    st.warning(a)
    else:
        st.warning("Por favor, suba os dois arquivos para iniciar.")
