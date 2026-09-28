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
st.set_page_config(page_title="Validador de Ofertas CRM - Nagumo", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📝 Validador de Ofertas CRM - Nagumo")
st.markdown("Validação estrita de preços cruzando a planilha Excel com o e-mail HTML na ordem correta.")

# --- INTERFACE DO USUÁRIO ---
cluster = st.selectbox("1. Qual cluster você deseja validar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha de Ofertas (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail (HTML)", type=["html"])

if st.button("🔍 Validar Preços"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Validando preços..."):
            
            # --- LER EXCEL ---
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
                st.error("❌ Erro: Coluna PLU não encontrada na planilha.")
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
                    if "PLU" in val_main: col_plu_idx = i
                    elif "NAGUMO" in val_main: current_cluster = val_main
                
                sub_val = str(row_sub[i]).upper() if pd.notna(row_sub[i]) else ""
                
                if "NAGUMO SP" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val: sp_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: sp_oferta_idx = i
                    elif "CART" in sub_val or "MEU NAGUMO" in sub_val: sp_cartao_idx = i
                elif "NAGUMO RIO" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val: rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: rio_oferta_idx = i
                    elif "CART" in sub_val or "MEU NAGUMO" in sub_val: rio_cartao_idx = i

            idx_preco = sp_preco_idx if cluster == "Lojas (SP)" else rio_preco_idx
            idx_oferta = sp_oferta_idx if cluster == "Lojas (SP)" else rio_oferta_idx
            idx_cartao = sp_cartao_idx if cluster == "Lojas (SP)" else rio_cartao_idx

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            produtos_excel = []
            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                plu = limpar_plu(row.iloc[col_plu_idx])
                if not plu: continue
                
                nome = str(row.iloc[col_plu_idx + 1]) if col_plu_idx + 1 < len(row) else "Produto"
                de = limpar_preco(row.iloc[idx_preco]) if idx_preco != -1 and idx_preco < len(row) else None
                oferta = limpar_preco(row.iloc[idx_oferta]) if idx_oferta != -1 and idx_oferta < len(row) else None
                cartao = limpar_preco(row.iloc[idx_cartao]) if idx_cartao != -1 and idx_cartao < len(row) else None
                
                produtos_excel.append({'plu': plu, 'nome': nome, 'de': de, 'oferta': oferta, 'cartao': cartao})

            # --- LER HTML ---
            arquivo_html.seek(0)
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Extrair os blocos de produtos do HTML em ordem estrita
            links = soup.find_all('a', title=True)
            blocos = links if links else soup.find_all('table', attrs={'width': True})
            
            produtos_html = []
            for el in blocos:
                container = el.find_parent('table') or el
                txt = str(container)
                
                # Extrair Preço De
                m_de = re.search(r'DE\s*R\$\s*([\d,]+)', txt)
                de_val = float(m_de.group(1).replace(',', '.')) if m_de else None
                
                # Extrair Preço Oferta
                m_of = re.search(r'background-color:#D50037[^>]*>.*?R\$\s*([\d,]+)', txt, re.DOTALL)
                if not m_of:
                    m_of = re.search(r'R\$\s*([\d,]+)', txt)
                of_val = float(m_of.group(1).replace(',', '.')) if m_of else None
                
                produtos_html.append({'de': de_val, 'oferta': of_val})

            # --- CRUZAMENTO E VALIDAÇÃO ---
            st.markdown("### 📊 Relatório de Validação")
            
            erros = []
            validados = 0
            
            total_comparar = min(len(produtos_excel), len(produtos_html))
            for i in range(total_comparar):
                ex = produtos_excel[i]
                ht = produtos_html[i]
                validados += 1
                
                # Compara DE
                if ht['de'] is not None and ex['de'] is not None and abs(ht['de'] - ex['de']) > 0.01:
                    erros.append(f"**ERRO PREÇO (DE)** | Item {i+1}: **{ex['nome']}** (PLU: {ex['plu']}) | HTML: R${ht['de']:.2f} | Excel: R${ex['de']:.2f}")
                
                # Compara OFERTA
                if ht['oferta'] is not None and ex['oferta'] is not None and abs(ht['oferta'] - ex['oferta']) > 0.01:
                    erros.append(f"**ERRO OFERTA** | Item {i+1}: **{ex['nome']}** (PLU: {ex['plu']}) | HTML: R${ht['oferta']:.2f} | Excel: R${ex['oferta']:.2f}")

            if not erros:
                st.success(f"✅ Perfeito! Todos os {validados} produtos conferem rigorosamente entre o Excel e o HTML para o cluster **{cluster}**!")
            else:
                for e in erros:
                    st.error(e)
    else:
        st.warning("Por favor, faça o upload da planilha Excel e do arquivo HTML.")
