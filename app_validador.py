import streamlit as st
import pandas as pd
import re
from bs4 import BeautifulSoup
import unicodedata

# --- FUNÇÕES DE LIMPEZA E NORMALIZAÇÃO ---
def normalizar_texto(texto):
    if not texto:
        return ""
    # Remove acentos, espaços extras e deixa em maiúsculo para comparar descrições perfeitamente
    nfkd = unicodedata.normalize('NFKD', str(texto))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^A-Z0-9]', '', sem_acento.upper())

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
st.set_page_config(page_title="Validador e Injetor de PLU - Nagumo", layout="wide", page_icon="📝")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📝 Validador de Ofertas & Injetor de PLU - Nagumo")
st.markdown("O sistema cruza por descrição, injeta o PLU no HTML e valida os preços **De** e **Oferta**.")

# --- INTERFACE DO USUÁRIO ---
cluster = st.selectbox("1. Qual cluster você deseja validar/processar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha de Ofertas (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail Original (HTML)", type=["html"])

if st.button("🚀 Executar Validação e Injeção de PLU"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Analisando planilha, cruzando descrições e validando preços..."):
            
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
                
            # Mapeamento do cabeçalho duplo (SP / Rio)
            row_main = df_raw.iloc[linha_plu_idx].values
            row_sub = df_raw.iloc[linha_plu_idx + 1].values
            
            col_plu_idx = -1; col_prod_idx = -1
            sp_preco_idx = -1; sp_oferta_idx = -1
            rio_preco_idx = -1; rio_oferta_idx = -1
            
            current_cluster = ""
            for i in range(len(row_main)):
                if pd.notna(row_main[i]):
                    val_main = str(row_main[i]).upper()
                    if "PLU" in val_main: col_plu_idx = i
                    elif "PRODUTO" in val_main or "DESC" in val_main: col_prod_idx = i
                    elif "NAGUMO" in val_main: current_cluster = val_main
                
                sub_val = str(row_sub[i]).upper() if pd.notna(row_sub[i]) else ""
                
                if "NAGUMO SP" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val: sp_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: sp_oferta_idx = i
                elif "NAGUMO RIO" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val: rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: rio_oferta_idx = i

            # Se não achou col_prod_idx explicitamente, assume que é logo após o PLU
            if col_prod_idx == -1:
                col_prod_idx = col_plu_idx + 1

            idx_preco = sp_preco_idx if cluster == "Lojas (SP)" else rio_preco_idx
            idx_oferta = sp_oferta_idx if cluster == "Lojas (SP)" else rio_oferta_idx

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            # Mapear dados do Excel por descrição normalizada
            mapa_excel = {}
            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                plu = limpar_plu(row.iloc[col_plu_idx])
                if not plu: 
                    continue
                
                nome = str(row.iloc[col_prod_idx]) if col_prod_idx < len(row) else ""
                de = limpar_preco(row.iloc[idx_preco]) if idx_preco != -1 and idx_preco < len(row) else None
                oferta = limpar_preco(row.iloc[idx_oferta]) if idx_oferta != -1 and idx_oferta < len(row) else None
                
                chave_norm = normalizar_texto(nome)
                if chave_norm:
                    mapa_excel[chave_norm] = {
                        'plu': plu,
                        'nome_original': nome,
                        'de': de,
                        'oferta': oferta
                    }

            # --- LER HTML ---
            arquivo_html.seek(0)
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Encontrar blocos de produtos no HTML (links com title ou tabelas)
            links = soup.find_all('a', title=True)
            blocos = links if links else soup.find_all('table', attrs={'width': True})
            
            erros = []
            encontrados_count = 0
            
            for el in blocos:
                container = el.find_parent('table') or el
                txt_bloco = str(container)
                
                # Pega o título/descrição descrita no HTML (ex: atributo title ou texto interno)
                titulo_html = el.get('title', '')
                if not titulo_html:
                    titulo_html = el.get_text(strip=True)[:50]
                
                chave_html_norm = normalizar_texto(titulo_html)
                
                # Tenta cruzar a descrição do HTML com o dicionário do Excel
                item_excel = None
                for chave_ex, dados in mapa_excel.items():
                    # Verifica correspondência parcial robusta (se parte significativa da string bate)
                    if chave_ex in chave_html_norm or chave_html_norm in chave_ex or len(set(chave_ex.split()).intersection(set(chave_html_norm.split()))) >= 2:
                        item_excel = dados
                        break
                
                if item_excel:
                    encontrados_count += 1
                    plu = item_excel['plu']
                    
                    # 1. INJEÇÃO DO PLU NO HTML
                    el['id'] = plu
                    
                    # 2. EXTRAÇÃO DE PREÇOS DO HTML PARA COMPARAÇÃO
                    # Extrair Preço De
                    m_de = re.search(r'DE\s*R\$\s*([\d,]+)', txt_bloco)
                    de_html = float(m_de.group(1).replace(',', '.')) if m_de else None
                    
                    # Extrair Preço Oferta
                    m_of = re.search(r'background-color:#D50037[^>]*>.*?R\$\s*([\d,]+)', txt_bloco, re.DOTALL)
                    if not m_of:
                        m_of = re.search(r'R\$\s*([\d,]+)', txt_bloco)
                    oferta_html = float(m_of.group(1).replace(',', '.')) if m_of else None
                    
                    # 3. VALIDAÇÃO DOS PREÇOS (DE e OFERTA)
                    nome_prod = item_excel['nome_original']
                    
                    if de_html is not None and item_excel['de'] is not None and abs(de_html - item_excel['de']) > 0.01:
                        erros.append(f"**ERRO PREÇO (DE)** | **{nome_prod}** (PLU: {plu}) | HTML: R${de_html:.2f} | Excel: R${item_excel['de']:.2f}")
                    
                    if oferta_html is not None and item_excel['oferta'] is not None and abs(oferta_html - item_excel['oferta']) > 0.01:
                        erros.append(f"**ERRO OFERTA** | **{nome_prod}** (PLU: {plu}) | HTML: R${oferta_html:.2f} | Excel: R${item_excel['oferta']:.2f}")

            # --- EXIBIÇÃO DOS RESULTADOS ---
            st.markdown("### 📊 Relatório de Validação e Injeção")
            st.info(f"Produtos cruzados com sucesso entre Planilha e HTML: {encontrados_count}")
            
            if not erros:
                st.success(f"✅ Perfeito! Todos os preços conferem rigorosamente e os PLUs foram injetados com sucesso para o cluster **{cluster}**!")
            else:
                for e in erros:
                    st.error(e)
            
            # Botão para baixar o HTML com os PLUs injetados automaticamente
            novo_html_str = str(soup)
            st.download_button(
                label="📥 Baixar HTML com PLUs Injetados",
                data=novo_html_str,
                file_name=f"tabloide_{cluster.lower().replace(' ', '_')}_com_plus.html",
                mime="text/html"
            )
    else:
        st.warning("Por favor, faça o upload da planilha Excel e do arquivo HTML.")
