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

st.title("📝 Validador de Ofertas CRM (Versão 2.5)")
st.markdown("Validação automática por PLU com deteção inteligente de texto no HTML (Sem necessidade de IDs manuais).")

# --- INTERFACE DO UTILIZADOR ---
cluster = st.selectbox("1. Qual cluster você deseja validar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha c/ itens da oferta (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail (HTML)", type=["html"])

if st.button("🔍 Validar Preços"):
    if arquivo_excel and arquivo_html:
        with st.spinner("A analisar a planilha e a procurar os PLUs no HTML..."):
            
            arquivo_excel.seek(0)
            
            try:
                todas_abas = pd.read_excel(arquivo_excel, sheet_name=None, header=None)
            except Exception as e:
                st.error(f"Erro ao ler o Excel: {e}")
                st.stop()
                
            df_raw = None
            linha_plu_idx = -1
            aba_encontrada = ""
            
            for nome_aba, df_aba in todas_abas.items():
                for i, row in df_aba.head(20).iterrows():
                    row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                    if 'PLU' in row_str:
                        linha_plu_idx = i
                        df_raw = df_aba
                        aba_encontrada = nome_aba
                        break
                if linha_plu_idx != -1:
                    break
                    
            if df_raw is None or linha_plu_idx == -1:
                st.error("❌ Erro: Não foi encontrada a coluna PLU na planilha.")
                st.stop()
                
            st.info(f"✅ Cabeçalho detetado na aba: **{aba_encontrada}** (Linha do PLU: {linha_plu_idx + 1})")
                
            # Tratamento do cabeçalho duplo
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
                    if "PREÇO" in sub_val or "PRECO" in sub_val:
                        sp_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val:
                        sp_oferta_idx = i
                    elif "CART" in sub_val or "MEU NAGUMO" in sub_val:
                        sp_cartao_idx = i
                        
                elif "NAGUMO RIO" in current_cluster:
                    if "PREÇO" in sub_val or "PRECO" in sub_val:
                        rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val:
                        rio_oferta_idx = i
                    elif "CART" in sub_val or "MEU NAGUMO" in sub_val:
                        rio_cartao_idx = i

            if cluster == "Lojas (SP)":
                idx_preco = sp_preco_idx
                idx_oferta = sp_oferta_idx
                idx_cartao = sp_cartao_idx
            else:
                idx_preco = rio_preco_idx
                idx_oferta = rio_oferta_idx
                idx_cartao = rio_cartao_idx

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            # Mapeamento do Excel por PLU
            mapa_excel = {}
            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                    
                plu_limpo = limpar_plu(row.iloc[col_plu_idx])
                if not plu_limpo:
                    continue
                
                nome_produto = str(row.iloc[col_plu_idx + 1]) if col_plu_idx + 1 < len(row) else "Produto"
                de_val = limpar_preco(row.iloc[idx_preco]) if idx_preco != -1 and idx_preco < len(row) else None
                of_val = limpar_preco(row.iloc[idx_oferta]) if idx_oferta != -1 and idx_oferta < len(row) else None
                ca_val = limpar_preco(row.iloc[idx_cartao]) if idx_cartao != -1 and idx_cartao < len(row) else None
                
                mapa_excel[plu_limpo] = {
                    'nome': nome_produto,
                    'de': de_val,
                    'oferta': of_val,
                    'cartao': ca_val
                }

            # --- LER HTML E FAZER BUSCA INTELIGENTE POR TEXTO ---
            arquivo_html.seek(0)
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            erros = []
            alertas = []
            validados_count = 0
            
            for plu_h, dados_excel in mapa_excel.items():
                nome_h = dados_excel['nome']
                
                # Procura elemento por ID, atributo ou texto correspondente ao PLU
                el = soup.find(id=plu_h) or soup.find(attrs={"data-plu": plu_h})
                
                if not el:
                    # Busca avançada: Procura em tabelas ou células onde o número do PLU apareça como texto
                    for table in soup.find_all(['table', 'td', 'div']):
                        if plu_h in table.get_text():
                            el = table
                            # Tenta afunilar para a sub-tabela mais específica se existir
                            sub_tables = table.find_all('table')
                            for st_t in sub_tables:
                                if plu_h in st_t.get_text():
                                    el = st_t
                            break
                            
                if not el:
                    alertas.append(f"⚠️ **PLU {plu_h} ({nome_h})** não foi localizado no HTML.")
                    continue
                    
                validados_count += 1
                html_item_str = str(el)

                # Extração de preços do bloco HTML encontrado
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

                # Comparação rigorosa de preços
                if de_preco is not None and dados_excel['de'] is not None and abs(de_preco - dados_excel['de']) > 0.01:
                    erros.append(f"**ERRO PREÇO (DE)** | {nome_h} (PLU: {plu_h}) | HTML: R${de_preco:.2f} | Excel: R${dados_excel['de']:.2f}")
                
                if oferta_preco is not None and dados_excel['oferta'] is not None and abs(oferta_preco - dados_excel['oferta']) > 0.01:
                    erros.append(f"**ERRO OFERTA** | {nome_h} (PLU: {plu_h}) | HTML: R${oferta_preco:.2f} | Excel: R${dados_excel['oferta']:.2f}")
                
                if cartao_preco is not None:
                    if dados_excel['cartao'] is not None:
                        if abs(cartao_preco - dados_excel['cartao']) > 0.01:
                            erros.append(f"**ERRO CARTÃO** | {nome_h} (PLU: {plu_h}) | HTML: R${cartao_preco:.2f} | Excel: R${dados_excel['cartao']:.2f}")
                    else:
                        erros.append(f"**ERRO CARTÃO** | {nome_h} (PLU: {plu_h}) | HTML tem Cartão (R${cartao_preco:.2f}), mas no Excel está vazio.")

            # --- EXIBIÇÃO DOS RESULTADOS ---
            st.markdown("### Resultados da Validação (Por PLU)")
            
            if validados_count == 0:
                st.warning("⚠️ Nenhum PLU foi encontrado no HTML. Verifique se os números dos códigos constam nos ficheiros HTML.")
            else:
                if not erros and not alertas:
                    st.success(f"✅ Perfeito! Todos os {validados_count} produtos encontrados batem rigorosamente com o Excel para o cluster **{cluster}**.")
                
                for e in erros:
                    st.error(e)
                    
                for a in alertas:
                    st.warning(a)
    else:
        st.warning("Por favor, suba os dois arquivos para iniciar.")
