import streamlit as st
import pandas as pd
import re
from bs4 import BeautifulSoup
import unicodedata

# --- FUNÇÕES DE LIMPEZA E NORMALIZAÇÃO ---
def normalizar_texto(texto):
    if not texto:
        return ""
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
st.markdown("Validação detalhada item a item cruzando a planilha com o HTML.")

# --- INTERFACE DO USUÁRIO ---
cluster = st.selectbox("1. Qual cluster você deseja validar/processar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha de Ofertas (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail Original (HTML)", type=["html"])

if st.button("🚀 Executar Validação Completa"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Analisando e validando item a item..."):
            
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

            if col_prod_idx == -1:
                col_prod_idx = col_plu_idx + 1

            idx_preco = sp_preco_idx if cluster == "Lojas (SP)" else rio_preco_idx
            idx_oferta = sp_oferta_idx if cluster == "Lojas (SP)" else rio_oferta_idx

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

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
            
            links = soup.find_all('a', title=True)
            blocos = links if links else soup.find_all('table', attrs={'width': True})
            
            resultados = []
            
            for el in blocos:
                container = el.find_parent('table') or el
                txt_bloco = str(container)
                
                titulo_html = el.get('title', '')
                if not titulo_html:
                    titulo_html = el.get_text(strip=True)[:50]
                
                chave_html_norm = normalizar_texto(titulo_html)
                
                item_excel = None
                for chave_ex, dados in mapa_excel.items():
                    if chave_ex in chave_html_norm or chave_html_norm in chave_ex or len(set(chave_ex.split()).intersection(set(chave_html_norm.split()))) >= 2:
                        item_excel = dados
                        break
                
                if item_excel:
                    plu = item_excel['plu']
                    el['id'] = plu # Injeta o PLU no HTML
                    
                    # Extração de preços do HTML
                    m_de = re.search(r'DE\s*R\$\s*([\d,]+)', txt_bloco)
                    de_html = float(m_de.group(1).replace(',', '.')) if m_de else None
                    
                    m_of = re.search(r'background-color:#D50037[^>]*>.*?R\$\s*([\d,]+)', txt_bloco, re.DOTALL)
                    if not m_of:
                        m_of = re.search(r'R\$\s*([\d,]+)', txt_bloco)
                    oferta_html = float(m_of.group(1).replace(',', '.')) if m_of else None
                    
                    # Validação
                    status = "✅ OK"
                    detalhes = []
                    
                    de_ex = item_excel['de']
                    of_ex = item_excel['oferta']
                    
                    if de_html is not None and de_ex is not None:
                        if abs(de_html - de_ex) > 0.01:
                            status = "❌ Erro Preço DE"
                            detalhes.append(f"DE HTML R${de_html:.2f} != Planilha R${de_ex:.2f}")
                    
                    if oferta_html is not None and of_ex is not None:
                        if abs(oferta_html - of_ex) > 0.01:
                            status = "❌ Erro Oferta" if status == "✅ OK" else status + " | Erro Oferta"
                            detalhes.append(f"Oferta HTML R${oferta_html:.2f} != Planilha R${of_ex:.2f}")

                    resultados.append({
                        "Produto": item_excel['nome_original'],
                        "PLU": plu,
                        "Planilha (De / Oferta)": f"R$ {de_ex:.2f} / R$ {of_ex:.2f}" if de_ex is not None and of_ex is not None else "Não definido",
                        "HTML (De / Oferta)": f"R$ {de_html:.2f} / R$ {oferta_html:.2f}" if de_html is not None and oferta_html is not None else "Não lido",
                        "Status": status,
                        "Observação": " | ".join(detalhes) if detalhes else "Preços conferem perfeitamente"
                    })

            # --- EXIBIÇÃO DOS RESULTADOS EM TABELA ---
            st.markdown("### 📊 Relatório Detalhado de Validação")
            
            if resultados:
                df_res = pd.DataFrame(resultados)
                st.dataframe(df_res, use_container_width=True)
                
                erros_count = sum(1 for r in resultados if "❌" in r["Status"])
                if erros_count == 0:
                    st.success(f"✅ Validação bem-sucedida! Todos os {len(resultados)} produtos validados conferem rigorosamente entre a planilha e o HTML.")
                else:
                    st.error(f"⚠️ Atenção! Foram encontrados **{erros_count}** produto(s) com divergência de preços no e-mail!")
            else:
                st.warning("⚠️ Nenhum produto em comum foi cruzado entre a planilha e o HTML. Verifique se as descrições batem.")

            # Botão de download do HTML atualizado com os PLUs injetados
            novo_html_str = str(soup)
            st.download_button(
                label="📥 Baixar HTML com PLUs Injetados",
                data=novo_html_str,
                file_name=f"tabloide_{cluster.lower().replace(' ', '_')}_com_plus.html",
                mime="text/html"
            )
    else:
        st.warning("Por favor, faça o upload da planilha Excel e do arquivo HTML.")
