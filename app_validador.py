import streamlit as st
import pandas as pd
import re
from bs4 import BeautifulSoup

# --- FUNÇÕES DE LIMPEZA ---
def limpar_preco(valor):
    if pd.isna(valor) or str(valor).upper().strip() in ['', 'X', 'NAN']:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).upper().replace('R$', '').replace(' ', '').strip()
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
st.set_page_config(page_title="Validador Puro de Ofertas - Nagumo", layout="wide", page_icon="🔍")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🔍 Validador Puro de Ofertas por PLU (Nagumo SP & Rio)")
st.markdown("Validação estrita cruzando o HTML (com PLUs embutidos) com o novo layout de colunas da planilha.")

# --- INTERFACE DO USUÁRIO ---
cluster = st.selectbox("1. Qual cluster você deseja validar?", ["Lojas (SP)", "Rio"])

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha de Ofertas (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o HTML (com PLUs embutidos)", type=["html"])

if st.button("🚀 Executar Validação de Preços"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Validando preços por PLU com o novo layout..."):
            
            # --- LER EXCEL ---
            arquivo_excel.seek(0)
            todas_abas = pd.read_excel(arquivo_excel, sheet_name=None, header=None)
            
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
                
            # Mapeamento do novo cabeçalho estruturado (Linha principal com clusters e sublinha com Preço/Oferta)
            row_main = df_raw.iloc[linha_plu_idx].values
            row_sub = df_raw.iloc[linha_plu_idx + 1].values if linha_plu_idx + 1 < len(df_raw) else []
            
            col_plu_idx = -1; col_prod_idx = -1
            sp_preco_idx = -1; sp_oferta_idx = -1
            rio_preco_idx = -1; rio_oferta_idx = -1
            
            cluster_atual = ""
            for i in range(len(row_main)):
                val_main = str(row_main[i]).upper() if pd.notna(row_main[i]) else ""
                
                if "PLU" in val_main: 
                    col_plu_idx = i
                elif "PRODUTO" in val_main or "DESC" in val_main: 
                    col_prod_idx = i
                elif "NAGUMO SP" in val_main: 
                    cluster_atual = "SP"
                elif "NAGUMO RIO" in val_main: 
                    cluster_atual = "RIO"
                elif "NAGUMO MIXTER" in val_main: 
                    cluster_atual = "MIXTER"
                
                sub_val = str(row_sub[i]).upper() if i < len(row_sub) and pd.notna(row_sub[i]) else ""
                
                if cluster_atual == "SP":
                    if "PREÇO" in sub_val or "PRECO" in sub_val: 
                        sp_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: 
                        sp_oferta_idx = i
                elif cluster_atual == "RIO":
                    if "PREÇO" in sub_val or "PRECO" in sub_val: 
                        rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: 
                        rio_oferta_idx = i

            if col_prod_idx == -1:
                col_prod_idx = col_plu_idx + 1

            # Seleciona as colunas corretas conforme a escolha do usuário
            if cluster == "Lojas (SP)":
                idx_preco = sp_preco_idx
                idx_oferta = sp_oferta_idx
            else:
                idx_preco = rio_preco_idx
                idx_oferta = rio_oferta_idx

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            mapa_excel = {}
            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                plu = limpar_plu(row.iloc[col_plu_idx])
                if not plu: 
                    continue
                
                nome = str(row.iloc[col_prod_idx]) if col_prod_idx < len(row) else "Produto"
                de = limpar_preco(row.iloc[idx_preco]) if idx_preco != -1 and idx_preco < len(row) else None
                oferta = limpar_preco(row.iloc[idx_oferta]) if idx_oferta != -1 and idx_oferta < len(row) else None
                
                mapa_excel[plu] = {
                    'nome': nome,
                    'de': de,
                    'oferta': oferta
                }

            # --- LER HTML (SOMENTE LEITURA ESTRITA POR PLU) ---
            arquivo_html.seek(0)
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            resultados = []
            
            for plu, dados_ex in mapa_excel.items():
                # Procura no HTML pelo ID exato do PLU (ex: id="788274")
                tag_plu = soup.find(id=plu) or soup.find(attrs={"data-plu": plu})
                
                if not tag_plu:
                    # Busca alternativa caso o ID esteja parcial
                    for el in soup.find_all(True):
                        if el.get('id') and plu in str(el.get('id')):
                            tag_plu = el
                            break

                if not tag_plu:
                    resultados.append({
                        "PLU": plu,
                        "Produto": dados_ex['nome'],
                        "Planilha (De / Oferta)": f"R$ {dados_ex['de']:.2f} / R$ {dados_ex['oferta']:.2f}" if dados_ex['de'] and dados_ex['oferta'] else "N/D",
                        "HTML (De / Oferta)": "Não encontrado",
                        "Status": "⚠️ PLU Ausente no HTML",
                        "Detalhes": f"O PLU {plu} consta na planilha para o cluster {cluster}, mas não foi localizado no HTML."
                    })
                    continue
                
                container = tag_plu.find_parent('table') or tag_plu.find_parent('td') or tag_plu
                txt_bloco = str(container)
                
                # Extração cirúrgica dos preços do bloco HTML correspondente
                m_de = re.search(r'DE\s*R\$\s*([\d,]+)', txt_bloco)
                de_html = float(m_de.group(1).replace(',', '.')) if m_de else None
                
                m_of = re.search(r'background-color:#D50037[^>]*>.*?R\$\s*([\d,]+)', txt_bloco, re.DOTALL)
                if not m_of:
                    m_of = re.search(r'R\$\s*([\d,]+)', txt_bloco)
                oferta_html = float(m_of.group(1).replace(',', '.')) if m_of else None
                
                status = "✅ OK"
                erros_detalhes = []
                
                de_ex = dados_ex['de']
                of_ex = dados_ex['oferta']
                
                if de_html is not None and de_ex is not None:
                    if abs(de_html - de_ex) > 0.01:
                        status = "❌ Erro Preço DE"
                        erros_detalhes.append(f"DE HTML R${de_html:.2f} != Planilha R${de_ex:.2f}")
                
                if oferta_html is not None and of_ex is not None:
                    if abs(oferta_html - of_ex) > 0.01:
                        status = "❌ Erro Oferta" if status == "✅ OK" else status + " | Erro Oferta"
                        erros_detalhes.append(f"Oferta HTML R${oferta_html:.2f} != Planilha R${of_ex:.2f}")

                resultados.append({
                    "PLU": plu,
                    "Produto": dados_ex['nome'],
                    "Planilha (De / Oferta)": f"R$ {de_ex:.2f} / R$ {of_ex:.2f}" if de_ex is not None and of_ex is not None else "Não definido",
                    "HTML (De / Oferta)": f"R$ {de_html:.2f} / R$ {oferta_html:.2f}" if de_html is not None and oferta_html is not None else "Não lido",
                    "Status": status,
                    "Detalhes": " | ".join(erros_detalhes) if erros_detalhes else "Preços conferem"
                })

            st.markdown(f"### 📊 Relatório de Validação por PLU — Cluster: **{cluster}**")
            df_res = pd.DataFrame(resultados)
            st.dataframe(df_res, use_container_width=True)
            
            problemas = sum(1 for r in resultados if "❌" in r["Status"] or "⚠️" in r["Status"])
            if problemas == 0:
                st.success(f"✅ Validação bem-sucedida! Todos os {len(resultados)} PLUs conferem rigorosamente com o cluster **{cluster}**.")
            else:
                st.error(f"⚠️ Atenção! Foram encontrados **{problemas}** problemas de divergência ou ausência no HTML.")
    else:
        st.warning("Por favor, faça o upload da planilha Excel e do arquivo HTML.")
