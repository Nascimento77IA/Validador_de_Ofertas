import streamlit as st
import pandas as pd
import re
from bs4 import BeautifulSoup

# --- FUNÇÕES DE LIMPEZA E FILTRAGEM ---
def limpar_preco(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).upper().replace('R$', '').replace(' ', '').strip()
    if texto in ['', 'X', 'NAN', '-']:
        return "NAO_APLICA"
    if ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return "NAO_APLICA"

def limpar_plu(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith('.0'):
        texto = texto[:-2]
    return re.sub(r'\D', '', texto)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador Inteligente - Nagumo", layout="wide", page_icon="⚡")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚡ Validador Multi-Cluster com Filtro de Exclusão (SP & Rio)")
st.markdown("Ignora automaticamente produtos com **'-'** no cluster e valida estritamente apenas os itens válidos.")

# --- INTERFACE DO USUÁRIO ---
col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("1. Suba a Planilha de Ofertas (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("2. Suba o HTML (com PLUs embutidos)", type=["html"])

if st.button("🚀 Executar Validação Inteligente"):
    if arquivo_excel and arquivo_html:
        with st.spinner("Processando e aplicando regras de exclusão por cluster..."):
            
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
                
            row_main = df_raw.iloc[linha_plu_idx].values
            row_sub = df_raw.iloc[linha_plu_idx + 1].values if linha_plu_idx + 1 < len(df_raw) else []
            
            col_plu_idx = -1; col_prod_idx = -1
            sp_preco_idx = -1; sp_oferta_idx = -1
            rio_preco_idx = -1; rio_oferta_idx = -1
            
            cluster_atual = ""
            for i in range(len(row_main)):
                val_main = str(row_main[i]).upper() if pd.notna(row_main[i]) else ""
                
                if "PLU" in val_main: col_plu_idx = i
                elif "PRODUTO" in val_main or "DESC" in val_main: col_prod_idx = i
                elif "NAGUMO SP" in val_main: cluster_atual = "SP"
                elif "NAGUMO RIO" in val_main: cluster_atual = "RIO"
                elif "NAGUMO MIXTER" in val_main: cluster_atual = "MIXTER"
                
                sub_val = str(row_sub[i]).upper() if i < len(row_sub) and pd.notna(row_sub[i]) else ""
                
                if cluster_atual == "SP":
                    if "PREÇO" in sub_val or "PRECO" in sub_val: sp_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: sp_oferta_idx = i
                elif cluster_atual == "RIO":
                    if "PREÇO" in sub_val or "PRECO" in sub_val: rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: rio_oferta_idx = i

            if col_prod_idx == -1:
                col_prod_idx = col_plu_idx + 1

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            dados_sp = {}
            dados_rio = {}
            alertas_comerciais = []

            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                plu = limpar_plu(row.iloc[col_plu_idx])
                if not plu: 
                    continue
                
                nome = str(row.iloc[col_prod_idx]) if col_prod_idx < len(row) else "Produto"
                
                # Extração SP
                de_sp = limpar_preco(row.iloc[sp_preco_idx]) if sp_preco_idx != -1 and sp_preco_idx < len(row) else "NAO_APLICA"
                oferta_sp = limpar_preco(row.iloc[sp_oferta_idx]) if sp_oferta_idx != -1 and sp_oferta_idx < len(row) else "NAO_APLICA"
                
                # Extração Rio
                de_rio = limpar_preco(row.iloc[rio_preco_idx]) if rio_preco_idx != -1 and rio_preco_idx < len(row) else "NAO_APLICA"
                oferta_rio = limpar_preco(row.iloc[rio_oferta_idx]) if rio_oferta_idx != -1 and rio_oferta_idx < len(row) else "NAO_APLICA"
                
                # Blindagem comercial (apenas se o item for válido no cluster)
                if isinstance(oferta_sp, float) and isinstance(de_sp, float) and oferta_sp >= de_sp:
                    alertas_comerciais.append(f"🚨 **SP [Erro Comercial]** | {nome} (PLU: {plu}) | Oferta (R${oferta_sp:.2f}) >= De (R${de_sp:.2f})")
                
                if isinstance(oferta_rio, float) and isinstance(de_rio, float) and oferta_rio >= de_rio:
                    alertas_comerciais.append(f"🚨 **Rio [Erro Comercial]** | {nome} (PLU: {plu}) | Oferta (R${oferta_rio:.2f}) >= De (R${de_rio:.2f})")

                # Guarda apenas se tiver oferta válida (ignora o traço)
                if oferta_sp != "NAO_APLICA":
                    dados_sp[plu] = {'nome': nome, 'de': de_sp if isinstance(de_sp, float) else None, 'oferta': oferta_sp}
                
                if oferta_rio != "NAO_APLICA":
                    dados_rio[plu] = {'nome': nome, 'de': de_rio if isinstance(de_rio, float) else None, 'oferta': oferta_rio}

            # --- LER HTML ---
            arquivo_html.seek(0)
            html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
            soup = BeautifulSoup(html_content, 'html.parser')
            
            def validar_cluster(mapa_cluster, nome_cluster):
                resultados = []
                for plu, dados_ex in mapa_cluster.items():
                    tag_plu = soup.find(id=plu) or soup.find(attrs={"data-plu": plu})
                    if not tag_plu:
                        for el in soup.find_all(True):
                            if el.get('id') and plu in str(el.get('id')):
                                tag_plu = el
                                break

                    if not tag_plu:
                        resultados.append({
                            "PLU": plu, "Produto": dados_ex['nome'],
                            "Planilha (De / Oferta)": f"R$ {dados_ex['de']:.2f} / R$ {dados_ex['oferta']:.2f}" if dados_ex['de'] and dados_ex['oferta'] else "N/D",
                            "HTML (De / Oferta)": "Não encontrado",
                            "Status": "⚠️ PLU Ausente no HTML",
                            "Detalhes": f"Produto ativo na planilha de {nome_cluster} mas ausente no HTML."
                        })
                        continue
                    
                    container = tag_plu.find_parent('table') or tag_plu.find_parent('td') or tag_plu
                    txt_bloco = str(container)
                    
                    m_de = re.search(r'DE\s*R\$\s*([\d,]+)', txt_bloco)
                    de_html = float(m_de.group(1).replace(',', '.')) if m_de else None
                    
                    m_of = re.search(r'background-color:#D50037[^>]*>.*?R\$\s*([\d,]+)', txt_bloco, re.DOTALL)
                    if not m_of:
                        m_of = re.search(r'R\$\s*([\d,]+)', txt_bloco)
                    oferta_html = float(m_of.group(1).replace(',', '.')) if m_of else None
                    
                    status = "✅ OK"
                    erros_detalhes = []
                    de_ex, of_ex = dados_ex['de'], dados_ex['oferta']
                    
                    if de_html is not None and de_ex is not None and abs(de_html - de_ex) > 0.01:
                        status = "❌ Erro Preço DE"
                        erros_detalhes.append(f"DE HTML R${de_html:.2f} != Planilha R${de_ex:.2f}")
                    
                    if oferta_html is not None and of_ex is not None and abs(oferta_html - of_ex) > 0.01:
                        status = "❌ Erro Oferta" if status == "✅ OK" else status + " | Erro Oferta"
                        erros_detalhes.append(f"Oferta HTML R${oferta_html:.2f} != Planilha R${of_ex:.2f}")

                    resultados.append({
                        "PLU": plu, "Produto": dados_ex['nome'],
                        "Planilha (De / Oferta)": f"R$ {de_ex:.2f} / R$ {of_ex:.2f}" if de_ex is not None and of_ex is not None else "Não definido",
                        "HTML (De / Oferta)": f"R$ {de_html:.2f} / R$ {oferta_html:.2f}" if de_html is not None and oferta_html is not None else "Não lido",
                        "Status": status,
                        "Detalhes": " | ".join(erros_detalhes) if erros_detalhes else "Preços conferem"
                    })
                return resultados

            res_sp = validar_cluster(dados_sp, "Nagumo SP")
            res_rio = validar_cluster(dados_rio, "Nagumo Rio")

            # --- EXIBIÇÃO DE ALERTAS ---
            if alertas_comerciais:
                st.error("🚨 **ATENÇÃO: Inconsistências lógicas na planilha do comercial!**")
                for alerta in alertas_comerciais:
                    st.warning(alerta)
                st.markdown("---")

            # --- ABAS DE RESULTADOS ---
            aba_sp, aba_rio = st.tabs(["🏢 Cluster Nagumo SP (Lojas)", "🌴 Cluster Nagumo Rio"])

            with aba_sp:
                st.markdown(f"### Relatório SP ({len(dados_sp)} itens válidos na campanha)")
                df_sp = pd.DataFrame(res_sp)
                st.dataframe(df_sp, use_container_width=True)
                prob_sp = sum(1 for r in res_sp if "❌" in r["Status"] or "⚠️" in r["Status"])
                if prob_sp == 0:
                    st.success("✅ Todos os PLUs válidos em SP conferem perfeitamente com o HTML.")
                else:
                    st.error(f"⚠️ Foram encontrados {prob_sp} problemas em SP.")

            with aba_rio:
                st.markdown(f"### Relatório Rio ({len(dados_rio)} itens válidos na campanha)")
                df_rio = pd.DataFrame(res_rio)
                st.dataframe(df_rio, use_container_width=True)
                prob_rio = sum(1 for r in res_rio if "❌" in r["Status"] or "⚠️" in r["Status"])
                if prob_rio == 0:
                    st.success("✅ Todos os PLUs válidos no Rio conferem perfeitamente com o HTML.")
                else:
                    st.error(f"⚠️ Foram encontrados {prob_rio} problemas no Rio.")
    else:
        st.warning("Por favor, faça o upload da planilha Excel e do arquivo HTML.")
