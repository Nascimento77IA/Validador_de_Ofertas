import streamlit as st
import pandas as pd
import re
import math
from bs4 import BeautifulSoup

# --- FUNÇÕES DE LIMPEZA E ARREDONDAMENTO COMERCIAL ---
def limpar_preco(valor):
    if pd.isna(valor):
        return "NAO_APLICA"
    texto = str(valor).upper().replace('R$', '').replace(' ', '').strip()
    if texto in ['', 'X', 'NAN', '-']:
        return "NAO_APLICA"
    if ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return "NAO_APLICA"

def limpar_desconto(valor):
    if pd.isna(valor):
        return None
    
    # Se o pandas ler o resultado da fórmula do Excel como decimal (ex: 0.1165 para 12%)
    if isinstance(valor, (int, float)):
        val_float = float(valor)
        if 0 < val_float < 1:
            return math.floor((val_float * 100) + 0.5)
        return math.floor(val_float + 0.5)

    texto = str(valor).upper().replace('%', '').replace(' ', '').strip()
    if texto in ['', 'X', 'NAN', '-']:
        return None
    if ',' in texto:
        texto = texto.replace(',', '.')
    try:
        num = float(texto)
        if 0 < num < 1:
            return math.floor((num * 100) + 0.5)
        return math.floor(num + 0.5)
    except ValueError:
        return None

def limpar_plu(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith('.0'):
        texto = texto[:-2]
    return re.sub(r'\D', '', texto)

def arredondar_comercial(valor):
    # Regra comercial: >= .50 arredonda para cima, < .50 arredonda para baixo
    return math.floor(valor + 0.5)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador & Auditor de Descontos - Nagumo", layout="wide", page_icon="🎯")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎯 Validador Inteligente com Auditoria de Desconto (Regra Comercial)")
st.markdown("Valida preços por PLU, trata os traços (-) e audita as percentagens de desconto considerando fórmulas e arredondamento comercial.")

# --- INTERFACE DO UTILIZADOR ---
st.markdown("### 📁 Carregamento de Ficheiros")
col_exc, col_sp, col_rio = st.columns(3)

with col_exc:
    arquivo_excel = st.file_uploader("Planilha de Ofertas (Excel) *Obrigatório*", type=["xlsx"])
with col_sp:
    html_sp_file = st.file_uploader("HTML — Cluster SP (Opcional)", type=["html"])
with col_rio:
    html_rio_file = st.file_uploader("HTML — Cluster Rio (Opcional)", type=["html"])

st.markdown("---")

if st.button("🚀 Executar Validação e Auditoria"):
    if arquivo_excel and (html_sp_file or html_rio_file):
        with st.spinner("A processar dados, aplicar arredondamento comercial e validar HTMLs..."):
            
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
            sp_preco_idx = -1; sp_oferta_idx = -1; sp_desc_idx = -1
            rio_preco_idx = -1; rio_oferta_idx = -1; rio_desc_idx = -1
            
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
                    elif "%" in sub_val or "DESCONTO" in sub_val: sp_desc_idx = i
                elif cluster_atual == "RIO":
                    if "PREÇO" in sub_val or "PRECO" in sub_val: rio_preco_idx = i
                    elif "OFERTA" in sub_val and "CART" not in sub_val: rio_oferta_idx = i
                    elif "%" in sub_val or "DESCONTO" in sub_val: rio_desc_idx = i

            if col_prod_idx == -1:
                col_prod_idx = col_plu_idx + 1

            df_excel = df_raw.iloc[linha_plu_idx + 2:].reset_index(drop=True)

            registros_sp = []
            registros_rio = []
            alertas_comerciais = []

            for _, row in df_excel.iterrows():
                if col_plu_idx == -1 or col_plu_idx >= len(row) or pd.isna(row.iloc[col_plu_idx]):
                    continue
                plu = limpar_plu(row.iloc[col_plu_idx])
                if not plu: 
                    continue
                
                nome = str(row.iloc[col_prod_idx]) if col_prod_idx < len(row) else "Produto"
                
                # SP Valores
                val_de_sp = row.iloc[sp_preco_idx] if sp_preco_idx != -1 and sp_preco_idx < len(row) else None
                val_of_sp = row.iloc[sp_oferta_idx] if sp_oferta_idx != -1 and sp_oferta_idx < len(row) else None
                val_desc_sp = row.iloc[sp_desc_idx] if sp_desc_idx != -1 and sp_desc_idx < len(row) else None

                de_sp = limpar_preco(val_de_sp)
                oferta_sp = limpar_preco(val_of_sp)
                desc_plan_sp = limpar_desconto(val_desc_sp)
                
                # Rio Valores
                val_de_rio = row.iloc[rio_preco_idx] if rio_preco_idx != -1 and rio_preco_idx < len(row) else None
                val_of_rio = row.iloc[rio_oferta_idx] if rio_oferta_idx != -1 and rio_oferta_idx < len(row) else None
                val_desc_rio = row.iloc[rio_desc_idx] if rio_desc_idx != -1 and rio_desc_idx < len(row) else None

                de_rio = limpar_preco(val_de_rio)
                oferta_rio = limpar_preco(val_of_rio)
                desc_plan_rio = limpar_desconto(val_desc_rio)
                
                # SP - Auditoria
                if isinstance(oferta_sp, float) and isinstance(de_sp, float):
                    if oferta_sp >= de_sp:
                        alertas_comerciais.append(f"🚨 **SP [Erro Comercial]** | {nome} (PLU: {plu}) | Oferta (R${oferta_sp:.2f}) >= De (R${de_sp:.2f})")
                    else:
                        calc_exato = ((de_sp - oferta_sp) / de_sp) * 100
                        calc_desc = arredondar_comercial(calc_exato)
                        if desc_plan_sp is not None and abs(calc_desc - desc_plan_sp) > 0:
                            alertas_comerciais.append(f"⚠️ **SP [Divergência de Desconto]** | {nome} (PLU: {plu}) | Informado: {desc_plan_sp}% vs Calculado: {calc_desc}%")

                # Rio - Auditoria
                if isinstance(oferta_rio, float) and isinstance(de_rio, float):
                    if oferta_rio >= de_rio:
                        alertas_comerciais.append(f"🚨 **Rio [Erro Comercial]** | {nome} (PLU: {plu}) | Oferta (R${oferta_rio:.2f}) >= De (R${de_rio:.2f})")
                    else:
                        calc_exato = ((de_rio - oferta_rio) / de_rio) * 100
                        calc_desc = arredondar_comercial(calc_exato)
                        if desc_plan_rio is not None and abs(calc_desc - desc_plan_rio) > 0:
                            alertas_comerciais.append(f"⚠️ **Rio [Divergência de Desconto]** | {nome} (PLU: {plu}) | Informado: {desc_plan_rio}% vs Calculado: {calc_desc}%")

                registros_sp.append({'plu': plu, 'nome': nome, 'de': de_sp, 'oferta': oferta_sp, 'val_of_raw': val_of_sp})
                registros_rio.append({'plu': plu, 'nome': nome, 'de': de_rio, 'oferta': oferta_rio, 'val_of_raw': val_of_rio})

            def validar_cluster(registros, soup_obj, nome_cluster):
                resultados = []
                for item in registros:
                    plu = item['plu']
                    nome = item['nome']
                    de_ex = item['de']
                    of_ex = item['oferta']
                    val_of_raw = item['val_of_raw']

                    is_traco = pd.isna(val_of_raw) or str(val_of_raw).strip() in ['-', 'X', 'NAN', ''] or of_ex == "NAO_APLICA"

                    if is_traco:
                        resultados.append({
                            "PLU": plu,
                            "Produto": nome,
                            "Planilha (De / Oferta)": "Não aplicável (-)",
                            "HTML (De / Oferta)": "Ignorado",
                            "Status": "ℹ️ Não se aplica",
                            "Detalhes": f"Item com '-' na planilha, não pertence ao cluster {nome_cluster}."
                        })
                        continue

                    tag_plu = soup_obj.find(id=plu) or soup_obj.find(attrs={"data-plu": plu})
                    if not tag_plu:
                        for el in soup_obj.find_all(True):
                            if el.get('id') and plu in str(el.get('id')):
                                tag_plu = el
                                break

                    if not tag_plu:
                        resultados.append({
                            "PLU": plu,
                            "Produto": nome,
                            "Planilha (De / Oferta)": f"R$ {de_ex:.2f} / R$ {of_ex:.2f}" if isinstance(de_ex, float) and isinstance(of_ex, float) else "N/D",
                            "HTML (De / Oferta)": "Não encontrado",
                            "Status": "⚠️ PLU Ausente no HTML",
                            "Detalhes": f"Produto ativo na planilha de {nome_cluster} mas ausente no respetivo HTML."
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
                    
                    if de_html is not None and isinstance(de_ex, float) and abs(de_html - de_ex) > 0.01:
                        status = "❌ Erro Preço DE"
                        erros_detalhes.append(f"DE HTML R${de_html:.2f} != Planilha R${de_ex:.2f}")
                    
                    if oferta_html is not None and isinstance(of_ex, float) and abs(oferta_html - of_ex) > 0.01:
                        status = "❌ Erro Oferta" if status == "✅ OK" else status + " | Erro Oferta"
                        erros_detalhes.append(f"Oferta HTML R${oferta_html:.2f} != Planilha R${of_ex:.2f}")

                    resultados.append({
                        "PLU": plu,
                        "Produto": nome,
                        "Planilha (De / Oferta)": f"R$ {de_ex:.2f} / R$ {of_ex:.2f}" if isinstance(de_ex, float) and isinstance(of_ex, float) else "Não definido",
                        "HTML (De / Oferta)": f"R$ {de_html:.2f} / R$ {oferta_html:.2f}" if de_html is not None and oferta_html is not None else "Não lido",
                        "Status": status,
                        "Detalhes": " | ".join(erros_detalhes) if erros_detalhes else "Preços conferem"
                    })
                return resultados

            res_sp = []
            res_rio = []

            if html_sp_file:
                html_sp_file.seek(0)
                soup_sp = BeautifulSoup(html_sp_file.getvalue().decode('utf-8', errors='replace'), 'html.parser')
                res_sp = validar_cluster(registros_sp, soup_sp, "Nagumo SP")

            if html_rio_file:
                html_rio_file.seek(0)
                soup_rio = BeautifulSoup(html_rio_file.getvalue().decode('utf-8', errors='replace'), 'html.parser')
                res_rio = validar_cluster(registros_rio, soup_rio, "Nagumo Rio")

            # --- EXIBIÇÃO DE ALERTAS ---
            if alertas_comerciais:
                st.error("🚨 **ATENÇÃO: Inconsistências lógicas ou de desconto detetadas na planilha!**")
                for alerta in alertas_comerciais:
                    st.warning(alerta)
                st.markdown("---")

            # --- ABAS DINÂMICAS ---
            abas_nomes = []
            if html_sp_file: abas_nomes.append("🏢 Cluster Nagumo SP (Lojas)")
            if html_rio_file: abas_nomes.append("🌴 Cluster Nagumo Rio")

            if abas_nomes:
                tabs = st.tabs(abas_nomes)
                tab_idx = 0
                
                if html_sp_file:
                    with tabs[tab_idx]:
                        st.markdown(f"### Relatório SP ({len(res_sp)} itens processados)")
                        df_sp = pd.DataFrame(res_sp)
                        st.dataframe(df_sp, use_container_width=True)
                        prob_sp = sum(1 for r in res_sp if "❌" in r["Status"] or "⚠️" in r["Status"])
                        if prob_sp == 0:
                            st.success("✅ Nenhum erro real encontrado em SP!")
                        else:
                            st.error(f"⚠️ Foram encontrados {prob_sp} problemas reais em SP.")
                    tab_idx += 1

                if html_rio_file:
                    with tabs[tab_idx]:
                        st.markdown(f"### Relatório Rio ({len(res_rio)} itens processados)")
                        df_rio = pd.DataFrame(res_rio)
                        st.dataframe(df_rio, use_container_width=True)
                        prob_rio = sum(1 for r in res_rio if "❌" in r["Status"] or "⚠️" in r["Status"])
                        if prob_rio == 0:
                            st.success("✅ Nenhum erro real encontrado no Rio!")
                        else:
                            st.error(f"⚠️ Foram encontrados {prob_rio} problemas reais no Rio.")
            else:
                st.warning("Carregue pelo menos um ficheiro HTML para realizar a validação.")
    else:
        st.warning("Por favor, faça o upload da Planilha Excel e de pelo menos um dos ficheiros HTML.")
