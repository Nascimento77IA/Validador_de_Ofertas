import streamlit as st
import pandas as pd
import re
import unicodedata
import difflib
import html

# --- FUNÇÕES DE LIMPEZA ---
def limpar_preco(valor):
    if pd.isna(valor):
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

def limpar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.lower()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    
    # SEPARA LETRAS DE NÚMEROS (ex: OVOS500G vira OVOS 500 G)
    texto = re.sub(r'([a-zA-Z])(\d)', r'\1 \2', texto)
    texto = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', texto)

    substituicoes = {
        r'\brefr\.': 'refrigerante ', r'\bcerv\.': 'cerveja ', r'\bsabon\.': 'sabonete ',
        r'\bmolho tom\.': 'molho de tomate ', r'\bmac\b|\bmac\.': 'macarrao ', r'\blv\.': 'longa vida ',
        r'\bref\.': 'refinado ', r'\bcond\.': 'condensado ', r'\btrad\.': 'tradicional ',
        r'\bsac\.': 'sache ', r'\bcg\.': 'congelada ', r'\blava r\.po\b': 'lava roupas em po ',
        r'\bguar\.': 'guarana ', r'\b1lt\b': '1l', r'\b2lt\b': '2l', r'\blt\b': 'lata ',
        r'\bc/': 'com ', r'\bc\\': 'com ', 
        r'pet': '', r'sachet': 'sache', r'pouch': '', r'\bunids?\.?': '',
    }
    for padrao, subst in substituicoes.items():
        texto = re.sub(padrao, subst, texto)
    texto = re.sub(r'[^\w\s]', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador de Ofertas", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:opsz,wght@6..12,800&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, button, span, div {
        font-family: 'Nunito Sans', sans-serif;
        font-weight: 800;
    }
    
    [data-testid="stIconMaterial"], .material-icons, svg, [class*="icon"] {
        font-family: "Material Symbols Rounded", "Material Icons", sans-serif !important;
        font-weight: normal !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📝 Validador de Ofertas CRM - Nagumo/BWCA")
st.markdown("Faça o upload dos arquivos da campanha para validar automaticamente os preços de Lojas (SP) ou Rio.")

# --- INTERFACE DO USUÁRIO ---
praca_escolhida = st.selectbox("1. Qual cluster você deseja validar?", ["Lojas (SP)", "Rio"]).lower()

col1, col2 = st.columns(2)
with col1:
    arquivo_excel = st.file_uploader("2. Suba a Planilha c/ itens da oferta (Excel)", type=["xlsx"])
with col2:
    arquivo_html = st.file_uploader("3. Suba o E-mail (HTML)", type=["html"])

if st.button("🔍 Validar Preços"):
    if not arquivo_excel or not arquivo_html:
        st.warning("Por favor, faça o upload dos dois arquivos antes de continuar.")
        st.stop()

    with st.spinner("Analisando arquivos..."):
        # ==========================================
        # LER EXCEL
        # ==========================================
        df_excel = pd.read_excel(arquivo_excel, header=None)
        
        if "sp" in praca_escolhida or "lojas" in praca_escolhida:
            praca_str = "NAGUMO SP"
        else:
            praca_str = "NAGUMO RIO"
            
        linha_praca = -1
        coluna_praca = -1

        for r in range(min(10, len(df_excel))):
            for c in range(len(df_excel.columns)):
                val = str(df_excel.iloc[r, c]).upper()
                val = re.sub(r'\s+', ' ', val).strip()
                if praca_str in val:
                    linha_praca = r
                    coluna_praca = c
                    break
            if linha_praca != -1:
                break

        if linha_praca == -1:
            st.error(f"ERRO: Não encontrei '{praca_str}' no cabeçalho. Verifique o arquivo Excel.")
            st.stop()

        linha_cabecalho = linha_praca + 1
        col_preco = -1
        col_oferta = -1
        col_cartao = -1

        limite_busca = min(coluna_praca + 4, len(df_excel.columns))
        for c in range(coluna_praca, limite_busca):
            val_coluna = str(df_excel.iloc[linha_cabecalho, c]).upper()
            if col_preco == -1 and ("PREÇO" in val_coluna or "PRECO" in val_coluna):
                col_preco = c
            elif col_oferta == -1 and "OFERTA" in val_coluna and "CART" not in val_coluna and "NAGUMO" not in val_coluna:
                col_oferta = c
            elif col_cartao == -1 and ("CARTÃO" in val_coluna or "CARTAO" in val_coluna or "MEU NAGUMO" in val_coluna):
                col_cartao = c

        if col_preco == -1 or col_oferta == -1:
            st.error(f"ERRO: Achei '{praca_str}', mas não encontrei as colunas 'PREÇO' ou 'OFERTA' embaixo dela.")
            st.stop()

        col_produto = 1
        for c in range(len(df_excel.columns)):
            val_prod1 = str(df_excel.iloc[linha_praca, c]).upper()
            val_prod2 = str(df_excel.iloc[linha_cabecalho, c]).upper()
            if "PRODUTO" in val_prod1 or "PRODUTO" in val_prod2:
                col_produto = c
                break

        mapa_excel = {}
        for index, row in df_excel.iloc[linha_cabecalho + 1:].iterrows():
            nome_excel_original = str(row.iloc[col_produto])
            if pd.isna(nome_excel_original) or not nome_excel_original.strip() or nome_excel_original.lower() == 'nan':
                continue
                
            chave_excel = limpar_texto(nome_excel_original)
            de_ex = limpar_preco(row.iloc[col_preco]) if col_preco < len(row) else None
            of_ex = limpar_preco(row.iloc[col_oferta]) if col_oferta < len(row) else None
            ca_ex = limpar_preco(row.iloc[col_cartao]) if col_cartao != -1 and col_cartao < len(row) else None
                
            mapa_excel[chave_excel] = {'nome_original': nome_excel_original.strip(), 'de': de_ex, 'oferta': of_ex, 'cartao': ca_ex}

        # ==========================================
        # LER HTML
        # ==========================================
        html_content = arquivo_html.getvalue().decode('utf-8', errors='replace')
        produtos_html = []
        matches = re.finditer(r'color:#003865;">(.*?)</td>', html_content, re.DOTALL)

        for m in matches:
            nome = html.unescape(m.group(1)).strip()
            nome = re.sub(r'\s+', ' ', nome)
            if '%%FirstName%%' in nome or 'Ofertas' in nome:
                continue
            
            rest_of_html = html_content[m.end():m.end()+3000] 
            
            m_de = re.search(r'DE\s*R\$\s*([\d,]+)', rest_of_html)
            de_preco = float(m_de.group(1).replace(',', '.')) if m_de else None
                
            m_oferta = re.search(r'bgcolor="#D50037"[^>]*>.*?R\$\s*([\d,]+)', rest_of_html, re.DOTALL)
            oferta_preco = float(m_oferta.group(1).replace(',', '.')) if m_oferta else None
                
            m_cartao = re.search(r'Pre&#231;o Exclusivo.*?R\$\s*([\d,]+).*?Cart&#227;o Nagumo', rest_of_html, re.DOTALL)
            cartao_preco = float(m_cartao.group(1).replace(',', '.')) if m_cartao else None

            produtos_html.append({'nome_original': nome, 'nome_limpo': limpar_texto(nome), 'de': de_preco, 'oferta': oferta_preco, 'cartao': cartao_preco})

        # ==========================================
        # CRUZAMENTO E VALIDAÇÃO (Similaridade Pura)
        # ==========================================
        erros = []
        alertas = []

        for p_html in produtos_html:
            nome_h = p_html['nome_original']
            chave_h = p_html['nome_limpo']
            dados_excel = None
            
            if chave_h in mapa_excel:
                dados_excel = mapa_excel[chave_h]
            else:
                melhor_score = 0
                melhor_chave = None
                
                for k_excel in mapa_excel.keys():
                    # Calcula a porcentagem de semelhança entre as frases inteiras
                    similaridade = difflib.SequenceMatcher(None, chave_h, k_excel).ratio()
                    
                    if similaridade > melhor_score:
                        melhor_score = similaridade
                        melhor_chave = k_excel
                        
                # Se as frases forem pelo menos 45% parecidas, ele assume que é o mesmo produto e não gera alerta amarelo
                if melhor_chave and melhor_score > 0.45:
                    dados_excel = mapa_excel[melhor_chave]

            if dados_excel:
                if p_html['de'] is not None and dados_excel['de'] is not None and abs(p_html['de'] - dados_excel['de']) > 0.01:
                    erros.append(f"**ERRO PREÇO (DE)** | {nome_h} | HTML: R${p_html['de']:.2f} | Excel: R${dados_excel['de']:.2f}")
                if p_html['oferta'] is not None and dados_excel['oferta'] is not None and abs(p_html['oferta'] - dados_excel['oferta']) > 0.01:
                    erros.append(f"**ERRO OFERTA** | {nome_h} | HTML: R${p_html['oferta']:.2f} | Excel: R${dados_excel['oferta']:.2f}")
                if p_html['cartao'] is not None:
                    if dados_excel['cartao'] is not None:
                        if abs(p_html['cartao'] - dados_excel['cartao']) > 0.01:
                            erros.append(f"**ERRO CARTÃO** | {nome_h} | HTML: R${p_html['cartao']:.2f} | Excel: R${dados_excel['cartao']:.2f}")
                    else:
                        erros.append(f"**ERRO CARTÃO** | {nome_h} | HTML tem Cartão (R${p_html['cartao']:.2f}), mas no Excel está vazio.")
            else:
                alertas.append(f"Produto do HTML **'{nome_h}'** não encontrou correspondência no Excel.")

        # ==========================================
        # RESULTADOS NA TELA
        # ==========================================
        st.subheader("Resultados da Validação")
        if not erros and not alertas:
            st.success("Tudo certo! Valores do HTML batem perfeitamente com o Excel. Pode disparar! 🚀")
        
        if erros:
            for e in erros:
                st.error(e)
                
        if alertas:
            for a in alertas:
                st.warning(a)
