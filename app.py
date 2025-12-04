import streamlit as st
import pandas as pd
import re
import io
import fitz  # PyMuPDF
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURAÇÃO GERAL
# ==========================================
st.set_page_config(
    page_title="Super Conciliador v2.6 (Visual Pro)",
    layout="wide",
    page_icon="💸",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. FUNÇÕES DE LIMPEZA E CHAVES
# ==========================================

def gerar_chave_padronizada(texto_conta):
    """
    Padroniza a chave tentando ser inteligente com códigos contábeis longos.
    """
    if not isinstance(texto_conta, str):
        return None
    
    texto_conta = texto_conta.strip()
    
    # 1. Lógica para Códigos Contábeis Longos
    if '.' in texto_conta and len(texto_conta) > 15:
        partes = texto_conta.split('.')
        maior_parte = ""
        for p in partes:
            limpo = re.sub(r'\D', '', p)
            if len(limpo) > len(maior_parte):
                maior_parte = limpo
        if len(maior_parte) > 4:
            texto_conta = maior_parte

    # 2. Lógica Padrão
    elif '/' in texto_conta:
        texto_conta = texto_conta.split('/')[-1]
            
    # 3. Limpeza Final
    parte_numerica = re.sub(r'\D', '', texto_conta)
    
    if not parte_numerica: return None
        
    ultimos_7_digitos = parte_numerica[-7:]
    return ultimos_7_digitos.zfill(7)

def limpar_valor_monetario(valor_str):
    if not isinstance(valor_str, str): return 0.0
    valor_upper = valor_str.upper()
    eh_negativo = 'D' in valor_upper or '-' in valor_str or 'DEB' in valor_upper
    limpo = re.sub(r'[^\d,\.]', '', valor_str)
    try:
        if not limpo: return 0.0
        if ',' in limpo and '.' in limpo:
             limpo = limpo.replace('.', '').replace(',', '.')
        elif ',' in limpo:
             limpo = limpo.replace(',', '.')
        
        valor_float = float(limpo)
        return -valor_float if eh_negativo else valor_float
    except ValueError:
        return 0.0

def formatar_moeda_br(valor):
    """Formata float para string BRL: 1.234,56"""
    if pd.isna(valor): return "0,00"
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================================
# 2. MOTOR DE LEITURA DE PDF
# ==========================================

def extrair_pdf_melhorado(arquivo, tipo_extrato):
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto_completo = ""
        for pag in doc:
            texto_completo += pag.get_text() + "\n"
        doc.close()
        
        # --- CONTA ---
        conta_encontrada = "N/A"
        padroes_conta = [
            r"Conta:\s*(\d{4}\/\d{3,4}\/[\d\-]+)", 
            r"Conta\s*Vinculada:\s*(\d{4}\/\d{3,4}\/[\d\-]+)",
            r"Conta\s*Corrente\s*[:\s]*([\d\.\-\/]+)",
            r"Conta\s*[:\s]*([\d\.\-\/]+)",
            r"Agência.*?Conta.*?([\d\.\-]{5,})"
        ]
        for p in padroes_conta:
            match = re.search(p, texto_completo, re.IGNORECASE)
            if match:
                conta_raw = match.group(1).strip()
                if len(re.sub(r'\D', '', conta_raw)) > 4:
                    conta_encontrada = conta_raw
                    break
        
        # --- VALORES ---
        saldo_final = 0.0
        rendimento_total = 0.0
        linhas = texto_completo.split('\n')
        
        for linha in linhas:
            linha_upper = linha.upper().strip()
            
            gatilhos_saldo = ["SALDO FINAL", "SALDO TOTAL", "SALDO ATUAL", "SALDO EM", "SALDO LÍQUIDO", "SALDO BRUTO", "VALOR LIQUIDO", "TOTAL DISPONIVEL", "POSICAO EM"]
            ignorar = ["ANTERIOR", "BLOQUEADO", "PROVISORIO", "RENDIMENTO"]
            
            if any(g in linha_upper for g in gatilhos_saldo) and not any(i in linha_upper for i in ignorar):
                match_val = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})([CD]?)", linha_upper)
                if match_val:
                    valor_str = match_val.group(1)
                    letra = match_val.group(2)
                    sinal = "D" if letra == 'D' or " D" in linha_upper or "DEB" in linha_upper else "C"
                    v = limpar_valor_monetario(f"{valor_str} {sinal}")
                    if v != 0: saldo_final = v

            if tipo_extrato == 'INV':
                gatilhos_rend = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTO LIQUIDO"]
                if any(g in linha_upper for g in gatilhos_rend) and "ACUMULADO" not in linha_upper:
                    match_val = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
                    if match_val:
                        valor = limpar_valor_monetario(match_val.group(1))
                        if valor < 100000000: rendimento_total += valor

        if saldo_final == 0.0 and ("NAO HOUVE MOVIMENTO" in texto_completo.upper() or "SEM MOVIMENTO" in texto_completo.upper()):
             match_ant = re.search(r"(?:SALDO ANTERIOR|SALDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE | re.DOTALL)
             if match_ant: saldo_final = limpar_valor_monetario(match_ant.group(1))

        if saldo_final == 0.0 and tipo_extrato == 'INV':
            match_last = re.findall(r"(?:TOTAL|SALDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE)
            if match_last: saldo_final = limpar_valor_monetario(match_last[-1])

        return {"Conta": conta_encontrada, "Saldo": saldo_final, "Rendimento": rendimento_total, "Texto_Raw": texto_completo[:300]}
    except Exception as e:
        return {"Conta": "Erro", "Saldo": 0.0, "Rendimento": 0.0, "Texto_Raw": str(e)}

# ==========================================
# 3. LEITURA CONTÁBIL
# ==========================================

def processar_contabil(arquivo, tipo='SALDO'):
    if arquivo is None: return pd.DataFrame()
    try:
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1, dtype=str)
        col_chave = None
        possiveis_chaves = ['Domicílio bancário', 'Conta', 'Nº Conta', 'Descrição', 'Conta Contabil']
        for col in df.columns:
            for p in possiveis_chaves:
                if p.lower() in str(col).lower():
                    col_chave = col; break
        
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            for col in df.columns:
                for p in possiveis_chaves:
                    if p.lower() in str(col).lower():
                        col_chave = col; break
        
        if not col_chave: return pd.DataFrame()

        col_valor = None
        possiveis_valores = ['Saldo Final', 'Saldo Atual', 'Movimento', 'Valor']
        for col in df.columns:
            if 'anterior' in str(col).lower(): continue
            for p in possiveis_valores:
                if p.lower() in str(col).lower():
                    col_valor = col; break
        
        if not col_valor: return pd.DataFrame()

        df['Chave Primaria'] = df[col_chave].apply(gerar_chave_padronizada)
        df = df.dropna(subset=['Chave Primaria'])
        df['Valor_Numerico'] = df[col_valor].astype(str).apply(limpar_valor_monetario)
        
        if tipo == 'SALDO':
            if any('contábil' in str(c).lower() for c in df.columns):
                col_contabil = next(c for c in df.columns if 'contábil' in str(c).lower())
                df_pivot = df.pivot_table(index='Chave Primaria', columns=col_contabil, values='Valor_Numerico', aggfunc='sum').reset_index()
                
                col_mov = next((c for c in df_pivot.columns if '1111119' in str(c) or 'Conta Movimento' in str(c) or 'MOVIMENTO' in str(c).upper()), None)
                col_app = next((c for c in df_pivot.columns if '1111150' in str(c) or 'Aplicação' in str(c) or 'APLICACAO' in str(c).upper()), None)
                
                df_res = pd.DataFrame()
                df_res['Chave Primaria'] = df_pivot['Chave Primaria']
                df_res['Saldo_Contabil_CC'] = df_pivot[col_mov].fillna(0) if col_mov else 0.0
                df_res['Saldo_Contabil_Aplic'] = df_pivot[col_app].fillna(0) if col_app else 0.0
                
                desc = df[['Chave Primaria', col_chave]].drop_duplicates(subset='Chave Primaria')
                df_res = df_res.merge(desc, on='Chave Primaria', how='left')
                df_res.rename(columns={col_chave: 'Descrição'}, inplace=True)
                return df_res
            else:
                df_agrup = df.groupby('Chave Primaria')['Valor_Numerico'].sum().reset_index()
                df_agrup.rename(columns={'Valor_Numerico': 'Saldo_Contabil_CC'}, inplace=True)
                df_agrup['Saldo_Contabil_Aplic'] = 0.0
                desc = df[['Chave Primaria', col_chave]].drop_duplicates(subset='Chave Primaria')
                df_agrup = df_agrup.merge(desc, on='Chave Primaria', how='left')
                df_agrup.rename(columns={col_chave: 'Descrição'}, inplace=True)
                return df_agrup

        elif tipo == 'RENDIMENTO':
            df_agrup = df.groupby('Chave Primaria')['Valor_Numerico'].sum().reset_index()
            df_agrup.rename(columns={'Valor_Numerico': 'Rendimento_Contabil'}, inplace=True)
            return df_agrup
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 4. CONSOLIDAÇÃO
# ==========================================

def executar_processo(file_saldos, file_rendim, lista_extratos_cc, lista_extratos_inv):
    df_saldos = processar_contabil(file_saldos, 'SALDO')
    df_rendim = processar_contabil(file_rendim, 'RENDIMENTO')
    
    if df_saldos.empty:
        st.error("Erro na leitura do CSV de Saldos.")
        return pd.DataFrame(), pd.DataFrame()

    df_contabil = df_saldos
    if not df_rendim.empty:
        df_contabil = pd.merge(df_saldos, df_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        df_contabil['Rendimento_Contabil'] = 0.0

    dados_banco = []
    log_leitura = []

    for f in lista_extratos_cc:
        res = extrair_pdf_melhorado(f, 'CC')
        chave = gerar_chave_padronizada(res['Conta'])
        if chave: dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': res['Saldo'], 'Saldo_Banco_Aplic': 0.0, 'Rendimento_Banco': 0.0})
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': chave, 'Saldo': res['Saldo'], 'Tipo': 'Conta Corrente'})

    for f in lista_extratos_inv:
        res = extrair_pdf_melhorado(f, 'INV')
        chave = gerar_chave_padronizada(res['Conta'])
        if chave: dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': 0.0, 'Saldo_Banco_Aplic': res['Saldo'], 'Rendimento_Banco': res['Rendimento']})
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': chave, 'Saldo_Aplic': res['Saldo'], 'Rendimento': res['Rendimento'], 'Tipo': 'Investimento'})

    df_log = pd.DataFrame(log_leitura)
    
    if dados_banco:
        df_banco = pd.DataFrame(dados_banco).groupby('Chave Primaria').sum().reset_index()
    else:
        df_banco = pd.DataFrame(columns=['Chave Primaria', 'Saldo_Banco_CC', 'Saldo_Banco_Aplic', 'Rendimento_Banco'])

    df_final = pd.merge(df_contabil, df_banco, on='Chave Primaria', how='outer').fillna(0)

    if 'Descrição' in df_final.columns:
        df_final['Descrição'] = df_final['Descrição'].fillna('CONTA SEM DESCRIÇÃO')
    else:
        df_final['Descrição'] = 'CONTA SEM DESCRIÇÃO'

    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']

    # --- REORDENAR E FILTRAR COLUNAS ---
    cols = ['Descrição', 'Chave Primaria', 
            'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
            'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
            'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento']
    colunas_finais = [c for c in cols if c in df_final.columns]
    
    return df_final[colunas_finais], df_log

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True) # Index true porque vamos usar MultiIndex
        ws = writer.sheets['Sheet1']
        for i in range(1, 20):
             ws.column_dimensions[get_column_letter(i)].width = 18
    return output.getvalue()

# ==========================================
# 5. INTERFACE (VISUAL PRO)
# ==========================================

st.title("💸 Super Conciliador v2.6 (Visual Pro)")
st.markdown("### Suporte a Caixa e BB (Agrupamento Visual e Formatação BR)")

st.markdown("---")
st.header("1. Arquivos da Contabilidade")
col_cont1, col_cont2 = st.columns(2)
with col_cont1: f_saldos = st.file_uploader("📂 Saldos Contábeis (CSV)", type='csv')
with col_cont2: f_rendim = st.file_uploader("📂 Rendimentos (CSV)", type='csv')

st.markdown("---")
st.header("2. Extratos Bancários")
col_bb, col_caixa = st.columns(2)
with col_bb:
    st.subheader("🔵 Banco do Brasil")
    f_bb_cc = st.file_uploader("BB - Conta Corrente", type='pdf', accept_multiple_files=True, key="bb_cc")
    f_bb_inv = st.file_uploader("BB - Aplicações/Invest", type='pdf', accept_multiple_files=True, key="bb_inv")
with col_caixa:
    st.subheader("🟠 Caixa Econômica")
    f_caixa_cc = st.file_uploader("Caixa - Conta Corrente", type='pdf', accept_multiple_files=True, key="cx_cc")
    f_caixa_inv = st.file_uploader("Caixa - Aplicações/Invest", type='pdf', accept_multiple_files=True, key="cx_inv")

st.markdown("---")
if st.button("Executar Conciliação", type="primary"):
    
    tem_banco = (f_bb_cc or f_bb_inv or f_caixa_cc or f_caixa_inv)
    
    if f_saldos and tem_banco:
        with st.spinner("Processando..."):
            
            lista_final_cc = []
            if f_bb_cc: lista_final_cc.extend(f_bb_cc)
            if f_caixa_cc: lista_final_cc.extend(f_caixa_cc)
            
            lista_final_inv = []
            if f_bb_inv: lista_final_inv.extend(f_bb_inv)
            if f_caixa_inv: lista_final_inv.extend(f_caixa_inv)
            
            df_final, df_log = executar_processo(f_saldos, f_rendim, lista_final_cc, lista_final_inv)
            
            if not df_final.empty:
                st.success("Conciliação Realizada!")
                
                # --- PREPARAÇÃO VISUAL (MULTI-INDEX) ---
                # Criamos um novo DataFrame só para exibição bonita
                df_display = df_final.copy()
                
                # Mapa de Renomeação para criar Hierarquia
                mapa_colunas = {
                    'Descrição': ('Dados', 'Descrição'),
                    'Chave Primaria': ('Dados', 'Conta Reduzida'),
                    'Saldo_Contabil_CC': ('CONTA CORRENTE', 'Contabilidade'),
                    'Saldo_Banco_CC': ('CONTA CORRENTE', 'Extrato Banco'),
                    'Diferenca_Saldo_CC': ('CONTA CORRENTE', 'Diferença'),
                    'Saldo_Contabil_Aplic': ('APLICAÇÃO', 'Contabilidade'),
                    'Saldo_Banco_Aplic': ('APLICAÇÃO', 'Extrato Banco'),
                    'Diferenca_Saldo_Aplic': ('APLICAÇÃO', 'Diferença'),
                    'Rendimento_Contabil': ('RENDIMENTO', 'Contabilidade'),
                    'Rendimento_Banco': ('RENDIMENTO', 'Extrato Banco'),
                    'Diferenca_Rendimento': ('RENDIMENTO', 'Diferença')
                }
                
                # Filtra colunas que realmente existem
                cols_existentes = [c for c in df_display.columns if c in mapa_colunas]
                df_display = df_display[cols_existentes]
                
                # Aplica o MultiIndex nas colunas
                df_display.columns = pd.MultiIndex.from_tuples([mapa_colunas[c] for c in df_display.columns])
                
                # Identifica colunas numéricas para aplicar a função BR
                numeric_cols = df_display.select_dtypes(include=['float', 'int']).columns
                
                # Aplica formatação BR
                df_formatado = df_display.copy()
                for col in numeric_cols:
                    df_formatado[col] = df_formatado[col].apply(formatar_moeda_br)

                # --- EXIBIÇÃO ---
                tab1, tab2, tab3 = st.tabs(["📊 Tabela Formatada", "⚠️ Divergências", "🕵️ Log"])
                
                with tab1:
                    st.dataframe(df_formatado, use_container_width=True)
                    st.download_button("Baixar Excel", to_excel(df_display), "conciliacao_pro.xlsx")
                
                with tab2:
                    # Filtro de divergência usando a estrutura original para facilidade
                    filtro = (df_final['Diferenca_Saldo_CC'].abs() > 0.01) | \
                             (df_final['Diferenca_Saldo_Aplic'].abs() > 0.01) | \
                             (df_final['Diferenca_Rendimento'].abs() > 0.01)
                    
                    df_div = df_formatado[filtro]
                    if df_div.empty: st.info("Sem divergências!")
                    else: st.dataframe(df_div, use_container_width=True)
                
                with tab3:
                    st.dataframe(df_log)
            else:
                st.error("Erro na consolidação.")
    else:
        st.warning("Faltam arquivos.")
