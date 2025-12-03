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
    page_title="Super Conciliador Financeiro",
    layout="wide",
    page_icon="💰",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. FUNÇÕES AUXILIARES (CORE)
# ==========================================

def gerar_chave_padronizada(texto_conta):
    """Padroniza a chave (últimos 7 dígitos). Ex: '7695-3' -> '0007695'"""
    if isinstance(texto_conta, str):
        parte_numerica = re.sub(r'\D', '', texto_conta)
        if not parte_numerica: return None
        # Garante pelo menos 7 digitos
        ultimos_7_digitos = parte_numerica[-7:]
        return ultimos_7_digitos.zfill(7)
    return None

def gerar_chave_contabil(texto_conta):
    """Extrai chave do Domicílio bancário da contabilidade."""
    if not isinstance(texto_conta, str): return None
    try:
        # Tenta dividir por traço (padrão Domicílio: Banco-Agencia-Conta)
        partes = texto_conta.split('-')
        if len(partes) > 1:
            # Pega o último segmento que parece ser a conta
            conta_raw = partes[-1] if len(partes[-1]) > 3 else partes[-2]
            conta_numerica = re.sub(r'\D', '', conta_raw)
            return conta_numerica[-7:].zfill(7)
        else:
            # Se não tiver traço, tenta pegar só números
            conta_numerica = re.sub(r'\D', '', texto_conta)
            return conta_numerica[-7:].zfill(7)
    except:
        return None

def limpar_valor_monetario(valor_str):
    """Transforma strings financeiras (R$ 1.000,00 D) em float."""
    if not isinstance(valor_str, str): return 0.0
    
    eh_negativo = 'D' in valor_str.upper() or '-' in valor_str
    limpo = re.sub(r'[^\d,\.]', '', valor_str)
    
    try:
        if ',' in limpo and '.' in limpo:
             limpo = limpo.replace('.', '').replace(',', '.')
        elif ',' in limpo:
             limpo = limpo.replace(',', '.')
        
        if not limpo: return 0.0
        valor_float = float(limpo)
        return -valor_float if eh_negativo else valor_float
    except ValueError:
        return 0.0

# ==========================================
# 2. PROCESSAMENTO DE PDFS
# ==========================================

def extrair_pdf_cc_generico(arquivo, banco):
    """Lê saldo de Conta Corrente."""
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto = "".join([pag.get_text() for pag in doc])
        doc.close()
        
        # 1. Identificar Conta
        conta = "N/A"
        if banco == 'BB':
            m = re.search(r"Conta(?:\s+corrente)?[:\s]+([\d]+-[\dX])", texto, re.IGNORECASE)
            if not m: m = re.search(r"Conta corrente.*?([\d]+-[\dX])", texto, re.IGNORECASE)
            if m: conta = m.group(1)
        else: # Caixa
            m = re.search(r"Conta:.*?([\d/]+-?\d?)", texto)
            if not m: m = re.search(r"Conta\s+Vinculada:.*?([\d/]+-?\d?)", texto)
            if m: conta = m.group(1)
            
        # 2. Identificar Saldo Final CC
        padrao_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*([CD]?)"
        linhas = texto.split('\n')
        saldo_final = 0.0
        matches = []
        for linha in linhas:
            if "SALDO" in linha.upper() and ("ANTERIOR" not in linha.upper() or "SALDO FINAL" in linha.upper()):
                m = re.search(padrao_valor, linha)
                if m: matches.append(m)
        
        if matches:
            v, t = matches[-1].groups()
            saldo_final = limpar_valor_monetario(f"{v} {t}")
            
        return {"Conta": conta, "Saldo_CC": saldo_final, "Arquivo": arquivo.name}
    except Exception as e:
        return {"Conta": "Erro", "Saldo_CC": 0.0, "Arquivo": arquivo.name}

def extrair_pdf_investimento_generico(arquivo, banco):
    """Lê Saldo e Rendimentos de Aplicação."""
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto = "".join([pag.get_text() for pag in doc])
        doc.close()
        
        conta = "N/A"
        if banco == 'BB':
            m = re.search(r"Conta(?:\s+corrente|de crédito)?[:\s]+([\d]+-[\dX])", texto, re.IGNORECASE)
            if m: conta = m.group(1)
        else:
            m = re.search(r"Conta:.*?([\d/]+-?\d?)", texto)
            if m: conta = m.group(1)
        
        saldo_aplic = 0.0
        rendimento = 0.0
        
        linhas = texto.split('\n')
        padrao_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
        
        for linha in linhas:
            linha_upper = linha.upper()
            
            # Captura Saldo
            if ("SALDO ATUAL" in linha_upper or "SALDO BRUTO" in linha_upper or "SALDO FINAL" in linha_upper) and "ANTERIOR" not in linha_upper:
                m = re.search(padrao_valor, linha)
                if m: saldo_aplic = limpar_valor_monetario(m.group(1))

            # Captura Rendimento
            termos_rendimento = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTOS"]
            if any(t in linha_upper for t in termos_rendimento) and "ACUMULADO" not in linha_upper:
                m = re.search(padrao_valor, linha)
                if m:
                    valor = limpar_valor_monetario(m.group(1))
                    if valor < 10000000: rendimento += valor

        return {"Conta": conta, "Saldo_Aplic": saldo_aplic, "Rendimento": rendimento, "Arquivo": arquivo.name}
    except Exception:
        return {"Conta": "Erro", "Saldo_Aplic": 0.0, "Rendimento": 0.0, "Arquivo": arquivo.name}

# ==========================================
# 3. PROCESSAMENTO CONTÁBIL (ROBUSTO)
# ==========================================

def encontrar_coluna_chave(df):
    """Procura dinamicamente a coluna de conta/domicílio."""
    candidatos = ['Domicílio bancário', 'Domicilio', 'Conta', 'Nº Conta', 'Descrição']
    for cand in candidatos:
        for col in df.columns:
            if cand.lower() in str(col).lower():
                return col
    return None

def encontrar_coluna_valor(df):
    """Procura dinamicamente a coluna de valor."""
    candidatos = ['Saldo Final', 'Saldo Atual', 'Valor', 'Movimento', 'Saldo']
    for cand in candidatos:
        for col in df.columns:
            if cand.lower() in str(col).lower() and 'anterior' not in str(col).lower():
                return col
    return None

def processar_contabilidade_saldos(arquivo):
    try:
        # Tenta ler com header na linha 1 (padrão)
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1, dtype=str)
        col_chave = encontrar_coluna_chave(df)
        
        # Se não achou, tenta header na linha 0
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            col_chave = encontrar_coluna_chave(df)
            
        if not col_chave:
            st.error("Não encontrei a coluna 'Domicílio bancário' no arquivo de SALDOS.")
            return pd.DataFrame()
            
        df['Chave Primaria'] = df[col_chave].apply(gerar_chave_contabil)
        df = df.dropna(subset=['Chave Primaria'])
        
        # Pivotagem (Assumindo que existem colunas de Conta Contábil)
        # Se o arquivo for simples (sem conta contábil), adaptação necessária
        if 'Conta contábil' in df.columns:
            df['Saldo Final'] = df['Saldo Final'].astype(str).apply(limpar_valor_monetario)
            df_pivot = df.pivot_table(index='Chave Primaria', columns='Conta contábil', values='Saldo Final', aggfunc='sum').reset_index()
            
            # Procura colunas
            col_mov = next((c for c in df_pivot.columns if '111111901' in str(c) or 'Conta Movimento' in str(c)), None)
            col_app = next((c for c in df_pivot.columns if '1111150' in str(c) or 'Aplicação' in str(c)), None)
            
            df_res = pd.DataFrame()
            df_res['Chave Primaria'] = df_pivot['Chave Primaria']
            df_res['Saldo_Contabil_CC'] = df_pivot[col_mov].fillna(0) if col_mov else 0.0
            df_res['Saldo_Contabil_Aplic'] = df_pivot[col_app].fillna(0) if col_app else 0.0
            
            # Traz descrição
            desc_map = df[['Chave Primaria', col_chave]].drop_duplicates(subset=['Chave Primaria']).set_index('Chave Primaria')
            df_res = df_res.join(desc_map, on='Chave Primaria')
            df_res.rename(columns={col_chave: 'Domicílio bancário'}, inplace=True)
            
            return df_res
        else:
            # Caso seja um CSV simples sem estrutura de Balancete
            st.warning("Arquivo de Saldos não parece ser um Balancete (falta 'Conta contábil').")
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"Erro ao ler Contabilidade Saldos: {e}")
        return pd.DataFrame()

def processar_contabilidade_rendimentos(arquivo):
    """
    Função corrigida para achar a coluna automaticamente e evitar KeyError.
    """
    if arquivo is None: return pd.DataFrame()
    
    try:
        # Tenta header=1
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1, dtype=str)
        col_chave = encontrar_coluna_chave(df)
        
        # Tenta header=0 se falhar
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            col_chave = encontrar_coluna_chave(df)
            
        if not col_chave:
            st.warning("⚠️ Não foi possível identificar a coluna da CONTA no arquivo de Rendimentos. O arquivo será ignorado.")
            st.write("Colunas encontradas:", df.columns.tolist())
            return pd.DataFrame() # Retorna vazio mas não trava

        # Identifica coluna de valor
        col_valor = encontrar_coluna_valor(df)
        if not col_valor:
            st.warning(f"⚠️ Não foi possível identificar a coluna de VALOR no arquivo de Rendimentos. Colunas: {df.columns.tolist()}")
            return pd.DataFrame()

        # Processamento seguro
        df['Chave Primaria'] = df[col_chave].apply(gerar_chave_contabil)
        df = df.dropna(subset=['Chave Primaria'])
        
        df['Rendimento'] = df[col_valor].astype(str).apply(limpar_valor_monetario)
        
        # Agrupa
        df_agrupado = df.groupby('Chave Primaria')['Rendimento'].sum().reset_index()
        df_agrupado.rename(columns={'Rendimento': 'Rendimento_Contabil'}, inplace=True)
        
        return df_agrupado

    except Exception as e:
        st.error(f"Erro ao processar Rendimentos: {e}")
        return pd.DataFrame()

# ==========================================
# 4. MOTOR DE CONSOLIDAÇÃO
# ==========================================

def executar_consolidacao(files_saldos, files_rendimentos, files_cc_bb, files_cc_cef, files_inv_bb, files_inv_cef):
    
    # 1. Processar Contabilidade
    df_contabil_saldos = processar_contabilidade_saldos(files_saldos)
    df_contabil_rendim = processar_contabilidade_rendimentos(files_rendimentos)
    
    # Merge Contábil Seguro
    if df_contabil_saldos.empty:
        st.error("Falha ao processar arquivo de SALDOS. Verifique o formato.")
        return pd.DataFrame()

    if not df_contabil_rendim.empty:
        df_contabil_master = pd.merge(df_contabil_saldos, df_contabil_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        # Se rendimentos falhou ou não foi enviado, segue só com saldos
        df_contabil_master = df_contabil_saldos
        df_contabil_master['Rendimento_Contabil'] = 0.0

    # 2. Processar Extratos
    dados_banco = []

    def processar_lista(lista_files, tipo, banco):
        if not lista_files: return
        for f in lista_files:
            if tipo == 'CC':
                d = extrair_pdf_cc_generico(f, banco)
                dados_banco.append({
                    "Chave Primaria": gerar_chave_padronizada(d['Conta']),
                    "Saldo_Banco_CC": d['Saldo_CC'],
                    "Saldo_Banco_Aplic": 0.0,
                    "Rendimento_Banco": 0.0
                })
            elif tipo == 'INV':
                d = extrair_pdf_investimento_generico(f, banco)
                dados_banco.append({
                    "Chave Primaria": gerar_chave_padronizada(d['Conta']),
                    "Saldo_Banco_CC": 0.0,
                    "Saldo_Banco_Aplic": d['Saldo_Aplic'],
                    "Rendimento_Banco": d['Rendimento']
                })

    processar_lista(files_cc_bb, 'CC', 'BB')
    processar_lista(files_cc_cef, 'CC', 'CEF')
    processar_lista(files_inv_bb, 'INV', 'BB')
    processar_lista(files_inv_cef, 'INV', 'CEF')

    if not dados_banco:
        st.warning("Nenhum dado válido extraído dos PDFs.")
        return pd.DataFrame()

    df_banco_raw = pd.DataFrame(dados_banco)
    # Filtra chaves nulas
    df_banco_raw = df_banco_raw[df_banco_raw['Chave Primaria'].notna()]
    
    df_banco_consol = df_banco_raw.groupby('Chave Primaria').agg({
        'Saldo_Banco_CC': 'sum',
        'Saldo_Banco_Aplic': 'sum',
        'Rendimento_Banco': 'sum'
    }).reset_index()

    # 3. Cruzamento Final
    df_final = pd.merge(df_contabil_master, df_banco_consol, on='Chave Primaria', how='inner')

    # Cálculos
    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']

    cols = [
        'Domicílio bancário', 
        'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
        'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
        'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento'
    ]
    cols_existentes = [c for c in cols if c in df_final.columns]
    return df_final[cols_existentes]

def gerar_excel_download(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacao')
        ws = writer.sheets['Conciliacao']
        for i, col in enumerate(df.columns):
            ws.column_dimensions[get_column_letter(i+1)].width = 20
    return output.getvalue()

# ==========================================
# 5. INTERFACE
# ==========================================

st.title("Hub de Conciliação Financeira")
st.info("Sistema robusto para conciliação de Saldos e Rendimentos.")

st.markdown("### 1. Contabilidade")
col1, col2 = st.columns(2)
file_saldos = col1.file_uploader("Saldos Finais (CSV)", type="csv")
file_rendim = col2.file_uploader("Rendimentos (CSV)", type="csv")

st.markdown("### 2. Extratos Bancários")
col3, col4 = st.columns(2)
files_cc_bb = col3.file_uploader("CC - Banco do Brasil", type="pdf", accept_multiple_files=True)
files_inv_bb = col3.file_uploader("Invest - Banco do Brasil", type="pdf", accept_multiple_files=True)
files_cc_cef = col4.file_uploader("CC - Caixa", type="pdf", accept_multiple_files=True)
files_inv_cef = col4.file_uploader("Invest - Caixa", type="pdf", accept_multiple_files=True)

if st.button("🚀 Processar", type="primary"):
    if file_saldos and (files_cc_bb or files_inv_bb or files_cc_cef or files_inv_cef):
        with st.spinner("Conciliando..."):
            df_final = executar_consolidacao(file_saldos, file_rendim, files_cc_bb, files_cc_cef, files_inv_bb, files_inv_cef)
            
            if not df_final.empty:
                st.success("Concluído!")
                
                # Filtro Divergências
                mask = (df_final['Diferenca_Saldo_CC'].abs() > 0.01) | \
                       (df_final['Diferenca_Saldo_Aplic'].abs() > 0.01) | \
                       (df_final['Diferenca_Rendimento'].abs() > 0.01)
                
                df_div = df_final[mask]
                
                if df_div.empty:
                    st.balloons()
                    st.info("Tudo batendo! Zero divergências.")
                else:
                    st.warning(f"{len(df_div)} contas com divergência.")
                    st.dataframe(df_div.style.format("{:,.2f}"))
                
                st.download_button("Baixar Excel", gerar_excel_download(df_final), "conciliacao.xlsx")
                
                with st.expander("Ver dados completos"):
                    st.dataframe(df_final)
            else:
                st.warning("Não foi possível cruzar os dados. Verifique se os números das contas nos PDFs batem com o CSV.")
    else:
        st.error("Anexe pelo menos o arquivo de Saldos e um Extrato.")
