import streamlit as st
import pandas as pd
import re
import io
import fitz  # PyMuPDF
from openpyxl import Workbook
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
        # Garante pelo menos 7 digitos, preenchendo com zeros ou cortando
        ultimos_7_digitos = parte_numerica[-7:]
        return ultimos_7_digitos.zfill(7)
    return None

def gerar_chave_contabil(texto_conta):
    """Extrai chave do Domicílio bancário da contabilidade."""
    if not isinstance(texto_conta, str): return None
    try:
        partes = texto_conta.split('-')
        if len(partes) > 2:
            conta_numerica = re.sub(r'\D', '', partes[2])
            return conta_numerica[-7:].zfill(7)
    except:
        return None
    return None

def limpar_valor_monetario(valor_str):
    """
    Transforma strings financeiras (R$ 1.000,00 D) em float (-1000.00).
    """
    if not isinstance(valor_str, str): return 0.0
    
    eh_negativo = 'D' in valor_str.upper() or '-' in valor_str
    
    # Remove tudo que não é dígito, ponto ou vírgula
    limpo = re.sub(r'[^\d,\.]', '', valor_str)
    
    try:
        # Lógica BR: inverte pontuação
        if ',' in limpo and '.' in limpo:
             limpo = limpo.replace('.', '').replace(',', '.')
        elif ',' in limpo:
             limpo = limpo.replace(',', '.')
        
        valor_float = float(limpo)
        return -valor_float if eh_negativo else valor_float
    except ValueError:
        return 0.0

# ==========================================
# 2. PROCESSAMENTO DE PDFS (EXTRATORES)
# ==========================================

def extrair_pdf_cc_generico(arquivo, banco):
    """Lê saldo de Conta Corrente (BB ou Caixa)."""
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto = "".join([pag.get_text() for pag in doc])
        doc.close()
        
        # 1. Identificar Conta
        if banco == 'BB':
            match_conta = re.search(r"Conta(?:\s+corrente)?[:\s]+([\d]+-[\dX])", texto, re.IGNORECASE)
            if not match_conta: match_conta = re.search(r"Conta corrente.*?([\d]+-[\dX])", texto, re.IGNORECASE)
        else: # Caixa
            match_conta = re.search(r"Conta:.*?([\d/]+-?\d?)", texto)
            if not match_conta: match_conta = re.search(r"Conta\s+Vinculada:.*?([\d/]+-?\d?)", texto)
            
        conta = match_conta.group(1).strip() if match_conta else "N/A"
        
        # 2. Identificar Saldo Final CC
        # Busca o último padrão monetário associado a SALDO
        padrao_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*([CD]?)"
        linhas = texto.split('\n')
        saldo_final = 0.0
        
        # Varre linhas buscando a última ocorrência de saldo
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
    """
    Lê Saldo e Rendimentos de Aplicação Financeira.
    Foca em 'Saldo Atual/Bruto' e 'Rendimento Bruto/Líquido/Periodo'.
    """
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto = "".join([pag.get_text() for pag in doc])
        doc.close()
        
        # 1. Identificar Conta (Muitas vezes o extrato de investimento cita a conta corrente de crédito)
        conta = "N/A"
        if banco == 'BB':
            match_conta = re.search(r"Conta(?:\s+corrente|de crédito)?[:\s]+([\d]+-[\dX])", texto, re.IGNORECASE)
        else: # Caixa
            match_conta = re.search(r"Conta:.*?([\d/]+-?\d?)", texto)
        
        if match_conta: conta = match_conta.group(1).strip()
        
        # 2. Identificar Rendimentos e Saldo
        saldo_aplic = 0.0
        rendimento = 0.0
        
        linhas = texto.split('\n')
        padrao_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
        
        for linha in linhas:
            linha_upper = linha.upper()
            
            # --- Captura de Saldo de Aplicação ---
            # Prioriza termos como "Saldo Atual", "Saldo Bruto Final"
            if ("SALDO ATUAL" in linha_upper or "SALDO BRUTO" in linha_upper or "SALDO FINAL" in linha_upper) and "ANTERIOR" not in linha_upper:
                m = re.search(padrao_valor, linha)
                if m: 
                    # Atualiza sempre que acha, assumindo que o último é o saldo final do extrato
                    saldo_aplic = limpar_valor_monetario(m.group(1))

            # --- Captura de Rendimento ---
            # Termos comuns: "Rendimento Bruto", "Rentabilidade", "Rendimentos"
            # Cuidado para não pegar "Rendimento acumulado" se quisermos apenas o do mês. 
            # Assumindo "Rendimento no mês" ou similar.
            termos_rendimento = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTOS"]
            if any(t in linha_upper for t in termos_rendimento) and "ACUMULADO" not in linha_upper:
                m = re.search(padrao_valor, linha)
                if m:
                    # Rendimentos podem aparecer várias vezes (por fundo), aqui somamos
                    valor = limpar_valor_monetario(m.group(1))
                    # Filtro simples para evitar somar saldos enormes por engano como rendimento
                    if valor < 10000000: 
                        rendimento += valor

        return {"Conta": conta, "Saldo_Aplic": saldo_aplic, "Rendimento": rendimento, "Arquivo": arquivo.name}

    except Exception:
        return {"Conta": "Erro", "Saldo_Aplic": 0.0, "Rendimento": 0.0, "Arquivo": arquivo.name}

# ==========================================
# 3. PROCESSAMENTO CONTÁBIL (CSVs)
# ==========================================

def processar_contabilidade_saldos(arquivo):
    """Processa o CSV de Saldos Finais."""
    try:
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1)
        
        df['Chave Primaria'] = df['Domicílio bancário'].apply(gerar_chave_contabil)
        df = df.dropna(subset=['Chave Primaria'])
        
        # Limpa Saldo
        df['Saldo Final'] = df['Saldo Final'].astype(str).apply(limpar_valor_monetario)
        
        # Pivotar para separar Saldo Conta Movimento (1111101/901) e Aplicação (1111150)
        # Nota: Ajuste os códigos de conta contábil conforme seu plano de contas real
        df_pivot = df.pivot_table(index='Chave Primaria', columns='Conta contábil', values='Saldo Final', aggfunc='sum').reset_index()
        
        # Mapeamento dinâmico (tenta achar colunas que contêm o padrão da conta)
        col_mov = next((c for c in df_pivot.columns if '111111901' in str(c) or 'Conta Movimento' in str(c)), None)
        col_app = next((c for c in df_pivot.columns if '1111150' in str(c) or 'Aplicação' in str(c)), None)
        
        df_res = pd.DataFrame()
        df_res['Chave Primaria'] = df_pivot['Chave Primaria']
        
        if col_mov: df_res['Saldo_Contabil_CC'] = df_pivot[col_mov].fillna(0)
        else: df_res['Saldo_Contabil_CC'] = 0.0
            
        if col_app: df_res['Saldo_Contabil_Aplic'] = df_pivot[col_app].fillna(0)
        else: df_res['Saldo_Contabil_Aplic'] = 0.0
        
        # Traz descrição
        desc_map = df[['Chave Primaria', 'Domicílio bancário']].drop_duplicates(subset=['Chave Primaria']).set_index('Chave Primaria')
        df_res = df_res.join(desc_map, on='Chave Primaria')
        
        return df_res
    except Exception as e:
        st.error(f"Erro ao ler Contabilidade Saldos: {e}")
        return pd.DataFrame()

def processar_contabilidade_rendimentos(arquivo):
    """Processa o novo CSV de Rendimentos."""
    try:
        # Assumindo estrutura similar: Domicílio, Conta Contábil de VPA (Rendimentos), Valor
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1) # Ajuste header se necessário
        
        df['Chave Primaria'] = df['Domicílio bancário'].apply(gerar_chave_contabil)
        df = df.dropna(subset=['Chave Primaria'])
        
        # Procura coluna de valor (pode ser 'Saldo Atual', 'Valor', 'Movimento')
        # No CSV de rendimentos, geralmente é o movimento do período
        col_valor = next((c for c in df.columns if 'Saldo' in c or 'Valor' in c or 'Movimento' in c), None)
        
        if col_valor:
            df['Rendimento'] = df[col_valor].astype(str).apply(limpar_valor_monetario)
            # Agrupa por conta (caso haja múltiplos lançamentos de rendimento)
            df_agrupado = df.groupby('Chave Primaria')['Rendimento'].sum().reset_index()
            df_agrupado.rename(columns={'Rendimento': 'Rendimento_Contabil'}, inplace=True)
            return df_agrupado
        else:
            st.warning("Não encontrei coluna de Valor no arquivo de Rendimentos.")
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"Erro ao ler Contabilidade Rendimentos: {e}")
        return pd.DataFrame()

# ==========================================
# 4. MOTOR DE CONSOLIDAÇÃO
# ==========================================

def executar_consolidacao(files_saldos, files_rendimentos, files_cc_bb, files_cc_cef, files_inv_bb, files_inv_cef):
    
    # 1. Processar Contabilidade
    df_contabil_saldos = processar_contabilidade_saldos(files_saldos)
    df_contabil_rendim = processar_contabilidade_rendimentos(files_rendimentos)
    
    # Merge Contábil (Saldos + Rendimentos)
    if not df_contabil_saldos.empty:
        df_contabil_master = pd.merge(df_contabil_saldos, df_contabil_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        return pd.DataFrame()

    # 2. Processar Extratos (Iterar sobre listas de arquivos)
    dados_banco = []

    # Helper para loop
    def processar_lista(lista_files, tipo, banco):
        for f in lista_files:
            if tipo == 'CC':
                d = extrair_pdf_cc_generico(f, banco)
                dados_banco.append({
                    "Chave Primaria": gerar_chave_padronizada(d['Conta']),
                    "Saldo_Banco_CC": d['Saldo_CC'],
                    "Saldo_Banco_Aplic": 0.0,
                    "Rendimento_Banco": 0.0,
                    "Origem": f"{banco} CC"
                })
            elif tipo == 'INV':
                d = extrair_pdf_investimento_generico(f, banco)
                dados_banco.append({
                    "Chave Primaria": gerar_chave_padronizada(d['Conta']),
                    "Saldo_Banco_CC": 0.0,
                    "Saldo_Banco_Aplic": d['Saldo_Aplic'],
                    "Rendimento_Banco": d['Rendimento'],
                    "Origem": f"{banco} INV"
                })

    processar_lista(files_cc_bb, 'CC', 'BB')
    processar_lista(files_cc_cef, 'CC', 'CEF')
    processar_lista(files_inv_bb, 'INV', 'BB')
    processar_lista(files_inv_cef, 'INV', 'CEF')

    if not dados_banco:
        return pd.DataFrame()

    # Consolidar dados bancários (soma por chave, pois podemos ter CC e INV separados para mesma conta)
    df_banco_raw = pd.DataFrame(dados_banco)
    df_banco_consol = df_banco_raw.groupby('Chave Primaria').agg({
        'Saldo_Banco_CC': 'sum',
        'Saldo_Banco_Aplic': 'sum',
        'Rendimento_Banco': 'sum'
    }).reset_index()

    # 3. Cruzamento Final (Contábil + Banco)
    df_final = pd.merge(df_contabil_master, df_banco_consol, on='Chave Primaria', how='inner') # Inner para pegar apenas o que tem match

    # Cálculos de Diferença
    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']

    # Organização das Colunas
    cols = [
        'Domicílio bancário', 
        'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
        'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
        'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento'
    ]
    # Filtra colunas que existem (caso alguma falhe)
    cols_existentes = [c for c in cols if c in df_final.columns]
    
    return df_final[cols_existentes]

def gerar_excel_download(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacao Completa')
        ws = writer.sheets['Conciliacao Completa']
        # Ajuste de largura simples
        for i, col in enumerate(df.columns):
            ws.column_dimensions[get_column_letter(i+1)].width = 20
    return output.getvalue()

# ==========================================
# 5. INTERFACE DO USUÁRIO
# ==========================================

st.title("Hub de Conciliação Financeira")
st.markdown("### Saldos de Conta Corrente, Aplicações e Rendimentos")
st.info("O sistema cruza as informações contábeis com os extratos bancários de todas as fontes fornecidas.")

# --- ÁREA DE UPLOAD (Organizada em Colunas) ---
st.markdown("---")
st.subheader("1. Arquivos da Contabilidade")
col_cont1, col_cont2 = st.columns(2)
file_contabil_saldos = col_cont1.file_uploader("Relatório de SALDOS FINAIS (.csv)", type="csv")
file_contabil_rendim = col_cont2.file_uploader("Relatório de RENDIMENTOS (.csv)", type="csv")

st.markdown("---")
st.subheader("2. Extratos Bancários (PDFs)")

col_banco1, col_banco2 = st.columns(2)

with col_banco1:
    st.markdown("#### 🏛️ Banco do Brasil")
    files_cc_bb = st.file_uploader("Extratos Conta Corrente (BB)", type="pdf", accept_multiple_files=True, key="bb_cc")
    files_inv_bb = st.file_uploader("Extratos Aplicação (BB)", type="pdf", accept_multiple_files=True, key="bb_inv")

with col_banco2:
    st.markdown("#### 🏦 Caixa Econômica")
    files_cc_cef = st.file_uploader("Extratos Conta Corrente (CEF)", type="pdf", accept_multiple_files=True, key="cef_cc")
    files_inv_cef = st.file_uploader("Extratos Aplicação (CEF)", type="pdf", accept_multiple_files=True, key="cef_inv")

# --- BOTÃO DE AÇÃO ---
st.markdown("---")
if st.button("🚀 Executar Conciliação Unificada", type="primary"):
    
    # Validação Mínima: Precisa pelo menos dos saldos contábeis e algum extrato
    tem_banco = (files_cc_bb or files_inv_bb or files_cc_cef or files_inv_cef)
    
    if file_contabil_saldos and tem_banco:
        with st.spinner("Lendo arquivos, extraindo dados e conciliando..."):
            
            df_resultado = executar_consolidacao(
                file_contabil_saldos,
                file_contabil_rendim, # Pode ser None, o código trata
                files_cc_bb,
                files_cc_cef,
                files_inv_bb,
                files_inv_cef
            )
            
            if not df_resultado.empty:
                st.success("Conciliação finalizada!")
                
                # Resumo de Divergências
                st.subheader("⚠️ Resumo de Divergências")
                filtros_div = (df_resultado['Diferenca_Saldo_CC'].abs() > 0.01) | \
                              (df_resultado['Diferenca_Saldo_Aplic'].abs() > 0.01) | \
                              (df_resultado['Diferenca_Rendimento'].abs() > 0.01)
                
                df_div = df_resultado[filtros_div]
                
                if df_div.empty:
                    st.balloons()
                    st.info("Parabéns! Nenhuma divergência financeira encontrada.")
                else:
                    st.error(f"Foram encontradas {len(df_div)} contas com diferenças.")
                    st.dataframe(df_div.style.format("{:,.2f}"))
                
                # Tabela Completa
                with st.expander("Ver Tabela Completa de Todas as Contas"):
                    st.dataframe(df_resultado.style.format("{:,.2f}"))
                
                # Download
                st.download_button(
                    label="📥 Baixar Relatório Completo (Excel)",
                    data=gerar_excel_download(df_resultado),
                    file_name="Relatorio_Conciliacao_Unificado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            else:
                st.warning("O processamento terminou, mas nenhuma conta foi correspondida (match). Verifique se os números das contas nos PDFs correspondem aos do relatório contábil.")
    else:
        st.error("Por favor, anexe o Relatório de Saldos Contábeis e pelo menos um tipo de Extrato Bancário.")
