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
    page_title="Super Conciliador v2.2 (Caixa Fix + Outer Join)",
    layout="wide",
    page_icon="💸",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. FUNÇÕES DE LIMPEZA E CHAVES
# ==========================================

def gerar_chave_padronizada(texto_conta):
    """
    Padroniza a chave para cruzar os dados.
    """
    if isinstance(texto_conta, str):
        # Correção específica para Caixa (Agencia/Op/Conta)
        if '/' in texto_conta:
            texto_conta = texto_conta.split('/')[-1]
            
        # Remove letras, traços, espaços
        parte_numerica = re.sub(r'\D', '', texto_conta)
        
        if not parte_numerica: 
            return None
            
        # Pega os últimos 7 dígitos e garante zeros à esquerda
        ultimos_7_digitos = parte_numerica[-7:]
        return ultimos_7_digitos.zfill(7)
    return None

def limpar_valor_monetario(valor_str):
    """Transforma texto financeiro (ex: '1.500,20 D') em número float."""
    if not isinstance(valor_str, str): return 0.0
    
    valor_upper = valor_str.upper()
    eh_negativo = 'D' in valor_upper or '-' in valor_str or 'DEB' in valor_upper
    
    # Mantém apenas dígitos, vírgula e ponto
    limpo = re.sub(r'[^\d,\.]', '', valor_str)
    
    try:
        if not limpo: return 0.0
        
        # Formato Brasileiro (1.000,00)
        if ',' in limpo and '.' in limpo:
             limpo = limpo.replace('.', '').replace(',', '.')
        elif ',' in limpo:
             limpo = limpo.replace(',', '.')
        
        valor_float = float(limpo)
        return -valor_float if eh_negativo else valor_float
    except ValueError:
        return 0.0

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
        
        # --- 1. EXTRAÇÃO DA CONTA ---
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
        
        # --- 2. EXTRAÇÃO DE VALORES ---
        saldo_final = 0.0
        rendimento_total = 0.0
        
        linhas = texto_completo.split('\n')
        
        for linha in linhas:
            linha_upper = linha.upper().strip()
            
            # -- SALDO --
            gatilhos_saldo = ["SALDO FINAL", "SALDO TOTAL", "SALDO ATUAL", "SALDO EM"]
            ignorar = ["ANTERIOR", "BLOQUEADO", "PROVISORIO"]
            
            if any(g in linha_upper for g in gatilhos_saldo) and not any(i in linha_upper for i in ignorar):
                match_val = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})([CD]?)", linha_upper)
                if match_val:
                    valor_str = match_val.group(1)
                    letra = match_val.group(2)
                    sinal = "D" if letra == 'D' or " D" in linha_upper or "DEB" in linha_upper else "C"
                    saldo_final = limpar_valor_monetario(f"{valor_str} {sinal}")

            # -- RENDIMENTOS --
            if tipo_extrato == 'INV':
                gatilhos_rend = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTO LIQUIDO"]
                if any(g in linha_upper for g in gatilhos_rend) and "ACUMULADO" not in linha_upper:
                    match_val = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
                    if match_val:
                        valor = limpar_valor_monetario(match_val.group(1))
                        if valor < 100000000: 
                            rendimento_total += valor

        # Fallback para contas sem movimento
        if saldo_final == 0.0 and ("NAO HOUVE MOVIMENTO" in texto_completo.upper() or "SEM MOVIMENTO" in texto_completo.upper()):
             match_ant = re.search(r"(?:SALDO ANTERIOR|SALDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE | re.DOTALL)
             if match_ant:
                 saldo_final = limpar_valor_monetario(match_ant.group(1))

        return {
            "Conta": conta_encontrada,
            "Saldo": saldo_final,
            "Rendimento": rendimento_total,
            "Texto_Raw": texto_completo[:300] + "..."
        }
        
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
        possiveis_chaves = ['Domicílio bancário', 'Domicilio', 'Conta', 'Nº Conta', 'Descrição']
        for col in df.columns:
            for p in possiveis_chaves:
                if p.lower() in str(col).lower():
                    col_chave = col
                    break
        
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            for col in df.columns:
                for p in possiveis_chaves:
                    if p.lower() in str(col).lower():
                        col_chave = col
                        break
        
        if not col_chave: return pd.DataFrame()

        col_valor = None
        possiveis_valores = ['Saldo Final', 'Saldo Atual', 'Movimento', 'Valor']
        for col in df.columns:
            if 'anterior' in str(col).lower(): continue
            for p in possiveis_valores:
                if p.lower() in str(col).lower():
                    col_valor = col
                    break
        
        if not col_valor: return pd.DataFrame()

        df['Chave Primaria'] = df[col_chave].apply(gerar_chave_padronizada)
        df = df.dropna(subset=['Chave Primaria'])
        df['Valor_Numerico'] = df[col_valor].astype(str).apply(limpar_valor_monetario)
        
        if tipo == 'SALDO':
            if any('contábil' in str(c).lower() for c in df.columns):
                col_contabil = next(c for c in df.columns if 'contábil' in str(c).lower())
                df_pivot = df.pivot_table(index='Chave Primaria', columns=col_contabil, values='Valor_Numerico', aggfunc='sum').reset_index()
                
                # CÓDIGOS CONTÁBEIS (Ajuste aqui se necessário)
                col_mov = next((c for c in df_pivot.columns if '111111901' in str(c) or 'Conta Movimento' in str(c)), None)
                col_app = next((c for c in df_pivot.columns if '1111150' in str(c) or 'Aplicação' in str(c)), None)
                
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
        st.error(f"Erro Contábil: {e}")
        return pd.DataFrame()

# ==========================================
# 4. CONSOLIDAÇÃO (ATUALIZADO)
# ==========================================

def executar_processo(file_saldos, file_rendim, lista_extratos_cc, lista_extratos_inv):
    
    # Processa Contabilidade
    df_saldos = processar_contabil(file_saldos, 'SALDO')
    df_rendim = processar_contabil(file_rendim, 'RENDIMENTO')
    
    if df_saldos.empty:
        st.error("Não foi possível ler o arquivo de Saldos Contábeis.")
        return pd.DataFrame(), pd.DataFrame()

    # Junta as tabelas contábeis
    df_contabil = df_saldos
    if not df_rendim.empty:
        df_contabil = pd.merge(df_saldos, df_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        df_contabil['Rendimento_Contabil'] = 0.0

    # Processa Bancos
    dados_banco = []
    log_leitura = []

    # Extratos CC
    for f in lista_extratos_cc:
        res = extrair_pdf_melhorado(f, 'CC')
        chave = gerar_chave_padronizada(res['Conta'])
        if chave:
            dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': res['Saldo'], 'Saldo_Banco_Aplic': 0.0, 'Rendimento_Banco': 0.0})
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': chave, 'Saldo': res['Saldo'], 'Tipo': 'CC', 'Raw': res['Texto_Raw']})

    # Extratos Investimento
    for f in lista_extratos_inv:
        res = extrair_pdf_melhorado(f, 'INV')
        chave = gerar_chave_padronizada(res['Conta'])
        if chave:
            dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': 0.0, 'Saldo_Banco_Aplic': res['Saldo'], 'Rendimento_Banco': res['Rendimento']})
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': chave, 'Saldo_Aplic': res['Saldo'], 'Rendimento': res['Rendimento'], 'Tipo': 'INV', 'Raw': res['Texto_Raw']})

    df_log = pd.DataFrame(log_leitura)

    # Consolida Banco (CORREÇÃO DE ERRO SE LISTA VAZIA)
    if dados_banco:
        df_banco = pd.DataFrame(dados_banco).groupby('Chave Primaria').sum().reset_index()
    else:
        # Cria dataframe vazio com estrutura correta para não quebrar o merge
        df_banco = pd.DataFrame(columns=['Chave Primaria', 'Saldo_Banco_CC', 'Saldo_Banco_Aplic', 'Rendimento_Banco'])

    # Cruzamento Final (ALTERADO PARA OUTER JOIN)
    # Isso garante que contas que só existem no Banco OU só no CSV apareçam
    df_final = pd.merge(df_contabil, df_banco, on='Chave Primaria', how='outer').fillna(0)

    # Tratamento da coluna Descrição para contas que só vieram do Banco (não tem descrição no CSV)
    if 'Descrição' in df_final.columns:
        df_final['Descrição'] = df_final['Descrição'].replace(0, 'CONTA NÃO LOCALIZADA NO CSV')
        df_final['Descrição'] = df_final['Descrição'].fillna('CONTA NÃO LOCALIZADA NO CSV')
    else:
        df_final['Descrição'] = 'CONTA NÃO LOCALIZADA NO CSV'

    # Diferenças
    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']

    # Organização de Colunas
    cols = ['Descrição', 'Chave Primaria', 
            'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
            'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
            'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento']
    
    colunas_finais = [c for c in cols if c in df_final.columns]
    return df_final[colunas_finais], df_log

def to_excel(df):
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

st.title("💸 Super Conciliador v2.2")
st.markdown("### Suporte a Caixa (Ag/Op/Conta) e BB Integrado")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Contabilidade")
    f_saldos = st.file_uploader("Arquivo de Saldos (CSV)", type='csv')
    f_rendim = st.file_uploader("Arquivo de Rendimentos (CSV)", type='csv')

with col2:
    st.subheader("2. Extratos Bancários")
    f_cc = st.file_uploader("Conta Corrente (BB e Caixa)", type='pdf', accept_multiple_files=True)
    f_inv = st.file_uploader("Investimentos (BB e Caixa)", type='pdf', accept_multiple_files=True)

if st.button("Executar Conciliação", type="primary"):
    if f_saldos and (f_cc or f_inv):
        with st.spinner("Processando..."):
            df_final, df_log = executar_processo(f_saldos, f_rendim, f_cc if f_cc else [], f_inv if f_inv else [])
            
            if not df_final.empty:
                st.success("Processamento concluído!")
                
                tab1, tab2, tab3 = st.tabs(["📊 Resultado", "⚠️ Divergências", "🕵️ Raio-X"])
                
                with tab1:
                    st.dataframe(df_final.style.format("{:,.2f}"))
                    st.download_button("Baixar Excel", to_excel(df_final), "conciliacao_final.xlsx")
                
                with tab2:
                    # Filtra divergências maiores que 1 centavo
                    div = df_final[
                        (df_final['Diferenca_Saldo_CC'].abs() > 0.01) | 
                        (df_final['Diferenca_Saldo_Aplic'].abs() > 0.01) |
                        (df_final['Diferenca_Rendimento'].abs() > 0.01)
                    ]
                    if div.empty: st.info("Sem divergências materiais encontradas!")
                    else: st.dataframe(div.style.format("{:,.2f}"))
                
                with tab3:
                    st.dataframe(df_log)
            else:
                st.error("Ocorreu um erro ou não há dados correspondentes.")
                st.dataframe(df_log)
    else:
        st.warning("Carregue o arquivo de Saldos (CSV) e pelo menos um extrato bancário.")

st.markdown("---")
st.info("Nota: Os extratos devem estar separados por banco e tipo. O sistema usa os **últimos 7 dígitos** da conta para o cruzamento.")
